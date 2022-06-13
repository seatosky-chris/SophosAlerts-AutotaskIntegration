const {AutotaskRestApi} = require('@apigrate/autotask-restapi');
const fetch = require("node-fetch-commonjs");
const { RateLimit } = require('async-sema');
var fs = require('fs');
const orgMapping = require('../OrgMapping.json');
const upDownEvents = require('../UpDownEvents.json');

module.exports = async function (context, myTimer) {
    var timeStamp = new Date().toISOString();
    var lastRun = false
    var lastRunUnixTimestamp = false;
    var lastRunDir = null;


    if (fs.existsSync("C:/home/data")) {
        // linux / live azure function
        context.log("Using lastRun in C:/home/data");
        lastRunDir = "C:/home/data/lastRun.dat";
    } else {
        // windows
        context.log("Using local lastRun file");
        lastRunDir = "lastRun.dat";
    }

    try {
        await fs.promises.access(lastRunDir);
        lastRun = new Date(fs.readFileSync(lastRunDir, 'utf-8'));
    } catch (error) {
        // New run
        context.log.warn("This script has never been ran before.");
    }

    if (lastRun && lastRun instanceof Date) {
        lastRunUnixTimestamp = Math.floor(lastRun.getTime() / 1000);

        // if timestamp is more than 24 hours old, set to false (we can't go over 24 hours)
        var curTimeStamp = Math.round(new Date().getTime() / 1000);
        if (lastRunUnixTimestamp < (curTimeStamp - (24 * 3600))) {
            lastRunUnixTimestamp = false;
        }
    }
    
    let sophosToken = await getSophosToken(context);
    let sophosJWT = sophosToken.access_token;

    if (sophosJWT) {
        var sophosPartnerID = await getSophosPartnerID(context, sophosJWT);

        if (sophosPartnerID) {
            // Get list of tenants, we need to handle each on an individual basis
            var sophosTenants = await getSophosTenants(context, sophosJWT, sophosPartnerID);

            if (sophosTenants && sophosTenants.items) {
                await timeout(1000); // wait a second to prevent rate limiting
                let alerts = await getSophosSiemAlerts(context, sophosJWT, sophosTenants, lastRunUnixTimestamp);

                if (alerts && alerts.length > 0) {
                    var filteredAlerts = alerts.filter(alert => alert.severity != "low");
                    var upAlerts = alerts.filter(alert => alert.severity == "low" && Object.keys(upDownEvents).includes(alert.type));

                    // Connect to the Autotask API
                    const autotask = new AutotaskRestApi(
                        process.env.AUTOTASK_USER,
                        process.env.AUTOTASK_SECRET, 
                        process.env.AUTOTASK_INTEGRATION_CODE 
                    );
                    let api = await autotask.api();

                    var alertTenants = filteredAlerts.map(function(alert) {
                        return alert.customer_id;
                    });
                    alertTenants = [...new Set(alertTenants)]

                    let alertDevices = {};
                    for (i = 0; i < alertTenants.length; i++) {
                        var tenantID = alertTenants[i];
                        let sophosTenant = sophosTenants.items.filter(t => t.id == tenantID)[0];
                        let deviceIDs = filteredAlerts.filter(alert => alert.customer_id == tenantID).map(function(alert) {
                            return alert.data.endpoint_id;
                        });

                        var devices = await getSophosDevices(context, sophosJWT, sophosTenant, deviceIDs);
                        if (devices && devices.items) {
                            alertDevices[tenantID] = devices.items;
                        }
                    }

                    for (i = 0; i < filteredAlerts.length; i++) {
                        var alert = filteredAlerts[i];
                        // Go through each alert (that isn't low severity) and create a new ticket in Autotask for it
                        let sophosCompany = (sophosTenants.items.filter(tenant => tenant.id == alert.customer_id))[0].name;
                        let autotaskID = 0;
                        if (sophosCompany) {
                            autotaskID = orgMapping[sophosCompany];
                        }

                        // Get primary location and default contract
                        var contractID = await getAutotaskContractID(api, autotaskID);
                        var location = await getAutotaskLocation(api, autotaskID);

                        // Get related device if applicable
                        var customerDevices = alertDevices[alert.customer_id];
                        var alertDevice = customerDevices.filter(device => device.id == alert.data.endpoint_id)[0];
                        var deviceID = null;
                        if (alertDevice) {
                            deviceID = await getAutotaskDevice(api, autotaskID, alertDevice);
                        }

                        var title = `Sophos Alert: "${alert.description}"`;
                        if (!title.includes(alert.location)) {
                            title = title + ` on "${alert.location}"`;
                        }
                        var when = new Date(alert.when);
                        var description = `${alert.description} \nSeverity: ${alert.severity} \nDevice: ${alert.location}`;
                        if (alert.data && alert.data.source_info && alert.data.source_info.ip) {
                            description += `\nIP: ${alert.data.source_info.ip}`;
                        }
                        description += `\nEvent Type: ${alert.type} \nWhen: ${when.toLocaleDateString('en-us', { weekday:"long", year:"numeric", month:"short", day:"numeric"})} \n\nSee the Sophos portal for more details.`;

                        // See if there are any existing tickets of this type and for this device
                        let tickets = await searchAutotaskTickets(context, api, autotaskID, "Sophos Alert: ", alert.location, upDownEvents[alert.type]);

                        if (tickets && tickets.length > 0) {
                            // Existing ticket found, add notes
                            // get latest ticket
                            let existingTicket = tickets.reduce((a, b) => new Date(a.createDate) > new Date(b.createDate) ? a : b);

                            if (existingTicket) {
                                let updateNote = {
                                    "TicketID": existingTicket.id,
                                    "Title": "New Alert",
                                    "Description": description,
                                    "NoteType": 1,
                                    "Publish": 1
                                }
                                api.TicketNotes.create(existingTicket.id, updateNote);
                                context.log("New ticket note added on ticket id: " + existingTicket.id);
                            }
                        } else {
                            // No existing ticket found, create a new one
                            if (process.env.HOW_TO_DOCUMENTATION_LINK) {
                                description += '\n\nHow To Documentation: ' + process.env.HOW_TO_DOCUMENTATION_LINK;
                            }
                            
                            // Make a new ticket
                            let newTicket = {
                                CompanyID: autotaskID,
                                CompanyLocationID: (location ? location.id : 10),
                                Priority: alert.severity == 'medium' ? 3 : 2,
                                Status: 1,
                                QueueID: parseInt(process.env.TICKET_QueueID),
                                IssueType: parseInt(process.env.TICKET_IssueType),
                                SubIssueType: parseInt(process.env.TICKET_SubIssueType),
                                ServiceLevelAgreementID: parseInt(process.env.TICKET_ServiceLevelAgreementID),
                                ContractID: (contractID ? contractID : null),
                                Title: title,
                                Description: description
                            };
                            if (deviceID) {
                                newTicket.ConfigurationItemID = deviceID;
                            }

                            await createAutotaskTicket(context, api, newTicket);
                        }
                    };

                    // Close tickets on up alerts
                    for (i = 0; i < upAlerts.length; i++) {
                        var alert = upAlerts[i];
                        // Go through each up alert and find the relevant ticket in Autotask then self-heal it
                        let sophosCompany = (sophosTenants.items.filter(tenant => tenant.id == alert.customer_id))[0].name;
                        let autotaskID = 0;
                        if (sophosCompany) {
                            autotaskID = orgMapping[sophosCompany];
                        }
                        
                        let tickets = await searchAutotaskTickets(context, api, autotaskID, "Sophos Alert: ", alert.location, upDownEvents[alert.type]);
                        if (tickets && tickets.length > 0) {
                            // get latest ticket
                            let downTicket = tickets.reduce((a, b) => new Date(a.createDate) > new Date(b.createDate) ? a : b);

                            if (downTicket) {
                                let closingNote = {
                                    "TicketID": downTicket.id,
                                    "Title": "Self-Healing Update",
                                    "Description": "[Self-Healing] " + alert.description,
                                    "NoteType": 1,
                                    "Publish": 1
                                }
                                api.TicketNotes.create(downTicket.id, closingNote);

                                let closingTicket = {
                                    "id": downTicket.id,
                                    "Status": (downTicket.assignedResourceID ? 13 : 5)
                                }
                                api.Tickets.update(closingTicket);
                            }
                        }
                    }
                }

                try {
                    fs.writeFile(lastRunDir, timeStamp, function(err) {
                        if (err) throw err;
                    });
                    context.log("Updated lastRun.dat to: " + timeStamp);
                } catch (error) {
                    context.log.error("Could not update lastRun.dat: " + error);
                }
            }
        }
    }

    context.log('JavaScript timer trigger function ran!', timeStamp);
};

function timeout(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

async function getSophosToken(context) {
    let url = 'https://id.sophos.com/api/v2/oauth2/token';

    var authBody = new URLSearchParams({
        "grant_type": "client_credentials",
        "client_id": process.env.SOPHOS_CLIENT_ID,
        "client_secret": process.env.SOPHOS_SECRET,
        "scope": "token"
    });

    try {
        let sophosToken = await fetch(url, {
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded'
            },
            method: "POST",
            body: authBody
        });
        return await sophosToken.json();
    } catch (error) {
        context.log.error(error);
        return null;
    }
}

async function getSophosPartnerID(context, token) {
    let url = 'https://api.central.sophos.com/whoami/v1';

    try {
        let sophosPartnerInfo = await fetch(url, {
            headers: {
                Authorization: "Bearer " + token,
                method: "GET"
            }
        });
        let sophosPartnerInfoJson = await sophosPartnerInfo.json();
        return sophosPartnerInfoJson.id;
    } catch (error) {
        context.log.error(error);
        return null;
    }
}

async function getSophosTenants(context, token, partnerID) {
    let url = 'https://api.central.sophos.com/partner/v1/tenants?page';
    let sophosTenantsJson;

    try {
        let sophosTenants = await fetch(url + "Total=true", {
            headers: {
                Authorization: "Bearer " + token,
                "X-Partner-ID": partnerID,
                method: "GET"
            }
        });
        sophosTenantsJson =  await sophosTenants.json();
    } catch (error) {
        context.log.error(error);
    }

    if (sophosTenantsJson.pages && sophosTenantsJson.pages.total > 1) {
        var totalPages = sophosTenantsJson.pages.total;
        for (let i = 2; i <= totalPages; i++) {
            try {
                let sophosTenantsTemp = await fetch(url + "=" + i, {
                    headers: {
                        Authorization: "Bearer " + token,
                        "X-Partner-ID": partnerID,
                        method: "GET"
                    }
                });
                let sophosTenantsTempJson = await sophosTenantsTemp.json();
                sophosTenantsJson.items = sophosTenantsJson.items.concat(sophosTenantsTempJson.items);
            } catch (error) {
                context.log.error(error);
            }
        }
    }

    return sophosTenantsJson;
}

async function getSophosDevices(context, token, tenant, ids = null) {
    let url = tenant.apiHost + '/endpoint/v1/endpoints?pageSize=500';

    if (ids) {        
        ids.forEach(function(id) {
            url += "&ids=" + id;
        });
    }

    let fetchHeader = {
        method: "GET",
        headers: {
            Authorization: "Bearer " + token,
            "X-Tenant-ID": tenant.id,
            "Accept": "application/json"
        }
    };

    return fetch(url, fetchHeader)
        .then((response) => response.json());
}

async function getSophosSiemAlerts(context, token, tenants, fromDate = false) {
    let queryUrls = [];

    tenants.items.filter(t => t.status == 'active').forEach(function(tenant) {
        let url = 'https://api-' + tenant.dataRegion + '.central.sophos.com/siem/v1/alerts';
        if (fromDate && Number.isInteger(fromDate)) {
            url = url + '?from_date=' + fromDate + '&limit=1000';
        }

        let fetchHeader = {
            headers: {
                Authorization: "Bearer " + token,
                "X-Tenant-ID": tenant.id,
                method: "GET"
            }
        };

        queryUrls.push({url, fetchHeader});
    });

    const limit = RateLimit(8);
    const fetchFromApi = ({url, fetchHeader}) => {
        const response = fetch(url, fetchHeader)
            .then((res) => res.json())
            .catch((error) => context.log.error(error));
        return response;
    };

    let alerts = [];
    for (const query of queryUrls) {
        await limit();
        fetchFromApi(query).then((result) => {
            if (result && result.items) {
                alerts = alerts.concat(result.items);
            }
        });
    }

    return alerts;
}

async function getAutotaskLocation(autotaskAPI, autotaskID) {
    let locations = await autotaskAPI.CompanyLocations.query({
        filter: [
            {
                "op": "eq",
                "field": "CompanyID",
                "value": autotaskID
            }
        ],
        includeFields: [
            "id", "isActive", "isPrimary"
        ]
    });

    locations = locations.items.filter(location => location.isActive);

    var location;
    if (locations.length > 0) {
        location = locations.filter(location => location.isPrimary);
        location = location[0];
        if (!location) {
            location = locations[0];
        }
    } else {
        location = locations[0];
    }

    return location;
}

async function getAutotaskContractID(autotaskAPI, autotaskID) {
    var contractID = null;
    let contract = await autotaskAPI.Contracts.query({
        filter: [
            {
                "op": "and",
                "items": [
                    {
                        "op": "eq",
                        "field": "CompanyID",
                        "value": autotaskID
                    },
                    {
                        "op": "eq",
                        "field": "IsDefaultContract",
                        "value": true
                    }
                ]
            }
        ],
        includeFields: [ "id" ]
    });
    
    if (contract && contract.items.length > 0) {
        contractID = contract.items[0].id
    }
    return contractID;
}

async function getAutotaskDevice(autotaskAPI, autotaskID, deviceDetails) {
    var deviceID = null;
    let device = await autotaskAPI.ConfigurationItems.query({
        filter: [
            {
                "op": "and",
                "items": [
                    {
                        "op": "eq",
                        "field": "CompanyID",
                        "value": autotaskID
                    },
                    {
                        "op": "or",
                        "items": [
                            {
                                "op": "eq",
                                "field": "referenceTitle",
                                "value": deviceDetails.hostname
                            },
                            {
                                "op": "eq",
                                "field": "rmmDeviceAuditHostname",
                                "value": deviceDetails.hostname
                            },
                            {
                                "op": "eq",
                                "field": "rmmDeviceAuditDescription",
                                "value": deviceDetails.hostname
                            },
                            {
                                "op": "eq",
                                "field": "rmmDeviceAuditSNMPName",
                                "value": deviceDetails.hostname
                            }
                        ]
                    }
                ]
            }
        ]
    });

    
    if (device.items.length > 1) { 
        var filteredDevices = device.items.filter(function(device) {
            var rmmMacAddresses = device.rmmDeviceAuditMacAddress.replace(/^\[|\]$/gm,'').split(', ');
            var intersection = deviceDetails.macAddresses.filter(addr => rmmMacAddresses.includes(addr));
            return intersection.length > 0;
        });
        if (filteredDevices.length > 0) {
            device.items = filteredDevices;
        }

        if (device.items.length > 1) {  
            filteredDevices = device.items.filter(device => device.rmmDeviceAuditLastUser == deviceDetails.associatedPerson.viaLogin);
            if (filteredDevices.length > 0) {
                device.items = filteredDevices;
            }
        }

        if (device.items.length > 1) {
            filteredDevices = device.items.filter(device => deviceDetails.ipv4Addresses.includes(device.rmmDeviceAuditIPAddress));
            if (filteredDevices.length > 0) {
                device.items = filteredDevices;
            }
        }

        if (device.items.length > 1) {
            device.items.sort(function(a, b) { return b.lastSeen - a.lastSeen});
        }
    }

    if (device && device.length > 0) {
        deviceID = device.items[0].id;
    }
    return deviceID;
}

async function searchAutotaskTickets(context, autotaskAPI, companyID = false, titleStart = false, deviceName = false, eventType = false) {
    var ticketFilters = [];

    if (companyID) {
        ticketFilters.push({
            "op": "eq",
            "field": "CompanyID",
            "value": companyID
        });
    }
    if (titleStart) {
        ticketFilters.push({
            "op": "beginsWith",
            "field": "title",
            "value": titleStart
        });
    }
    if (deviceName) {
        ticketFilters.push({
            "op": "contains",
            "field": "description",
            "value": "Device: " + deviceName
        });
    }
    if (eventType) {
        ticketFilters.push({
            "op": "contains",
            "field": "description",
            "value": "Event Type: " + eventType
        });
    }
    ticketFilters.push({
        "op": "notExist",
        "field": "CompletedByResourceID"
    });
    ticketFilters.push({
        "op": "notExist",
        "field": "CompletedDate"
    });

    try {
        let tickets = await autotaskAPI.Tickets.query({
            "filter": [
                {
                    "op": "and",
                    "items": ticketFilters
                }
            ]
        });
        if (tickets && tickets.items) {
            return tickets.items;
        }
        return null;
    } catch (error) {
        context.log.error(error);
    }
}

async function createAutotaskTicket(context, autotaskAPI, newTicket) {
    var ticketID = null;
    try {
        result = await autotaskAPI.Tickets.create(newTicket);
        ticketID = result.itemId;
        if (!ticketID) {
            throw "No ticket ID";
        } else {
            context.log("New ticket created: " + ticketID);
        }
    } catch (error) {
        // Send an email to support if we couldn't create the ticket
        var mailBody = {
            From: {
                Email: process.env.EMAIL_FROM__Email,
                Name: process.env.EMAIL_FROM__Name
            },
            To: [
                {
                    Email: process.env.EMAIL_TO__Email,
                    Name: process.env.EMAIL_TO__Name
                }
            ],
            "Subject": newTicket.title,
            "HTMLContent": newTicket.description.replace(new RegExp('\r?\n','g'), "<br />")
        }

        try {
            let emailResponse = await fetch(process.env.EMAIL_API_ENDPOINT, {
                headers: {
                    'Accept': 'application/json',
                    'Content-Type': 'application/json',
                    'x-api-key': process.env.EMAIL_API_KEY
                },
                method: "POST",
                body: JSON.stringify(mailBody)
            });
            context.log.warn("Ticket creation failed. Backup email sent to support.");
        } catch (error) {
            context.log.error("Ticket creation failed. Sending an email as a backup also failed.");
            context.log.error(error);
        }
        ticketID = null;
    }
    return ticketID;
}