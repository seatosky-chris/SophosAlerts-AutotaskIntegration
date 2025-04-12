const { app } = require('@azure/functions');
const {AutotaskRestApi} = require('@apigrate/autotask-restapi');
const { RateLimit } = require('async-sema');
var fs = require('fs');
const orgMapping = require('../OrgMapping.json');
const upDownEvents = require('../UpDownEvents.json');
var idRegex = /ID: ([0-9a-fA-F]{8}\b-[0-9a-fA-F]{4}\b-[0-9a-fA-F]{4}\b-[0-9a-fA-F]{4}\b-[0-9a-fA-F]{12})(\n| )/gm

app.timer('SophosAlerts-AutotaskIntegration', {
    schedule: "0 */10 * * * *",
    handler: async (myTimer, context) => {
        var timeStamp = new Date().toISOString();
        var lastRun = false
        var lastRunUnixTimestamp = false;
        var lastRunDir = null;
        var ignoredAlertTypes = [];


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
            context.warn("This script has never been ran before.");
        }

        if (lastRun && lastRun instanceof Date) {
            lastRunUnixTimestamp = Math.floor(lastRun.getTime() / 1000);
            context.log("Last run: " + lastRunUnixTimestamp.toString());

            // if timestamp is more than 24 hours old, set to false (we can't go over 24 hours)
            var curTimeStamp = Math.round(new Date().getTime() / 1000);
            if (lastRunUnixTimestamp < (curTimeStamp - (24 * 3600))) {
                lastRunUnixTimestamp = false;
                context.log("Last run is more than 24 hours old")
            }
        }

        if (process.env.IGNORE_AlertTypes) {
            ignoredAlertTypes = process.env.IGNORE_AlertTypes.split(',');
            ignoredAlertTypes = ignoredAlertTypes.map(a => a.trim());
        }

        var skipSelfHealingTickets = [];
        if (process.env.SKIP_SelfHealing_TicketIDs) {
            skipSelfHealingTickets = process.env.SKIP_SelfHealing_TicketIDs.split(',')
            skipSelfHealingTickets = skipSelfHealingTickets.map(a => parseInt(a.trim()));
        }
        
        let sophosToken = await getSophosToken(context);

        let sophosJWT = false;
        if (sophosToken && sophosToken.access_token) {
            sophosJWT = sophosToken.access_token;
        }

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

                        // Verify the Autotask API key works (the library doesn't always provide a nice error message)
                        var useAutotaskAPI = true;
                        var autotaskTest = await autotask.Companies.get(0); // we need to do a call for the autotask module to get the zone info
                        try {
                            let fetchParms = {
                                method: 'GET',
                                headers: {
                                "Content-Type": "application/json",
                                "User-Agent": "Apigrate/1.0 autotask-restapi NodeJS connector"
                                }
                            };
                            fetchParms.headers.ApiIntegrationcode = process.env.AUTOTASK_INTEGRATION_CODE;
                            fetchParms.headers.UserName =  process.env.AUTOTASK_USER;
                            fetchParms.headers.Secret = process.env.AUTOTASK_SECRET;

                            let test_url = `${autotask.zoneInfo ? autotask.zoneInfo.url : autotask.base_url}V${autotask.version}/Companies/entityInformation`;
                            let response = await fetch(`${test_url}`, fetchParms);
                            if(!response.ok){
                                var result = await response.text();
                                if (!result) {
                                    result = `${response.status} - ${response.statusText}`;
                                }
                                throw result;
                            } else {
                                context.log(`Successfully connected to Autotask. (${response.status} - ${response.statusText})`)
                            }
                        } catch (error) {
                            if (error.startsWith("401")) {
                                error = `API Key Unauthorized. (${error})`
                            }
                            context.error(error);
                            useAutotaskAPI = false;
                        }

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

                            var when = new Date(alert.when);
                            var description = `${alert.description} \nSeverity: ${alert.severity} \nCompany: ${sophosCompany} \nDevice: ${alert.location}`;
                            if (alert.data && alert.data.source_info && alert.data.source_info.ip) {
                                description += `\nIP: ${alert.data.source_info.ip}`;
                            }
                            description += `\nEvent Type: ${alert.type} \nID: ${alert.id} \nWhen: ${when.toLocaleDateString('en-us', { weekday:"long", year:"numeric", month:"short", day:"numeric"})} \n\nSee the Sophos portal for more details.`;

                            // See if there are any existing tickets of this type and for this device
                            let tickets = null;
                            if (useAutotaskAPI) {
                                tickets = await searchAutotaskTickets(context, autotask, autotaskID, "Sophos Alert: ", alert.location, alert.type);
                            }

                            if (tickets && tickets.length > 0) {
                                // Existing ticket found, add notes
                                // get latest ticket
                                let existingTicket = tickets.reduce((a, b) => new Date(a.createDate) > new Date(b.createDate) ? a : b);

                                if (existingTicket) {
                                    if (!existingTicket.description.includes(alert.id)) {
                                        let updateNote = {
                                            "TicketID": existingTicket.id,
                                            "Title": "New Alert",
                                            "Description": description,
                                            "NoteType": 1,
                                            "Publish": 1
                                        }
                                        await autotask.TicketNotes.create(existingTicket.id, updateNote);
                                        context.log("New ticket note added on ticket id: " + existingTicket.id);
                                    } else {
                                        context.log("Skipped adding ticket note on ticket id (ticket is for this alert already):" + existingTicket.id);
                                    }
                                }
                            } else {
                                // No existing ticket found, create a new one
                                if (process.env.HOW_TO_DOCUMENTATION_LINK) {
                                    description += '\n\nHow To Documentation: ' + process.env.HOW_TO_DOCUMENTATION_LINK;
                                }

                                // Get primary location and default contract
                                var contractID = null;
                                var location = null;
                                if (useAutotaskAPI) {
                                    contractID = await getAutotaskContractID(autotask, autotaskID);
                                    location = await getAutotaskLocation(autotask, autotaskID);
                                }

                                // Get related device if applicable
                                var customerDevices = alertDevices[alert.customer_id];
                                var alertDevice = null;
                                if (customerDevices && customerDevices.length > 0) {
                                    alertDevice = customerDevices.filter(device => device.id == alert.data.endpoint_id)[0];
                                }
                                var deviceID = null;
                                if (useAutotaskAPI && alertDevice) {
                                    deviceID = await getAutotaskDevice(autotask, autotaskID, alertDevice);
                                }
                                var title = `Sophos Alert: "${alert.description}"`;
                                var includeAlertLocation = false;
                                if (!title.includes(alert.location)) {
                                    title = title + ` on "${alert.location}"`;
                                    includeAlertLocation = true;
                                }
                                var titleLength = title.length;
                                if (titleLength > 140) {
                                    // title is too long, lets cut it down to 140 characters
                                    var cutOff = titleLength - 140;
                                    var cutDescription = alert.description.substring(0, (alert.description.length - cutOff) - 3) + "...";
                                    var title = `Sophos Alert: "${cutDescription}"`;
                                    if (includeAlertLocation) {
                                        title = title + ` on "${alert.location}"`;
                                    }
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

                            await createAutotaskTicket(context, autotask, newTicket);
                            }
                        };

                        // Close tickets on up alerts
                        if (useAutotaskAPI) {
                            for (i = 0; i < upAlerts.length; i++) {
                                var alert = upAlerts[i];
                                context.log("Processing UP alert: " + alert.id)
                                // Go through each up alert and find the relevant ticket in Autotask then self-heal it
                                var sophosTenant = (sophosTenants.items.filter(tenant => tenant.id == alert.customer_id))[0];
                                let sophosCompany = sophosTenant.name;
                                let autotaskID = 0;
                                if (sophosCompany) {
                                    autotaskID = orgMapping[sophosCompany];
                                }
                                
                                let tickets = await searchAutotaskTickets(context, autotask, autotaskID, "Sophos Alert: ", alert.location, upDownEvents[alert.type]);
                                if (tickets && tickets.length > 0) {
                                    context.log("Existing Tickets: " + tickets.length)
                                    // get latest ticket
                                    let downTicket = tickets.reduce((a, b) => new Date(a.createDate) > new Date(b.createDate) ? a : b);

                                    if (downTicket) {
                                        if (skipSelfHealingTickets.includes(downTicket.id)) {
                                            context.log("Skipped self healing of ticket id: " + downTicket.id)
                                            continue;
                                        }
                                        
                                        let closingNote = {
                                            "TicketID": downTicket.id,
                                            "Title": "Self-Healing Update",
                                            "Description": "[Self-Healing] " + alert.description,
                                            "NoteType": 1,
                                            "Publish": 1
                                        }
                                        await autotask.TicketNotes.create(downTicket.id, closingNote);

                                        let closingTicket = {
                                            "id": downTicket.id,
                                            "Status": (downTicket.assignedResourceID ? 13 : 5)
                                        }
                                        await autotask.Tickets.update(closingTicket);

                                        // Close sophos down alert
                                        var alertIDMatches = idRegex.exec(downTicket.description);
                                        if (alertIDMatches) {
                                            var alertID = alertIDMatches[1];
                                            if (alertID) {
                                                closeSophosAlert(context, sophosJWT, sophosTenant, alertID);
                                                context.log("Closed the Sophos down alert.")
                                            }
                                        }

                                        // Close sophos up alert
                                        closeSophosAlert(context, sophosJWT, sophosTenant, alert.id);
                                        context.log("Closed the Sophos up alert.")
                                    } else {
                                        context.log("No latest down ticket found.")
                                    }
                                }
                            }

                            // Close tickets where the original alert no longer exists (closed but we don't get an up alert)
                            var allSophosAlertTickets = await searchAutotaskTickets(context, autotask, false, "Sophos Alert: ");
                            if (allSophosAlertTickets && allSophosAlertTickets.length > 0) {
                                for (i = 0; i < allSophosAlertTickets.length; i++) {
                                    var alertTicket = allSophosAlertTickets[i];
                                    var alertIDMatches = idRegex.exec(alertTicket.description);
                                    if (alertIDMatches) {
                                        var alertID = alertIDMatches[1];
                                        if (alertID) {
                                            var sophosCompanyName = getKeyByValue(orgMapping, alertTicket.companyID)
                                            var sophosTenant = (sophosTenants.items.filter(tenant => tenant.name == sophosCompanyName))[0];

                                            if (!sophosTenant || sophosTenant == undefined) {
                                                continue
                                            }

                                            sophosAlert = await getSophosAlert(context, sophosJWT, sophosTenant, alertID);
                                            context.log("Alert: " + sophosAlert);
                                            
                                            if (!sophosAlert || (sophosAlert.error && sophosAlert.error == "resourceNotFound")) {
                                                // Alert in Sophos has been closed, self-heal the related ticket
                                                let closingNote = {
                                                    "TicketID": alertTicket.id,
                                                    "Title": "Self-Healing Update",
                                                    "Description": "[Self-Healing] The Sophos alert is no longer open. Self-healing this ticket.",
                                                    "NoteType": 1,
                                                    "Publish": 1
                                                }
                                                await autotask.TicketNotes.create(alertTicket.id, closingNote);
                
                                                let closingTicket = {
                                                    "id": alertTicket.id,
                                                    "Status": (alertTicket.assignedResourceID ? 13 : 5)
                                                }
                                                await autotask.Tickets.update(closingTicket);
                                            }
                                        }
                                    }
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
                        context.error("Could not update lastRun.dat: " + error);
                    }
                }
            }
        }

        context.log('JavaScript timer trigger function ran!', timeStamp);
    }
});

function timeout(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

function getKeyByValue(object, value) {
    return Object.keys(object).find(key => object[key] == value);
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
        context.error(error);
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
        context.warn(error);
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
        context.error(error);
    }

    if (sophosTenantsJson && sophosTenantsJson.pages && sophosTenantsJson.pages.total > 1) {
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
                context.error(error);
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
    let retryUrls = [];

    try {
        tenants.items.filter(t => t !== undefined).filter(t => t.status && t.status == 'active').forEach(function(tenant) {
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
    } catch (err) {
        context.error(err); 
        context.warn(tenants.items[0]);
        context.warn(tenants.items);
    }

    const limit = RateLimit(12);
    const fetchFromApi = ({url, fetchHeader}, retry) => {
        const response = fetch(url, fetchHeader)
            .then((res) => res.json())
            .catch((error) => {
                if (!retry) {
                    retryUrls.push({url, fetchHeader});
                    context.warn(error);
                } else {
                    context.error(error); 
                }
                return;
            });
        return response;
    };

    let alerts = [];
    for (const query of queryUrls) {
        await limit();
        fetchFromApi(query, false).then((result) => {
            if (result && result.items) {
                alerts = alerts.concat(result.items);
            }
        });
    }

    if (retryUrls && retryUrls.length > 0) {
        for (const query of retryUrls) {
            await limit();
            fetchFromApi(query, true).then((result) => {
                if (result && result.items) {
                    alerts = alerts.concat(result.items);
                }
            });
        }
    }

    return alerts;
}

async function getSophosAlert(context, token, tenant, alertID) {
    let url = 'https://api-' + tenant.dataRegion + '.central.sophos.com/common/v1/alerts/' + alertID

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

async function closeSophosAlert(context, token, tenant, alertID) {
    // Marks the alert as acknowledged
    let url = 'https://api-' + tenant.dataRegion + '.central.sophos.com/common/v1/alerts/' + alertID + '/actions'

    let requestBody = {        
        action: "acknowledge",
        message: "Acknowledged by Autotask Integration"
    }

    let fetchHeader = {
        method: "POST",
        headers: {
            Authorization: "Bearer " + token,
            "X-Tenant-ID": tenant.id,
            "Accept": "application/json"
        },
        body: JSON.stringify(requestBody)
    };

    let response = fetch(url, fetchHeader)
        .then((response) => response.json());

    return;
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
            if (!device.rmmDeviceAuditMacAddress || !deviceDetails.macAddresses) {
                return false;
            }
            var rmmMacAddresses = device.rmmDeviceAuditMacAddress.replace(/^\[|\]$/gm,'').split(', ');
            var intersection = deviceDetails.macAddresses.filter(addr => rmmMacAddresses.includes(addr));
            return intersection.length > 0;
        });
        if (filteredDevices.length > 0) {
            device.items = filteredDevices;
        }

        if (device.items.length > 1 && deviceDetails.associatedPerson && deviceDetails.associatedPerson.viaLogin) {  
            filteredDevices = device.items.filter(device => device.rmmDeviceAuditLastUser == deviceDetails.associatedPerson.viaLogin);
            if (filteredDevices.length > 0) {
                device.items = filteredDevices;
            }
        }

        if (device.items.length > 1 && deviceDetails.ipv4Addresses) {
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
        context.error(error);
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
            "Subject": newTicket.Title,
            "HTMLContent": newTicket.Description.replace(new RegExp('\r?\n','g'), "<br />")
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
            context.warn("Ticket creation failed. Backup email sent to support.");
        } catch (error) {
            context.error("Ticket creation failed. Sending an email as a backup also failed.");
            context.error(error);
        }
        ticketID = null;
    }
    return ticketID;
}
