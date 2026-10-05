const { app } = require('@azure/functions');
const {AutotaskRestApi} = require('@apigrate/autotask-restapi');
const axios = require('axios');
const { BlobServiceClient } = require('@azure/storage-blob');
const { DefaultAzureCredential } = require('@azure/identity');
const orgMapping = require('../OrgMapping.json');
const upDownEvents = require('../UpDownEvents.json');
var idRegex = /ID: ([0-9a-fA-F]{8}\b-[0-9a-fA-F]{4}\b-[0-9a-fA-F]{4}\b-[0-9a-fA-F]{4}\b-[0-9a-fA-F]{12})(\n| )/m

app.timer('SophosAlerts_AutotaskIntegration', {
    schedule: "0 */20 * * * *",
    handler: async (myTimer, context) => {
        var timeStamp = new Date().toISOString();
        var lastRun = false
        var lastRunUnixTimestamp = false;
        var ignoredAlertTypes = [];
        let shouldUpdateLastRunCheckpoint = false;

        // Initialize the client
        const blobServiceClient = getBlobServiceClient();
        const containerClient = blobServiceClient.getContainerClient("function-state");
        const blockBlobClient = containerClient.getBlockBlobClient("lastRun.dat");
        const sophosMetadataCacheBlobClient = containerClient.getBlockBlobClient("sophosMetadata.json");
        const closedAlertsCheckBlobClient = containerClient.getBlockBlobClient("lastClosedAlertsCheck.dat");
        const sophosAlertQueryFailuresBlobClient = containerClient.getBlockBlobClient("sophosAlertQueryFailures.json");
        const autotaskLocationsCacheBlobClient = containerClient.getBlockBlobClient("autotaskLocations.json");
        const sophosDevicesCacheBlobClient = containerClient.getBlockBlobClient("sophosDevices.json");
        const autotaskDevicesCacheBlobClient = containerClient.getBlockBlobClient("autotaskDevices.json");

        context.log("Starting SophosAlerts_AutotaskIntegration function at: " + timeStamp);

        const updateLastRunCheckpoint = async () => {
            try {
                await containerClient.createIfNotExists();
                await blockBlobClient.uploadData(Buffer.from(timeStamp));
                context.log("Updated lastRun.dat in Blob Storage to: " + timeStamp);
            } catch (error) {
                context.error("Could not update lastRun.dat in Blob Storage: " + error);
            }
        };

        const updateClosedAlertsCheck = async () => {
            try {
                await containerClient.createIfNotExists();
                await closedAlertsCheckBlobClient.uploadData(Buffer.from(timeStamp));
                context.log("Updated lastClosedAlertsCheck.dat in Blob Storage to: " + timeStamp);
            } catch (error) {
                context.error("Could not update lastClosedAlertsCheck.dat in Blob Storage: " + error);
            }
        };

        try {
            try {
                const containerExists = await containerClient.exists();
                if (containerExists && (await blockBlobClient.exists())) {
                    const downloadResponse = await blockBlobClient.downloadToBuffer();
                    lastRun = new Date(downloadResponse.toString("utf-8"));
                    context.log("Last run read from Blob Storage: " + lastRun.toISOString());
                } else {
                    context.warn("This script has never been run before.");
                }
            } catch (error) {
                const errMessage = error && error.message ? error.message : String(error);
                if (errMessage.includes('not supported by Azurite') || errMessage.includes('skipApiVersionCheck')) {
                    context.error('Azurite API version mismatch detected. Local Azurite is older than the Azure Storage SDK version in this project.');
                    context.error('Suggested startup command: npx azurite --skipApiVersionCheck');
                    context.error('Full error: ' + errMessage);
                    return;
                }
                context.error("Error reading lastRun state from Blob Storage: " + error);
            }

            if (lastRun && !isNaN(lastRun.getTime())) {
                lastRunUnixTimestamp = Math.floor(lastRun.getTime() / 1000);

                // if timestamp is more than 24 hours old, reset (cannot exceed 24 hours)
                var curTimeStamp = Math.round(Date.now() / 1000);
                if (lastRunUnixTimestamp < (curTimeStamp - (24 * 3600))) {
                    lastRunUnixTimestamp = false;
                    context.log("Last run is more than 24 hours old");
                }
            }

            if (process.env.IGNORE_AlertTypes) {
                ignoredAlertTypes = process.env.IGNORE_AlertTypes.split(',');
                ignoredAlertTypes = ignoredAlertTypes.map(a => a.trim());
            }

            context.log("Starting Sophos alerts sync");
            const sophosRateLimiter = createSophosRateLimiter(context, 7);
            let sophosToken = await getSophosToken(context, sophosRateLimiter);

            let sophosJWT = false;
            if (sophosToken && sophosToken.access_token) {
                sophosJWT = sophosToken.access_token;
            }

            if (!sophosJWT) {
                context.warn("Sophos token unavailable. Skipping alert sync for this run.");
                return;
            }

            let sophosPartnerID;
            let sophosTenants;
            let usedSophosMetadataCache = false;
            const sophosMetadataCache = await readSophosMetadataCache(context, sophosMetadataCacheBlobClient);

            if (isFreshSophosMetadataCache(sophosMetadataCache)) {
                sophosPartnerID = sophosMetadataCache.partnerID;
                sophosTenants = sophosMetadataCache.tenants;
                usedSophosMetadataCache = true;
                context.log("Using cached Sophos partner ID and tenant list.");
            } else {
                sophosPartnerID = await getSophosPartnerID(context, sophosJWT, sophosRateLimiter);
                if (!sophosPartnerID) {
                    context.warn("Sophos partner ID unavailable. Skipping alert sync for this run.");
                    return;
                }

                // Get list of tenants, we need to handle each on an individual basis
                sophosTenants = await getSophosTenants(context, sophosJWT, sophosPartnerID, sophosRateLimiter);
                if (sophosTenants && sophosTenants.items && sophosTenants.items.length > 0) {
                    await writeSophosMetadataCache(context, sophosMetadataCacheBlobClient, {
                        cachedAt: new Date().toISOString(),
                        partnerID: sophosPartnerID,
                        tenants: sophosTenants
                    }, containerClient);
                }
            }

            if (!sophosTenants || !sophosTenants.items || sophosTenants.items.length === 0) {
                context.log("No Sophos tenants returned. Skipping alert sync for this run.");
                return;
            }

            const activeTenants = sophosTenants.items.filter(t => t && t.status && t.status === 'active');
            if (activeTenants.length === 0) {
                context.log("No active Sophos tenants found. Skipping alert sync for this run.");
                return;
            }

            if (!usedSophosMetadataCache) {
                await timeout(1000); // wait a second to prevent rate limiting after metadata requests
            }
            let alerts;
            try {
                alerts = await getSophosSiemAlerts(context, sophosJWT, sophosTenants, lastRunUnixTimestamp);
                shouldUpdateLastRunCheckpoint = true;
            } catch (error) {
                const failureState = await recordSophosAlertQueryFailure(
                    context,
                    sophosAlertQueryFailuresBlobClient,
                    containerClient
                );
                if (failureState.failureCount > 4) {
                    context.error("Sophos alerts query failed; leaving lastRun.dat unchanged so the time window is retried.");
                    context.error(error);
                } else {
                    context.warn(`Sophos alert query failure recorded (${failureState.failureCount}/5 in 24 hours; latest ${failureState.lastFailureAt}).`);
                }
                return;
            }

            if (!alerts || alerts.length === 0) {
            context.log("No Sophos alerts returned for this run. Skipping alert sync for this run.");
                return;
            }

            var filteredAlerts = deduplicateAlerts(alerts.filter(alert => alert && alert.severity && alert.severity != "low"));
            var upAlerts = deduplicateAlerts(alerts.filter(alert => alert && alert.severity == "low" && Object.keys(upDownEvents).includes(alert.type)));

            if (!shouldProcessActionableAlerts(filteredAlerts, upAlerts)) {
                context.log("No actionable Sophos alerts found. Skipping alert sync for this run.");
                return;
            }

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
                fetchParms.headers.UserName = process.env.AUTOTASK_USER;
                fetchParms.headers.Secret = process.env.AUTOTASK_SECRET;

                let test_url = `${autotask.zoneInfo ? autotask.zoneInfo.url : autotask.base_url}V${autotask.version}/Companies/entityInformation`;
                let response = await fetch(`${test_url}`, fetchParms);
                if (!response.ok) {
                    var result = await response.text();
                    if (!result) {
                        result = `${response.status} - ${response.statusText}`;
                    }
                    throw result;
                } else {
                    context.log(`Successfully connected to Autotask. (${response.status} - ${response.statusText})`);
                }
            } catch (error) {
                if (error.startsWith("401")) {
                    error = `API Key Unauthorized. (${error})`;
                }
                context.error(error);
                useAutotaskAPI = false;
            }

            const autotaskLocationsCache = useAutotaskAPI
                ? await readAutotaskLocationsCache(context, autotaskLocationsCacheBlobClient)
                : { companies: {} };
            const sophosDevicesCache = await readSophosDevicesCache(context, sophosDevicesCacheBlobClient);
            const autotaskDevicesCache = useAutotaskAPI
                ? await readAutotaskDevicesCache(context, autotaskDevicesCacheBlobClient)
                : { devices: {} };
            const ticketSearchCache = new Map();

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
                    tickets = await getCachedAutotaskTickets(
                        context,
                        autotask,
                        ticketSearchCache,
                        autotaskID,
                        "Sophos Alert: ",
                        alert.location,
                        alert.type
                    );
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
                            };
                            await autotask.TicketNotes.create(existingTicket.id, updateNote);
                            context.log("New ticket note added on ticket id: " + existingTicket.id);
                            await timeout(500); // wait a moment to prevent API throttling when creating multiple notes in a row
                        } else {
                            context.log("Skipped adding ticket note on ticket id (ticket is for this alert already):" + existingTicket.id);
                        }
                    }
                } else {
                    // No existing ticket found, create a new one
                    if (process.env.HOW_TO_DOCUMENTATION_LINK) {
                        description += '\n\nHow To Documentation: ' + process.env.HOW_TO_DOCUMENTATION_LINK;
                    }

                    // Get primary location
                    var location = null;
                    if (useAutotaskAPI) {
                        location = await getCachedAutotaskLocation(
                            context,
                            autotask,
                            autotaskID,
                            autotaskLocationsCache,
                            autotaskLocationsCacheBlobClient,
                            containerClient
                        );
                    }

                    // Get related device only when a new ticket actually needs one.
                    var customerDevices = [];
                    var endpointID = alert.data && alert.data.endpoint_id;
                    var sophosTenant = sophosTenants.items.filter(tenant => tenant.id == alert.customer_id)[0];
                    if (useAutotaskAPI && sophosTenant && endpointID) {
                        var devices = await getCachedSophosDevices(
                            context,
                            sophosJWT,
                            sophosTenant,
                            [endpointID],
                            sophosRateLimiter,
                            sophosDevicesCache,
                            sophosDevicesCacheBlobClient,
                            containerClient
                        );
                        customerDevices = devices && devices.items ? devices.items : [];
                    }
                    var alertDevice = null;
                    if (customerDevices && customerDevices.length > 0) {
                        alertDevice = customerDevices.filter(device => device.id == endpointID)[0];
                    }
                    var deviceID = null;
                    if (useAutotaskAPI && alertDevice) {
                        deviceID = await getCachedAutotaskDevice(
                            context,
                            autotask,
                            autotaskID,
                            alertDevice,
                            autotaskDevicesCache,
                            autotaskDevicesCacheBlobClient,
                            containerClient
                        );
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

                    if (title.includes("detected ransomware")) {
                        description += '\n\n\n!!! A related RANSOMWARE email has been sent to notifications@seatosky.com. Check the email for more info.';
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
                        Title: title,
                        Description: description
                    };
                    if (deviceID) {
                        newTicket.ConfigurationItemID = deviceID;
                    }

                    await createAutotaskTicket(context, autotask, newTicket);
                    ticketSearchCache.delete(getAutotaskTicketSearchKey(autotaskID, "Sophos Alert: ", alert.location, alert.type));
                }
            }

            // Close tickets on up alerts
            if (useAutotaskAPI) {
                for (i = 0; i < upAlerts.length; i++) {
                    var alert = upAlerts[i];
                    context.log("Processing UP alert: " + alert.id);
                    // Go through each up alert and find the relevant ticket in Autotask then self-heal it
                    var sophosTenant = (sophosTenants.items.filter(tenant => tenant.id == alert.customer_id))[0];
                    let sophosCompany = sophosTenant.name;
                    let autotaskID = 0;
                    if (sophosCompany) {
                        autotaskID = orgMapping[sophosCompany];
                    }

                    let tickets = await getCachedAutotaskTickets(
                        context,
                        autotask,
                        ticketSearchCache,
                        autotaskID,
                        "Sophos Alert: ",
                        alert.location,
                        upDownEvents[alert.type]
                    );
                    if (tickets && tickets.length > 0) {
                        context.log("Existing Tickets: " + tickets.length);
                        // get latest ticket
                        let downTicket = tickets.reduce((a, b) => new Date(a.createDate) > new Date(b.createDate) ? a : b);

                        if (downTicket) {
                            let closingNote = {
                                "TicketID": downTicket.id,
                                "Title": "Self-Healing Update",
                                "Description": "[Self-Healing] " + alert.description,
                                "NoteType": 1,
                                "Publish": 1
                            };
                            await autotask.TicketNotes.create(downTicket.id, closingNote);

                            let closingTicket = {
                                "id": downTicket.id,
                                "Status": (downTicket.assignedResourceID ? 13 : 5)
                            };
                            await autotask.Tickets.update(closingTicket);

                            // Close sophos down alert
                            var alertIDMatches = idRegex.exec(downTicket.description);
                            if (alertIDMatches) {
                                var alertID = alertIDMatches[1];
                                if (alertID) {
                                    closeSophosAlert(context, sophosJWT, sophosTenant, alertID, sophosRateLimiter);
                                    context.log("Closed the Sophos down alert.");
                                }
                            }

                            // Close sophos up alert
                            closeSophosAlert(context, sophosJWT, sophosTenant, alert.id, sophosRateLimiter);
                            context.log("Closed the Sophos up alert.");
                        } else {
                            context.log("No latest down ticket found.");
                        }
                    }
                }

                const lastClosedAlertsCheck = await readTimestampBlob(context, closedAlertsCheckBlobClient, "lastClosedAlertsCheck.dat");
                if (shouldRunClosedAlertsCheck(lastClosedAlertsCheck)) {
                    try {
                        // Close tickets where the original alert no longer exists (closed but we don't get an up alert)
                        var allSophosAlertTickets = await searchAutotaskTickets(context, autotask, false, "Sophos Alert: ");
                        if (allSophosAlertTickets && allSophosAlertTickets.length > 0) {
                            for (i = 0; i < allSophosAlertTickets.length; i++) {
                                var alertTicket = allSophosAlertTickets[i];
                                var alertIDMatches = idRegex.exec(alertTicket.description);
                                if (alertIDMatches) {
                                    var alertID = alertIDMatches[1];
                                    if (alertID) {
                                        const sophosCompanyName = getKeyByValue(orgMapping, alertTicket.companyID);
                                        const sophosTenant = (sophosTenants.items.filter(tenant => tenant.name == sophosCompanyName))[0];

                                        if (!sophosTenant || sophosTenant == undefined) {
                                            continue;
                                        }

                                        sophosAlert = await getSophosAlert(context, sophosJWT, sophosTenant, alertID, sophosRateLimiter);

                                        if (!sophosAlert || (sophosAlert.error && sophosAlert.error == "resourceNotFound")) {
                                            // Alert in Sophos has been closed, self-heal the related ticket
                                            let closingNote = {
                                                "TicketID": alertTicket.id,
                                                "Title": "Self-Healing Update",
                                                "Description": "[Self-Healing] The Sophos alert is no longer open. Self-healing this ticket.",
                                                "NoteType": 1,
                                                "Publish": 1
                                            };
                                            await autotask.TicketNotes.create(alertTicket.id, closingNote);

                                            let closingTicket = {
                                                "id": alertTicket.id,
                                                "Status": (alertTicket.assignedResourceID ? 13 : 5)
                                            };
                                            await autotask.Tickets.update(closingTicket);
                                        }
                                    }
                                }
                            }
                        }
                        await updateClosedAlertsCheck();
                    } catch (error) {
                        context.error("Closed Sophos alert sweep failed; it will be retried on the next run.");
                        context.error(error);
                    }
                } else {
                    context.log("Skipping closed Sophos alert sweep; it ran less than two hours ago.");
                }
            }
        } finally {
            if (shouldUpdateLastRunCheckpoint) {
                await updateLastRunCheckpoint();
            } else {
                context.log("Sophos alert query did not complete successfully; lastRun.dat was not updated.");
            }
        }
    }
});

function isLocalDevelopment() {
  if (process.env.AZURE_FUNCTIONS_ENVIRONMENT === 'Development') {
    return true;
  }

  const isAzureRuntime =
    !!process.env.WEBSITE_SITE_NAME ||
    !!process.env.WEBSITE_INSTANCE_ID ||
    !!process.env.AzureWebJobsStorage__blobServiceUri ||
    !!process.env.AzureWebJobsStorage__queueServiceUri;

  return !isAzureRuntime;
}

function getBlobServiceClient() {
    // Check if running locally
    const isLocalDev = isLocalDevelopment();

    // Local Development (Always force connection string / Azurite)
    if (isLocalDev) {
        return BlobServiceClient.fromConnectionString(
            process.env.AzureWebJobsStorage || "UseDevelopmentStorage=true"
        );
    }

    // Flex Consumption in Azure (Managed Identity)
    if (process.env.AzureWebJobsStorage__blobServiceUri) {
        const credential = new DefaultAzureCredential();

        return new BlobServiceClient(
            process.env.AzureWebJobsStorage__blobServiceUri,
            credential
        );
    }
    
    // Standard Consumption in Azure (Connection String)
    if (process.env.AzureWebJobsStorage) {
        return BlobServiceClient.fromConnectionString(
            process.env.AzureWebJobsStorage
        );
    }

    throw new Error("No valid storage environment variables found.");
}

function timeout(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

function createSophosRateLimiter(context, requestsPerSecond = 10) {
    const minIntervalMs = Math.ceil(1000 / requestsPerSecond);
    let lastRequestAt = 0;

    return async function waitForSlot() {
        const now = Date.now();
        const elapsedMs = now - lastRequestAt;
        const waitMs = Math.max(0, minIntervalMs - elapsedMs);

        if (waitMs > 0) {
            await timeout(waitMs);
        }

        lastRequestAt = Date.now();
    };
}

function createAdaptiveSophosRetryState() {
    return {
        recentSuccesses: 0,
        recent429s: 0,
        consecutive429s: 0,
        lastStatus: null
    };
}

function getAdaptiveSophosRetryMaxAttempts(state = {}, options = {}) {
    const baseMaxAttempts = Math.max(1, Number(options.baseMaxAttempts) || 3);
    const maxSafeAttempts = Math.max(baseMaxAttempts, Number(options.maxSafeAttempts) || 8);
    const recentSuccesses = Math.max(0, Number(state.recentSuccesses) || 0);
    const consecutive429s = Math.max(0, Number(state.consecutive429s) || 0);
    const recent429s = Math.max(0, Number(state.recent429s) || 0);

    if (consecutive429s >= 3 || recent429s >= 3) {
        return baseMaxAttempts;
    }

    if (recentSuccesses >= 2) {
        const opportunity = Math.min(maxSafeAttempts - baseMaxAttempts, Math.floor(recentSuccesses / 2));
        return Math.min(maxSafeAttempts, baseMaxAttempts + opportunity);
    }

    return baseMaxAttempts;
}

function shouldStopSophosTenantLoop(state = {}, options = {}) {
    const consecutive429s = Math.max(0, Number(state.consecutive429s) || 0);
    const recent429s = Math.max(0, Number(state.recent429s) || 0);
    const recentSuccesses = Math.max(0, Number(state.recentSuccesses) || 0);
    const minConsecutive429s = Math.max(2, Number(options.minConsecutive429s) || 3);
    const minRecent429s = Math.max(2, Number(options.minRecent429s) || 3);

    return (consecutive429s >= minConsecutive429s) || (recent429s >= minRecent429s && recentSuccesses === 0);
}

async function sleepWithContext(context, ms, reason) {
    if (reason) {
        context && context.log && context.log(`Sophos ${reason}: waiting ${ms}ms before retrying`);
    }
    await timeout(ms);
}

async function runSophosRequestWithRetry(context, label, requestFn, options = {}) {
    const adaptiveState = options.state || createAdaptiveSophosRetryState();
    const baseMaxAttempts = Math.max(1, Number(options.maxAttempts) || 3);
    const maxAttempts = getAdaptiveSophosRetryMaxAttempts(adaptiveState, {
        baseMaxAttempts,
        maxSafeAttempts: Math.max(baseMaxAttempts, Number(options.maxSafeAttempts) || 8)
    });
    const backoffMs = Math.max(0, Number(options.backoffMs) || 2000);

    let lastError = null;
    for (let attempt = 1; attempt <= maxAttempts; attempt++) {
        try {
            const result = await requestFn(attempt);
            adaptiveState.recentSuccesses = (Number(adaptiveState.recentSuccesses) || 0) + 1;
            adaptiveState.recent429s = Math.max(0, (Number(adaptiveState.recent429s) || 0) - 1);
            adaptiveState.consecutive429s = 0;
            adaptiveState.lastStatus = 'success';
            return result;
        } catch (error) {
            lastError = error;
            const status = error && error.response ? error.response.status : null;

            if (status !== 429) {
                adaptiveState.recentSuccesses = 0;
                adaptiveState.recent429s = 0;
                adaptiveState.consecutive429s = 0;
                adaptiveState.lastStatus = status || 'error';
                throw error;
            }

            adaptiveState.recent429s = (Number(adaptiveState.recent429s) || 0) + 1;
            adaptiveState.consecutive429s = (Number(adaptiveState.consecutive429s) || 0) + 1;
            adaptiveState.recentSuccesses = 0;
            adaptiveState.lastStatus = 429;

            if (attempt < maxAttempts) {
                context && context.warn && context.warn(`Sophos ${label} hit 429 on attempt ${attempt}/${maxAttempts}. Retrying in ${backoffMs}ms.`);
                await sleepWithContext(context, backoffMs, '429 backoff');
                continue;
            }

            context && context.warn && context.warn(`Sophos ${label} hit 429 on final attempt ${attempt}/${maxAttempts}. Stopping retries.`);
            return null;
        }
    }

    return null;
}

function shouldProcessActionableAlerts(alerts = [], upAlerts = []) {
    if (!Array.isArray(alerts) || !Array.isArray(upAlerts)) {
        return false;
    }

    const hasNonLowAlerts = alerts.some(alert => alert && alert.severity && alert.severity !== 'low');
    const hasUpAlerts = upAlerts.some(alert => alert && alert.type && Object.keys(upDownEvents).includes(alert.type));

    return hasNonLowAlerts || hasUpAlerts;
}

function deduplicateAlerts(alerts = []) {
    const seenAlertIDs = new Set();

    return alerts.filter(alert => {
        if (!alert || !alert.id) {
            return true;
        }

        if (seenAlertIDs.has(alert.id)) {
            return false;
        }

        seenAlertIDs.add(alert.id);
        return true;
    });
}

function getAutotaskTicketSearchKey(companyID, titleStart, deviceName, eventType) {
    return JSON.stringify({ companyID, titleStart, deviceName, eventType });
}

async function getCachedAutotaskTickets(context, autotaskAPI, cache, companyID, titleStart, deviceName, eventType) {
    const cacheKey = getAutotaskTicketSearchKey(companyID, titleStart, deviceName, eventType);
    if (cache.has(cacheKey)) {
        return cache.get(cacheKey);
    }

    const tickets = await searchAutotaskTickets(context, autotaskAPI, companyID, titleStart, deviceName, eventType);
    if (tickets !== undefined) {
        cache.set(cacheKey, tickets);
    }

    return tickets;
}

function getKeyByValue(object, value) {
    return Object.keys(object).find(key => object[key] == value);
}

function isFreshSophosMetadataCache(cache, now = Date.now()) {
    if (!cache || !cache.cachedAt || !cache.partnerID || !cache.tenants || !Array.isArray(cache.tenants.items)) {
        return false;
    }

    const cachedAt = Date.parse(cache.cachedAt);
    return Number.isFinite(cachedAt) && cachedAt <= now && now - cachedAt < 24 * 60 * 60 * 1000;
}

function shouldRunClosedAlertsCheck(lastCheck, now = Date.now()) {
    if (!(lastCheck instanceof Date) || isNaN(lastCheck.getTime())) {
        return true;
    }

    return lastCheck.getTime() <= now && now - lastCheck.getTime() >= 2 * 60 * 60 * 1000;
}

function isFreshAutotaskLocationCacheEntry(entry, now = Date.now()) {
    if (!entry || !Object.prototype.hasOwnProperty.call(entry, "location") || !entry.cachedAt) {
        return false;
    }

    const cachedAt = Date.parse(entry.cachedAt);
    return Number.isFinite(cachedAt) && cachedAt <= now && now - cachedAt < 7 * 24 * 60 * 60 * 1000;
}

function isFreshDeviceCacheEntry(entry, now = Date.now()) {
    if (!entry || !entry.cachedAt) {
        return false;
    }

    const cachedAt = Date.parse(entry.cachedAt);
    return Number.isFinite(cachedAt) && cachedAt <= now && now - cachedAt < 24 * 60 * 60 * 1000;
}

function updateSophosAlertQueryFailureState(state, now = Date.now()) {
    const cutoff = now - 24 * 60 * 60 * 1000;
    const failures = Array.isArray(state && state.failures)
        ? state.failures.filter(timestamp => {
            const failureTime = Date.parse(timestamp);
            return Number.isFinite(failureTime) && failureTime > cutoff && failureTime <= now;
        })
        : [];
    const lastFailureAt = new Date(now).toISOString();
    failures.push(lastFailureAt);

    return {
        failureCount: failures.length,
        lastFailureAt,
        failures
    };
}

async function recordSophosAlertQueryFailure(context, blobClient, containerClient, now = Date.now()) {
    let previousState = null;
    try {
        if (await blobClient.exists()) {
            const downloadResponse = await blobClient.downloadToBuffer();
            previousState = JSON.parse(downloadResponse.toString("utf-8"));
        }
    } catch (error) {
        context.warn("Could not read Sophos alert query failure state; starting a new 24-hour count: " + error);
    }

    const failureState = updateSophosAlertQueryFailureState(previousState, now);
    try {
        await containerClient.createIfNotExists();
        await blobClient.uploadData(Buffer.from(JSON.stringify(failureState)), {
            blobHTTPHeaders: { blobContentType: "application/json" }
        });
    } catch (error) {
        context.warn("Could not persist Sophos alert query failure state: " + error);
    }

    return failureState;
}

async function readSophosDevicesCache(context, blobClient) {
    try {
        if (!(await blobClient.exists())) {
            return { tenants: {} };
        }

        const downloadResponse = await blobClient.downloadToBuffer();
        const cache = JSON.parse(downloadResponse.toString("utf-8"));
        return cache && cache.tenants ? cache : { tenants: {} };
    } catch (error) {
        context.warn("Could not read Sophos device cache; refreshing devices as needed: " + error);
        return { tenants: {} };
    }
}

async function writeSophosDevicesCache(context, blobClient, cache, containerClient) {
    try {
        await containerClient.createIfNotExists();
        await blobClient.uploadData(Buffer.from(JSON.stringify(cache)), {
            blobHTTPHeaders: { blobContentType: "application/json" }
        });
        context.log("Updated Sophos device cache.");
    } catch (error) {
        context.warn("Could not update Sophos device cache; continuing without cache: " + error);
    }
}

async function getCachedSophosDevices(context, token, tenant, ids, rateLimiter, cache, blobClient, containerClient) {
    const deviceIDs = [...new Set((ids || []).filter(Boolean).map(String))];
    const tenantKey = String(tenant.id);
    const cachedEntry = cache.tenants[tenantKey];
    const hasFreshCache = isFreshDeviceCacheEntry(cachedEntry);
    const cachedDevices = hasFreshCache && cachedEntry.devices ? cachedEntry.devices : {};
    const missingDeviceIDs = deviceIDs.filter(deviceID => !Object.prototype.hasOwnProperty.call(cachedDevices, deviceID));

    if (missingDeviceIDs.length === 0) {
        return {
            items: deviceIDs.map(deviceID => cachedDevices[deviceID]).filter(Boolean)
        };
    }

    const devices = await getSophosDevices(context, token, tenant, missingDeviceIDs, rateLimiter);
    const updatedDevices = { ...cachedDevices };
    if (devices && Array.isArray(devices.items)) {
        for (const device of devices.items) {
            if (device && device.id !== undefined && device.id !== null) {
                updatedDevices[String(device.id)] = device;
            }
        }
    }

    cache.tenants[tenantKey] = {
        cachedAt: new Date().toISOString(),
        devices: updatedDevices
    };
    await writeSophosDevicesCache(context, blobClient, cache, containerClient);

    return {
        items: deviceIDs.map(deviceID => updatedDevices[deviceID]).filter(Boolean)
    };
}

async function readAutotaskDevicesCache(context, blobClient) {
    try {
        if (!(await blobClient.exists())) {
            return { devices: {} };
        }

        const downloadResponse = await blobClient.downloadToBuffer();
        const cache = JSON.parse(downloadResponse.toString("utf-8"));
        return cache && cache.devices ? cache : { devices: {} };
    } catch (error) {
        context.warn("Could not read Autotask device cache; refreshing devices as needed: " + error);
        return { devices: {} };
    }
}

async function writeAutotaskDevicesCache(context, blobClient, cache, containerClient) {
    try {
        await containerClient.createIfNotExists();
        await blobClient.uploadData(Buffer.from(JSON.stringify(cache)), {
            blobHTTPHeaders: { blobContentType: "application/json" }
        });
        context.log("Updated Autotask device cache.");
    } catch (error) {
        context.warn("Could not update Autotask device cache; continuing without cache: " + error);
    }
}

function getAutotaskDeviceCacheKey(autotaskID, deviceDetails) {
    const normalize = value => Array.isArray(value) ? [...value].sort() : (value || "");
    return JSON.stringify({
        companyID: autotaskID,
        hostname: deviceDetails.hostname || "",
        macAddresses: normalize(deviceDetails.macAddresses),
        login: deviceDetails.associatedPerson && deviceDetails.associatedPerson.viaLogin || "",
        ipv4Addresses: normalize(deviceDetails.ipv4Addresses)
    });
}

async function getCachedAutotaskDevice(context, autotaskAPI, autotaskID, deviceDetails, cache, blobClient, containerClient) {
    const deviceKey = getAutotaskDeviceCacheKey(autotaskID, deviceDetails);
    const cachedEntry = cache.devices[deviceKey];

    if (isFreshDeviceCacheEntry(cachedEntry) && Object.prototype.hasOwnProperty.call(cachedEntry, "deviceID")) {
        context.log("Using cached Autotask device for company " + autotaskID + ".");
        return cachedEntry.deviceID;
    }

    const deviceID = await getAutotaskDevice(autotaskAPI, autotaskID, deviceDetails);
    cache.devices[deviceKey] = {
        cachedAt: new Date().toISOString(),
        deviceID: deviceID || null
    };
    await writeAutotaskDevicesCache(context, blobClient, cache, containerClient);
    return deviceID;
}

async function readAutotaskLocationsCache(context, blobClient) {
    try {
        if (!(await blobClient.exists())) {
            return { companies: {} };
        }

        const downloadResponse = await blobClient.downloadToBuffer();
        const cache = JSON.parse(downloadResponse.toString("utf-8"));
        return cache && cache.companies ? cache : { companies: {} };
    } catch (error) {
        context.warn("Could not read Autotask location cache; refreshing locations as needed: " + error);
        return { companies: {} };
    }
}

async function writeAutotaskLocationsCache(context, blobClient, cache, containerClient) {
    try {
        await containerClient.createIfNotExists();
        await blobClient.uploadData(Buffer.from(JSON.stringify(cache)), {
            blobHTTPHeaders: { blobContentType: "application/json" }
        });
        context.log("Updated Autotask location cache.");
    } catch (error) {
        context.warn("Could not update Autotask location cache; continuing without cache: " + error);
    }
}

async function getCachedAutotaskLocation(context, autotaskAPI, autotaskID, cache, blobClient, containerClient) {
    const companyKey = String(autotaskID);
    const cachedEntry = cache.companies[companyKey];

    if (isFreshAutotaskLocationCacheEntry(cachedEntry)) {
        context.log("Using cached Autotask location for company " + companyKey + ".");
        return cachedEntry.location;
    }

    const location = await getAutotaskLocation(autotaskAPI, autotaskID);
    cache.companies[companyKey] = {
        cachedAt: new Date().toISOString(),
        location: location || null
    };
    await writeAutotaskLocationsCache(context, blobClient, cache, containerClient);
    return location;
}

async function readTimestampBlob(context, blobClient, blobName) {
    try {
        if (!(await blobClient.exists())) {
            return null;
        }

        const downloadResponse = await blobClient.downloadToBuffer();
        return new Date(downloadResponse.toString("utf-8"));
    } catch (error) {
        context.warn(`Could not read ${blobName}; running the check: ${error}`);
        return null;
    }
}

async function readSophosMetadataCache(context, blobClient) {
    try {
        if (!(await blobClient.exists())) {
            return null;
        }

        const downloadResponse = await blobClient.downloadToBuffer();
        return JSON.parse(downloadResponse.toString("utf-8"));
    } catch (error) {
        context.warn("Could not read Sophos metadata cache; refreshing from Sophos: " + error);
        return null;
    }
}

async function writeSophosMetadataCache(context, blobClient, cache, containerClient) {
    try {
        await containerClient.createIfNotExists();
        await blobClient.uploadData(Buffer.from(JSON.stringify(cache)), {
            blobHTTPHeaders: { blobContentType: "application/json" }
        });
        context.log("Updated Sophos partner and tenant metadata cache.");
    } catch (error) {
        context.warn("Could not update Sophos metadata cache; continuing without cache: " + error);
    }
}

async function getSophosToken(context, rateLimiter = null) {
    if (rateLimiter) {
        await rateLimiter();
    }

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

        if (!sophosToken.ok) {
            throw new Error(`Error getting Sophos token! Error: ${sophosToken.status} ${sophosToken.statusText}`)
        }

        return await sophosToken.json();
    } catch (error) {
        context.error(error);
        return null;
    }
}

async function getSophosPartnerID(context, token, rateLimiter = null) {
    if (rateLimiter) {
        await rateLimiter();
    }

    let url = 'https://api.central.sophos.com/whoami/v1';

    try {
        let sophosPartnerInfo = await fetch(url, {
            method: "GET",
            headers: {
                Authorization: "Bearer " + token,
            }
        });

        if (!sophosPartnerInfo.ok) {
            throw new Error(`Error getting partner ID! Error: ${sophosPartnerInfo.status} ${sophosPartnerInfo.statusText}`)
        }

        let sophosPartnerInfoJson = await sophosPartnerInfo.json();
        return sophosPartnerInfoJson.id;
    } catch (error) {
        context.warn(error);
        return null;
    }
}

async function getSophosTenants(context, token, partnerID, rateLimiter = null) {
    if (rateLimiter) {
        await rateLimiter();
    }

    let url = 'https://api.central.sophos.com/partner/v1/tenants?page';
    let sophosTenantsJson;

    try {
        let sophosTenants = await fetch(url + "Total=true", {
            method: "GET",
            headers: {
                Authorization: "Bearer " + token,
                "X-Partner-ID": partnerID
            }
        });

        if (!sophosTenants.ok) {
            throw new Error(`Error getting Sophos tenants! Error: ${sophosTenants.status} ${sophosTenants.statusText}`)
        }

        sophosTenantsJson =  await sophosTenants.json();
    } catch (error) {
        context.error(error);
    }

    if (sophosTenantsJson && sophosTenantsJson.pages && sophosTenantsJson.pages.total > 1) {
        var totalPages = sophosTenantsJson.pages.total;
        for (let i = 2; i <= totalPages; i++) {
            try {
                let sophosTenantsTemp = await fetch(url + "=" + i, {
                    method: "GET",
                    headers: {
                        Authorization: "Bearer " + token,
                        "X-Partner-ID": partnerID
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

async function getSophosDevices(context, token, tenant, ids = null, rateLimiter = null) {
    if (rateLimiter) {
        await rateLimiter();
    }

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

async function getSophosSiemAlerts(context, token, tenants, fromDate = false, rateLimiter = null) {
    let queryUrls = [];
    let retryUrls = [];
    let queryFailure = false;
    const sophosRateLimiter = rateLimiter || createSophosRateLimiter(context, 7);
    const adaptiveRetryState = createAdaptiveSophosRetryState();

    try {
        tenants.items.filter(t => t !== undefined).filter(t => t.status && t.status == 'active').forEach(function(tenant) {
            let url = 'https://api-' + tenant.dataRegion + '.central.sophos.com/siem/v1/alerts';
            if (fromDate && Number.isInteger(fromDate)) {
                url = url + '?from_date=' + fromDate + '&limit=1000';
            }

            let fetchHeader = {
                method: "GET",
                headers: {
                    Authorization: "Bearer " + token,
                    "X-Tenant-ID": tenant.id
                }
            };
            let axiosHeader = {
                headers: {
                    Authorization: "Bearer " + token,
                    "X-Tenant-ID": tenant.id,
                }
            };

            queryUrls.push({url, fetchHeader, axiosHeader});
        });
    } catch (err) {
        context.error(err);
        queryFailure = true;
    }

    let alerts = [];
    for (const query of queryUrls) {
        if (shouldStopSophosTenantLoop(adaptiveRetryState)) {
            context.warn('Sophos API is continuing to throttle; stopping tenant alert retrieval for this cycle.');
            queryFailure = true;
            break;
        }

        await sophosRateLimiter();

        try {
            const requestFn = async () => {
                const result = await axios.get(query.url, query.axiosHeader);
                if (result && result.data && result.data.items) {
                    return result.data.items;
                }
                return [];
            };

            const items = await runSophosRequestWithRetry(context, `SIEM alerts for ${query.url}`, requestFn, {
                maxAttempts: 3,
                maxSafeAttempts: 8,
                backoffMs: 2000,
                state: adaptiveRetryState
            });

            if (Array.isArray(items)) {
                alerts = alerts.concat(items);
            } else {
                queryFailure = true;
            }
        } catch (error) {
            queryFailure = true;
            retryUrls.push(query);
            context.log("Got error:" + error);
            context.warn(error);

            if (error && error.response) {
                context.log(error.response.data);
                context.log(error.response.status);
                context.log(error.response.headers);
            } else if (error && error.request) {
                context.log(error.request);
            } else if (error) {
                context.log('Error', error.message);
            }
            if (error && error.config) {
                context.log(error.config);
            }
        }
    }

    if (retryUrls && retryUrls.length > 0) {
        for (const query of retryUrls) {
            if (shouldStopSophosTenantLoop(adaptiveRetryState)) {
                context.warn('Sophos API is still throttling during retry requests; stopping retry tenant loop.');
                queryFailure = true;
                break;
            }

            await sophosRateLimiter();

            try {
                const requestFn = async () => {
                    const response = await fetch(query.url, query.fetchHeader);
                    if (response && response.status === 429) {
                        const error = new Error('TooManyRequests');
                        error.response = { status: 429 };
                        throw error;
                    }
                    if (!response.ok) {
                        const error = new Error(`Sophos alerts request failed: ${response.status} ${response.statusText}`);
                        error.response = { status: response.status };
                        throw error;
                    }

                    const parsedJson = await response.json();
                    return parsedJson && parsedJson.items ? parsedJson.items : [];
                };

                const items = await runSophosRequestWithRetry(context, `retry SIEM alerts for ${query.url}`, requestFn, {
                    maxAttempts: 3,
                    maxSafeAttempts: 8,
                    backoffMs: 2000,
                    state: adaptiveRetryState
                });

                if (Array.isArray(items)) {
                    alerts = alerts.concat(items);
                } else {
                    queryFailure = true;
                }
            } catch (error) {
                queryFailure = true;
                context.error(error);
            }
        }
    }

    if (queryFailure) {
        throw new Error('Sophos alerts query failed; lastRun.dat must not be advanced.');
    }

    return alerts;
}

async function getSophosAlert(context, token, tenant, alertID, rateLimiter = null) {
    if (rateLimiter) {
        await rateLimiter();
    }

    const url = `https://api-${tenant.dataRegion}.central.sophos.com/common/v1/alerts/${alertID}`;

    const fetchHeader = {
        method: "GET",
        headers: {
            "Authorization": `Bearer ${token}`,
            "X-Tenant-ID": tenant.id,
            "Accept": "application/json"
        }
    };

    try {
        const response = await fetch(url, fetchHeader);

        if (response.status === 204 || !response.ok) {
            return null; 
        }

        const text = await response.text();

        return text ? JSON.parse(text) : null;
    } catch (error) {
        context.error(`Error fetching Sophos alert ${alertID}:`, error);
        return null;
    }
}

async function closeSophosAlert(context, token, tenant, alertID, rateLimiter = null) {
    if (rateLimiter) {
        await rateLimiter();
    }

    // Marks the alert as acknowledged
    const url = `https://api-${tenant.dataRegion}.central.sophos.com/common/v1/alerts/${alertID}/actions`;

    const requestBody = {        
        action: "acknowledge",
        message: "Acknowledged by Autotask Integration"
    }

    const fetchHeader = {
        method: "POST",
        headers: {
            "Authorization": `Bearer ${token}`,
            "X-Tenant-ID": tenant.id,
            "Accept": "application/json"
        },
        body: JSON.stringify(requestBody)
    };

    try {
        const response = await fetch(url, fetchHeader);

        if (response.status === 204 || !response.ok) {
            return null;
        }

        const text = await response.text();
        return text ? JSON.parse(text) : null;
    } catch (error) {
        context.error(`Error closing Sophos alert ${alertID}:`, error);
        return null;
    }
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

    if (device && device.items && device.items.length > 0) {
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

module.exports = {
    createSophosRateLimiter,
    createAdaptiveSophosRetryState,
    getAdaptiveSophosRetryMaxAttempts,
    shouldStopSophosTenantLoop,
    runSophosRequestWithRetry,
    shouldProcessActionableAlerts,
    getSophosSiemAlerts,
    isFreshSophosMetadataCache,
    shouldRunClosedAlertsCheck,
    isFreshAutotaskLocationCacheEntry,
    isFreshDeviceCacheEntry,
    updateSophosAlertQueryFailureState,
    getAutotaskDeviceCacheKey,
    deduplicateAlerts,
    getAutotaskTicketSearchKey
};
