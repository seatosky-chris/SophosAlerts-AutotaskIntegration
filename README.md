# Sophos Alerts - Autotask Integration

This script is used to watch for Sophos Alerts and handles ticket creation in Autotask. This script is designed to be ran in an Azure function. It runs on a timer and every 10 minutes it will check for any new alerts that have been created in Sophos since the timer last ran. When an alert comes up, the script will create a new ticket in Autotask via the API. If the script cannot connect to the Autotask API, it will fallback to sending an email to an address of your choice. It will only create tickets for "medium" and "high" level alerts (not "low").

### Development Testing
Note that for testing purposes this function uses the Azurite DB Emulator. You must have this extension installed in VS Code and before doing any testing, you must start Azurite.

### Organization Mapping
For the script to work, you must map each organization in Sophos to the corresponding Company ID in Autotask. This is done in the OrgMapping.json document. Each line should map the full name of the organization in Sophos to the Autotask Company ID.

### Script Configuration
- Setup an Autotask API account with access to READ Companies, Locations, Contracts & ConfigurationItems, and to READ/WRITE Tickets and TicketNotes. Fill in the Autotask configuration with API account details in local.settings.json.
- Setup a Sophos API Key, fill in the API Key details in local.settings.json. This needs to be created at the Partner level in Sophos.
- Configure the Email forwarder details in local.settings.json. (See my Email Forwarder script.) This could also be configured to use something like SendGrid instead but the script may require minor modifications.
- Configure the default ticket options in local.settings.json. These are details on Queue ID, Issue Type, Sub Issue Type, and the Service Level Agreement ID that the new ticket will be created with.
- Setup the OrgMapping.json file with an entry for each company in Sophos that you want to integrate with Autotask. See the above "Organization Mapping" section.
- Push this to an Azure Function and ensure the environment variables get updated.
- The script will now run every 10 minutes to check for new alerts.

### Troubleshooting
When testing, I found that the Autotask package was unable to connect to the Autotask server due to the OpenSSL error: `unsafe legacy renegotiation disabled`. It appears this is related to a fix in OpenSSL that Autotask likely hasn't implemented on their server. I was able to fix this by using an older version of NodeJS, specifically version **16.13.0**.