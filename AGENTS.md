# AGENTS.md

## Project overview

This repo is an Azure Functions timer app that polls Sophos for alerts and creates or updates tickets in Autotask. The main runtime entry point is [SophosAlerts_AutotaskIntegration/index.js](SophosAlerts_AutotaskIntegration/index.js).

## Important repo conventions

- Use Node.js for local development; the function is a timer trigger and not a typical web app.
- Local testing requires Azurite. Before debugging, start it with:
  - `npx azurite --skipApiVersionCheck`
- Persist runtime state in Azure Blob Storage, not in local files. The current checkpoint is `function-state/lastRun.dat`, the daily Sophos partner/tenant metadata cache is `function-state/sophosMetadata.json`, and the two-hour closed-alert sweep checkpoint is `function-state/lastClosedAlertsCheck.dat`; all are read and written through `BlobServiceClient` in [SophosAlerts_AutotaskIntegration/index.js](SophosAlerts_AutotaskIntegration/index.js). Local development uses Azurite through `AzureWebJobsStorage` or `UseDevelopmentStorage=true`; do not replace this with filesystem-based state or add local runtime-state files.
- The project uses environment settings from `local.settings.json` (copy from `local.settings.json.template`), and organization mapping from `OrgMapping.json` (copy from `OrgMapping.json.template`).
- This project intentionally filters out `low` severity Sophos alerts and uses a self-healing flow for `up` events.

## Files to know

- [README.md](README.md): setup and troubleshooting guidance.
- [package.json](package.json): scripts and dependencies.
- [local.settings.json.template](local.settings.json.template): required environment variables.
- [OrgMapping.json.template](OrgMapping.json.template): mapping format for Sophos companies to Autotask company IDs.
- [test/sophosRateLimiter.test.js](test/sophosRateLimiter.test.js): unit tests for rate limiting and retry behavior.

## Validation commands

Run the project checks with:

- `npm install`
- `node --test`

If you are testing the Azure Function locally:

- `func start`
- ensure Azurite is already running with `npx azurite --skipApiVersionCheck`

## Working rules for AI coding agents

- Prefer small, targeted edits in [SophosAlerts_AutotaskIntegration/index.js](SophosAlerts_AutotaskIntegration/index.js) and keep behavior aligned with the existing timer-trigger design.
- Preserve the current rate-limit and retry patterns for Sophos API calls; they are intentional and covered by tests.
- When changing environment variables or config behavior, update both the template files and the README so the setup remains consistent.
- Do not add broad refactors to this repo without a clear reason; the app is compact and relies on careful API sequencing.
