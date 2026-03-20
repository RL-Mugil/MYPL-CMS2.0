# Metayage Portal Backend README

Google Apps Script backend for the Metayage IP operations platform.

This backend is the business-logic layer behind the GitHub Pages frontend. It reads and writes Google Sheets, stores files in Google Drive, manages role-based access, and exposes the web-app API consumed through the Cloudflare Worker.

## Backend Scope

Main backend areas in this folder:

- `00_setup.gs`
  - setup helpers, shared config, base utilities
- `01_ClientManagement.gs`
  - client creation and updates
- `02_CaseManagement.gs`
  - case creation, update, and case-level logic
- `03_InvoiceSystem.gs`
  - invoice generation and finance operations
- `04_Automation.gs`
  - scheduled and automation logic
- `05_ClientPortal.gs`
  - web app entry points and API action routing
- `06_DocManager.gs`
  - Drive/document helpers
- `07_Addons.gs`
  - role access, inbox, threads, approvals, document requests, notifications, expense claims, bulk import, client-portal extensions

## Architecture

1. User signs in through Clerk in the frontend.
2. Frontend calls the Cloudflare Worker.
3. Worker forwards trusted requests to this Apps Script web app.
4. `05_ClientPortal.gs` routes actions to backend functions.
5. Backend reads and writes Google Sheets and Google Drive.

## Data Dependencies

This backend depends on:

- Google Sheets as the main database
- Google Drive for client folders, case folders, uploaded documents, and bill images
- Script Properties for cache/version flags and setup values
- Script Cache for portal caching and session caching

## Important Runtime Rules

- Deploy the Apps Script as a Web App
- Execute as: `Me`
- Who has access: `Anyone`

This backend is designed to be called by the Worker. Opening the GAS `/exec` URL directly in a browser usually returns:

```json
{ "error": "No action specified" }
```

That response is normal and means the web app is reachable.

## Key API Notes

Main entry points:

- `doGet(e)`
- `doPost(e)`
- `handleRequest_(e)`
- `getPortalData(action, params)`

Authentication/session functions:

- `portalLogin`
- `clerkLogin_`
- `validateSession`
- `portalLogout`

Recent important backend behavior includes:

- effective-role handling using primary role plus additional roles
- staff/client portal split based on effective roles
- organization-linked client access fallback
- client-safe threads/messages filtering
- direct document upload
- expense bill upload
- bulk DocketTrak import

## Deployment Steps

### 1. Update source files

When backend changes are made, update the required `.gs` files in the Apps Script project.

Most portal changes typically involve one or more of:

- `05_ClientPortal.gs`
- `07_Addons.gs`
- `02_CaseManagement.gs`
- `01_ClientManagement.gs`
- `03_InvoiceSystem.gs`

### 2. Save the Apps Script project

In Apps Script:

- paste the updated code
- save all changed script files
- confirm there are no syntax errors

### 3. Create a new deployment version

In the Apps Script editor:

1. Click `Deploy`
2. Click `Manage deployments`
3. Edit the web app deployment or create a new version
4. Select the new version
5. Confirm:
   - Execute as: `Me`
   - Who has access: `Anyone`
6. Deploy

### 4. Keep the same Web App URL if possible

Use the existing deployment when possible so the Worker does not need a URL change.

Only update the Worker config if the backend endpoint itself changes.

### 5. Verify the Worker still points to the correct GAS URL

If the Worker uses an environment variable or hardcoded GAS endpoint, verify it still matches the active deployment.

### 6. Hard refresh the frontend after deploy

Because the frontend caches some data and sessions, after backend deployment:

- hard refresh the browser
- or open in incognito for first verification

## Recommended Deployment Order

### Frontend-only changes

If the change is only HTML/CSS/theme/UI and does not require new backend actions:

- deploy frontend only

### Backend-only changes

If the change adds or changes backend actions, access rules, or Sheets/Drive behavior:

- deploy GAS first

### Full-stack changes

If frontend depends on new backend actions:

1. deploy GAS
2. verify Worker target/config
3. deploy frontend
4. test end to end

## Post-Deploy Smoke Test

After each backend deployment, verify:

### General

- GAS `/exec` endpoint returns a JSON response
- Worker can reach GAS
- no `404`, `405`, or permission error

### Login

- staff login works
- client login works
- mixed-role routing works correctly

### Staff Portal

- dashboard loads
- cases load
- documents load
- notifications load
- inbox/threads load

### Client Portal

- org-linked client users can see correct cases
- visible documents load
- client-visible threads load
- notifications load

### Feature-Specific Checks

If touched, verify:

- bulk DocketTrak import
- approvals
- circle soft delete
- organization-linked access
- document upload
- expense bill upload
- role-gated visibility

## Common Problems

### `No action specified`

Meaning:

- endpoint is live
- request reached GAS
- browser opened the URL directly without an API action

This is not a failure.

### Session expired

Usually means:

- stale cached token
- login/session mismatch
- backend deploy happened while an old session was still open

Fix:

- sign out
- sign in again

### Login routes to wrong portal

Check:

- primary role
- additional roles
- effective-role logic in frontend
- client-side vs internal role mapping

### Client employee sees no cases

Check:

- user `ORG_ID`
- client `ORG_ID`
- case `ORG_ID`
- user status is active
- client-side role is correct

### Feature visible but action fails

Usually one of:

- frontend deployed, backend not deployed
- backend deployed, Worker still points to old target
- role allows page visibility but backend access denies action

## Safe Change Guidelines

- do not remove old actions unless frontend is updated everywhere
- prefer additive changes for new portal features
- use soft delete where operational history matters
- keep client portal access strictly filtered for client-visible data only
- test mixed-role users after auth or routing changes

## Notes For This Project

Important backend behaviors currently expected by the frontend:

- any internal effective role opens staff portal
- only pure client-side users open client portal
- organization-linked client users can inherit access through org mapping
- notifications support read and soft delete
- circles use soft delete and vanish from frontend
- threads and inbox are separate concepts
- DocketTrak import creates records under the selected existing client

## Suggested Maintenance Practice

For every backend release:

1. save changed `.gs` files
2. deploy new web app version
3. verify Worker target
4. hard refresh frontend
5. run smoke test for login, cases, documents, and notifications
