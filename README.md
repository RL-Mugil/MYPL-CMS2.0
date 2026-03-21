# Metayage Portal

Frontend for the Metayage IP operations platform.

This repo is intended to be hosted on GitHub Pages and talk to a Cloudflare Worker, which then forwards trusted requests to Google Apps Script (GAS). The frontend does not talk directly to Sheets or Drive.

## What This Includes

- Staff portal
  - dashboard
  - cases
  - documents
  - workflow board
  - attorney workspace
  - approvals
  - tasks
  - threads
  - notifications
  - management
  - bulk DocketTrak import
- Client portal
  - overview
  - case search and filters
  - document library and upload
  - document request tracking
  - invoice filters
  - notifications
  - client-visible threads
  - smart search
- Shared auth flow with Clerk
- Dark and light theme support

## Architecture

1. User signs in with Clerk in the browser.
2. Frontend exchanges the verified email for an app session through the Worker.
3. Worker validates the request and forwards it to GAS.
4. GAS reads and writes Sheets and Drive data.
5. Frontend renders staff or client UI based on role and additional roles.

## Repo Structure

- `index.html`
  - staff login page
- `dashboard.html`
  - staff portal app
- `client-login.html`
  - client login page
- `client-portal.html`
  - client portal app
- `api.js`
  - Worker/GAS API client
- `auth.js`
  - Clerk bootstrapping and app-session handling
- `config.js`
  - environment-specific frontend config
- `style.css`
  - shared styling for staff and client portals

## Auth and Role Model

The portal supports:

- `Super Admin`
- `Admin`
- `Galvanizer`
- `Staff`
- `Attorney`
- `Client Admin`
- `Client Employee`
- `Individual Client`

Important:

- portal routing uses both `role` and `additionalRoles`
- any user with an internal role is routed to the staff portal
- only pure client-side users are routed to the client portal

## Configuration

Frontend runtime config is in `config.js`.

Current config values are selected by environment:

- `apiBaseByEnv`
- `clerkPublishableKeyByEnv`
- `errorEndpointByEnv`

Before production use, verify:

- Worker URL is correct
- Clerk publishable key is correct
- GitHub Pages domain is allowed in Clerk

## Deployment

### Frontend

Push these files to the GitHub Pages branch/repo:

- `index.html`
- `dashboard.html`
- `client-login.html`
- `client-portal.html`
- `api.js`
- `auth.js`
- `config.js`
- `style.css`

### Backend

This frontend expects a live Worker and GAS deployment.

Relevant GAS files live outside this repo, including:

- `05_ClientPortal.gs`
- `07_Addons.gs`

If frontend changes depend on new actions or access rules, deploy a new GAS version as well.

## Main Features

### Staff Portal

- role-aware navigation
- finance and management access control
- case creation and editing
- bulk case updates
- DocketTrak Excel import
- document upload
- workflow board
- galvanizer queue
- approvals and document workflow
- inbox and threads
- notifications with clear actions

### Client Portal

- organization/client-aware access
- searchable and filterable cases
- searchable and filterable invoices
- searchable documents
- client document upload
- document request visibility
- client-visible thread conversations
- smart search
- notifications with clear actions

## DocketTrak Bulk Import

The staff portal supports bulk import from DocketTrak exports.

Current behavior:

- select an existing client code first
- upload the Excel file
- import both `Patent` and `Trademark` rows
- imported rows are created under the selected client
- duplicate checks are applied during import
- progress is shown during upload/import

## Safety Notes

- frontend sanitizes API payloads before rendering
- client portal only shows client-visible threads/messages
- client portal document access is restricted to accessible clients
- notification deletes are soft-delete operations
- most portal data calls are cached client-side for performance

## Manual Test Checklist

### Login

- staff login opens `dashboard.html`
- client login opens `client-portal.html`
- mixed-role users route correctly based on effective roles

### Staff Portal

- dashboard loads
- cases load
- documents load
- finance visibility matches permission
- management visibility matches permission
- galvanizer queue visibility matches permission

### Client Portal

- overview loads
- cases load for the correct client/org
- documents open
- document upload works
- invoices load when finance access is allowed
- notifications clear correctly
- threads open and reply correctly

## Recommended Next Improvements

- move Worker source into version control
- add staging config values in `config.js`
- add CI checks for HTML inline-script syntax
- add smoke tests for login routing and role-based navigation
- document GAS deployment steps in a separate backend README
