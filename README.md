# Metayage Portal

This frontend is intended to run on GitHub Pages and talk to the Cloudflare Worker, not directly to GAS.

Current public URLs:

- GitHub Pages: `https://rl-mugil.github.io/metayage-portal/`
- Cloudflare Worker: `https://metayage-proxy.mugilvannan.workers.dev`
- GAS Web App: `https://script.google.com/macros/s/AKfycbxBD_oT9Jv_MJkt77n2BJxserwv84iTVOS472DncMyYEVcdA_2g6oPqEj3zIb2nkeSfZA/exec`

## Architecture

1. User signs in with Clerk on the frontend.
2. Frontend calls the Cloudflare Worker.
3. Worker verifies the Clerk identity or token.
4. Worker forwards trusted requests to GAS.
5. GAS reads and writes Sheets/Drive data.

## Local file overview

- `config.js`: environment config for development, staging, production
- `api.js`: frontend API client and runtime error reporting
- `auth.js`: Clerk auth bootstrap and app session handling
- `index.html`: staff login
- `client-login.html`: client login
- `dashboard.html`: staff app
- `client-portal.html`: client app

## What was changed

- Frontend now expects the Worker as the API base.
- Login goes through one shared auth path in `requireAuth()`.
- `PASSWORD_HASH` is no longer returned to the browser from GAS user APIs.
- New users can be created without a legacy local password.
- New admin/client bootstrap users are no longer given default local passwords.
- Critical create flows in GAS use `LockService` to reduce duplicate IDs on concurrent writes.
- Invoice column mapping was corrected so government fees no longer overwrite payment dates.
- Minimal CI was added in `.github/workflows/ci.yml`.
- Runtime error logging hooks were added in `api.js`.

## Manual checks you must do

### 1. Verify the Cloudflare Worker is really checking Clerk

The Worker source is not in this workspace, so this check must be done manually.

Your Worker must do all of the following:

1. Accept frontend requests only from your GitHub Pages domain.
2. Read the Clerk session proof from the request.
3. Verify that proof with Clerk server-side.
4. Extract the verified user email and identity from Clerk.
5. Ignore any raw email sent directly by the browser unless it is matched to verified Clerk identity.
6. Forward only trusted requests to GAS.
7. Add CORS headers only for your allowed frontend origins.

If your Worker is still just accepting `{ email }` from the browser and forwarding it, that still needs fixing.

### 2. Clerk dashboard setup

1. Open Clerk Dashboard.
2. Go to domain or allowed origin settings.
3. Add `https://rl-mugil.github.io`.
4. If you later use a custom domain, add that too.
5. Keep using your `pk_test_...` key only for testing.
6. Before real production use, switch to a live Clerk publishable key.

### 3. GitHub Pages deployment

1. Open your GitHub repo for the portal.
2. Commit and push the updated files.
3. Confirm GitHub Pages is enabled for the repository.
4. Open the live site and test both login pages.

### 4. GAS deployment

1. Open the Apps Script project.
2. Paste or sync the updated GAS files:
   - `00_setup.gs`
   - `01_ClientManagement.gs`
   - `02_CaseManagement.gs`
   - `03_InvoiceSystem.gs`
   - `05_ClientPortal.gs`
3. Create a new deployment version for the GAS web app.
4. Keep the Worker pointing to the latest GAS deployment URL if your Worker stores it as config.

## End-to-end test checklist

### Staff flow

1. Open `https://rl-mugil.github.io/metayage-portal/`
2. Sign in as an Admin or Attorney through Clerk.
3. Confirm redirect to `dashboard.html`.
4. Confirm dashboard, cases, documents, finance load.
5. If Admin, confirm management tab appears.
6. Create or edit a user without entering a password and confirm it still saves.

### Client flow

1. Open `https://rl-mugil.github.io/metayage-portal/client-login.html`
2. Sign in as a Client through Clerk.
3. Confirm redirect to `client-portal.html`.
4. Confirm dashboard, cases, documents, invoices load.
5. Send a contact message and confirm it reaches your admin mailbox.

### Legacy local-password flow

1. Use an old user that already has a local password hash stored.
2. Confirm legacy login still works if you still need it.
3. Confirm a newly created user without a local password cannot use local password login.

## Next recommended steps

1. Move Worker source into version control.
2. Add staging Worker and staging GitHub Pages environment.
3. Replace large `innerHTML` UI building blocks with safer DOM rendering over time.
4. Add a real monitoring sink for `errorEndpoint`.
5. Later, move core multi-tenant data out of Sheets/GAS.
