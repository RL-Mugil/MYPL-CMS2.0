/**
 * ============================================================
 * gasUpdates.gs — Changes needed in 05_ClientPortal.gs
 * to support the new GitHub Pages + Clerk frontend
 * ============================================================
 *
 * WHAT TO DO:
 *  Open your GAS project → 05_ClientPortal.gs
 *  Make the two changes described below.
 *
 * CHANGE 1 — Add clerkLogin_ function (new function)
 * ─────────────────────────────────────────────────
 * Add this anywhere in 05_ClientPortal.gs (e.g. after portalLogout):
 */

/**
 * Called by the Clerk-authenticated frontend.
 * Clerk has already verified the user's email/password — we just
 * look up the user in our USER_ROLES sheet and create a GAS session token.
 */
function clerkLogin_(email) {
  if (!email) return { success: false, message: 'Email required.' };

  var sheet = getSheet_('USERS');
  var data  = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    // columns: USER_ID[0], EMAIL[1], FULL_NAME[2], ROLE[3], CLIENT_ID[4],
    //          DEPARTMENT[5], STATUS[6], PASSWORD_HASH[7], CREATED_DATE[8]
    if (data[i][1] === email && data[i][6] === 'Active') {
      var token = Utilities.getUuid();
      var cache = CacheService.getScriptCache();
      cache.put('session_' + token, JSON.stringify({
        email:    email,
        role:     data[i][3],
        clientId: data[i][4],
        name:     data[i][2],
        userId:   data[i][0]
      }), 21600); // 6-hour session

      logActivity_('CLERK_LOGIN', 'USER', data[i][0], 'Login via Clerk/GitHub Pages');

      return {
        success: true,
        token:   token,
        role:    data[i][3],
        name:    data[i][2]
      };
    }
  }

  return { success: false, message: 'Email not found in system or account inactive.' };
}


/**
 * CHANGE 2 — Add clerkLogin case to getPortalData()
 * ──────────────────────────────────────────────────
 * Inside the existing getPortalData(action, params) function,
 * find the block that starts with:
 *
 *   if (action === "login") {
 *     return portalLogin(params.email, params.password);
 *   }
 *
 * Add these TWO lines immediately AFTER that block:
 *
 *   if (action === "clerkLogin") {
 *     return clerkLogin_(params.email);
 *   }
 *
 * So it looks like:
 *
 *   if (action === "login") {
 *     return portalLogin(params.email, params.password);
 *   }
 *   if (action === "clerkLogin") {            // ← ADD THIS
 *     return clerkLogin_(params.email);        // ← ADD THIS
 *   }                                          // ← ADD THIS
 *   if (action === "logout") {
 *     return portalLogout(params.token);
 *   }
 *
 * That's it! No other changes needed in GAS.
 */


/**
 * CHANGE 3 — Update doGet() to handle CORS for GitHub Pages
 * ──────────────────────────────────────────────────────────
 * Replace your existing doGet() function with this version,
 * which handles both GET and POST and adds CORS headers:
 */
function doGet(e) {
  return handleRequest_(e);
}

function doPost(e) {
  return handleRequest_(e);
}

function handleRequest_(e) {
  var output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  // Handle CORS preflight — GitHub Pages needs this
  var result;
  try {
    var payload;

    // POST with JSON body (how our api.js sends requests)
    if (e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    }
    // GET with URL params (fallback)
    else if (e.parameter && e.parameter.action) {
      payload = {
        action: e.parameter.action,
        params: e.parameter.params ? JSON.parse(e.parameter.params) : {}
      };
    }
    else {
      result = { error: 'No action specified' };
      output.setContent(JSON.stringify(result));
      return output;
    }

    result = getPortalData(payload.action, payload.params || {});
  } catch (err) {
    result = { error: 'Request error: ' + err.message };
  }

  output.setContent(JSON.stringify(result));
  return output;
}


/**
 * ─────────────────────────────────────────────────────────────────────
 * IMPORTANT NOTE ABOUT CORS
 * ─────────────────────────────────────────────────────────────────────
 *
 * Google Apps Script Web Apps do NOT support true CORS preflight responses.
 * The api.js file uses Content-Type: 'text/plain' to avoid triggering
 * a CORS preflight, which works for simple cross-origin POST requests.
 *
 * If you get CORS errors in the browser console, you have two options:
 *
 * Option A — Use a Cloudflare Worker proxy (see Cloudflare setup):
 *   Point portal.yourfirm.com → Cloudflare Worker → GAS URL
 *   The Worker adds CORS headers to all responses.
 *
 * Option B — Use JSONP-style GET requests (simpler fallback):
 *   Change api.js to use fetch with mode: 'no-cors' and URL params.
 *   Downside: you can't read the response with no-cors mode.
 *   Better to use the Cloudflare Worker approach.
 *
 * ─────────────────────────────────────────────────────────────────────
 * DEPLOYMENT CHECKLIST
 * ─────────────────────────────────────────────────────────────────────
 *
 * GAS side:
 *  [x] Add clerkLogin_ function
 *  [x] Add clerkLogin case to getPortalData()
 *  [x] Replace doGet with doGet + doPost + handleRequest_
 *  [ ] Re-deploy as Web App (New Deployment or New Version)
 *       Execute as: Me
 *       Who has access: Anyone
 *  [ ] Copy the new Web App URL
 *
 * Frontend side:
 *  [ ] Set CLERK_PUBLISHABLE_KEY in auth.js
 *  [ ] Set GAS_URL in api.js
 *  [ ] Push all files to GitHub repository
 *  [ ] Enable GitHub Pages (Settings → Pages → main branch / root)
 *  [ ] (Optional) Set custom domain in GitHub Pages settings
 *  [ ] Add your custom domain to Clerk → Settings → Domains
 *
 * Testing:
 *  [ ] Open index.html → sign in as staff → should land on dashboard.html
 *  [ ] Open client-login.html → sign in as client → should land on client-portal.html
 *  [ ] Verify Dashboard charts load correctly
 *  [ ] Verify Cases page loads
 *  [ ] Verify Management tab visible for Admin only
 */
