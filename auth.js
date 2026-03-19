/**
 * auth.js - Clerk authentication integration for the portal.
 *
 * Clerk handles user auth. The verified identity is then exchanged for
 * an app session through the Worker-backed API.
 */

const CLERK_PUBLISHABLE_KEY = (window.APP_CONFIG && window.APP_CONFIG.clerkPublishableKey) || '';

let _clerkInstance = null;
let _gasSession = null;

function getCurrentPageUrl() {
  const url = new URL(window.location.href);
  url.search = '';
  url.hash = '';
  return url.toString();
}

async function loadClerk() {
  return new Promise((resolve, reject) => {
    if (window.Clerk) {
      resolve(window.Clerk);
      return;
    }
    if (!CLERK_PUBLISHABLE_KEY) {
      reject(new Error('Clerk publishable key is not configured.'));
      return;
    }
    const script = document.createElement('script');
    script.src = 'https://cdn.jsdelivr.net/npm/@clerk/clerk-js@5/dist/clerk.browser.js';
    script.setAttribute('data-clerk-publishable-key', CLERK_PUBLISHABLE_KEY);
    script.onload = () => resolve(window.Clerk);
    script.onerror = reject;
    document.head.appendChild(script);
  });
}

async function initClerk() {
  const Clerk = await loadClerk();
  const currentUrl = getCurrentPageUrl();
  await Clerk.load({
    afterSignInUrl: currentUrl,
    afterSignUpUrl: currentUrl,
  });
  _clerkInstance = Clerk;
  return Clerk;
}

function getClerkUser() {
  return _clerkInstance?.user || null;
}

function getClerkEmail() {
  const user = getClerkUser();
  return user?.primaryEmailAddress?.emailAddress || null;
}

function getGasSession() {
  if (_gasSession) return _gasSession;
  const stored = sessionStorage.getItem('mg_session');
  if (stored) {
    try {
      _gasSession = JSON.parse(stored);
      return _gasSession;
    } catch {
      return null;
    }
  }
  return null;
}

function setGasSession(session) {
  _gasSession = session;
  sessionStorage.setItem('mg_session', JSON.stringify(session));
}

function clearGasSession() {
  _gasSession = null;
  sessionStorage.removeItem('mg_session');
  try {
    Object.keys(sessionStorage).forEach((key) => {
      if (key.startsWith('mg_api_cache_v1:')) {
        sessionStorage.removeItem(key);
      }
    });
  } catch {}
}

async function requireAuth(allowedRoles) {
  const Clerk = await initClerk();

  if (!Clerk.user) return null;

  const email = getClerkEmail();
  if (!email) return null;

  const existing = getGasSession();
  if (existing && existing.email === email) {
    if (allowedRoles && !allowedRoles.includes(existing.role)) {
      return null;
    }
    return existing;
  }

  try {
    const result = await window.API.clerkLogin(email);
    if (!result.success) return null;

    const session = { ...result, email };
    setGasSession(session);

    if (allowedRoles && !allowedRoles.includes(session.role)) {
      return null;
    }
    return session;
  } catch (e) {
    console.error('App session exchange failed:', e);
    return null;
  }
}

async function mountSignIn(containerId, options = {}) {
  const Clerk = await initClerk();
  if (Clerk.user) return true;
  const currentUrl = getCurrentPageUrl();

  Clerk.mountSignIn(document.getElementById(containerId), {
    fallbackRedirectUrl: currentUrl,
    signInFallbackRedirectUrl: currentUrl,
    signUpFallbackRedirectUrl: currentUrl,
    appearance: {
      variables: {
        colorPrimary: '#6366f1',
        colorBackground: '#13131e',
        colorText: '#e2e2ef',
        colorTextSecondary: '#9191b4',
        colorInputBackground: '#08080f',
        colorInputText: '#e2e2ef',
        borderRadius: '10px',
        fontFamily: '"DM Sans", sans-serif',
      },
      elements: {
        card: { boxShadow: 'none', background: 'transparent' },
        headerTitle: { fontFamily: '"Syne", sans-serif', fontWeight: '700' },
        formButtonPrimary: {
          background: '#6366f1',
          '&:hover': { background: '#818cf8' },
        },
      },
    },
    ...options,
  });
  return false;
}

async function signOut() {
  clearGasSession();
  if (_clerkInstance) {
    await _clerkInstance.signOut();
  }
}

window.Auth = {
  initClerk,
  requireAuth,
  mountSignIn,
  signOut,
  getClerkUser,
  getClerkEmail,
  getGasSession,
  setGasSession,
  clearGasSession,
};
