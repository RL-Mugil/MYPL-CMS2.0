/**
 * auth.js - Clerk authentication integration for the portal.
 *
 * Clerk handles user auth. The verified identity is then exchanged for
 * an app session through the Worker-backed API.
 */

const CLERK_PUBLISHABLE_KEY = (window.APP_CONFIG && window.APP_CONFIG.clerkPublishableKey) || '';

let _clerkInstance = null;
let _gasSession = null;
let _lastAuthError = '';

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
  _lastAuthError = '';
  sessionStorage.removeItem('mg_session');
  try {
    Object.keys(sessionStorage).forEach((key) => {
      if (key.startsWith('mg_api_cache_v1:')) {
        sessionStorage.removeItem(key);
      }
    });
  } catch {}
  try {
    Object.keys(localStorage).forEach((key) => {
      if (key.startsWith('mg_api_cache_v1:')) {
        localStorage.removeItem(key);
      }
    });
  } catch {}
}

function getEffectiveSessionRoles(session) {
  const roles = [];
  if (session?.role) roles.push(session.role);
  String(session?.additionalRoles || '')
    .split(',')
    .map((value) => value.trim())
    .filter(Boolean)
    .forEach((role) => {
      if (!roles.includes(role)) roles.push(role);
    });
  return roles;
}

function sessionMatchesAllowedRoles(session, allowedRoles) {
  if (!allowedRoles || !allowedRoles.length) return true;
  const effectiveRoles = getEffectiveSessionRoles(session);
  return effectiveRoles.some((role) => allowedRoles.includes(role));
}

async function requireAuth(allowedRoles) {
  const Clerk = await initClerk();

  if (!Clerk.user) return null;

  const email = getClerkEmail();
  if (!email) return null;

  const existing = getGasSession();
  const matchesRealUser = existing && existing.email === email;
  const matchesImpersonationOwner = existing
    && existing.impersonatedByUserId
    && String(existing.originalEmail || '').toLowerCase() === String(email || '').toLowerCase();
  if (existing && (matchesRealUser || matchesImpersonationOwner)) {
    if (!existing.userId && window.API && typeof window.API.getUserInfo === 'function') {
      try {
        const info = await window.API.getUserInfo();
        const hydrated = { ...existing, ...info };
        setGasSession(hydrated);
        if (!sessionMatchesAllowedRoles(hydrated, allowedRoles)) {
          return null;
        }
        return hydrated;
      } catch {}
    }
    if (!sessionMatchesAllowedRoles(existing, allowedRoles)) {
      return null;
    }
    return existing;
  }

  try {
    const result = await window.API.clerkLogin(email);
    if (!result.success) {
      _lastAuthError = result.message || 'App session exchange failed.';
      return null;
    }

    const session = { ...result, email };
    setGasSession(session);
    _lastAuthError = '';

    if (!sessionMatchesAllowedRoles(session, allowedRoles)) {
      _lastAuthError = 'This account is authenticated but has no portal role mapping.';
      return null;
    }
    return session;
  } catch (e) {
    console.error('App session exchange failed:', e);
    _lastAuthError = e && e.message ? e.message : 'App session exchange failed.';
    return null;
  }
}

function getLastAuthError() {
  return _lastAuthError || '';
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
  if (window.StreamMessaging && typeof window.StreamMessaging.disconnect === 'function') {
    try { await window.StreamMessaging.disconnect(); } catch {}
  }
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
  getEffectiveSessionRoles,
  getGasSession,
  setGasSession,
  clearGasSession,
  getLastAuthError,
};
