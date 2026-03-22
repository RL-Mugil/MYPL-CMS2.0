/**
 * auth.js - Clerk authentication integration for the portal.
 *
 * Clerk handles user auth. The verified identity is then exchanged for
 * an app session through the Worker-backed API.
 */

const CLERK_PUBLISHABLE_KEY = (window.APP_CONFIG && window.APP_CONFIG.clerkPublishableKey) || '';
const CACHE_PREFIX_PATTERN = /^mg_api_cache_v\d+:/;

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
    fallbackRedirectUrl: currentUrl,
    signInFallbackRedirectUrl: currentUrl,
    signUpFallbackRedirectUrl: currentUrl,
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
      _gasSession = normalizePortalSession(JSON.parse(stored));
      return _gasSession;
    } catch {
      return null;
    }
  }
  return null;
}

function setGasSession(session) {
  _gasSession = normalizePortalSession(session);
  sessionStorage.setItem('mg_session', JSON.stringify(_gasSession));
}

function clearGasSession() {
  _gasSession = null;
  sessionStorage.removeItem('mg_session');
  try {
    Object.keys(sessionStorage).forEach((key) => {
      if (CACHE_PREFIX_PATTERN.test(key)) {
        sessionStorage.removeItem(key);
      }
    });
  } catch {}
  try {
    Object.keys(localStorage).forEach((key) => {
      if (CACHE_PREFIX_PATTERN.test(key)) {
        localStorage.removeItem(key);
      }
    });
  } catch {}
}

function normalizePortalSession(session) {
  if (!session || typeof session !== 'object') return null;
  return {
    ...session,
    email: session.email || '',
    role: session.role || '',
    additionalRoles: session.additionalRoles || '',
    name: session.name || session.email || '',
    clientId: session.clientId || '',
    orgId: session.orgId || '',
    userId: session.userId || '',
    canViewFinance: session.canViewFinance || '',
    reportsTo: session.reportsTo || '',
    isImpersonating: !!session.isImpersonating,
    originalEmail: session.originalEmail || '',
    originalUserId: session.originalUserId || '',
    originalName: session.originalName || '',
    impersonatedByUserId: session.impersonatedByUserId || '',
    impersonatedByEmail: session.impersonatedByEmail || '',
  };
}

function getEffectiveSessionRoles(session = getGasSession()) {
  const roles = [];
  if (session && session.role) roles.push(session.role);
  String(session && session.additionalRoles || '')
    .split(',')
    .map((value) => value.trim())
    .filter(Boolean)
    .forEach((role) => {
      if (!roles.includes(role)) roles.push(role);
    });
  return roles;
}

function sessionMatchesClerkIdentity(session, clerkEmail) {
  const normalizedClerkEmail = String(clerkEmail || '').trim().toLowerCase();
  if (!session || !normalizedClerkEmail) return false;
  if (String(session.email || '').trim().toLowerCase() === normalizedClerkEmail) return true;
  return !!session.isImpersonating && String(session.originalEmail || '').trim().toLowerCase() === normalizedClerkEmail;
}

function sessionAllowsRoles(session, allowedRoles) {
  if (!allowedRoles || !allowedRoles.length) return true;
  const effectiveRoles = getEffectiveSessionRoles(session);
  return effectiveRoles.some((role) => allowedRoles.includes(role));
}

async function fetchSessionInfo(token) {
  const apiBase = (window.APP_CONFIG && window.APP_CONFIG.apiBase) || '';
  if (!apiBase || !token) return null;
  try {
    const response = await fetch(apiBase, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action: 'getUserInfo', params: { token } }),
      redirect: 'follow',
    });
    if (!response.ok) return null;
    const data = await response.json();
    if (!data || data.error || data.sessionExpired) return null;
    return data;
  } catch {
    return null;
  }
}

async function requireAuth(allowedRoles) {
  const Clerk = await initClerk();

  const existing = getGasSession();
  if (existing && existing.token) {
    const live = await fetchSessionInfo(existing.token);
    if (live) {
      const hydrated = normalizePortalSession({ ...existing, ...live, token: existing.token });
      setGasSession(hydrated);
      if (sessionAllowsRoles(hydrated, allowedRoles)) {
        return hydrated;
      }
    } else {
      clearGasSession();
    }
  }

  if (!Clerk.user) return null;

  const email = getClerkEmail();
  if (!email) return null;

  const current = getGasSession();
  if (current && sessionMatchesClerkIdentity(current, email) && sessionAllowsRoles(current, allowedRoles)) {
    return current;
  }

  try {
    const result = await window.API.clerkLogin(email);
    if (!result.success) return null;

    const session = normalizePortalSession({ ...result, email: result.email || email });
    setGasSession(session);

    if (!sessionAllowsRoles(session, allowedRoles)) {
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
  getEffectiveSessionRoles,
};
