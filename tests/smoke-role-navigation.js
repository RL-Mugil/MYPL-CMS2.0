const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const root = path.resolve(__dirname, '..');
const html = fs.readFileSync(path.join(root, 'dashboard.html'), 'utf8');
const scriptMatches = [...html.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/gi)]
  .map((match) => match[1])
  .filter((content) => content.trim());

if (!scriptMatches.length) {
  throw new Error('No inline script found in dashboard.html');
}

const inlineScript = scriptMatches[scriptMatches.length - 1];
const start = inlineScript.indexOf('function hasRoleAtLeast(role, minimumRole) {');
const end = inlineScript.indexOf('(async () => {');

if (start === -1 || end === -1 || end <= start) {
  throw new Error('Could not extract dashboard role helpers.');
}

const context = {
  SESSION: null
};
vm.createContext(context);
vm.runInContext(
  `${inlineScript.slice(start, end)}
   this.getEffectiveSessionRoles = getEffectiveSessionRoles;
   this.sessionHasRoleAtLeast = sessionHasRoleAtLeast;
   this.sessionHasAnyRole = sessionHasAnyRole;
   this.canAccessFinance = canAccessFinance;
   this.canAccessManagement = canAccessManagement;
   this.isInternalUser = isInternalUser;`,
  context,
  { filename: 'dashboard.html' }
);

function evaluate(session) {
  context.SESSION = session;
  return {
    internal: context.isInternalUser(),
    management: context.canAccessManagement(),
    finance: context.canAccessFinance(),
    staffPlus: context.sessionHasRoleAtLeast('Staff'),
    adminPlus: context.sessionHasRoleAtLeast('Admin'),
    galvanizerAny: context.sessionHasAnyRole(['Super Admin', 'Admin', 'Galvanizer'])
  };
}

function expect(label, session, expected) {
  const actual = evaluate(session);
  for (const [key, value] of Object.entries(expected)) {
    assert.strictEqual(actual[key], value, `${label} ${key} expected ${value}, got ${actual[key]}`);
  }
}

expect('Super Admin', { role: 'Super Admin', additionalRoles: '', canViewFinance: '' }, {
  internal: true,
  management: true,
  finance: true,
  staffPlus: true,
  adminPlus: true,
  galvanizerAny: true
});

expect('Galvanizer', { role: 'Galvanizer', additionalRoles: '', canViewFinance: '' }, {
  internal: true,
  management: false,
  finance: false,
  staffPlus: true,
  adminPlus: false,
  galvanizerAny: true
});

expect('Attorney', { role: 'Attorney', additionalRoles: '', canViewFinance: '' }, {
  internal: true,
  management: false,
  finance: false,
  staffPlus: false,
  adminPlus: false,
  galvanizerAny: false
});

expect('Client Admin', { role: 'Client Admin', additionalRoles: '', canViewFinance: '' }, {
  internal: false,
  management: true,
  finance: false,
  staffPlus: false,
  adminPlus: false,
  galvanizerAny: false
});

expect('Client Employee + Staff', { role: 'Client Employee', additionalRoles: 'Staff', canViewFinance: '' }, {
  internal: true,
  management: false,
  finance: false,
  staffPlus: true,
  adminPlus: false,
  galvanizerAny: false
});

expect('Staff with finance flag', { role: 'Staff', additionalRoles: '', canViewFinance: 'Yes' }, {
  internal: true,
  management: false,
  finance: true,
  staffPlus: true,
  adminPlus: false,
  galvanizerAny: false
});

console.log('Role-navigation smoke tests passed.');
