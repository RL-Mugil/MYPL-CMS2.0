const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const root = path.resolve(__dirname, '..');

function extractInlineScript(file) {
  const html = fs.readFileSync(path.join(root, file), 'utf8');
  const matches = [...html.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/gi)]
    .map((match) => match[1])
    .filter((content) => content.trim());
  if (!matches.length) {
    throw new Error(`No inline script found in ${file}`);
  }
  return matches[matches.length - 1];
}

function loadLoginHelpers(file) {
  const script = extractInlineScript(file);
  const start = script.indexOf('const INTERNAL_ROLES =');
  const end = script.indexOf('(async () => {');
  if (start === -1 || end === -1 || end <= start) {
    throw new Error(`Could not extract role helpers from ${file}`);
  }
  const context = {};
  vm.createContext(context);
  vm.runInContext(
    `${script.slice(start, end)}
     this.getEffectiveSessionRoles = getEffectiveSessionRoles;
     this.hasInternalAccess = hasInternalAccess;
     this.hasClientAccess = hasClientAccess;`,
    context,
    { filename: file }
  );
  return context;
}

function assertRoute(label, helpers, session, expected) {
  const internal = helpers.hasInternalAccess(session);
  const client = helpers.hasClientAccess(session);
  const actual = internal ? 'staff' : client ? 'client' : 'none';
  assert.strictEqual(actual, expected, `${label} expected ${expected}, got ${actual}`);
}

const staffLogin = loadLoginHelpers('index.html');
const clientLogin = loadLoginHelpers('client-login.html');

const cases = [
  ['Super Admin', { role: 'Super Admin', additionalRoles: '' }, 'staff'],
  ['Admin', { role: 'Admin', additionalRoles: '' }, 'staff'],
  ['Galvanizer', { role: 'Galvanizer', additionalRoles: '' }, 'staff'],
  ['Staff', { role: 'Staff', additionalRoles: '' }, 'staff'],
  ['Attorney', { role: 'Attorney', additionalRoles: '' }, 'staff'],
  ['Client Admin', { role: 'Client Admin', additionalRoles: '' }, 'client'],
  ['Client Employee', { role: 'Client Employee', additionalRoles: '' }, 'client'],
  ['Individual Client', { role: 'Individual Client', additionalRoles: '' }, 'client'],
  ['Admin + Galvanizer', { role: 'Admin', additionalRoles: 'Galvanizer' }, 'staff'],
  ['Client Employee + Galvanizer', { role: 'Client Employee', additionalRoles: 'Galvanizer' }, 'staff'],
  ['Client Admin + Staff', { role: 'Client Admin', additionalRoles: 'Staff' }, 'staff'],
  ['Unmapped role', { role: 'Observer', additionalRoles: '' }, 'none']
];

for (const [label, session, expected] of cases) {
  assertRoute(`${label} via staff login`, staffLogin, session, expected);
  assertRoute(`${label} via client login`, clientLogin, session, expected);
}

console.log('Role-routing smoke tests passed.');
