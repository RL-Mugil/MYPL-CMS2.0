const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const htmlFiles = [
  'index.html',
  'client-login.html',
  'client-portal.html',
  'dashboard.html'
];

function extractScripts(html) {
  const matches = [...html.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/gi)];
  return matches.map((match) => match[1]).filter((content) => content.trim());
}

let hadError = false;

for (const file of htmlFiles) {
  const abs = path.join(root, file);
  const html = fs.readFileSync(abs, 'utf8');
  const scripts = extractScripts(html);

  if (!scripts.length) {
    console.log(`[warn] No inline scripts found in ${file}`);
    continue;
  }

  scripts.forEach((content, index) => {
    try {
      new vm.Script(content, { filename: `${file}#inline-${index + 1}` });
      console.log(`[ok] ${file} inline script ${index + 1}`);
    } catch (error) {
      hadError = true;
      console.error(`[fail] ${file} inline script ${index + 1}`);
      console.error(error.message);
    }
  });
}

if (hadError) {
  process.exit(1);
}
