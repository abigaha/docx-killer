/* Uses the installed Node version of Mammoth; no Python mammoth dependency. */
'use strict';
const fs = require('node:fs');
const path = require('node:path');

function loadMammoth() {
  const candidates = process.env.MAMMOTH_MODULE ? [process.env.MAMMOTH_MODULE] : [
    'mammoth',
    ...(process.env.APPDATA ? [path.join(process.env.APPDATA, 'npm', 'node_modules', 'mammoth')] : []),
    path.join(path.dirname(process.execPath), 'node_modules', 'mammoth'),
  ];
  for (const candidate of candidates) {
    let resolved;
    try { resolved = require.resolve(candidate); }
    catch (error) { if (error.code === 'MODULE_NOT_FOUND') continue; throw error; }
    let packageDirectory = path.dirname(resolved);
    while (!fs.existsSync(path.join(packageDirectory, 'package.json'))) {
      const parent = path.dirname(packageDirectory);
      if (parent === packageDirectory) throw new Error('Cannot locate Mammoth package metadata');
      packageDirectory = parent;
    }
    const metadata = JSON.parse(fs.readFileSync(path.join(packageDirectory, 'package.json'), 'utf8'));
    return {mammoth: require(resolved), version: metadata.version, path: resolved};
  }
  throw new Error('Mammoth not found. Run npm root -g and set MAMMOTH_MODULE to the mammoth package directory.');
}

async function main(args) {
  const loaded = loadMammoth();
  if (args.length === 1 && args[0] === '--doctor') {
    console.log(JSON.stringify({version: loaded.version, path: loaded.path}));
    return;
  }
  const [input, output, mode] = args;
  if (!input || !output || args.length > 3 || (mode && mode !== '--text')) {
    throw new Error('Usage: node mammoth_extract.cjs input.docx output.html [--text]');
  }
  if (path.resolve(input) === path.resolve(output) || fs.existsSync(output)) {
    throw new Error('Use a new output path; existing files are never overwritten');
  }
  const options = {externalFileAccess: false, includeEmbeddedStyleMap: false};
  const result = mode === '--text'
    ? await loaded.mammoth.extractRawText({path: input}, options)
    : await loaded.mammoth.convertToHtml({path: input}, options);
  // CSP is defense in depth, not a sanitizer. Never embed raw output in a trusted app.
  const value = mode === '--text' ? result.value : '<!doctype html>\n<html><head><meta charset="utf-8">' +
    '<meta http-equiv="Content-Security-Policy" content="default-src \'none\'; img-src data:; style-src \'none\'; base-uri \'none\'; form-action \'none\'">' +
    '<title>DOCX semantic preview</title></head><body>' + result.value + '</body></html>\n';
  fs.writeFileSync(output, value, {encoding: 'utf8', flag: 'wx'});
  console.log(JSON.stringify({output: path.resolve(output), mode: mode === '--text' ? 'text' : 'semantic-html', messages: result.messages, sanitized: false, layoutFidelity: false}));
}

main(process.argv.slice(2)).catch(error => {
  console.error(JSON.stringify({error: error.message}));
  process.exitCode = 1;
});
