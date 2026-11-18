'use strict';

/**
 * Opens the home page of the SharePoint site in the user's default browser (own profile, already
 * signed in) with the debug parameters of the extension:
 *
 *   <siteUrl>/?loadSPFX=true&debugManifestsFile=https://localhost:4321/temp/manifests.js&customActions={...}
 *
 * Component ID and properties come from the manifest and sharepoint/assets/elements.xml, so the
 * debug session uses the same configuration as the package; only the log level is set to "verbose".
 *
 * Usage: node tools/open-debug-page.js <siteUrl>   (started by .vscode/launch.json, "gulp serve" must run)
 */
const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const MANIFEST = path.join(ROOT, 'src/extensions/feedback/FeedbackApplicationCustomizer.manifest.json');
const ELEMENTS = path.join(ROOT, 'sharepoint/assets/elements.xml');
const MANIFESTS_FILE = 'https://localhost:4321/temp/manifests.js';

function fail(message) {
  console.error(message);
  process.exit(1);
}

function readComponentId() {
  return JSON.parse(fs.readFileSync(MANIFEST, 'utf8')).id;
}

/** ClientSideComponentProperties of the custom action (XML-encoded JSON). */
function readProperties() {
  const xml = fs.readFileSync(ELEMENTS, 'utf8');
  const match = /ClientSideComponentProperties="([^"]*)"/.exec(xml);
  if (!match) {
    return {};
  }
  const json = match[1]
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&');
  return JSON.parse(json);
}

/** Default browser of the operating system. spawn without shell: "&" and "%" in the URL stay untouched. */
function openInDefaultBrowser(url) {
  const [command, args] = process.platform === 'win32'
    ? ['rundll32', ['url.dll,FileProtocolHandler', url]]
    : [process.platform === 'darwin' ? 'open' : 'xdg-open', [url]];
  const child = spawn(command, args, { detached: true, stdio: 'ignore' });
  child.on('error', (error) => fail('Browser could not be started: ' + error.message));
  child.unref();
}

const siteUrl = (process.argv[2] || '').trim().replace(/\/+$/, '');
if (!/^https:\/\/[^/]+/i.test(siteUrl) || siteUrl.indexOf('${') >= 0) {
  fail('Site URL missing or invalid. Set "spfx.siteUrl" in .vscode/settings.json (see settings.example.json).');
}

const customActions = {
  [readComponentId()]: {
    location: 'ClientSideExtension.ApplicationCustomizer',
    properties: Object.assign(readProperties(), { logLevel: 'verbose' })
  }
};

// The site URL shows the home page (welcome page) of the site.
const url = siteUrl + '/?loadSPFX=true'
  + '&debugManifestsFile=' + encodeURIComponent(MANIFESTS_FILE)
  + '&customActions=' + encodeURIComponent(JSON.stringify(customActions));

console.log('Opening ' + url);
openInDefaultBrowser(url);
