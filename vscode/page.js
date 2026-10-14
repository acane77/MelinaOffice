/* The app page a custom editor shows. The HTML on disk uses relative links (../common/…,
 * js/…). A webview document has no folder of its own, so a <base> points those links at
 * the app's folder as a webview resource. The host script runs before suite.js. */
'use strict';

function attr(s) {
  return String(s).replace(/&/g, '&amp;').replace(/"/g, '&quot;');
}

/** CSP plus the vscode host, inserted after the charset so the existing look script and
 *  stylesheets still run, now resolved against the app folder. */
function renderPage(opts) {
  const base = String(opts.baseHref || '').endsWith('/') ? String(opts.baseHref) : String(opts.baseHref || '') + '/';
  const csp = [
    "default-src 'none'",
    `img-src ${opts.cspSource} data: blob:`,
    `media-src ${opts.cspSource} data: blob:`,
    `font-src ${opts.cspSource} data:`,
    `style-src ${opts.cspSource} 'unsafe-inline'`,
    `script-src ${opts.cspSource} 'unsafe-inline' 'unsafe-eval'`,
    `connect-src ${opts.cspSource}`,
  ].join('; ');
  const host = JSON.stringify({
    kind: 'vscode',
    platform: opts.platform || 'linux',
    version: 1,
    fileName: opts.fileName || '',
  }).replace(/</g, '\\u003c');
  const inject = `<meta http-equiv="Content-Security-Policy" content="${attr(csp)}">` +
    `<base href="${attr(base)}">` +
    `<script>window.__VO_HOST__=${host};window.__VO_VSCODE__=acquireVsCodeApi();</script>`;
  let html = String(opts.html || '').replace(/<link\s+rel=["']manifest["'][^>]*>\s*/gi, '');
  if (/<meta\s+charset[^>]*>/i.test(html)) html = html.replace(/<meta\s+charset[^>]*>/i, (m) => m + inject);
  else html = html.replace(/<head[^>]*>/i, (m) => m + inject);
  return html;
}

function missingPage(message) {
  const text = String(message || 'VibeOffice could not find its pages.');
  return '<!doctype html><html><head><meta charset="utf-8"></head><body><p>' +
    text.replace(/&/g, '&amp;').replace(/</g, '&lt;') + '</p></body></html>';
}

module.exports = { renderPage, missingPage };
