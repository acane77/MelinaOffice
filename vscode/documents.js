/* Which open editor an agent may script, and the shape of a VB result.
 * No vscode module: the extension passes a snapshot of the active tab group. */
'use strict';

const guide = require('./vbs-guide');

const NO_DOCUMENT = 'No Melina Office document is open. Open the file in the editor, then ask again.';
const NOT_READY = 'The file is not open in the editor yet.';
const SOURCE_MAX = 64 * 1024;
const ARGS_MAX = 32;
const VBS_TIMEOUT_MS = 60000;

function appFromViewType(viewType) {
  const s = String(viewType || '');
  if (s.indexOf('melinaoffice.') !== 0) return '';
  const app = s.slice('melinaoffice.'.length);
  return app === 'ecrivaine' || app === 'matheuse' || app === 'oratrice' ? app : '';
}

/** A value the agent can read. A VB object comes back as undefined so the caller can say so. */
function plainValue(value, depth) {
  if (value == null) return null;
  if (typeof value === 'string' || typeof value === 'boolean') return value;
  if (typeof value === 'number') return Number.isFinite(value) ? value : null;
  const level = depth || 0;
  if (!Array.isArray(value) || level > 8) return undefined;
  const out = [];
  for (let i = 0; i < value.length; i++) {
    const item = plainValue(value[i], level + 1);
    if (item === undefined) return undefined;
    out.push(item);
  }
  return out;
}

/**
 * tabs: the active editor group's tabs, { active, viewType, uri }.
 * open: documents whose webview is alive, { uri, name, app }.
 * Only the active tab is eligible. Another Melina Office tab is named in the error.
 */
function selectDocument(tabs, open) {
  const list = (Array.isArray(open) ? open : []).filter((d) => d && d.uri);
  const names = list.map((d) => d.name).filter(Boolean);
  if (!list.length) return { ok: false, error: NO_DOCUMENT, open: [] };
  const active = (Array.isArray(tabs) ? tabs : []).find((t) => t && t.active);
  const app = active && appFromViewType(active.viewType);
  if (app && active.uri) {
    const hit = list.find((d) => d.uri === active.uri);
    if (hit) return { ok: true, app: hit.app || app, uri: hit.uri, name: hit.name || '', open: names };
    return { ok: false, error: NOT_READY, open: names };
  }
  return {
    ok: false,
    error: 'The active editor is not a Melina Office document. Click one of these tabs, then ask again: ' + names.join(', ') + '.',
    open: names,
  };
}

function validateRun(body) {
  if (!body || typeof body.source !== 'string' || !String(body.source).trim()) return 'source is required.';
  if (body.source.length > SOURCE_MAX) return 'source is longer than 64 KiB.';
  const mode = body.mode == null || body.mode === '' ? 'module' : body.mode;
  if (mode !== 'module' && mode !== 'line') return 'mode must be module or line.';
  if (mode === 'line' && /[\r\n]/.test(body.source)) return 'mode line takes one statement.';
  if (body.entry != null && body.entry !== '') {
    if (typeof body.entry !== 'string' || body.entry.length > 128) return 'entry must be a procedure name.';
  }
  if (body.args != null) {
    if (!Array.isArray(body.args) || body.args.length > ARGS_MAX) return 'args must be an array of at most 32 values.';
    for (let i = 0; i < body.args.length; i++) {
      const a = body.args[i];
      const t = typeof a;
      if (t !== 'string' && t !== 'number' && t !== 'boolean') return 'args must be strings, numbers, or booleans.';
      if (t === 'number' && !Number.isFinite(a)) return 'args must be finite numbers.';
      if (t === 'string' && a.length > 8192) return 'an argument is too long.';
    }
  }
  return '';
}

function failure(app, fileName, message, extra) {
  const error = Object.assign({ message: message }, extra || {});
  return { ok: false, app: app || '', fileName: fileName || '', prints: [], value: null, error };
}

function normalizeResult(msg, picked) {
  const raw = msg && msg.value;
  const plain = plainValue(raw);
  const dropped = raw != null && plain === undefined;
  const err = msg && msg.error;
  const out = {
    ok: !!(msg && msg.ok),
    app: (msg && msg.app) || (picked && picked.app) || '',
    fileName: (msg && msg.fileName) || (picked && picked.name) || '',
    prints: Array.isArray(msg && msg.prints) ? msg.prints.map((p) => String(p)) : [],
    value: plain === undefined ? null : plain,
    error: err ? {
      message: String(err.message || err),
    } : null,
  };
  if (out.error && err && err.number != null && err.number !== '') out.error.number = err.number;
  if (out.error && err && err.code) out.error.code = String(err.code);
  const note = dropped
    ? 'The result was a VB object. Use Debug.Print to return a string, number, or Boolean.'
    : (msg && msg.note ? String(msg.note) : '');
  if (note) out.note = note;
  return out;
}

/**
 * body.tool is melinaoffice_dom_guide or melinaoffice_run_vbs.
 * run(picked, call) performs the script and returns the page's message.
 */
async function dispatch(body, tabs, open, run) {
  const tool = body && body.tool;
  const picked = selectDocument(tabs, open);
  if (!picked.ok) return failure('', '', picked.error);
  if (tool === 'melinaoffice_dom_guide') {
    return { ok: true, app: picked.app, fileName: picked.name, guide: guide.forApp(picked.app) };
  }
  if (tool !== 'melinaoffice_run_vbs') return failure(picked.app, picked.name, 'Unknown tool.');
  const problem = validateRun(body);
  if (problem) return failure(picked.app, picked.name, problem);
  const call = {
    source: body.source,
    entry: body.entry ? String(body.entry) : '',
    args: Array.isArray(body.args) ? body.args : [],
    mode: body.mode === 'line' ? 'line' : 'module',
  };
  const msg = await run(picked, call);
  return normalizeResult(msg, picked);
}

module.exports = {
  NO_DOCUMENT,
  NOT_READY,
  SOURCE_MAX,
  VBS_TIMEOUT_MS,
  appFromViewType,
  plainValue,
  selectDocument,
  validateRun,
  normalizeResult,
  dispatch,
};
