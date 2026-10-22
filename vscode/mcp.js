/* MCP server for the editor's agent. stdio is newline-delimited JSON-RPC.
 * Each tool call is forwarded to the extension host, which owns the open document.
 * Started by the extension with ELECTRON_RUN_AS_NODE=1. Docs: docs/vscode.md. */
'use strict';

const http = require('http');
const guide = require('./vbs-guide');
const docs = require('./documents');

const TOOLS = [
  {
    name: 'vibeoffice_dom_guide',
    description: 'Read the Visual Basic object model for the VibeOffice document in the active editor. Call this before vibeoffice_run_vbs. The text is only for that app (Quire, Ledger, or Lectern) and names the syntax and the gaps versus Microsoft VBA.',
    inputSchema: { type: 'object', properties: {}, additionalProperties: false },
    annotations: { title: 'VibeOffice DOM guide', readOnlyHint: true },
  },
  {
    name: 'vibeoffice_run_vbs',
    description: 'Run VB6 on the VibeOffice document in the active editor tab. mode "module" (default) compiles source as a standard module and runs Public Sub Main, or the public procedure named by entry. mode "line" runs one Immediate statement; start a read with ? or Debug.Print. The script can read and edit that document. It cannot open another file.',
    inputSchema: {
      type: 'object',
      additionalProperties: false,
      required: ['source'],
      properties: {
        source: { type: 'string', description: 'VB6 source. A module, or one statement when mode is line. At most 64 KiB.' },
        entry: { type: 'string', description: 'Public procedure on the module. Omit it to run Public Sub Main.' },
        args: {
          type: 'array',
          description: 'Arguments for entry. Strings, numbers, and booleans.',
          items: {},
        },
        mode: { type: 'string', enum: ['module', 'line'], description: 'module runs a standard module. line runs one Immediate statement.' },
      },
    },
    annotations: { title: 'Run VB6 on the open document', readOnlyHint: false, destructiveHint: true },
  },
];

function toolResult(payload) {
  return {
    content: [{ type: 'text', text: JSON.stringify(payload, null, 2) }],
    isError: !payload || payload.ok === false,
  };
}

function createSession(call) {
  return {
    async request(msg) {
      if (!msg || msg.jsonrpc !== '2.0' || typeof msg.method !== 'string') {
        if (msg && msg.id != null) return { jsonrpc: '2.0', id: msg.id, error: { code: -32600, message: 'Invalid request' } };
        return null;
      }
      const id = msg.id;
      const reply = (result) => (id == null ? null : { jsonrpc: '2.0', id, result });
      const fail = (code, message) => (id == null ? null : { jsonrpc: '2.0', id, error: { code, message } });
      if (msg.method === 'initialize') {
        const asked = msg.params && msg.params.protocolVersion;
        const protocolVersion = /^20\d{2}-\d{2}-\d{2}$/.test(asked || '') ? asked : '2024-11-05';
        return reply({
          protocolVersion,
          capabilities: { tools: { listChanged: false } },
          serverInfo: { name: 'vibeoffice', version: '1.0.3' },
          instructions: guide.INSTRUCTIONS,
        });
      }
      if (msg.method === 'notifications/initialized' || msg.method === 'initialized') return null;
      if (msg.method === 'ping') return reply({});
      if (msg.method === 'tools/list') return reply({ tools: TOOLS });
      if (msg.method === 'tools/call') {
        const params = msg.params || {};
        const name = params.name;
        const args = params.arguments || {};
        if (name !== 'vibeoffice_dom_guide' && name !== 'vibeoffice_run_vbs') {
          return reply(toolResult({ ok: false, error: { message: 'Unknown tool.' } }));
        }
        if (name === 'vibeoffice_run_vbs') {
          const problem = docs.validateRun(args);
          if (problem) return reply(toolResult({ ok: false, error: { message: problem } }));
        }
        try {
          const payload = await call({
            tool: name,
            source: args.source,
            entry: args.entry,
            args: args.args,
            mode: args.mode,
          });
          return reply(toolResult(payload));
        } catch (e) {
          return reply(toolResult({ ok: false, error: { message: e && e.message ? e.message : String(e) } }));
        }
      }
      return fail(-32601, 'Method not found');
    },
  };
}

function postJson(url, token, body) {
  return new Promise((resolve, reject) => {
    let u;
    try { u = new URL(url); }
    catch (e) { reject(e); return; }
    const data = Buffer.from(JSON.stringify(body));
    const req = http.request({
      hostname: u.hostname,
      port: u.port,
      path: u.pathname,
      method: 'POST',
      headers: {
        'content-type': 'application/json',
        'content-length': data.length,
        'x-vibeoffice-token': token,
      },
    }, (res) => {
      const chunks = [];
      res.on('data', (chunk) => chunks.push(chunk));
      res.on('end', () => {
        const text = Buffer.concat(chunks).toString('utf8');
        if (res.statusCode !== 200) {
          reject(new Error(text || ('The editor bridge returned ' + res.statusCode)));
          return;
        }
        try { resolve(JSON.parse(text)); }
        catch (e) { reject(e); }
      });
    });
    req.on('error', reject);
    req.end(data);
  });
}

function bridgeCall(body) {
  const url = process.env.VIBEOFFICE_BRIDGE_URL;
  const token = process.env.VIBEOFFICE_BRIDGE_TOKEN;
  if (!url || !token) {
    return Promise.resolve({
      ok: false,
      error: { message: 'The VibeOffice editor is not connected.' },
    });
  }
  return postJson(url, token, body);
}

function startStdio(call) {
  const session = createSession(call || bridgeCall);
  let buf = '';
  process.stdin.setEncoding('utf8');
  process.stdin.on('data', (chunk) => {
    buf += chunk;
    let nl;
    while ((nl = buf.indexOf('\n')) >= 0) {
      const line = buf.slice(0, nl).replace(/\r$/, '');
      buf = buf.slice(nl + 1);
      if (!line.trim()) continue;
      let msg;
      try { msg = JSON.parse(line); }
      catch (e) { continue; }
      session.request(msg).then((res) => {
        if (res) process.stdout.write(JSON.stringify(res) + '\n');
      }, (e) => {
        if (msg.id != null) {
          process.stdout.write(JSON.stringify({
            jsonrpc: '2.0',
            id: msg.id,
            error: { code: -32603, message: e && e.message ? e.message : String(e) },
          }) + '\n');
        }
      });
    }
  });
}

if (require.main === module) startStdio(bridgeCall);

module.exports = { TOOLS, createSession, postJson, bridgeCall, startStdio };
