/* A localhost endpoint the MCP process calls. The token is the only credential.
 * The listener is 127.0.0.1, so another machine cannot reach the open document. */
'use strict';

const http = require('http');
const crypto = require('crypto');

const BODY_MAX = 256 * 1024;

function createBridge(handler) {
  const token = crypto.randomBytes(24).toString('hex');
  const server = http.createServer((req, res) => {
    if (req.method !== 'POST' || req.url !== '/call') {
      res.writeHead(404);
      res.end();
      return;
    }
    if (req.headers['x-melinaoffice-token'] !== token) {
      res.writeHead(403);
      res.end();
      return;
    }
    const chunks = [];
    let size = 0;
    let tooBig = false;
    req.on('data', (chunk) => {
      size += chunk.length;
      if (size > BODY_MAX) {
        tooBig = true;
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });
    req.on('error', () => {});
    req.on('end', () => {
      if (tooBig) return;
      let body;
      try { body = JSON.parse(Buffer.concat(chunks).toString('utf8') || '{}'); }
      catch (e) {
        res.writeHead(400, { 'content-type': 'application/json' });
        res.end(JSON.stringify({ ok: false, error: { message: 'The request was not JSON.' } }));
        return;
      }
      Promise.resolve().then(() => handler(body)).then((result) => {
        res.writeHead(200, { 'content-type': 'application/json' });
        res.end(JSON.stringify(result == null ? { ok: false, error: { message: 'No result.' } } : result));
      }, (e) => {
        res.writeHead(200, { 'content-type': 'application/json' });
        res.end(JSON.stringify({ ok: false, error: { message: e && e.message ? e.message : String(e) } }));
      });
    });
  });
  return new Promise((resolve, reject) => {
    server.on('error', reject);
    server.listen(0, '127.0.0.1', () => {
      const addr = server.address();
      resolve({
        token,
        url: 'http://127.0.0.1:' + addr.port + '/call',
        close() {
          return new Promise((done) => {
            if (typeof server.closeAllConnections === 'function') server.closeAllConnections();
            server.close(() => done());
          });
        },
      });
    });
  });
}

module.exports = { createBridge, BODY_MAX };
