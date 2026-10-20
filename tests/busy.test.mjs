// The waiting window (ui.busy): a modal with the app's name, no close button and no buttons,
// covering a window already open.
//   node --test tests/busy.test.mjs
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

function load() {
  function Node() {}
  const els = [];
  function detach(el) {
    const i = els.indexOf(el);
    if (i >= 0) els.splice(i, 1);
  }
  function element(tag) {
    const el = Object.create(Node.prototype);
    el.tagName = String(tag || '').toUpperCase();
    el.style = {};
    el.dataset = {};
    el.children = [];
    el.attrs = {};
    el.className = '';
    el.hidden = false;
    el.isConnected = false;
    const classes = () => new Set(String(el.className || '').split(/\s+/).filter(Boolean));
    el.classList = {
      add(c) { const s = classes(); s.add(c); el.className = [...s].join(' '); },
      remove(c) { const s = classes(); s.delete(c); el.className = [...s].join(' '); },
      toggle(c, on) {
        const s = classes();
        const next = on === undefined ? !s.has(c) : !!on;
        if (next) s.add(c); else s.delete(c);
        el.className = [...s].join(' ');
        return next;
      },
      contains(c) { return classes().has(c); },
    };
    el.setAttribute = (k, v) => { el.attrs[k] = v == null ? null : String(v); if (k === 'class') el.className = v; };
    el.getAttribute = (k) => (k === 'class' ? el.className : el.attrs[k]);
    el.appendChild = (c) => { if (c) { el.children.push(c); c.parentNode = el; } return c; };
    el.append = (...cs) => { for (const c of cs) if (c) el.appendChild(c); };
    el.remove = () => { el.isConnected = false; detach(el); };
    el.addEventListener = () => {};
    el.removeEventListener = () => {};
    el.blur = () => {};
    el.focus = () => {};
    el.select = () => {};
    el.closest = () => null;
    el.contains = () => false;
    el.getBoundingClientRect = () => ({ width: 120, height: 28, left: 0, top: 0, right: 120, bottom: 28 });
    el.querySelector = () => null;
    el.querySelectorAll = () => [];
    return el;
  }
  function connected() {
    const out = [];
    const stack = els.filter((el) => el.isConnected);
    while (stack.length) {
      const el = stack.shift();
      out.push(el);
      if (el.children) stack.push(...el.children);
    }
    return out;
  }
  const body = element('body');
  body.appendChild = (el) => { els.push(el); el.isConnected = true; el.parentNode = body; return el; };
  body.insertBefore = (el, ref) => {
    detach(el);
    const i = els.indexOf(ref);
    if (i < 0) els.push(el); else els.splice(i, 0, el);
    el.isConnected = true;
    el.parentNode = body;
    return el;
  };
  const document = {
    addEventListener() {},
    documentElement: { style: {}, dataset: {} },
    body,
    activeElement: null,
    createElement: (tag) => element(tag),
    createElementNS: (ns, tag) => element(tag),
    createTextNode: (t) => { const n = Object.create(Node.prototype); n.text = t; return n; },
    querySelector: () => null,
    querySelectorAll: () => [],
    getElementById: () => null,
  };
  const sandbox = {
    console, Math, Date, JSON, Object, Array, String, Number, Error, Promise, Set, Map, CustomEvent,
    parseInt, parseFloat, isNaN, NaN, Infinity, undefined,
    setTimeout, clearTimeout,
    requestAnimationFrame: (fn) => fn(),
    navigator: { platform: 'Linux', userAgent: 'node' },
    addEventListener() {}, removeEventListener() {}, dispatchEvent() {},
    getSelection: () => ({ rangeCount: 0 }),
    DOMRect: function DOMRect() {},
    innerWidth: 800, innerHeight: 600,
    document, Node,
  };
  sandbox.window = sandbox;
  sandbox.globalThis = sandbox;
  sandbox.L = { APP_ID: 'quire', APP_NAME: 'Quire', APP: 'Quire' };
  vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(path.join(root, 'public/common/core.js'), 'utf8'), sandbox, { filename: 'core.js' });
  vm.runInContext(fs.readFileSync(path.join(root, 'public/common/ui.js'), 'utf8'), sandbox, { filename: 'ui.js' });
  return { L: sandbox.L, connected };
}

const { L, connected } = load();
const has = (el, c) => !!(el.classList && el.classList.contains(c)) || String(el.className || '').split(/\s+/).includes(c);
const busyWins = () => connected().filter((el) => has(el, 'busy') && has(el, 'dlg'));
const inside = (el) => {
  const out = [];
  const stack = (el.children || []).slice();
  while (stack.length) {
    const n = stack.shift();
    out.push(n);
    if (n.children) stack.push(...n.children);
  }
  return out;
};

test('the waiting window is a modal titled with the app, with no close button and no buttons', () => {
  const opening = L.ui.dialog({
    title: 'Starting Quire',
    canClose: false,
    buttons: [{ label: 'Cancel', cancel: true }],
    body: L.h('p', { text: 'Opening the file. Please wait...' }),
  });
  L.ui.busy(true, 'Opening a.docx...');
  const wins = busyWins();
  assert.equal(wins.length, 1);
  const bits = inside(wins[0]);
  assert.equal(bits.find((n) => has(n, 'dlg-ttl')).textContent, 'Quire');
  assert.equal(bits.some((n) => has(n, 'dlg-x')), false);
  assert.equal(bits.some((n) => has(n, 'btn')), false);
  assert.equal(bits.find((n) => has(n, 'busy-t')).textContent, 'Opening a.docx...');
  const z = (el) => parseInt(el.style.zIndex, 10);
  assert.ok(z(wins[0].parentNode) > z(opening.overlay));
  L.ui.busy(true, 'Creating PDF...');
  assert.equal(busyWins().length, 1);
  assert.equal(inside(busyWins()[0]).find((n) => has(n, 'busy-t')).textContent, 'Creating PDF...');
  L.ui.busy(false);
  assert.equal(busyWins().length, 0);
  assert.equal(opening.overlay.isConnected, true);
  opening.close(null);
});
