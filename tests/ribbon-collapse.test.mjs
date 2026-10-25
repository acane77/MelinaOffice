// A collapsed ribbon shows a contextual tab without dropping its panel over the document (common/ribbon.js).
//   node --test tests/ribbon-collapse.test.mjs
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

function load() {
  function Node() {}
  function names(el) { return String(el.className || '').split(/\s+/).filter(Boolean); }
  function markConnected(node) {
    node.isConnected = true;
    for (const c of node.children || []) markConnected(c);
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
    el.classList = {
      add(...cs) { el.className = [...new Set([...names(el), ...cs])].join(' '); },
      remove(...cs) { const drop = new Set(cs); el.className = names(el).filter((c) => !drop.has(c)).join(' '); },
      toggle(c, on) {
        const has = names(el).includes(c);
        const next = on === undefined ? !has : !!on;
        if (next) el.classList.add(c); else el.classList.remove(c);
        return next;
      },
      contains(c) { return names(el).includes(c); },
    };
    el.setAttribute = (k, v) => {
      el.attrs[k] = v == null ? null : String(v);
      if (String(k).startsWith('data-')) {
        const prop = k.slice(5).replace(/-([a-z])/g, (_, ch) => ch.toUpperCase());
        el.dataset[prop] = el.attrs[k];
      }
    };
    el.getAttribute = (k) => el.attrs[k];
    el.appendChild = (c) => {
      if (!c) return c;
      el.children.push(c);
      c.parentNode = el;
      if (el.isConnected) markConnected(c);
      return c;
    };
    el.append = (...cs) => { for (const c of cs) if (c) el.appendChild(c); };
    el.prepend = (...cs) => {
      for (let i = cs.length - 1; i >= 0; i--) {
        const c = cs[i];
        if (!c) continue;
        el.children.unshift(c);
        c.parentNode = el;
        if (el.isConnected) markConnected(c);
      }
    };
    el.remove = () => { el.isConnected = false; };
    el.addEventListener = () => {};
    el.removeEventListener = () => {};
    el.blur = () => {};
    el.focus = () => {};
    el.closest = () => null;
    el.contains = () => false;
    el.matches = () => false;
    el.getBoundingClientRect = () => ({ width: 80, height: 22, left: 0, top: 0, right: 80, bottom: 22 });
    Object.defineProperty(el, 'firstChild', { get() { return el.children[0] || null; } });
    Object.defineProperty(el, 'textContent', {
      get() { return ''; },
      set() { el.children.length = 0; },
    });
    el.querySelectorAll = (sel) => {
      const out = [];
      const walk = (n) => {
        for (const c of n.children || []) {
          if (sel === '.rb-tab' && names(c).includes('rb-tab')) out.push(c);
          walk(c);
        }
      };
      walk(el);
      return out;
    };
    el.querySelector = (sel) => el.querySelectorAll(sel)[0] || null;
    return el;
  }
  const cmdbars = element('div');
  cmdbars.className = 'cmdbars';
  cmdbars.isConnected = true;
  const body = element('body');
  body.appendChild(cmdbars);
  const listeners = [];
  const document = {
    addEventListener(type, fn) { listeners.push({ type, fn }); },
    removeEventListener(type, fn) {
      const i = listeners.findIndex((x) => x.type === type && x.fn === fn);
      if (i >= 0) listeners.splice(i, 1);
    },
    documentElement: { style: {}, dataset: { layout: 'ribbon' } },
    body,
    activeElement: null,
    createElement: (tag) => element(tag),
    createElementNS: (ns, tag) => element(tag),
    createTextNode: (t) => { const n = Object.create(Node.prototype); n.text = t; return n; },
    querySelector: (sel) => (sel === '.cmdbars' ? cmdbars : null),
    querySelectorAll: () => [],
    getElementById: () => null,
  };
  const store = new Map();
  const sandbox = {
    console, Math, Date, JSON, Object, Array, String, Number, Error, Promise, Set, Map,
    parseInt, parseFloat, isNaN, NaN, Infinity, undefined,
    setTimeout, clearTimeout,
    requestAnimationFrame: (fn) => fn(),
    navigator: { platform: 'Linux', userAgent: 'node' },
    addEventListener() {}, removeEventListener() {}, dispatchEvent() {},
    getSelection: () => ({ rangeCount: 0 }),
    DOMRect: function DOMRect() {},
    KeyboardEvent: function KeyboardEvent() {},
    CustomEvent: function CustomEvent() {},
    innerWidth: 1200, innerHeight: 800,
    document, Node,
    localStorage: {
      getItem: (k) => (store.has(k) ? store.get(k) : null),
      setItem: (k, v) => { store.set(k, String(v)); },
      removeItem: (k) => { store.delete(k); },
    },
  };
  sandbox.window = sandbox;
  sandbox.globalThis = sandbox;
  sandbox.L = { APP_ID: 'ecrivaine', APP_NAME: 'Ecrivaine', APP: 'Ecrivaine', t: (s) => s };
  vm.createContext(sandbox);
  for (const file of ['public/common/core.js', 'public/common/ui.js', 'public/common/search.js', 'public/common/ribbon.js']) {
    vm.runInContext(fs.readFileSync(path.join(root, file), 'utf8'), sandbox, { filename: file });
  }
  return { L: sandbox.L, localStorage: sandbox.localStorage, ribbon: () => cmdbars.children.find((c) => names(c).includes('ribbon')) };
}

function tabIds(ribbon, cls) {
  const out = [];
  const walk = (n) => {
    for (const c of n.children || []) {
      const cs = String(c.className || '').split(/\s+/);
      if (cs.includes('rb-tab') && (!cls || cs.includes(cls))) out.push(c.dataset.tab);
      walk(c);
    }
  };
  walk(ribbon);
  return out;
}

test('a collapsed ribbon shows a new contextual tab without opening its panel', () => {
  const { L, localStorage, ribbon } = load();
  const ctx = { picture: false, header: false };
  L.ribbonSpec = () => ({
    tabs: [
      { id: 'home', label: '&Home', groups: [{ label: 'Font', items: ['bold'] }] },
      { id: 'insert', label: '&Insert', groups: [{ label: 'Illustrations', items: ['picture'] }] },
    ],
    contextual: [
      { id: 'picture', set: 'Picture Tools', when: () => ctx.picture, tabs: [
        { id: 'pictureFormat', label: 'Format', groups: [{ label: 'Size', items: ['picCrop'] }] },
      ] },
      { id: 'headerFooter', set: 'Header & Footer Tools', open: true, when: () => ctx.header, tabs: [
        { id: 'headerDesign', label: 'Design', groups: [{ label: 'Header', items: ['header'] }] },
      ] },
    ],
  });
  localStorage.setItem('ecrivaine.ribbonCollapsed', 'true');
  L.ribbonUI.init();
  const bar = ribbon();
  assert.equal(bar.classList.contains('collapsed'), true);
  assert.equal(bar.classList.contains('peek'), false);
  assert.deepEqual(tabIds(bar, 'on'), ['home']);

  ctx.header = true;
  L.bus.emit('ui-refresh');
  assert.equal(bar.classList.contains('collapsed'), true);
  assert.equal(bar.classList.contains('peek'), false);
  assert.ok(tabIds(bar).includes('headerDesign'));
  assert.deepEqual(tabIds(bar, 'on'), ['headerDesign']);

  L.ribbonUI.select('headerDesign');
  assert.equal(bar.classList.contains('peek'), true);

  L.ribbonUI.collapse(true);
  assert.equal(bar.classList.contains('collapsed'), true);
  assert.equal(bar.classList.contains('peek'), false);
  ctx.header = false;
  L.bus.emit('ui-refresh');
  assert.equal(bar.classList.contains('peek'), false);
  assert.deepEqual(tabIds(bar, 'on'), ['home']);

  L.ribbonUI.select('insert');
  L.ribbonUI.collapse(true);
  ctx.picture = true;
  L.bus.emit('ui-refresh');
  assert.equal(bar.classList.contains('collapsed'), true);
  assert.equal(bar.classList.contains('peek'), false);
  assert.ok(tabIds(bar).includes('pictureFormat'));
  assert.deepEqual(tabIds(bar, 'on'), ['pictureFormat']);

  L.ribbonUI.collapse(false);
  assert.equal(bar.classList.contains('collapsed'), false);
  assert.equal(bar.classList.contains('peek'), false);
  assert.deepEqual(tabIds(bar, 'on'), ['pictureFormat']);
});
