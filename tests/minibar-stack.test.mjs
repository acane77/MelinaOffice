// The formatting bar's layer: above every window, under every menu (common/ui.js, common/minibar.js).
//   node --test tests/minibar-stack.test.mjs
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
    const classes = new Set();
    el.classList = {
      add(c) { classes.add(c); },
      remove(c) { classes.delete(c); },
      toggle(c, on) {
        const next = on === undefined ? !classes.has(c) : !!on;
        if (next) classes.add(c); else classes.delete(c);
        return next;
      },
      contains(c) { return classes.has(c); },
    };
    el.setAttribute = (k, v) => { el.attrs[k] = v == null ? null : String(v); };
    el.getAttribute = (k) => el.attrs[k];
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
  function find(sel) {
    const stack = els.slice();
    while (stack.length) {
      const el = stack.shift();
      if (el.isConnected && sel === '.minibar' && String(el.className || '').split(/\s+/).includes('minibar')) return el;
      if (el.children) stack.push(...el.children);
    }
    return null;
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
    querySelector: find,
    querySelectorAll: () => [],
    getElementById: () => null,
  };
  const sandbox = {
    console, Math, Date, JSON, Object, Array, String, Number, Error, Promise, Set, Map,
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
  vm.runInContext(fs.readFileSync(path.join(root, 'public/common/minibar.js'), 'utf8'), sandbox, { filename: 'minibar.js' });
  sandbox.L.minibar.mount(() => []);
  return { L: sandbox.L, bar: () => find('.minibar'), document };
}

const { L, bar, document } = load();
const z = (el) => parseInt(el.style.zIndex, 10);

test('with no window open the bar sits at 2000, and a menu stays above it', () => {
  L.minibar.show(20, 40);
  assert.equal(L.minibar.visible(), true);
  assert.equal(z(bar()), 2000);
  const menu = L.ui.openMenu([{ label: 'Font' }], { x: 10, y: 10 });
  assert.equal(z(menu), 6000);
  assert.ok(z(menu) > z(bar()));
  L.ui.closeMenus();
});

test('the bar stays one above the front window, and a menu stays above the bar', () => {
  L.minibar.show(20, 40);
  const a = L.ui.dialog({ title: 'Find', modeless: true, buttons: [] });
  assert.ok(z(bar()) > z(a.overlay));
  const b = L.ui.dialog({ title: 'Replace', modeless: true, buttons: [] });
  assert.ok(z(b.overlay) > z(a.overlay));
  assert.equal(z(bar()), z(b.overlay) + 1);
  const menu = L.ui.openMenu([{ label: 'Size' }], { x: 10, y: 10 });
  assert.ok(z(menu) > z(bar()));
  L.ui.closeMenus();
  b.close();
  a.close();
  assert.equal(z(bar()), 2000);
});

test('a menu rises above a bar that has reached the menu floor', () => {
  L.minibar.show(20, 40);
  bar().style.zIndex = '6100';
  const menu = L.ui.openMenu([{ label: 'Colour' }], { x: 10, y: 10 });
  assert.equal(z(menu), 6120);
  L.ui.closeMenus();
});

test('a hidden bar does not lift the menu', () => {
  L.minibar.hide();
  bar().style.zIndex = '9000';
  const menu = L.ui.openMenu([{ label: 'Font' }], { x: 10, y: 10 });
  assert.equal(z(menu), 6000);
  L.ui.closeMenus();
});

test('a Golden Gate menu is capped to the viewport so a long list scrolls', () => {
  for (const look of ['goldengate', 'goldengatedark', 'vscodegg']) {
    document.documentElement.dataset.look = look;
    const menu = L.ui.openMenu([{ label: 'Calibri' }, { label: 'Cambria' }], { x: 10, y: 40 });
    assert.equal(menu.style.position, 'relative');
    assert.equal(menu.style.maxHeight, '592px');
    assert.equal(menu.style.overflowY, 'auto');
    assert.equal(menu.parentNode.className, 'gg-menu-glass');
    L.ui.closeMenus();
  }
  document.documentElement.dataset.look = '';
  const plain = L.ui.openMenu([{ label: 'Calibri' }], { x: 10, y: 40 });
  assert.equal(plain.style.maxHeight || '', '');
  assert.notEqual(plain.parentNode.className, 'gg-menu-glass');
  L.ui.closeMenus();
});

test('a modal dialog hides the bar instead of leaving it above the dialog', () => {
  L.minibar.show(20, 40);
  const d = L.ui.dialog({ title: 'Go To' });
  assert.equal(L.minibar.visible(), false);
  assert.equal(bar().hidden, true);
  assert.ok(z(d.overlay) >= 2000);
  d.close();
});
