// The status-bar zoom maths (common/ui.js): a wheel notch, a two-finger pinch, and the scroll
// that keeps the point under the fingers in place.
//   node --test tests/zoom.test.mjs
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

function load() {
  const listeners = [];
  const document = {
    addEventListener(type, fn) { listeners.push([type, fn]); },
    documentElement: { style: {} },
    body: { appendChild() {}, classList: { add() {}, remove() {}, toggle() {} } },
    createElement: () => node(),
    createElementNS: () => node(),
    createTextNode: (t) => ({ text: t }),
    querySelector: () => null,
    querySelectorAll: () => [],
  };
  function node() {
    return {
      style: {}, dataset: {}, children: [], attrs: {}, className: '',
      classList: { add() {}, remove() {}, toggle() {}, contains: () => false },
      setAttribute(k, v) { this.attrs[k] = String(v); },
      getAttribute(k) { return this.attrs[k]; },
      appendChild(c) { this.children.push(c); return c; },
      addEventListener() {}, removeEventListener() {}, blur() {}, focus() {}, select() {},
      closest: () => null, contains: () => false,
    };
  }
  const sandbox = {
    console, Math, Date, JSON, Object, Array, String, Number, Error, parseInt, parseFloat, isNaN, NaN, Infinity, undefined,
    setTimeout, clearTimeout, requestAnimationFrame: (fn) => fn(),
    navigator: { platform: 'Linux', userAgent: 'node' },
    addEventListener() {}, removeEventListener() {}, dispatchEvent() {},
    document, Node: function Node() {},
  };
  sandbox.window = sandbox;
  sandbox.globalThis = sandbox;
  sandbox.L = { APP_ID: 'quire', APP_NAME: 'Quire', APP: 'Quire' };
  vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(path.join(root, 'public/common/core.js'), 'utf8'), sandbox, { filename: 'core.js' });
  vm.runInContext(fs.readFileSync(path.join(root, 'public/common/ui.js'), 'utf8'), sandbox, { filename: 'ui.js' });
  return sandbox.L.ui;
}

const ui = load();

test('a wheel notch zooms out, a pinch apart zooms in by the inverse', () => {
  const out = ui.wheelZoomFactor(100, 0);
  const inn = ui.wheelZoomFactor(-100, 0);
  assert.ok(out < 1 && out > 0.8, out);
  assert.ok(Math.abs(inn - 1 / out) < 1e-9);
  assert.equal(ui.wheelZoomFactor(0, 0), 1);
  const line = ui.wheelZoomFactor(3, 1);
  assert.ok(Math.abs(line - ui.wheelZoomFactor(96, 0)) < 1e-9);
});

test('two fingers: apart scales up, together scales down, a shared slide pans', () => {
  const apart = ui.pinchStep({ x: 0, y: 0 }, { x: 10, y: 0 }, { x: 0, y: 0 }, { x: 20, y: 0 });
  assert.equal(apart.scale, 2);
  assert.equal(apart.dx, 5);
  const together = ui.pinchStep({ x: 0, y: 0 }, { x: 40, y: 0 }, { x: 10, y: 0 }, { x: 30, y: 0 });
  assert.equal(together.scale, 0.5);
  const slide = ui.pinchStep({ x: 0, y: 0 }, { x: 20, y: 0 }, { x: 8, y: 6 }, { x: 28, y: 6 });
  assert.equal(slide.scale, 1);
  assert.equal(slide.dx, 8);
  assert.equal(slide.dy, 6);
});

test('zooming keeps the point under the pointer, and a sideways move follows the fingers', () => {
  const zoomed = ui.zoomScroll({ left: 0, top: 0 }, 1, 2, { x: 100, y: 40 }, { x: 100, y: 40 });
  assert.equal(zoomed.left, 100);
  assert.equal(zoomed.top, 40);
  const panned = ui.zoomScroll({ left: 40, top: 10 }, 1, 1, { x: 0, y: 0 }, { x: 30, y: -4 });
  assert.equal(panned.left, 10);
  assert.equal(panned.top, 14);
});
