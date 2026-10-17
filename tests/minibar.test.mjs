// The formatting bar's fade and placement (common/minibar.js). No page: the two functions are pure.
//   node --test tests/minibar.test.mjs
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

function load() {
  const sandbox = { console, Math };
  sandbox.window = sandbox;
  sandbox.globalThis = sandbox;
  sandbox.L = {};
  vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(path.join(root, 'public/common/minibar.js'), 'utf8'), sandbox, { filename: 'minibar.js' });
  return sandbox.L.minibar;
}

const mb = load();

test('starts at 70%, solid on the bar, gone at 50px, and a fade-out stays gone', () => {
  assert.equal(mb.fade(0, false, false).opacity, mb.REST);
  assert.equal(mb.fade(0, false, false).latched, false);
  const mid = mb.fade(25, false, false);
  assert.ok(Math.abs(mid.opacity - mb.REST * 0.5) < 1e-12);
  assert.equal(mid.latched, false);
  const near = mb.fade(10, false, false);
  assert.ok(near.opacity > mid.opacity);   /* coming back toward the open point raises it */
  const gone = mb.fade(50, false, false);
  assert.equal(gone.opacity, 0);
  assert.equal(gone.latched, true);
  assert.equal(mb.fade(80, false, false).latched, true);
  const on = mb.fade(80, true, false);
  assert.equal(on.opacity, 1);
  assert.equal(on.latched, false);
  const stuck = mb.fade(0, true, true);
  assert.equal(stuck.opacity, 0);
  assert.equal(stuck.latched, true);
});

test('the bar sits up-right of the pointer and stays inside the window', () => {
  const open = mb.place({ x: 200, y: 180 }, { w: 100, h: 40 }, { w: 800, h: 600 });
  assert.equal(open.x, 200 + mb.GAP);
  assert.equal(open.y, 180 - 40 - mb.GAP);
  const top = mb.place({ x: 20, y: 10 }, { w: 100, h: 40 }, { w: 800, h: 600 });
  assert.equal(top.x, 20 + mb.GAP);
  assert.equal(top.y, 2);
  const edge = mb.place({ x: 790, y: 180 }, { w: 100, h: 40 }, { w: 800, h: 600 });
  assert.equal(edge.x, 800 - 100 - 2);
  assert.equal(edge.y, 180 - 40 - mb.GAP);
  const tight = mb.place({ x: 30, y: 30 }, { w: 200, h: 80 }, { w: 100, h: 60 });
  assert.equal(tight.x, 2);
  assert.equal(tight.y, 2);
});
