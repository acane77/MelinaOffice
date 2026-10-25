// The assistant (public/common/scout/), checked in each app: Help ▸ Show the Assistant brings her in with
// the welcome's choices (introducing herself the first time each day); the right-click menu is the app's own; she
// acts out what the user asks for (Print, Find; saving once the file is written) and works along with a long
// wait, with manners (one act waits for the one under way). A window in the background freezes her, and once
// hidden she is gone. Then, in Ecrivaine: she comes in by herself on a desktop, not on an iPhone, in an
// automated browser (tools/shot.mjs) or on a day the user hid her. Leaves the settings as it found them.
//
//   node --test tests/scout.test.mjs   (needs Chromium with CDP at 127.0.0.1:9222, as tools/shot.mjs, or CDP=… for one of
//   its own: serves on :8775)
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import path from 'node:path';

const root = path.resolve(path.dirname(new URL(import.meta.url).pathname), '..');
/* on an origin of its own (unless BASE says otherwise): the assistant's settings are shared by an origin's windows,
   so apps open on :8760 would answer this test's changes to them (the storage event) and change them back */
const env = { ...process.env, BASE: process.env.BASE || 'http://127.0.0.1:8775' };

const HELLO = /^Good (morning|afternoon|evening)! I'm Nikki\. I'll keep you company while you work\. What would you like to do\?$/;

const CHECK = `
const KEYS = ['melinaoffice.assistant', 'scout.lastGreetedDay', 'scout.lastHiddenDay'], KEY = KEYS[0];
const kept = JSON.parse(sessionStorage.getItem('scout.test.kept'));   // by tests/scout/clean.js
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const ui = L.ui, out = {};
const until = async (f, ms) => { const t0 = Date.now(); while (!f() && Date.now() - t0 < ms) await sleep(30); return f(); };
try {
  for (const k of KEYS) localStorage.removeItem(k);
  out.auto = L.assistant.shown();
  const help = L.app.menuBar.menus.find((m) => L.ui.plainAmp(m.label) === 'Help');
  out.inHelp = !!help && help.items.includes('showAssistant');
  const sel = help && help.items.find((x) => x && x.label === '&Select Assistant');
  out.hasSelect = !!sel;
  out.selectLabels = sel && sel.sub.map((x) => x.label);
  out.labelOff = ui.cmds.showAssistant.label;
  ui.exec('showAssistant');
  out.firstAct = L.assistant.scout().current && L.assistant.scout().current.name;
  await sleep(4000);
  const s = L.assistant.scout();
  out.labelOn = ui.cmds.showAssistant.label;
  out.stored = JSON.parse(localStorage.getItem(KEY)).on;
  out.visible = !s.hidden && getComputedStyle(s.el).visibility !== 'hidden';
  out.said = document.querySelector('.scout-balloon.is-on .scout-text').textContent;
  out.welcome = [...document.querySelectorAll('.scout-balloon.is-on .scout-choices button')].map((b) => b.textContent);
  /* the right-click menu */
  const r = s.el.getBoundingClientRect();
  s.el.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: r.left + r.width / 2, clientY: r.top + r.height / 2 }));
  out.menu = [...document.querySelectorAll('.menu .m-lbl')].map((x) => x.textContent);
  out.seps = document.querySelectorAll('.menu .m-sep').length;
  out.lastIsHide = document.querySelector('.menu').lastElementChild.previousElementSibling.classList.contains('m-sep');
  ui.closeMenus();
  /* shown again the same day: a wave, nothing said */
  s.stop();
  L.assistant.show(true);
  out.againAct = s.current && s.current.name;
  await sleep(2500);
  out.saidAgain = document.querySelector('.scout-balloon.is-on') ? document.querySelector('.scout-balloon.is-on .scout-text').textContent : null;
  /* a click on her: the choices */
  s.o.onclick();
  await sleep(1500);
  out.onClick = document.querySelector('.scout-balloon.is-on .scout-text') && document.querySelector('.scout-balloon.is-on .scout-text').textContent;
  out.choices = [...document.querySelectorAll('.scout-balloon.is-on .scout-choices button')].map((b) => b.textContent);
  s.stop();
  /* acting out what the user does, from the suite's signals */
  const act = async (fn) => { s.stop(); await sleep(250); fn(); await sleep(150); return s.current && s.current.name; };
  const command = (id) => window.dispatchEvent(new CustomEvent('vo-command', { detail: { id } }));
  /* Save and Save As once the file is written (L.saveFile, its download link held back here) */
  const click = HTMLAnchorElement.prototype.click;
  HTMLAnchorElement.prototype.click = function () {};
  const written = () => L.saveFile('scout-test.txt', new Blob(['x']));
  try {
    out.onPrint = await act(() => command('print'));
    out.printWritten = await act(written);
    out.onExport = await act(() => command('exportPDF'));
    out.onSaveAs = await act(() => command('saveAs'));
    out.onSave = await act(written);
  } finally { HTMLAnchorElement.prototype.click = click; }
  out.onWork = await act(() => ui.busy(true, 'Opening...'));
  await sleep(400);
  out.working = s.current && s.current.name;
  ui.busy(false);
  out.onFind = await act(() => command('find'));
  out.findAgain = await act(() => command('find'));
  let told = null;
  const listen = (e) => { told = e.detail.id; };
  addEventListener('vo-command', listen);
  ui.exec('selectAll');
  removeEventListener('vo-command', listen);
  out.told = told;
  s.stop();
  s.ask('Shall I?', ['Yes']);
  await sleep(300);
  command('insertPicture');
  await sleep(150);
  out.whileAsking = s.current && s.current.name;
  s.stop();
  const now = () => s.current && s.current.name + (s.current.loop ? ' loop' : '');
  /* one act waits for the one under way, a newer one in its place; the act she is at already, nothing new */
  await sleep(250);
  command('print');
  await sleep(150);
  command('print'); command('find'); command('insertPicture');
  out.waitWhile = now();
  out.waiting = s._nx && s._nx.name;
  await until(() => now() !== 'Print', 6000);
  out.waitedFor = now();
  s.stop();
  const hasFocus = document.hasFocus;
  document.hasFocus = () => false;
  window.dispatchEvent(new Event('blur'));
  await sleep(400);
  out.frozenAway = s.frozen;
  document.hasFocus = hasFocus;
  window.dispatchEvent(new Event('focus'));
  out.frozenBack = s.frozen;
  ui.exec('showAssistant');
  await sleep(150);
  out.goodbye = s.current && s.current.name;
  await sleep(3700);
  out.hidden = s.hidden;
  out.storedOff = JSON.parse(localStorage.getItem(KEY)).on;
  out.hiddenDay = localStorage.getItem('scout.lastHiddenDay');
  out.today = new Date().toLocaleDateString('sv');
} finally {
  KEYS.forEach((k, i) => { if (kept[i] == null) localStorage.removeItem(k); else localStorage.setItem(k, kept[i]); });
  sessionStorage.removeItem('scout.test.kept');
}
return JSON.stringify(out);
`;

for (const app of ['ecrivaine', 'matheuse', 'oratrice']) {
  test(`${app}: Nikki welcomes, has her menu, and freezes in the background or hidden`, () => {
    const out = execFileSync('node', [path.join(root, 'tools/shot.mjs'), app + '/', JSON.stringify([{ js: CHECK, wait: 200 }]), '--pre', path.join(root, 'tests/scout/clean.js')], { encoding: 'utf8', env });
    const line = out.split('\n').find((l) => l.startsWith('js -> '));
    assert.ok(line, out);
    const r = JSON.parse(JSON.parse(line.slice(6)));
    assert.equal(r.auto, false, 'not by herself in an automated browser');
    assert.ok(r.inHelp, 'Show the Assistant is in the Help menu');
    assert.ok(r.hasSelect, 'Select Assistant is in the Help menu');
    assert.deepEqual(r.selectLabels, ['&Nikki']);
    assert.equal(r.labelOff, 'Show the &Assistant');
    assert.equal(r.labelOn, 'Hide the &Assistant');
    assert.equal(r.firstAct, 'Greeting', 'shown from the menu, she waves');
    assert.equal(r.stored, true, 'shown is kept for the suite');
    assert.ok(r.visible, 'she is on the window');
    assert.match(r.said, HELLO, 'the first time today');
    assert.deepEqual(r.welcome, r.choices, 'with the same choices as a click');
    assert.equal(r.againAct, 'Greeting', 'shown again the same day: she waves');
    assert.equal(r.saidAgain, null, 'and says nothing');
    assert.equal(r.onClick, 'What would you like to do?', 'a click on her asks');
    assert.equal(r.choices[0], 'Find a command by name');
    assert.ok(r.choices.includes('Look something up in Help') && r.choices.includes('Send me a wink') && r.choices.at(-1) === 'Nothing for now', r.choices.join(', '));
    assert.deepEqual(r.menu.slice(0, 2), ['Animate!', 'Search commands']);
    assert.match(r.menu[2], /Help$/);
    assert.equal(r.menu.at(-1), 'Hide');
    assert.equal(r.seps, 2, 'separators: after Animate!, before Hide');
    assert.ok(r.lastIsHide, 'a separator just before Hide');
    assert.equal(r.onPrint, 'Print', 'Print');
    assert.equal(r.printWritten, null, 'a PDF written after Print is no save');
    assert.equal(r.onExport, null, 'an export: nothing when asked');
    assert.equal(r.onSaveAs, null, 'Save As: nothing while its dialog opens');
    assert.equal(r.onSave, 'Save', 'Save As: once the file is written');
    assert.equal(r.onWork, null, 'nothing for a short wait');
    assert.equal(r.working, 'Processing', 'working along once the wait goes on');
    assert.equal(r.onFind, 'Searching', 'Find');
    assert.equal(r.findAgain, 'Searching', 'the same act again: no wait');
    assert.equal(r.told, 'selectAll', 'ui.exec tells vo-command');
    assert.equal(r.whileAsking, null, 'nothing while she asks something');
    assert.equal(r.waitWhile, 'Print', 'the act under way plays out');
    assert.equal(r.waiting, 'GetArtsy', 'one act waits, the newest');
    assert.equal(r.waitedFor, 'GetArtsy', 'and plays next');
    assert.equal(r.frozenAway, true, 'the window in the background freezes her');
    assert.equal(r.frozenBack, false, 'she carries on when it comes back');
    assert.equal(r.goodbye, 'GoodBye', 'she waves goodbye');
    assert.ok(r.hidden, 'Hide the Assistant hides her');
    assert.equal(r.storedOff, false, 'hidden is kept for the suite');
    assert.equal(r.hiddenDay, r.today, 'and the day she was hidden');
    assert.match(out, /no errors, no external requests/);
  });
}

/* comes in by herself in a person's desktop browser; not on iOS or Android, nor on a day the user hid her
   (tools/shot.mjs is automated, so scripts before the page's own make it a person's browser and set the day) */
const UNASKED = `
const first = L.assistant.scout() && L.assistant.scout().current;
await new Promise((r) => setTimeout(r, 4500));
const s = L.assistant.scout(), said = document.querySelector('.scout-balloon.is-on .scout-text');
const out = { mobile: L.isMobile, webdriver: navigator.webdriver, shown: L.assistant.shown(), visible: !!s && !s.hidden, first: first && first.name,
  said: said ? said.textContent : null, choices: document.querySelectorAll('.scout-balloon.is-on .scout-choices button').length };
L.assistant.hide(true);
const kept = JSON.parse(sessionStorage.getItem('scout.test.kept'));
['melinaoffice.assistant', 'scout.lastGreetedDay', 'scout.lastHiddenDay'].forEach((k, i) => { if (kept[i] == null) localStorage.removeItem(k); else localStorage.setItem(k, kept[i]); });
sessionStorage.removeItem('scout.test.kept');
return JSON.stringify(out);
`;
/* [case, scripts before the page's, comes in, mobile, first act] */
for (const [name, pre, want, mobile, first] of [
  ['a desktop', ['clean', 'unasked'], true, false, 'Greeting'],
  ['a desktop, greeted today already', ['clean', 'unasked', 'greeted-today'], true, false, 'Show'],
  ['an iPhone', ['clean', 'unasked', 'iphone'], false, true],
  ['a desktop where she was hidden today', ['clean', 'unasked', 'hidden-today'], false, false],
  ['a desktop where she was hidden another day', ['clean', 'unasked', 'hidden-before'], true, false, 'Greeting'],
]) {
  test(`Nikki comes in by herself on ${name}: ${want}${first ? ', with ' + first : ''}`, () => {
    const files = pre.map((f) => path.join(root, 'tests/scout', f + '.js')).join(',');
    const out = execFileSync('node', [path.join(root, 'tools/shot.mjs'), 'ecrivaine/', JSON.stringify([{ js: UNASKED, wait: 100 }]), '--pre', files], { encoding: 'utf8', env });
    const line = out.split('\n').find((l) => l.startsWith('js -> '));
    assert.ok(line, out);
    const r = JSON.parse(JSON.parse(line.slice(6)));
    assert.equal(r.webdriver, false);
    assert.equal(r.mobile, mobile);
    assert.equal(r.shown, want);
    assert.equal(r.visible, want);
    if (first) assert.equal(r.first, first, 'her first act');
    if (first === 'Greeting') {
      assert.match(r.said, HELLO);
      assert.ok(r.choices > 3, 'with the choices');
    } else assert.equal(r.said, null, 'nothing said');
    assert.match(out, /no errors, no external requests/);
  });
}
