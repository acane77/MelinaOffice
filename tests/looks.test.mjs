// The look settings (public/common/looks.css) stay complete and the chrome uses them:
//   node --test tests/looks.test.mjs
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';

const root = new URL('../public/', import.meta.url);
const read = (p) => fs.readFileSync(new URL(p, root), 'utf8');
const looks = read('common/looks.css');
const styleOf = (html) => [...html.matchAll(/<style>([\s\S]*?)<\/style>/g)].map((m) => m[1]).join('\n');
const SHEETS = {
  'common/ui.css': read('common/ui.css'),
  'common/scout/scout.css': read('common/scout/scout.css'),
  'ledger/index.html': styleOf(read('ledger/index.html')),
  'quire/index.html': styleOf(read('quire/index.html')),
  'lectern/index.html': styleOf(read('lectern/index.html')),
  'index.html': styleOf(read('index.html')),
};
/* set per element by the apps for document content (and by scout.js for his size and balloon tip), not by a look */
const DOCUMENT_VARS = new Set(['--tabc', '--olvl', '--bu-c', '--bu-f', '--bu-s', '--bu-w', '--icon-swatch', '--scout-w', '--scout-tip']);

const names = (css) => new Set([...css.matchAll(/(--[a-z0-9-]+)\s*:/g)].map((m) => m[1]));
/** look id → the settings its blocks set (a block's selector may list several looks) */
const settingsOf = () => {
  const out = new Map();
  for (const m of looks.matchAll(/^([^\n{}]*data-look[^\n{}]*)\{([^}]*)\}/gm)) {
    if (!/^\s*(:root|\[data-look)/.test(m[1]) || /\)\s+[.#a-z]/.test(m[1])) continue;   // value blocks, not shape rules
    for (const id of m[1].matchAll(/data-look="([a-z0-9]+)"/g)) out.set(id[1], new Set([...(out.get(id[1]) || []), ...names(m[2])]));
  }
  return out;
};

test('every look sets every setting', () => {
  const all = settingsOf();
  const classic = all.get('classic');
  assert.ok(classic && classic.size > 50, 'Classic block found');
  const ids = [...read('common/suite.js').matchAll(/\{ id: '([a-z0-9]+)', name:/g)].map((m) => m[1]);
  assert.ok(ids.length >= 3, 'looks listed in suite.js');
  for (const id of ids) {
    assert.ok(all.has(id), `look "${id}" has settings in looks.css`);
    const missing = [...classic].filter((n) => !all.get(id).has(n));
    assert.deepEqual(missing, [], `look "${id}" falls back to Classic for ${missing.join(', ')}`);
  }
});

test('every setting the chrome uses is defined', () => {
  const defined = names(looks);
  for (const [file, css] of Object.entries(SHEETS)) {
    const used = new Set([...css.matchAll(/var\((--[a-z0-9-]+)/g)].map((m) => m[1]));
    /* --t-…: text an app sets in the UI's language for its CSS to draw (common/i18n.js), not a look setting */
    const missing = [...used].filter((n) => !defined.has(n) && !DOCUMENT_VARS.has(n) && !n.startsWith('--t-'));
    assert.deepEqual(missing, [], `${file} uses undefined settings`);
  }
});

test('the shared chrome has no colours of its own', () => {
  /* colour swatches and proofing marks show content colours, not the look */
  const allowed = /^\.(cp-sw|cc-|cpk-sw|sg-bad|sg-gram)/;
  for (const file of ['common/ui.css', 'common/scout/scout.css']) {
    const bad = SHEETS[file].split('\n')
      .filter((l) => /#[0-9a-fA-F]{3,8}\b|rgba?\(/.test(l) && !allowed.test(l.trim()));
    assert.deepEqual(bad, [], file);
  }
});

test('opening and about content uses each app colour and does not restyle the frame', () => {
  assert.match(looks, /--brand-quire:\s*#2D5999/);
  assert.match(looks, /--brand-ledger:\s*#137E43/);
  assert.match(looks, /--brand-lectern:\s*#B8492B/);
  assert.equal([...looks.matchAll(/--brand-quire:/g)].length, 4);
  const ui = read('common/ui.css');
  assert.match(ui, /\.dlg\.brand \.dlg-body,\n\.dlg\.brand \.dlg-foot \{ background-color: var\(--brand-bg\); \}/);
  assert.doesNotMatch(ui, /\.dlg\.brand \.dlg-(?:title|body|foot)[^}]*\b(?:margin|border|padding|box-shadow)\s*:/);
});

test('Windows 7 is the stored default, and its default button keeps the blue highlight', () => {
  const suite = read('common/suite.js');
  assert.match(suite, /localStorage\.getItem\('vibeoffice\.look'\) \|\| 'win7'/);
  assert.match(suite, /id: 'win7', name: N\('Windows 7'\)/);
  for (const page of ['index.html', 'quire/index.html', 'ledger/index.html', 'lectern/index.html']) {
    assert.match(read(page), /localStorage\.getItem\('vibeoffice\.look'\)\|\|'win7'/);
  }
  /* the reference defined the default-button colours and a pulse, then applied neither:
     the keyframes rule was not valid, and nothing selected the default button */
  assert.match(looks, /--w7-btn-default-line:\s*#5586a3/);
  assert.match(looks, /--w7-btn-default:\s*linear-gradient\([^)]*#c3dcea/);
  assert.match(looks, /\[data-look="win7"\] \.btn\.primary:not\(:disabled\)[\s\S]*?background:\s*var\(--w7-btn-default\)/);
  assert.match(looks, /@keyframes win7-pulse[\s\S]*?rgba\(52, 222, 255/);
  /* balloon face and shadow on tooltips and the assistant; the tail is his alone */
  assert.match(looks, /\[data-look="win7"\] :is\(\.tooltip, \.scout-balloon\)[\s\S]*?background:\s*var\(--w7-bl-bg\)[\s\S]*?box-shadow:\s*var\(--w7-bl-sd\)/);
  assert.match(looks, /--w7-bl-bg:\s*linear-gradient\(to bottom, #fff, #ebebeb\)/);
  assert.match(looks, /--w7-bl-sd:\s*5px 5px 3px -3px #0006/);
  assert.doesNotMatch(looks, /\[data-look="win7"\][^{]*\.tooltip[^{]*::before/);
  assert.match(looks, /\[data-look="win7"\] \.scout-balloon::before[\s\S]*?background-image:\s*var\(--w7-blt-up\)/);
});

test('Windows XP keeps the reference images and the transparent veil', () => {
  assert.match(read('common/suite.js'), /id: 'xp', name: N\('Windows XP'\)/);
  assert.match(looks, /--xp-close:\s*url\("data:image\/svg\+xml/);
  assert.match(looks, /--xp-dd:\s*url\("data:image\/svg\+xml/);
  assert.match(looks, /--xp-sb-up:\s*url\("data:image\/svg\+xml/);
  assert.match(looks, /--xp-slider:\s*url\("data:image\/svg\+xml/);
  assert.match(looks, /\[data-look="xp"\] \.dlg-overlay:not\(\.modeless\) \{ background: transparent; \}/);
  assert.match(looks, /\[data-look="xp"\] \{[^}]*--ui-font:\s*"Microsoft Yahei UI", "Segoe UI", system-ui, var\(--ui-cjk, sans-serif\)/);
});

test('disabled and read-only text fields use the dialog control colour', () => {
  const ui = read('common/ui.css');
  assert.match(ui, /\.dlg input\[type=text\]:read-only:not\(\.combo-in\), \.dlg textarea:read-only,[\s\S]*?background:\s*var\(--dlg\); color:\s*var\(--dlg-ink\)/);
  assert.match(ui, /\.dlg input\[type=text\]:disabled:not\(\.combo-in\), \.dlg textarea:disabled,[\s\S]*?color:\s*var\(--dis-ink\)[\s\S]*?user-select:\s*none; pointer-events:\s*none/);
  /* an editable field keeps the look's own face; a disabled field is :read-only, so it is left off that face */
  assert.match(looks, /\[data-look="win7"\] \.dlg input\[type=text\]:not\(:read-only\)[\s\S]*?background:\s*#fff/);
  assert.match(looks, /\[data-look="xp"\] \.dlg input\[type=text\]:not\(:read-only\)[\s\S]*?background:\s*#fff/);
});

test('the pages pick the look before the first paint and load looks.css first', () => {
  for (const page of ['index.html', 'quire/index.html', 'ledger/index.html', 'lectern/index.html']) {
    const html = read(page);
    const script = html.indexOf("localStorage.getItem('vibeoffice.look')");
    const sheet = html.indexOf('common/looks.css');
    assert.ok(script > 0 && script < sheet, `${page}: the look is picked before looks.css loads`);
    assert.ok(sheet > 0 && (html.indexOf('common/ui.css') < 0 || sheet < html.indexOf('common/ui.css')), `${page}: looks.css before ui.css`);
  }
});
