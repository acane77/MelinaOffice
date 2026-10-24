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
const DOCUMENT_VARS = new Set(['--tabc', '--olvl', '--bu-c', '--bu-f', '--bu-s', '--bu-w', '--icon-swatch', '--scout-w', '--scout-tip', '--mt', '--mb', '--ws-pad']);

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
  assert.equal([...looks.matchAll(/--brand-quire:/g)].length, 8);
  const ui = read('common/ui.css');
  assert.match(ui, /\.dlg\.brand \.dlg-body,\n\.dlg\.brand \.dlg-foot \{ background-color: var\(--brand-bg\); \}/);
  assert.doesNotMatch(ui, /\.dlg\.brand \.dlg-(?:title|body|foot)[^}]*\b(?:margin|border|padding|box-shadow)\s*:/);
});

test('Golden Gate is the stored default, and the Windows 7 default button keeps the blue highlight', () => {
  const suite = read('common/suite.js');
  assert.match(suite, /const LOOKS = \[\{ id: 'goldengate', name: N\('Golden Gate'\),[\s\S]*?\n    \{ id: 'goldengatedark', name: N\('Golden Gate \(Dark\)'\),[\s\S]*?\n    \{ id: 'vscodegg', name: N\('Golden Gate \(Follow VS Code Color Theme\)'\),[\s\S]*?\n    \{ id: 'ribbon2010', name: N\('Ribbon 2010'\),[\s\S]*?\n    \{ id: 'win7', name: N\('Classic 7'\), sep: true,[\s\S]*?\n    \{ id: 'xp', name: N\('Classic XP'\),[\s\S]*?\n    \{ id: 'classic', name: N\('Classic'\),/);
  assert.match(suite, /id === 'paper' \|\| id === 'paper2016'/);
  assert.doesNotMatch(suite, /\{ id: 'paper'/);
  assert.match(suite, /const defaultLook = \(\) => \(vscodeHost\(\) \? 'vscodegg' : 'goldengate'\)/);
  assert.match(suite, /localStorage\.getItem\('vibeoffice\.look'\) \|\| defaultLook\(\)/);
  assert.match(suite, /if \(id === 'vscodegg' && !vscodeHost\(\)\) return 'goldengate'/);
  assert.match(suite, /return defaultLook\(\)/);
  assert.match(suite, /id: 'goldengate', name: N\('Golden Gate'\)/);
  for (const page of ['quire/index.html', 'ledger/index.html', 'lectern/index.html']) {
    assert.match(read(page), /localStorage\.getItem\('vibeoffice\.look'\)\|\|\(v\?'vscodegg':'goldengate'\)/);
    assert.match(read(page), /l==='paper'\|\|l==='paper2016'/);
    assert.match(read(page), /\{l=v\?'vscodegg':'goldengate'/);
    assert.match(read(page), /if\(l==='vscodegg'&&!v\)l='goldengate'/);
  }
  assert.match(read('../tools/ledger/build/template.html'), /localStorage\.getItem\('vibeoffice\.look'\)\|\|\(v\?'vscodegg':'goldengate'\)/);
  assert.match(read('../tools/ledger/build/template.html'), /l==='paper'\|\|l==='paper2016'/);
  assert.match(read('common/ui.js'), /if \(l\.sep\) items\.push\('-'\)/);
  assert.match(read('../desktop/src/host.rs'), /localStorage\.setItem\('vibeoffice\.look', 'goldengate'\)/);
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
  assert.match(read('common/suite.js'), /id: 'xp', name: N\('Classic XP'\)/);
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

test('the app pages pick the look before the first paint and load looks.css first', () => {
  for (const page of ['quire/index.html', 'ledger/index.html', 'lectern/index.html']) {
    const html = read(page);
    const script = html.indexOf("localStorage.getItem('vibeoffice.look')");
    const sheet = html.indexOf('common/looks.css');
    assert.ok(script > 0 && script < sheet, `${page}: the look is picked before looks.css loads`);
    assert.ok(sheet > 0 && (html.indexOf('common/ui.css') < 0 || sheet < html.indexOf('common/ui.css')), `${page}: looks.css before ui.css`);
    assert.match(html, /l==='ribbon2010'\|\|l==='goldengate'\|\|l==='goldengatedark'\|\|l==='vscodegg'/);
    assert.match(html, /v=window\.__VO_HOST__&&window\.__VO_HOST__\.kind==='vscode'/);
    assert.match(html, /l==='vscodegg'&&!v/);
    assert.match(html, /\|\|\(v\?'vscodegg':'goldengate'\)/);
  }
});

test('Ribbon Golden Gate copies its window motion and leaves Ribbon 2010 alone', () => {
  assert.match(read('common/suite.js'), /id: 'goldengate', name: N\('Golden Gate'\)/);
  assert.match(looks, /@keyframes goldengate-open[\s\S]*?scale\(0\.94\)/);
  assert.match(looks, /@keyframes goldengate-close[\s\S]*?scale\(0\.94\)/);
  assert.match(looks, /\[data-look="goldengate"\] \.dlg\.is-opening \{ animation: goldengate-open 250ms cubic-bezier\(0\.16, 1, 0\.3, 1\) both; \}/);
  assert.match(read('common/ui.js'), /look !== 'goldengate'/);
  assert.match(read('common/ui.js'), /animationName === 'goldengate-' \+ name/);
  assert.match(looks, /background: #ff5c60/);
  assert.match(looks, /\[data-look="goldengate"\] \.dlg-title::after \{ content: none; \}/);
  assert.match(looks, /--gg-hover: rgba\(0,0,0,\.08\)/);
  assert.match(looks, /--gg-on: rgba\(0,0,0,\.11\)/);
  assert.match(looks, /--gg-down: rgba\(0,0,0,\.16\)/);
  assert.match(looks, /\[data-look="ribbon2010"\] \.rb-tab\.file \{/);
  assert.match(looks, /\[data-look="ribbon2010"\] \.dlg\.is-opening \{ animation: win7-open /);
  assert.doesNotMatch(looks, /\[data-look="goldengate"\][^{]*\.ico[^{]*filter:\s*grayscale/);
  assert.match(read('common/suite.js'), /id: 'goldengatedark', name: N\('Golden Gate \(Dark\)'\)/);
  assert.match(looks, /\[data-look="goldengatedark"\] \{[^}]*color-scheme:\s*dark/);
  assert.match(looks, /@keyframes goldengatedark-open[\s\S]*?scale\(0\.94\)/);
  assert.match(looks, /\[data-look="goldengatedark"\] \.dlg\.is-opening \{ animation: goldengatedark-open 250ms cubic-bezier\(0\.16, 1, 0\.3, 1\) both; \}/);
  assert.match(looks, /\[data-look="goldengatedark"\][\s\S]*?--gg-hover:\s*rgba\(255,255,255,\.08\)/);
  assert.match(looks, /\[data-look="goldengatedark"\][\s\S]*?--page-canvas:\s*#141416/);
  assert.match(looks, /\[data-look="goldengatedark"\][\s\S]*?--hdr-a:\s*#ffffff/);
  assert.match(read('common/suite.js'), /id: 'vscodegg', name: N\('Golden Gate \(Follow VS Code Color Theme\)'\)/);
  assert.match(looks, /\[data-look="vscodegg"\][\s\S]*?--vsc-editor:\s*var\(--vscode-editor-background/);
  assert.match(looks, /--vsc-bar:\s*var\(--vscode-titleBar-activeBackground/);
  assert.match(looks, /@keyframes vscodegg-open[\s\S]*?scale\(0\.94\)/);
  assert.match(read('common/ui.js'), /animationName === 'vscodegg-' \+ name/);
  assert.doesNotMatch(looks, /:is\(\[data-look="goldengate"\], \[data-look="vscodegg"\]\)/);
});
