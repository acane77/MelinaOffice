// The task-pane reference names the objects in each DOM specification.
//   node --test tests/vb-dom-ref.test.mjs
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';

const root = new URL('../', import.meta.url);
const read = (p) => fs.readFileSync(new URL(p, root), 'utf8');

/* section titles that are rules, not object names */
const META = /^(Rules|Status and gaps|Story index|Text index|Objects with no document operation|Enumerations|Constants|Worked calls)$/;

function objectNames(md) {
  const names = [];
  for (const m of md.matchAll(/^## (.+)$/gm)) {
    const heading = m[1].trim().replace(/ \(not implemented\)$/, '');
    if (META.test(heading)) continue;
    for (const name of heading.match(/\b[A-Z][A-Za-z0-9]+\b/g) || []) names.push(heading + ': ' + name);
  }
  return names;
}

const APPS = [
  ['ecrivaine', 'docs/ecrivaine_dom_api_specifications.md', 'public/ecrivaine/js/vb-dom-ref.js', 'public/ecrivaine/index.html'],
  ['matheuse', 'docs/matheuse_dom_api_specifications.md', 'public/matheuse/js/vb-dom-ref.js', 'public/matheuse/index.html'],
  ['oratrice', 'docs/oratrice_dom_api_specifications.md', 'public/oratrice/js/vb-dom-ref.js', 'public/oratrice/index.html'],
];

for (const [app, spec, catalog, page] of APPS) {
  test(app + ' reference names every spec object', () => {
    const src = read(catalog);
    const missing = objectNames(read(spec)).filter((row) => !src.includes(row.split(': ')[1]));
    assert.deepEqual(missing, []);
  });

  test(app + ' page loads the reference after the task pane', () => {
    const html = read(page);
    const panes = html.indexOf('js/panes.js');
    const shared = html.indexOf('common/vb-dom-ref.js');
    const own = html.indexOf('js/vb-dom-ref.js');
    assert.ok(panes >= 0 && shared > panes && own > shared, 'script order');
  });
}

test('the console button opens the task pane', () => {
  for (const file of ['public/ecrivaine/js/dom-console.js', 'public/matheuse/js/dom-console.js', 'public/oratrice/js/dom-console.js']) {
    const src = read(file);
    assert.match(src, /label: 'DOM API References'[\s\S]{0,120}L\.vbDomRef\.show\(\)/);
    assert.doesNotMatch(src, /window\.open\('https:\/\/learn\.microsoft\.com/);
  }
});

test('each specification points at its reference', () => {
  assert.match(read('docs/ecrivaine_dom_api_specifications.md'), /public\/ecrivaine\/js\/vb-dom-ref\.js/);
  assert.match(read('docs/matheuse_dom_api_specifications.md'), /public\/matheuse\/js\/vb-dom-ref\.js/);
  assert.match(read('docs/oratrice_dom_api_specifications.md'), /public\/oratrice\/js\/vb-dom-ref\.js/);
});
