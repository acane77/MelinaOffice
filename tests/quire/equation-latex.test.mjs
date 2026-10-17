// OMML to LaTeX, and LaTeX back to the OMML Word writes. The binomial is the equation in
// test_documents/Equation_Test.docx. node --test tests/quire/equation-latex.test.mjs
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';

function decode(s) {
  return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&amp;/g, '&');
}
function parseXml(xml) {
  const root = { nodeType: 1, localName: 'root', nodeName: 'root', attributes: [], childNodes: [], textContent: '' };
  const stack = [root];
  const re = /<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<\/([^\s>]+)>|<([^\s>/]+)([^>]*?)\/>|<([^\s>/]+)([^>]*)>|([^<]+)/g;
  let m;
  while ((m = re.exec(xml))) {
    if (m[1]) {
      const name = m[1].replace(/^.*:/, '');
      while (stack.length > 1 && stack[stack.length - 1].localName !== name) stack.pop();
      if (stack.length > 1) stack.pop();
    } else if (m[2]) {
      const attrs = [];
      const ar = /([^\s=]+)="([^"]*)"/g;
      let a;
      while ((a = ar.exec(m[3] || ''))) attrs.push({ localName: a[1].replace(/^.*:/, ''), value: decode(a[2]) });
      const node = { nodeType: 1, localName: m[2].replace(/^.*:/, ''), nodeName: m[2], attributes: attrs, childNodes: [], textContent: '' };
      stack[stack.length - 1].childNodes.push(node);
    } else if (m[4]) {
      const attrs = [];
      const ar = /([^\s=]+)="([^"]*)"/g;
      let a;
      while ((a = ar.exec(m[5] || ''))) attrs.push({ localName: a[1].replace(/^.*:/, ''), value: decode(a[2]) });
      const node = { nodeType: 1, localName: m[4].replace(/^.*:/, ''), nodeName: m[4], attributes: attrs, childNodes: [], textContent: '' };
      stack[stack.length - 1].childNodes.push(node);
      stack.push(node);
    } else if (m[6]) {
      const t = decode(m[6]);
      if (!t.trim()) continue;
      stack[stack.length - 1].textContent += t;
    }
  }
  return root.childNodes[0];
}
function DOMParser() {}
DOMParser.prototype.parseFromString = function (xml) {
  if (!xml || !String(xml).trim().startsWith('<')) {
    const bad = { nodeType: 1, localName: 'parsererror', nodeName: 'parsererror', attributes: [], childNodes: [], textContent: '' };
    return { documentElement: bad, getElementsByTagName: () => [bad] };
  }
  return { documentElement: parseXml(xml), getElementsByTagName: () => [] };
};
function mkel(tag) {
  const e = {
    tag, attrs: {}, children: [], text: '',
    setAttribute(k, v) { this.attrs[k] = v; },
    appendChild(c) { this.children.push(c); return c; },
    append(...cs) { cs.forEach((c) => this.appendChild(c)); },
  };
  Object.defineProperty(e, 'textContent', { get() { return this.text; }, set(v) { this.text = v; } });
  return e;
}

const L = {};
const sandbox = {
  window: { L },
  console,
  DOMParser,
  document: { createElementNS: (_ns, tag) => mkel(tag) },
};
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(new URL('../../public/quire/js/omml.js', import.meta.url), 'utf8'), sandbox);
vm.runInContext(fs.readFileSync(new URL('../../public/quire/js/latexprev.js', import.meta.url), 'utf8'), sandbox);

const BINOM = '<m:oMathPara><m:oMath><m:sSup><m:sSupPr><m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr></m:ctrlPr></m:sSupPr><m:e><m:d><m:dPr><m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr></m:ctrlPr></m:dPr><m:e><m:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><m:t>1+x</m:t></m:r></m:e></m:d></m:e><m:sup><m:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><m:t>n</m:t></m:r></m:sup></m:sSup><m:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><m:t>=1+</m:t></m:r><m:f><m:fPr><m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr></m:ctrlPr></m:fPr><m:num><m:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><m:t>nx</m:t></m:r></m:num><m:den><m:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><m:t>1!</m:t></m:r></m:den></m:f><m:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><m:t>+</m:t></m:r><m:f><m:fPr><m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr></m:ctrlPr></m:fPr><m:num><m:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><m:t>n</m:t></m:r><m:d><m:dPr><m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr></m:ctrlPr></m:dPr><m:e><m:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><m:t>n-1</m:t></m:r></m:e></m:d><m:sSup><m:sSupPr><m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr></m:ctrlPr></m:sSupPr><m:e><m:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><m:t>x</m:t></m:r></m:e><m:sup><m:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><m:t>2</m:t></m:r></m:sup></m:sSup></m:num><m:den><m:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><m:t>2!</m:t></m:r></m:den></m:f><m:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><m:t>+\u2026</m:t></m:r></m:oMath></m:oMathPara>';
const EXPECTED = '\\left(1+x\\right)^{n}=1+\\frac{nx}{1!}+\\frac{n\\left(n-1\\right)x^{2}}{2!}+\\ldots';

const texts = (xml) => [...xml.matchAll(/<m:t(?:\s[^>]*)?>([^<]*)<\/m:t>/g)].map((m) => m[1]);

test('the binomial in Equation_Test.docx becomes Office LaTeX and writes back the same dialect', () => {
  const latex = L.omml.toLatex(BINOM);
  assert.equal(latex, EXPECTED);
  assert.ok(L.latex.toMathML(latex));
  const xml = L.latex.toOMML(latex, { para: true });
  assert.match(xml, /^<m:oMathPara><m:oMath>/);
  assert.equal(L.omml.toLatex(xml), EXPECTED);
  assert.deepEqual(texts(xml), ['1+x', 'n', '=1+', 'nx', '1!', '+', 'n', 'n-1', 'x', '2', '2!', '+\u2026']);
  assert.match(xml, /<m:sSup>/);
  assert.match(xml, /<m:d>/);
  assert.match(xml, /<m:f>/);
  assert.match(xml, /w:ascii="Cambria Math"/);
  assert.equal(xml.includes('<m:begChr'), false);
});

test('roots, subscripts, sums, integrals and matrices', () => {
  const sqrt = L.latex.toOMML('\\sqrt{x}+\\sqrt[3]{y}');
  assert.match(sqrt, /<m:degHide m:val="1"\/>/);
  assert.match(sqrt, /<m:deg>[\s\S]*<m:t>3<\/m:t>/);
  assert.equal(L.omml.toLatex(sqrt), '\\sqrt{x}+\\sqrt[3]{y}');

  const sub = L.latex.toOMML('x_{i}^{2}');
  assert.match(sub, /<m:sSubSup>/);
  assert.equal(L.omml.toLatex(sub), 'x_{i}^{2}');

  const sum = L.latex.toOMML('\\sum_{i=1}^{n} x');
  assert.match(sum, /<m:chr m:val="∑"\/>/);
  assert.equal(L.omml.toLatex(sum), '\\sum_{i=1}^{n} x');

  const integ = L.latex.toOMML('\\int_0^1 x');
  assert.match(integ, /<m:chr m:val="∫"\/>/);
  assert.equal(L.omml.toLatex(integ), '\\int_{0}^{1} x');

  const mat = L.latex.toOMML('\\matrix{a&b\\\\c&d}');
  assert.equal([...mat.matchAll(/<m:mr>/g)].length, 2);
  assert.equal(L.omml.toLatex(mat), '\\matrix{a&b\\\\c&d}');
});

test('linear LaTeX is kept and UnicodeMath a/b becomes a fraction', () => {
  const linear = '<m:oMath><m:r><m:t>x=\\frac{1}{2}</m:t></m:r></m:oMath>';
  assert.equal(L.omml.toLatex(linear), 'x=\\frac{1}{2}');
  const slash = '<m:oMath><m:r><m:t>a/b</m:t></m:r></m:oMath>';
  assert.equal(L.omml.toLatex(slash), '\\frac{a}{b}');
  const plain = '<m:oMath><m:r><m:t>1+x</m:t></m:r></m:oMath>';
  assert.equal(L.omml.toLatex(plain), '1+x');
});

test('LaTeX that does not parse does not produce OMML', () => {
  assert.equal(L.latex.toOMML('\\frac{'), null);
  assert.equal(L.latex.toOMML('x^'), null);
  assert.equal(L.latex.toOMML(''), null);
  assert.equal(L.latex.toMathML('x^'), null);
  assert.equal(L.latex.toOMML('x^2', { para: false }).includes('oMathPara'), false);
});

test('Change to Display wraps an inline equation and Change to Inline unwraps it', () => {
  const inline = '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>a/b</m:t></m:r></m:oMath>';
  const display = L.omml.asDisplay(inline);
  assert.match(display, /^<m:oMathPara xmlns:m="http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/math"><m:oMath>/);
  assert.match(display, /<\/m:oMath><\/m:oMathPara>$/);
  assert.equal(L.omml.asInline(display), inline);
  assert.equal(L.omml.asDisplay(display), display);
  const numbered = '<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMathParaPr><m:jc m:val="centerGroup"/></m:oMathParaPr><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></m:oMathPara>';
  assert.equal(L.omml.asInline(numbered), '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>x</m:t></m:r></m:oMath>');
  assert.equal(L.omml.asInline(inline), inline);
});
