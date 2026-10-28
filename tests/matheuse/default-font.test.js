const test = require('node:test');
const assert = require('node:assert/strict');
require('../../public/common/core.js');
require('../../public/matheuse/js/model.js');

test('a workbook started with a font uses it for the theme heading and body fonts', () => {
  const wb = new L.model.Workbook({ font: { name: '等线', sz: 10 } });
  assert.equal(wb.styles.get(0).font.name, '等线');
  assert.deepEqual([wb.theme.major, wb.theme.minor, wb.theme.majorEA, wb.theme.minorEA], ['等线', '等线', '等线', '等线']);
  const plain = new L.model.Workbook();
  assert.equal(plain.theme.major, 'Cambria');
  assert.equal(plain.theme.minor, 'Calibri');
  assert.equal(plain.theme.majorEA, undefined);
});
