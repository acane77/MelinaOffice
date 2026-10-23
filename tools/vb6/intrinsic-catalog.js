/** Names the headless virtual machine installs in its library.
 *  Data only: this does not construct a machine or call a procedure.
 *  Kept beside createLibrary (third_party/VB6/src/runtime/library.js). */
import { BUILTIN_SIGNATURES, signatureParameters } from '../../third_party/VB6/src/runtime/signatures.js';
import { VB_CONSTANTS } from '../../third_party/VB6/src/runtime/constants.js';

const STRING_FORMS = ['Error', 'Left', 'Right', 'Mid', 'Trim', 'LTrim', 'RTrim', 'UCase', 'LCase', 'Space', 'String', 'Chr', 'ChrW', 'Str', 'Hex', 'Oct', 'Format', 'Input', 'Dir', 'Environ', 'Command'];
const VOID = new Set('Randomize Beep Kill Reset ChDir MkDir RmDir SetAttr FileCopy SaveSetting DeleteSetting'.split(' '));
const RETURNS = {
  CBool: 'Boolean', CByte: 'Byte', CInt: 'Integer', CLng: 'Long', CSng: 'Single', CDbl: 'Double', CCur: 'Currency', CDate: 'Date', CStr: 'String', CDec: 'Variant',
  MsgBox: 'VbMsgBoxResult', InputBox: 'String', Now: 'Date', Date: 'Date', Time: 'Date', DateSerial: 'Date', TimeSerial: 'Date', DateValue: 'Date', TimeValue: 'Date', DateAdd: 'Date', DateDiff: 'Long', DatePart: 'Integer', Timer: 'Single',
  RGB: 'ColorConstants', QBColor: 'ColorConstants', LBound: 'Long', UBound: 'Long', Len: 'Long', LenB: 'Long', Asc: 'Integer', AscW: 'Integer', InStr: 'Long', InStrRev: 'Long', StrComp: 'Integer', TypeName: 'String', VarType: 'Integer',
  CreateObject: 'Object', GetObject: 'Object', LoadResString: 'String', LoadResData: 'Byte', LoadResPicture: 'Object',
};
const TYPED = {
  Atn: 'Double', Cos: 'Double', Exp: 'Double', Log: 'Double', Sin: 'Double', Sqr: 'Double', Tan: 'Double', Sgn: 'Integer', Rnd: 'Single', Val: 'Double',
  Year: 'Integer', Month: 'Integer', Day: 'Integer', Hour: 'Integer', Minute: 'Integer', Second: 'Integer', Weekday: 'VbDayOfWeek', MonthName: 'String', WeekdayName: 'String',
  StrReverse: 'String', Join: 'String', Replace: 'String', Split: 'String', Filter: 'String', FormatNumber: 'String', FormatCurrency: 'String', FormatPercent: 'String', FormatDateTime: 'String',
  EOF: 'Boolean', LOF: 'Long', Loc: 'Long', Seek: 'Long', FileLen: 'Long', FileDateTime: 'Date', FileAttr: 'Long', GetAttr: 'Integer', FreeFile: 'Integer', CurDir: 'String', Dir: 'String', Environ: 'String', Command: 'String', GetSetting: 'String',
  DoEvents: 'Integer', Erl: 'Long', DDB: 'Double', FV: 'Double', IPmt: 'Double', IRR: 'Double', MIRR: 'Double', NPer: 'Double', NPV: 'Double', Pmt: 'Double', PPmt: 'Double', PV: 'Double', Rate: 'Double', SLN: 'Double', SYD: 'Double',
};

function procedureType(name) {
  if (VOID.has(name)) return 'Void';
  if (TYPED[name]) return TYPED[name];
  if (RETURNS[name]) return RETURNS[name];
  if (/^Is/.test(name)) return 'Boolean';
  return 'Variant';
}

function procedure(name, args, type) {
  return { name: name, kind: 'procedure', type: type || procedureType(name), params: signatureParameters(args || '') };
}

const OBJECTS = [
  {
    name: 'Debug', kind: 'module', type: 'Debug',
    members: [{ name: 'Print', kind: 'procedure', params: [{ name: 'outputlist', optional: true }] }],
  },
  {
    name: 'Err', kind: 'module', type: 'ErrObject',
    members: [
      { name: 'Number', kind: 'member', type: 'Long' },
      { name: 'Description', kind: 'member', type: 'String' },
      { name: 'Source', kind: 'member', type: 'String' },
      { name: 'HelpFile', kind: 'member', type: 'String' },
      { name: 'HelpContext', kind: 'member', type: 'Long' },
      { name: 'LastDLLError', kind: 'member', type: 'Long' },
      { name: 'Clear', kind: 'procedure', params: [] },
      { name: 'Raise', kind: 'procedure', params: [{ name: 'number' }, { name: 'source', optional: true }, { name: 'description', optional: true }, { name: 'helpfile', optional: true }, { name: 'helpcontext', optional: true }] },
    ],
  },
];

export function intrinsicCatalog() {
  const procedures = [];
  for (const name of Object.keys(BUILTIN_SIGNATURES)) procedures.push(procedure(name, BUILTIN_SIGNATURES[name]));
  if (!procedures.some(item => item.name === 'CallByName')) {
    procedures.push({
      name: 'CallByName', kind: 'procedure', type: 'Variant', variadic: true,
      params: [{ name: 'object' }, { name: 'procname' }, { name: 'calltype' }],
    });
  }
  for (const name of STRING_FORMS) {
    const base = procedures.find(item => item.name === name);
    if (base) procedures.push({ name: name + '$', kind: 'procedure', type: 'String', params: base.params });
  }
  const constants = [];
  for (const name of Object.keys(VB_CONSTANTS)) {
    constants.push({ name: name, kind: 'constant', type: typeof VB_CONSTANTS[name] === 'string' ? 'String' : 'Long' });
  }
  return { procedures: procedures, constants: constants, objects: OBJECTS };
}
