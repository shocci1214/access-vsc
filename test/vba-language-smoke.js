// vba-language.js の解析部分の単体テスト（VS Code なしで動く）
// 実行: set ELECTRON_RUN_AS_NODE=1 && Code.exe test/vba-language-smoke.js
'use strict';
const Module = require('module');
const path = require('path');
const origLoad = Module._load;
Module._load = function (req, ...rest) {
  if (req === 'vscode') return {};
  return origLoad.call(this, req, ...rest);
};
const L = require(path.join(__dirname, '..', 'vba-language.js'));

let failures = 0;
const check = (name, cond, detail) => {
  console.log(`${cond ? 'PASS' : 'FAIL'} ${name}${cond || detail === undefined ? '' : ` -> ${JSON.stringify(detail)}`}`);
  if (!cond) failures++;
};

const mod = [
  'Option Compare Database',
  'Option Explicit',
  '',
  'Public Const TAX_RATE As Double = 0.1',
  'Private m_count As Long',
  'Public Type 住所',
  '    郵便番号 As String',
  '    市区町村 As String',
  'End Type',
  'Public Enum Color',
  '    Red = 1',
  'End Enum',
  'Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long',
  '',
  "' 税込金額を返す",
  "' 1 円未満は切り捨て",
  'Public Function 税込金額(ByVal price As Currency, _',
  '                         Optional ByVal rate As Double = 0.1) As Currency',
  "    税込金額 = Int(price * (1 + rate)) ' コメント",
  'End Function',
  '',
  'Private Sub Helper()',
  '    Dim s As String: s = "Function Fake() ' + "'" + 'not a comment"',
  'End Sub',
  '',
  'Property Get Count() As Long',
  '    Count = m_count',
  'End Property',
].join('\r\n');

const p = L.parseText(mod, 'C:\\x\\shop.accdb.src\\modules\\modTax.bas');
check('module name', p.module === 'modTax', p.module);
check('proc count', p.procs.length === 4, p.procs.map((x) => x.name));
const tax = p.procs.find((x) => x.name === '税込金額');
check('japanese proc', !!tax);
check('continuation params', tax && /Optional ByVal rate As Double = 0\.1/.test(tax.params), tax && tax.params);
check('return type', tax && tax.returnType === 'Currency', tax && tax.returnType);
check('doc comment', tax && tax.doc === '税込金額を返す\n1 円未満は切り捨て', tax && tax.doc);
check('proc range', tax && tax.startLine === 16 && tax.endLine === 19, tax && [tax.startLine, tax.endLine]);
check('private scope', p.procs.find((x) => x.name === 'Helper').scope === 'Private');
check('property', p.procs.find((x) => x.name === 'Count').kind === 'Property Get');
check('declare', p.procs.some((x) => x.name === 'GetTickCount' && /Declare/.test(x.kind)));
check('string with Function inside is not a proc', !p.procs.some((x) => x.name === 'Fake'));
const names = p.decls.map((d) => `${d.kind}:${d.name}`);
check('decls', names.includes('Const:TAX_RATE') && names.includes('Variable:m_count') && names.includes('Type:住所') && names.includes('Enum:Color'), names);
check('type members', p.decls.find((d) => d.name === '住所').members.map((m) => m.name).join(',') === '郵便番号,市区町村');

const form = ['Version =21', 'Begin Form', '    Caption ="顧客"', 'End', 'CodeBehindForm', 'Attribute VB_GlobalNameSpace = False', 'Option Compare Database', '', 'Private Sub cmdOK_Click()', '    MsgBox "OK"', 'End Sub'].join('\r\n');
const f = L.parseText(form, 'C:\\x\\shop.accdb.src\\forms\\frmCustomers.txt');
check('form module name', f.module === 'Form_frmCustomers', f.module);
check('form code start', f.codeStart === 5, f.codeStart);
check('form proc', f.procs.length === 1 && f.procs[0].name === 'cmdOK_Click' && f.procs[0].startLine === 8, f.procs);

check('stripLine comment', L.stripLine("x = 1 ' c") === 'x = 1    ');
check('stripLine string', L.stripLine('s = "a\'b"') === 's = "   "');
check('splitParams', JSON.stringify(L.splitParams('a As Long, Optional b(1 To 2) As String')) === JSON.stringify(['a As Long', 'Optional b(1 To 2) As String']));
check('wordRe japanese', L.wordRe('税込金額', 'i').test('x = 税込金額(100)') && !L.wordRe('税込', 'i').test('税込金額'));
check('wordRe ascii boundary', L.wordRe('Tax', 'i').test('a = tax + 1') && !L.wordRe('Tax', 'i').test('TaxRate'));

process.exitCode = failures ? 1 : 0;
console.log(failures ? `${failures} failure(s)` : 'all passed');
