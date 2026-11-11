// vba-format.js のテスト（VS Code 不要）: node test/format-smoke.js
'use strict';
const fs = require('fs');
const path = require('path');
const { formatVba } = require('../vba-format');
const dir = path.join(__dirname, 'format');
let fail = 0;
const ok = (c, name) => { console.log(`${c ? 'PASS' : 'FAIL'} ${name}`); if (!c) fail++; };
const read = (f) => fs.readFileSync(path.join(dir, f), 'utf8');

const input = read('basic.in.bas');
const expected = read('basic.out.bas');
const got = formatVba(input);
ok(got === expected, 'basic indentation');
if (got !== expected && got) {
  const a = got.split('\n'); const b = expected.split('\n');
  for (let i = 0; i < Math.max(a.length, b.length); i++) if (a[i] !== b[i]) { console.log(`  line ${i + 1}\n  got: ${JSON.stringify(a[i])}\n  exp: ${JSON.stringify(b[i])}`); break; }
}
ok(formatVba(expected) === expected, 'idempotent');
ok(formatVba(input.replace(/\n/g, '\r\n')) === expected.replace(/\n/g, '\r\n'), 'keeps CRLF');
ok(formatVba(read('unbalanced.in.bas')) === null, 'unbalanced blocks are left alone (null)');
ok(formatVba('Sub A()\nx = 1\nEnd Sub\n', { indent: '\t' }) === 'Sub A()\n\tx = 1\nEnd Sub\n', 'tab indent');
const form = 'Version =21\nBegin Form\n    Caption ="x"\nEnd\nCodeBehindForm\nAttribute VB_GlobalNameSpace = False\nPrivate Sub cmdOK_Click()\nDoCmd.Close\nEnd Sub\n';
ok(formatVba(form, { file: 'forms/frm.txt' }) === form.replace('\nDoCmd.Close', '\n    DoCmd.Close'), 'form: only the code part');
const cls = 'VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1  \'True\nEND\nAttribute VB_Name = "C"\nPublic Property Get Item() As Long\nAttribute Item.VB_UserMemId = 0\nItem = 1\nEnd Property\n';
ok(formatVba(cls) === cls.replace('\nItem = 1', '\n    Item = 1'), 'class header and Attribute lines are kept');
ok(formatVba('Sub A()\nDoCmd.OpenForm "f"\nFormat$(1)\nEnd Sub\n') === 'Sub A()\n    DoCmd.OpenForm "f"\n    Format$(1)\nEnd Sub\n', 'DoCmd / Format are not Do / For');
ok(formatVba('Sub A()\nFor i = 1 To 2\nFor j = 1 To 2\nNext j, i\nEnd Sub\n') === 'Sub A()\n    For i = 1 To 2\n        For j = 1 To 2\n    Next j, i\nEnd Sub\n', 'Next j, i closes two loops');
console.log(fail ? `${fail} failed` : 'all passed');
process.exit(fail ? 1 : 0);
