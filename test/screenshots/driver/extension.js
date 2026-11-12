// ドキュメント用スクリーンショットの自動撮影（テスト専用。撮影用の VS Code にだけ入れる）
// チュートリアル（docs/tutorial-basic.html・intermediate・advanced）の手順どおりに
// 練習用サンプル（カフェ こもれび）を操作して、場面ごとに撮る。
// ワークスペース直下の .shot.json に { out, capture, marker, cli } を置くと動く。
'use strict';
const vscode = require('vscode');
const cp = require('child_process');
const path = require('path');
const fs = require('fs');

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
let cfg;
let root;

function log(m) {
  fs.appendFileSync(path.join(cfg.out, 'log.txt'), `${new Date().toISOString()} ${m}\n`);
}

function run(file, args) {
  return new Promise((resolve) => {
    const p = cp.spawn('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', file, ...args], { windowsHide: true, cwd: root });
    let out = '';
    p.stdout.on('data', (d) => (out += d));
    p.stderr.on('data', (d) => (out += d));
    p.on('close', () => resolve(out));
  });
}

async function cap(name, wait = 1500) {
  await sleep(wait);
  const r = await run(cfg.capture, ['-Out', path.join(cfg.out, `${name}.png`), '-Marker', cfg.marker]);
  log(`${name}: ${r.trim()}`);
}

const ex = (c, ...a) =>
  Promise.resolve(vscode.commands.executeCommand(c, ...a)).catch((e) => {
    log(`command ${c} failed: ${e && e.message}`);
  });
const src = (...p) => path.join(root, 'cafe.accdb.src', ...p);
const db = () => path.join(root, 'cafe.accdb');
const crlf = (s) => s.replace(/\r?\n/g, '\r\n');

async function openFile(p) {
  const doc = await vscode.workspace.openTextDocument(p);
  return vscode.window.showTextDocument(doc, { preview: false });
}

// ファイル全体を書き換えて保存（エディタ経由なので「保存」の動きも再現される）
async function rewrite(p, fn) {
  const ed = await openFile(p);
  const before = ed.document.getText();
  const after = fn(before);
  if (after === before) throw new Error(`no change: ${path.basename(p)}`);
  const all = new vscode.Range(ed.document.positionAt(0), ed.document.positionAt(before.length));
  await ed.edit((e) => e.replace(all, after));
  await ed.document.save();
  return ed;
}

// 指定した文字のある行を画面の中央に表示して選択する
function reveal(ed, needle) {
  const i = ed.document.getText().indexOf(needle);
  if (i < 0) return;
  const a = ed.document.positionAt(i);
  const b = ed.document.positionAt(i + needle.length);
  ed.selection = new vscode.Selection(a, b);
  ed.revealRange(new vscode.Range(a, b), vscode.TextEditorRevealType.InCenter);
}

async function reset() {
  await ex('notifications.clearAll');
  for (const doc of vscode.workspace.textDocuments.filter((d) => d.isUntitled)) {
    await vscode.window.showTextDocument(doc, { preview: false });
    await ex('workbench.action.revertAndCloseActiveEditor');
  }
  await ex('workbench.action.closeAllEditors');
  await ex('workbench.action.closePanel');
  await ex('workbench.view.extension.accessVsc');
}

const TOOLS = crlf(`Option Compare Database
Option Explicit

' スタンプカードのメッセージ（10 個でドリンク 1 杯無料）
Public Function StampMessage(ByVal stamps As Long) As String
    If stamps >= 10 Then
        StampMessage = "おめでとうございます！ドリンク 1 杯無料です"
    Else
        StampMessage = "あと " & (10 - stamps) & " 個でドリンク 1 杯無料！"
    End If
End Function
`);
const SPLIT_BAD = crlf(`
' 割り勘（1 人分。1 円未満は切り上げ）
Public Function SplitBill(ByVal total As Currency, ByVal people As Long) As Currency
    SplitBill = -Int(-total / people)
End Function
`);
const SPLIT_GOOD = crlf(`Public Function SplitBill(ByVal total As Currency, ByVal people As Long) As Currency
    If people <= 0 Then
        SplitBill = 0
        Exit Function
    End If
    SplitBill = -Int(-total / people)
End Function`);
const RANKING = crlf(`SELECT p.ProductName, Sum(d.Qty) AS TotalQty
FROM OrderDetails AS d
  INNER JOIN Products AS p ON d.ProductID = p.ID
GROUP BY p.ProductName
ORDER BY Sum(d.Qty) DESC;
`);
const TAXLINE = crlf(`

' レシートの税率ごとの内訳（インボイス用）。例: 8%対象 960円（消費税 76円）
Public Function TaxLine(ByVal orderId As Long, ByVal takeout As Boolean) As String
    Dim amount As Currency
    amount = Nz(DSum("Amount", "qrySalesDetail", "OrderID=" & orderId & " AND Takeout=" & takeout), 0)
    TaxLine = Format(TaxRate(takeout), "0%") & "対象 " & Format(amount, "#,##0") & "円（消費税 " & Format(Int(amount * TaxRate(takeout)), "#,##0") & "円）"
End Function
`);
const FORM_EVENTS = crlf(`

Private Sub cboProduct_AfterUpdate()
    UpdateTotals
End Sub

Private Sub txtQty_AfterUpdate()
    UpdateTotals
End Sub

Private Sub chkTakeout_AfterUpdate()
    UpdateTotals
End Sub
`);
const TEST_CAFE = crlf(`Option Compare Database
Option Explicit

Public Sub Test_TaxIncluded_EatIn()
    Assert.AreEqual 550, TaxIncluded(500, False)
End Sub

Public Sub Test_TaxIncluded_Takeout()
    Assert.AreEqual 540, TaxIncluded(500, True)
End Sub

Public Sub Test_TaxIncluded_RoundDown()
    Assert.AreEqual 106, TaxIncluded(99, True), "1 円未満は切り捨て（106.92 → 106）"
End Sub

' 注文 No.1：店内 1,530 円（10%）＋ 持ち帰り 960 円（8%）
Public Sub Test_ReceiptTotal_Mixed()
    Assert.AreEqual 2719, ReceiptTotal(1)
End Sub

' 注文 No.4：店内だけ 930 円
Public Sub Test_ReceiptTotal_EatInOnly()
    Assert.AreEqual 1023, ReceiptTotal(4)
End Sub
`);
const RECEIPT_TOTAL = crlf(`

' レシートの合計（税込）。税率ごとに合計してから、それぞれ 1 回だけ端数を切り捨てる
Public Function ReceiptTotal(ByVal orderId As Long) As Currency
    Dim reduced As Currency
    Dim standard As Currency
    reduced = Nz(DSum("Amount", "qrySalesDetail", "OrderID=" & orderId & " AND Takeout=True"), 0)
    standard = Nz(DSum("Amount", "qrySalesDetail", "OrderID=" & orderId & " AND Takeout=False"), 0)
    ReceiptTotal = reduced + standard + Int(reduced * 0.08) + Int(standard * 0.1)
End Function
`);
const TEST_FORM = crlf(`Option Compare Database
Option Explicit

' 店内は 10%：カフェラテ 500 円 × 2 = 1,000 円 → 1,100 円
Public Sub Test_EatIn_UsesStandardRate()
    UI.OpenForm "frmOrder"
    UI.TypeText "cboProduct", "カフェラテ"
    UI.TypeText "txtQty", "2"
    UI.Click "cmdCalc"
    Assert.AreEqual 1100, UI.Value("txtTotal")
End Sub

' テイクアウトは軽減税率 8%：1,000 円 → 1,080 円
Public Sub Test_Takeout_UsesReducedRate()
    UI.OpenForm "frmOrder"
    UI.TypeText "cboProduct", "カフェラテ"
    UI.TypeText "txtQty", "2"
    UI.Click "chkTakeout"
    UI.Click "cmdCalc"
    Assert.AreEqual 1080, UI.Value("txtTotal")
End Sub

' ボタンを押すと、今日のラッキーメニューが出る
Public Sub Test_LuckyMenu()
    UI.OpenForm "frmOrder"
    UI.Click "cmdPick"
    Assert.IsTrue UI.Form.lblPick.Caption Like "【*】今日のラッキーメニューは「*」！"
End Sub

' 各テストの後でフォームを閉じる
Public Sub TestTeardown()
    UI.CloseForm "frmOrder"
End Sub
`);

const BUG = 'TaxIncluded(Me.txtSubtotal, False)';
const FIX = 'TaxIncluded(Me.txtSubtotal, Nz(Me.chkTakeout, False))';
const reportArg = (name) => ({ objName: name, kind: 'reports', resourceUri: vscode.Uri.file(src('reports', `${name}.txt`)) });
const HEIGHT_RE = /(\r\n {12}ForceNewPage =2\r\n {12}Height =)950/;

// レシートに内訳 2 行と登録番号を足す（txtSubtotal と lblTitle のかたまりを複製）
function addInvoiceLines(t) {
  const block = (name, kind) => {
    const re = new RegExp(`^ {16}Begin ${kind}\\r\\n(?:(?!^ {16}End\\r\\n)[\\s\\S])*?Name ="${name}"[\\s\\S]*?^ {16}End\\r\\n`, 'm');
    const m = re.exec(t);
    return m ? m[0] : '';
  };
  const noGuid = (b) => b.replace(/^\s+GUID = Begin\r\n\s+0x[0-9a-f]+\r\n\s+End\r\n/m, '');
  const sub = block('txtSubtotal', 'TextBox');
  const title = block('lblTitle', 'Label');
  const b8 = noGuid(sub).replace('Name ="txtSubtotal"', 'Name ="txtTax8"').replace('ControlSource ="=Sum([Amount])"', 'ControlSource ="=TaxLine([OrderID],True)"').replace('Left =2600', 'Left =100').replace('Width =2100', 'Width =4600').replace('Top =100', 'Top =900').replace('TextAlign =3', 'TextAlign =1');
  const b10 = b8.replace('Name ="txtTax8"', 'Name ="txtTax10"').replace('TaxLine([OrderID],True)', 'TaxLine([OrderID],False)').replace('Top =900', 'Top =1250');
  const reg = noGuid(title).replace('Name ="lblTitle"', 'Name ="lblRegNo"').replace('Caption ="領収書"', 'Caption ="登録番号 T1234567890123"').replace('Top =600', 'Top =1650');
  return t.replace(sub, sub + b8 + b10 + reg);
}

async function steps() {
  // ================= 初級編 =================
  await reset();
  await cap('b01-first', 3000);

  await ex('accessVsc.export');
  await cap('b02-exported', 2500);

  await reset();
  const cafe = await openFile(src('modules', 'modCafe.bas'));
  reveal(cafe, 'Public Function PickOfTheDay');
  await ex('accessVsc.runProcedure', 'PickOfTheDay', 'Optional ByVal d As Date = 0', ['2026/04/01']);
  await cap('b03-lucky', 2000);

  await reset();
  fs.writeFileSync(src('modules', 'modTools.bas'), TOOLS);
  await openFile(src('modules', 'modTools.bas'));
  await ex('accessVsc.refresh');
  await cap('b04-new-module', 2500);
  await ex('accessVsc.import');
  await ex('notifications.clearAll');
  await ex('accessVsc.runProcedure', 'StampMessage', 'ByVal stamps As Long', [7]);
  await cap('b05-stamp-result', 2000);

  await reset();
  await rewrite(src('modules', 'modTools.bas'), (t) => t.replace('(10 - stamps)', '(10 - stanps)'));
  await ex('accessVsc.import');
  await cap('b06-compile-error', 2500);
  await rewrite(src('modules', 'modTools.bas'), (t) => t.replace('(10 - stanps)', '(10 - stamps)') + SPLIT_BAD);
  await ex('accessVsc.import');
  await ex('notifications.clearAll');
  await ex('accessVsc.runProcedure', 'SplitBill', 'ByVal total As Currency, ByVal people As Long', [2719, 0]);
  await ex('workbench.action.closePanel');
  await cap('b07-runtime-error', 2500);
  await rewrite(src('modules', 'modTools.bas'), (t) => t.replace(/Public Function SplitBill[\s\S]*?End Function/, SPLIT_GOOD));
  await ex('accessVsc.import');

  await reset();
  fs.writeFileSync(src('queries', 'qryRanking.sql'), RANKING);
  await ex('accessVsc.import');
  await ex('accessVsc.runSql', { objName: 'qryRanking', kind: 'queries' });
  await sleep(1000);
  await ex('markdown.showPreview');
  await cap('b08-ranking', 2500);

  await reset();
  await ex('accessVsc.previewReport', reportArg('rptMonthlySales'));
  await ex('notifications.clearAll');
  await cap('b09-monthly', 3000);

  // ================= 中級編 =================
  const formFile = src('forms', 'frmOrder.txt');
  const formOriginal = fs.readFileSync(formFile, 'utf8');
  await reset();
  const fed = await openFile(formFile);
  reveal(fed, BUG);
  await cap('m01-form-code', 2500);

  await rewrite(formFile, (t) => {
    let s = t.replace(BUG, FIX);
    for (const n of ['cboProduct', 'txtQty', 'chkTakeout']) s = s.replace(new RegExp(`(\\r\\n\\s+)Name ="${n}"`), `$1Name ="${n}"$1AfterUpdate ="[Event Procedure]"`);
    return s.replace(/\s*$/, '') + FORM_EVENTS;
  });
  await ex('accessVsc.import');
  const origCopy = path.join(root, '.accessvsc', 'frmOrder-before.txt');
  fs.mkdirSync(path.dirname(origCopy), { recursive: true });
  fs.writeFileSync(origCopy, formOriginal);
  await reset();
  await ex('vscode.diff', vscode.Uri.file(origCopy), vscode.Uri.file(formFile), 'frmOrder.txt（改修前 ↔ 今）');
  await sleep(1500);
  await ex('workbench.action.compareEditor.nextChange');
  await cap('m02-diff', 2500);

  await reset();
  await rewrite(src('modules', 'modCafe.bas'), (t) => t.replace(/\s*$/, '') + TAXLINE);
  await ex('accessVsc.import');
  await reset();
  const red = await rewrite(src('reports', 'rptReceipt.txt'), (t) => addInvoiceLines(t));
  reveal(red, 'Name ="txtTax8"');
  await ex('workbench.actions.view.problems');
  await cap('m03-height-warning', 4000);
  await rewrite(src('reports', 'rptReceipt.txt'), (t) => t.replace(HEIGHT_RE, '$12100'));
  await ex('accessVsc.import');
  await reset();
  await ex('accessVsc.previewReport', reportArg('rptReceipt'));
  await ex('notifications.clearAll');
  await cap('m04-receipt', 3000);

  await reset();
  const ag = ex('accessVsc.setupAgent');
  await cap('m05-setup-agent', 1500);
  await ex('workbench.action.closeQuickOpen');
  await ag;

  // ================= 上級編 =================
  // 不具合がある状態から始めるため、フォームを元に戻す
  fs.writeFileSync(formFile, formOriginal);
  await ex('accessVsc.import');
  await reset();
  await run(cfg.cli, ['setup-ci', '-DbPath', db(), '-GitInit', '-Compress']);
  await ex('accessVsc.import');
  await openFile(path.join(root, '.github', 'workflows', 'access-ci.yml'));
  await ex('workbench.view.explorer');
  await ex('workbench.files.action.showActiveFileInExplorer');
  await cap('a01-setup-ci', 2500);

  await reset();
  fs.writeFileSync(src('modules', 'Test_Cafe.bas'), TEST_CAFE);
  await openFile(src('modules', 'Test_Cafe.bas'));
  await ex('accessVsc.import');
  await cap('a02-tdd-red', 2500);
  await rewrite(src('modules', 'modCafe.bas'), (t) => t.replace(/\s*$/, '') + RECEIPT_TOTAL);
  await ex('accessVsc.import');

  await reset();
  fs.writeFileSync(src('modules', 'Test_OrderForm.bas'), TEST_FORM);
  await ex('accessVsc.import');
  await openFile(src('modules', 'Test_OrderForm.bas'));
  await ex('accessVsc.runTests');
  await sleep(1500);
  await ex('testing.showMostRecentOutput');
  await cap('a03-ui-test', 2500);
  await rewrite(formFile, (t) => t.replace(BUG, FIX));
  await ex('accessVsc.import');

  await reset();
  await run(cfg.cli, ['snapshot', '-DbPath', db(), '-Update', '-Compress']);
  await rewrite(src('reports', 'rptReceipt.txt'), (t) => t.replace('Caption ="カフェ こもれび"', 'Caption ="カフェ こもれび 駅前店"'));
  await ex('accessVsc.import');
  await ex('workbench.action.closeAllEditors');
  await ex('accessVsc.snapshot');
  await ex('notifications.clearAll');
  await cap('a04-snapshot', 3000);

  // ================= リファレンス用 =================
  await reset();
  fs.writeFileSync(src('modules', 'modLintDemo.bas'), crlf('Option Compare Database\n\nPublic Function Smile() As String\n    Smile = "スマイル ☕"\nEnd Function\n'));
  const led = await openFile(src('modules', 'modLintDemo.bas'));
  await led.edit((e) => e.insert(led.document.lineAt(led.document.lineCount - 1).range.end, '\r\n'));
  await led.document.save();
  await ex('workbench.actions.view.problems');
  await cap('r-lint', 4000);
  await ex('workbench.action.closeAllEditors');
  fs.unlinkSync(src('modules', 'modLintDemo.bas'));

  await reset();
  const mn = ex('accessVsc.menu');
  await cap('r-menu', 1500);
  await ex('workbench.action.closeQuickOpen');
  await mn;
  await reset();
  await cap('r-sidebar', 2000);
}

async function activate() {
  const f = vscode.workspace.workspaceFolders;
  if (!f || !f.length) return;
  root = f[0].uri.fsPath;
  const cfgPath = path.join(root, '.shot.json');
  if (!fs.existsSync(cfgPath)) return;
  cfg = JSON.parse(fs.readFileSync(cfgPath, 'utf8').replace(/^﻿/, ''));
  setTimeout(async () => {
    try {
      log('start');
      await steps();
      log('done');
    } catch (e) {
      log(`ERROR ${e && e.stack ? e.stack : e}`);
    } finally {
      fs.writeFileSync(path.join(cfg.out, 'done.txt'), 'ok');
    }
  }, 6000);
}

module.exports = { activate, deactivate() {} };
