// ドキュメント用スクリーンショットの自動撮影（テスト専用。撮影用の VS Code にだけ入れる）
// チュートリアル（docs/tutorial.html）の手順どおりに練習用サンプルを操作して、場面ごとに撮る。
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
const src = (...p) => path.join(root, 'shop.accdb.src', ...p);
const db = () => path.join(root, 'shop.accdb');

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

const TAX = 'Option Compare Database\r\nOption Explicit\r\n\r\n\' 税込金額（消費税 10%、1 円未満切り捨て）\r\nPublic Function TaxIncluded(ByVal price As Currency) As Currency\r\n    TaxIncluded = Int(price * 1.1)\r\nEnd Function\r\n';
const UNIT_BAD = '\r\n\' 単価（金額 ÷ 数量）\r\nPublic Function UnitPrice(ByVal amount As Currency, ByVal qty As Long) As Currency\r\n    UnitPrice = amount / qty\r\nEnd Function\r\n';
const UNIT_GOOD = 'Public Function UnitPrice(ByVal amount As Currency, ByVal qty As Long) As Currency\r\n    If qty = 0 Then\r\n        UnitPrice = 0\r\n        Exit Function\r\n    End If\r\n    UnitPrice = amount / qty\r\nEnd Function';
const QUERY = 'SELECT c.CustomerName, Sum(p.Price * o.Qty) AS Total\r\nFROM (Orders AS o\r\n  INNER JOIN Customers AS c ON o.CustomerID = c.ID)\r\n  INNER JOIN Products AS p ON o.ProductID = p.ID\r\nGROUP BY c.CustomerName\r\nORDER BY Sum(p.Price * o.Qty) DESC;\r\n';
const TESTS = 'Option Compare Database\r\nOption Explicit\r\n\r\nPublic Sub Test_TaxIncluded_100yen()\r\n    Assert.AreEqual 110, TaxIncluded(100)\r\nEnd Sub\r\n\r\nPublic Sub Test_TaxIncluded_RoundDown()\r\n    Assert.AreEqual 109, TaxIncluded(99), "1 円未満は切り捨て"\r\nEnd Sub\r\n\r\nPublic Sub Test_UnitPrice_ZeroQty()\r\n    Assert.AreEqual 0, UnitPrice(1000, 0)\r\nEnd Sub\r\n';
const report = () => ({ objName: 'rptOrders', kind: 'reports', resourceUri: vscode.Uri.file(src('reports', 'rptOrders.txt')) });

async function steps() {
  // 0. 最初の画面
  await reset();
  await cap('t01-first', 3000);

  // 1. 取り出す
  await ex('accessVsc.export');
  await cap('t02-exported', 2500);

  // 2. はじめての関数
  await reset();
  fs.writeFileSync(src('modules', 'modTax.bas'), TAX);
  await openFile(src('modules', 'modTax.bas'));
  await ex('accessVsc.refresh');
  await cap('t03-new-module', 2500);
  await ex('accessVsc.import');
  await cap('t04-imported', 2000);
  await ex('notifications.clearAll');
  const argsPrompt = ex('accessVsc.runProcedure', 'TaxIncluded', 'ByVal price As Currency');
  await cap('t05-run-args', 1500);
  await ex('workbench.action.closeQuickOpen');
  await argsPrompt;
  await ex('accessVsc.runProcedure', 'TaxIncluded', 'ByVal price As Currency', [100]);
  await cap('t06-run-result', 2000);

  // 3. わざと間違える（コンパイル エラー → 実行時エラー）
  await reset();
  await rewrite(src('modules', 'modTax.bas'), (t) => t.replace('Int(price * 1.1)', 'Int(prise * 1.1)'));
  await ex('accessVsc.import');
  await cap('t07-compile-error', 2500);
  await rewrite(src('modules', 'modTax.bas'), (t) => t.replace('Int(prise * 1.1)', 'Int(price * 1.1)') + UNIT_BAD);
  await ex('accessVsc.import');
  await ex('notifications.clearAll');
  await ex('accessVsc.runProcedure', 'UnitPrice', 'ByVal amount As Currency, ByVal qty As Long', [1000, 0]);
  await ex('workbench.action.closePanel');
  await cap('t08-runtime-error', 2500);
  await rewrite(src('modules', 'modTax.bas'), (t) => t.replace(/Public Function UnitPrice[\s\S]*?End Function/, UNIT_GOOD));
  await ex('accessVsc.import');

  // 4. クエリを作る
  await reset();
  fs.writeFileSync(src('queries', 'qryCustomerTotal.sql'), QUERY);
  await ex('accessVsc.import');
  await ex('accessVsc.runSql', { objName: 'qryCustomerTotal', kind: 'queries' });
  await sleep(1000);
  await ex('markdown.showPreview');
  await cap('t09-query-result', 2500);

  // 5. 帳票を直す
  await reset();
  await openFile(src('reports', 'rptOrders.txt'));
  await cap('t11-report-source', 2500);
  await rewrite(src('reports', 'rptOrders.txt'), (t) =>
    t.replace('Caption ="受注一覧"', 'Caption ="受注一覧（2026年4月）"').replace(/Width =1700((?:\r?\n(?!\s*(?:Begin|End)\b)[^\r\n]*){1,5}?\r?\n\s*Name ="(?:lbl|txt)ProductName")/g, 'Width =2400$1')
  );
  await ex('accessVsc.import');
  await ex('workbench.action.closeAllEditors');
  await ex('accessVsc.previewReport', report());
  await ex('notifications.clearAll');
  await cap('t12-report-overlap', 3000);
  await rewrite(src('reports', 'rptOrders.txt'), (t) => t.replace(/Left =5800/g, 'Left =6500').replace(/Left =6700/g, 'Left =7400').replace(/^(    Width =)8000(\r?)$/m, (m, a, b) => `${a}8800${b}`));
  await ex('accessVsc.import');
  await ex('workbench.action.closeAllEditors');
  await ex('accessVsc.previewReport', report());
  await ex('notifications.clearAll');
  await cap('t13-report-fixed', 3000);
  const ps = ex('accessVsc.pageSetup', report());
  await cap('t14-pagesetup', 2000);
  await ex('workbench.action.closeQuickOpen');
  await ps;

  // 6. テスト（1 つわざと失敗）
  await reset();
  await run(cfg.cli, ['setup-ci', '-DbPath', db(), '-GitInit', '-Compress']);
  fs.writeFileSync(src('modules', 'Test_Tax.bas'), TESTS);
  await ex('accessVsc.import');
  await openFile(src('modules', 'Test_Tax.bas'));
  await ex('accessVsc.runTests');
  await sleep(1500);
  await ex('testing.showMostRecentOutput');
  await cap('t15-tests', 2500);

  // コードチェック（リファレンス用）
  await reset();
  fs.writeFileSync(src('modules', 'modLintDemo.bas'), 'Option Compare Database\r\n\r\nPublic Function Smile() As String\r\n    Smile = "スマイル 😀"\r\nEnd Function\r\n');
  const led = await openFile(src('modules', 'modLintDemo.bas'));
  await led.edit((e) => e.insert(led.document.lineAt(led.document.lineCount - 1).range.end, '\r\n'));
  await led.document.save();
  await ex('workbench.actions.view.problems');
  await cap('t16-lint', 4000);
  await ex('workbench.action.closeAllEditors');
  fs.unlinkSync(src('modules', 'modLintDemo.bas'));

  // 8. 帳票の見た目の比較（基準を作ってから、タイトルを変える）
  await reset();
  await run(cfg.cli, ['snapshot', '-DbPath', db(), '-Update', '-Compress']);
  await rewrite(src('reports', 'rptOrders.txt'), (t) => t.replace('Caption ="受注一覧（2026年4月）"', 'Caption ="受注一覧（2026年4月分・確定）"'));
  await ex('accessVsc.import');
  await ex('workbench.action.closeAllEditors');
  await ex('accessVsc.snapshot');
  await ex('notifications.clearAll');
  await cap('t17-snapshot', 3000);

  // 7. AI 連携の設定（選択画面）
  await reset();
  const ag = ex('accessVsc.setupAgent');
  await cap('t18-setup-agent', 1500);
  await ex('workbench.action.closeQuickOpen');
  await ag;

  // メニュー・サイドバー全体（リファレンス用）
  const mn = ex('accessVsc.menu');
  await cap('t19-menu', 1500);
  await ex('workbench.action.closeQuickOpen');
  await mn;
  await reset();
  await cap('t20-sidebar', 2000);
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
