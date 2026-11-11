// Access VSC の結合テスト（本物の VS Code の中で、編集支援などの動きを確かめる）
// ワークスペース直下の .itest.json に { out } を置くと動き、out/results.json に結果を書く。
'use strict';
const vscode = require('vscode');
const fs = require('fs');
const path = require('path');

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function run(root, out) {
  const results = [];
  const check = (name, cond, detail) => results.push({ name, pass: !!cond, detail: cond ? undefined : detail });
  const src = (...p) => path.join(root, 'shop.accdb.src', ...p);
  const uri = (...p) => vscode.Uri.file(src(...p));
  const exec = (cmd, ...args) => vscode.commands.executeCommand(cmd, ...args);
  const posOf = async (u, needle, offset = 0) => {
    const doc = await vscode.workspace.openTextDocument(u);
    const i = doc.getText().indexOf(needle);
    if (i < 0) throw new Error(`not found: ${needle}`);
    return doc.positionAt(i + offset);
  };

  // 拡張機能が有効になるのを待つ
  const ext = vscode.extensions.getExtension('local.access-vscode');
  if (ext && !ext.isActive) await ext.activate();
  await sleep(3000);

  const tax = uri('modules', 'modTax.bas');
  const test = uri('modules', 'Test_Tax.bas');
  await vscode.window.showTextDocument(tax);

  // アウトライン
  const syms = (await exec('vscode.executeDocumentSymbolProvider', tax)) || [];
  const symNames = syms.map((s) => s.name);
  check('outline: procedures', symNames.includes('TaxIncluded') && symNames.includes('合計金額') && symNames.includes('TAX_RATE'), symNames);

  // ワークスペースのシンボル
  const ws = (await exec('vscode.executeWorkspaceSymbolProvider', 'Tax')) || [];
  check('workspace symbols', ws.some((s) => s.name === 'TaxIncluded' && s.containerName === 'modTax'), ws.map((s) => `${s.containerName}.${s.name}`));

  // 定義へ移動（別モジュールから / 日本語 / Module.Proc / ローカル変数）
  const p1 = await posOf(test, 'TaxIncluded(100)', 2);
  const d1 = (await exec('vscode.executeDefinitionProvider', test, p1)) || [];
  const d1loc = d1[0] && (d1[0].uri || d1[0].targetUri);
  check('definition: cross module', d1loc && /modTax\.bas$/i.test(d1loc.fsPath), d1.map((x) => (x.uri || x.targetUri).fsPath));
  const p2 = await posOf(test, '合計金額(', 1);
  const d2 = (await exec('vscode.executeDefinitionProvider', test, p2)) || [];
  check('definition: japanese name', d2.length && /modTax\.bas$/i.test((d2[0].uri || d2[0].targetUri).fsPath), d2.length);
  const p3 = await posOf(test, 'modTax.TaxIncluded', 8);
  const d3 = (await exec('vscode.executeDefinitionProvider', test, p3)) || [];
  check('definition: Module.Proc', d3.length === 1, d3.length);
  const p4 = await posOf(tax, 'subtotal * ', 2);
  const d4 = (await exec('vscode.executeDefinitionProvider', tax, p4)) || [];
  const d4line = d4[0] && (d4[0].range || d4[0].targetRange).start.line;
  const taxDoc = await vscode.workspace.openTextDocument(tax);
  check('definition: local variable', d4.length && /Dim subtotal/.test(taxDoc.lineAt(d4line).text), d4line);

  // 参照の検索（コメントと文字列は除く）
  const p5 = await posOf(tax, 'Function TaxIncluded', 9);
  const refs = (await exec('vscode.executeReferenceProvider', tax, p5)) || [];
  const refFiles = new Set(refs.map((r) => path.basename(r.uri.fsPath)));
  check('references: across files', refFiles.has('modTax.bas') && refFiles.has('Test_Tax.bas'), [...refFiles]);
  const testDoc = await vscode.workspace.openTextDocument(test);
  check('references: skip comments/strings', !refs.some((r) => r.uri.fsPath === test.fsPath && /'|"/.test(testDoc.lineAt(r.range.start.line).text.slice(0, r.range.start.character))), refs.length);

  // ホバー
  const hv = (await exec('vscode.executeHoverProvider', test, p1)) || [];
  const hvText = hv.map((h) => h.contents.map((c) => c.value || c).join(' ')).join(' ');
  check('hover: signature and doc comment', /TaxIncluded\(ByVal price As Currency\)/.test(hvText) && /税込金額/.test(hvText), hvText);
  const p6 = await posOf(tax, 'Nz(', 1);
  const hv2 = (await exec('vscode.executeHoverProvider', tax, p6)) || [];
  check('hover: builtin', hv2.length && /Null/.test(hv2.map((h) => h.contents.map((c) => c.value || c).join(' ')).join(' ')), hv2.length);

  // 補完（Module. のあと）
  const p7 = await posOf(test, 'modTax.', 7);
  const comp = await exec('vscode.executeCompletionItemProvider', test, p7, '.');
  const compLabels = (comp && comp.items ? comp.items : []).map((i) => (typeof i.label === 'string' ? i.label : i.label.label));
  check('completion: module members', compLabels.includes('TaxIncluded') && compLabels.includes('合計金額') && !compLabels.includes('PrivateHelper'), compLabels.slice(0, 20));

  // 引数のヒント
  const p8 = await posOf(test, 'TaxIncluded(100)', 'TaxIncluded('.length);
  const sh = await exec('vscode.executeSignatureHelpProvider', test, p8, '(');
  check('signature help', sh && sh.signatures && sh.signatures[0] && /price As Currency/.test(sh.signatures[0].label), sh && sh.signatures && sh.signatures.map((s) => s.label));

  // 折りたたみ
  const folds = (await exec('vscode.executeFoldingRangeProvider', tax)) || [];
  check('folding: procedures', folds.length >= 3, folds.length);

  // スニペット（package.json に登録されているか）
  const pj = ext && ext.packageJSON;
  check('snippets registered', pj && pj.contributes && Array.isArray(pj.contributes.snippets) && pj.contributes.snippets.length >= 1, pj && pj.contributes && pj.contributes.snippets);

  // 保存したときのコードチェック（暗黙の Variant）
  const lintFile = src('modules', 'modLintIt.bas');
  fs.writeFileSync(lintFile, 'Option Compare Database\r\nOption Explicit\r\n\r\nPublic Function Bad(x)\r\n    Dim unused As Long\r\n    Bad = x\r\nEnd Function\r\n');
  const ldoc = await vscode.workspace.openTextDocument(lintFile);
  const led = await vscode.window.showTextDocument(ldoc);
  await led.edit((e) => e.insert(new vscode.Position(ldoc.lineCount - 1, 0), '\r\n'));
  await ldoc.save();
  let diags = [];
  for (let i = 0; i < 30 && !diags.length; i++) {
    await sleep(1000);
    diags = vscode.languages.getDiagnostics(ldoc.uri).filter((d) => d.source === 'Access lint');
  }
  const codes = diags.map((d) => d.code);
  check('lint on save: implicit-variant / unused / return', codes.includes('implicit-variant') && codes.includes('unused-variable') && codes.includes('implicit-variant-return'), codes);
  fs.unlinkSync(lintFile);

  fs.writeFileSync(path.join(out, 'results.json'), JSON.stringify(results, null, 2));
}

async function activate() {
  const f = vscode.workspace.workspaceFolders;
  if (!f || !f.length) return;
  const root = f[0].uri.fsPath;
  const cfgPath = path.join(root, '.itest.json');
  if (!fs.existsSync(cfgPath)) return;
  const cfg = JSON.parse(fs.readFileSync(cfgPath, 'utf8').replace(/^﻿/, ''));
  setTimeout(async () => {
    try {
      await run(root, cfg.out);
    } catch (e) {
      fs.writeFileSync(path.join(cfg.out, 'error.txt'), String((e && e.stack) || e));
    } finally {
      fs.writeFileSync(path.join(cfg.out, 'done.txt'), 'ok');
    }
  }, 5000);
}

module.exports = { activate, deactivate() {} };
