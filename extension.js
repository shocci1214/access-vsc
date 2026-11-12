// Access VSC - VS Code 拡張機能本体（依存パッケージなし）
// Access の操作は scripts/access.ps1（PowerShell + COM）に任せ、ここでは UI と AI 連携を担当する。
'use strict';

const vscode = require('vscode');
const cp = require('child_process');
const path = require('path');
const fs = require('fs');
const os = require('os');
const crypto = require('crypto');

const MANAGED = {
  modules: ['.bas', '.cls'],
  forms: ['.txt'],
  reports: ['.txt'],
  macros: ['.txt'],
  queries: ['.sql'],
  tables: ['.json'],
};
const READONLY_KINDS = new Set(['tables']);
const MANIFEST = '.accessvsc.json';
const AGENT_BEGIN = '<!-- access-vsc:begin -->';
const AGENT_END = '<!-- access-vsc:end -->';

let ctx;
let output;
let diagnostics;
let statusItem;
let queue = Promise.resolve();
let watcher;
let statusTimer;
let saveTimer;
const pendingSaves = new Set();
const mcpChanged = new vscode.EventEmitter();
const lensChanged = new vscode.EventEmitter();

// ------------------------------------------------------------
// 設定・パス
// ------------------------------------------------------------
function cfg() {
  return vscode.workspace.getConfiguration('accessVsc');
}

function workspaceRoot() {
  const f = vscode.workspace.workspaceFolders;
  return f && f.length ? f[0].uri.fsPath : undefined;
}

function binDir() {
  return path.join(process.env.LOCALAPPDATA || os.homedir(), 'access-vscode', 'bin');
}

const cliPath = () => path.join(binDir(), 'access.ps1');
const mcpServerPath = () => path.join(binDir(), 'mcp-server.ps1');
const powershell = () => cfg().get('powershellPath') || 'powershell.exe';

// スクリプトを固定の場所（%LOCALAPPDATA%\access-vscode\bin）へコピーする。
// 拡張機能を更新してもパスが変わらないので、.mcp.json や CLAUDE.md から参照できる。
function installScripts(from = path.join(ctx.extensionPath, 'scripts'), to = binDir()) {
  fs.mkdirSync(to, { recursive: true });
  for (const ent of fs.readdirSync(from, { withFileTypes: true })) {
    const src = path.join(from, ent.name);
    const dst = path.join(to, ent.name);
    if (ent.isDirectory()) {
      installScripts(src, dst); // templates/ など
      continue;
    }
    const data = fs.readFileSync(src);
    let same = false;
    try {
      same = fs.readFileSync(dst).equals(data);
    } catch (_) { /* not yet installed */ }
    if (!same) fs.writeFileSync(dst, data);
  }
}

function configuredDb() {
  const v = cfg().get('database');
  if (!v) return undefined;
  const root = workspaceRoot();
  return path.isAbsolute(v) ? v : root ? path.join(root, v) : v;
}

async function findDatabases() {
  const uris = await vscode.workspace.findFiles('**/*.{accdb,mdb}', '{**/.accessvsc/**,**/node_modules/**}', 50);
  return uris.map((u) => u.fsPath).sort();
}

async function getDb(interactive = true) {
  const configured = configuredDb();
  if (configured && fs.existsSync(configured)) return configured;
  const found = await findDatabases();
  if (found.length === 1) return found[0];
  if (!interactive) return undefined;
  let db;
  if (found.length === 0) {
    const pick = await vscode.window.showOpenDialog({
      filters: { 'Access データベース': ['accdb', 'mdb'] },
      canSelectMany: false,
      openLabel: 'このデータベースを使う',
    });
    if (!pick) return undefined;
    db = pick[0].fsPath;
  } else {
    const item = await vscode.window.showQuickPick(
      found.map((f) => ({ label: path.basename(f), description: vscode.workspace.asRelativePath(f), f })),
      { placeHolder: '使用する Access データベースを選択してください' }
    );
    if (!item) return undefined;
    db = item.f;
  }
  await saveDbSetting(db);
  return db;
}

async function saveDbSetting(db) {
  const root = workspaceRoot();
  const rel = root ? path.relative(root, db) : '';
  const value = root && rel && !rel.startsWith('..') && !path.isAbsolute(rel) ? rel : db;
  const target = root ? vscode.ConfigurationTarget.Workspace : vscode.ConfigurationTarget.Global;
  await cfg().update('database', value, target);
}

function srcFor(db) {
  const v = cfg().get('sourceFolder');
  if (!v) return db + '.src';
  return path.isAbsolute(v) ? v : path.join(path.dirname(db), v);
}

// ------------------------------------------------------------
// 同期状態（PowerShell を起動せずに JS で計算する）
// ------------------------------------------------------------
function sha256(file) {
  return crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex');
}

function relOf(src, file) {
  const rel = path.relative(src, file);
  if (!rel || rel.startsWith('..') || path.isAbsolute(rel)) return undefined;
  return rel.split(path.sep).join('/');
}

function kindOf(rel) {
  if (!rel) return undefined;
  if (rel === 'relationships.json') return { kind: 'relationships', readOnly: true };
  if (rel === 'references.json') return { kind: 'references', readOnly: true };
  const parts = rel.split('/');
  if (parts.length !== 2) return undefined;
  const exts = MANAGED[parts[0]];
  if (!exts || !exts.includes(path.extname(parts[1]).toLowerCase())) return undefined;
  return { kind: parts[0], readOnly: READONLY_KINDS.has(parts[0]), name: decodeName(path.basename(parts[1], path.extname(parts[1]))) };
}

function decodeName(s) {
  return s.replace(/%([0-9A-Fa-f]{2})/g, (_, h) => String.fromCharCode(parseInt(h, 16)));
}

function listSourceFiles(src) {
  const list = [];
  for (const kind of Object.keys(MANAGED)) {
    const dir = path.join(src, kind);
    if (!fs.existsSync(dir)) continue;
    for (const name of fs.readdirSync(dir)) {
      const full = path.join(dir, name);
      if (MANAGED[kind].includes(path.extname(name).toLowerCase()) && fs.statSync(full).isFile()) list.push(`${kind}/${name}`);
    }
  }
  if (fs.existsSync(path.join(src, 'relationships.json'))) list.push('relationships.json');
  if (fs.existsSync(path.join(src, 'references.json'))) list.push('references.json');
  return list;
}

function syncStatus(src) {
  const st = { hasManifest: false, changed: [], added: [], missing: [], readOnlyChanged: [] };
  let files = {};
  try {
    const m = JSON.parse(fs.readFileSync(path.join(src, MANIFEST), 'utf8').replace(/^﻿/, ''));
    files = m.files || {};
    st.hasManifest = true;
  } catch (_) { /* no manifest yet */ }
  const current = fs.existsSync(src) ? listSourceFiles(src) : [];
  for (const rel of current) {
    const k = kindOf(rel);
    const full = path.join(src, ...rel.split('/'));
    if (!files[rel]) (k.readOnly ? st.readOnlyChanged : st.added).push(rel);
    else if (sha256(full) !== files[rel].hash) (k.readOnly ? st.readOnlyChanged : st.changed).push(rel);
  }
  for (const rel of Object.keys(files)) {
    const k = kindOf(rel);
    if (k && !k.readOnly && !current.includes(rel)) st.missing.push(rel);
  }
  return st;
}

// ------------------------------------------------------------
// CLI 実行（1 度に 1 つずつ、順番に）
// ------------------------------------------------------------
// ------------------------------------------------------------
// 常駐プロセス（access.ps1 serve）：PowerShell の起動を省き、Access を開いたままにして使い回す。
//   1 行 1 リクエストの JSON を送り、1 行の JSON を受け取る。止まったら作り直す
// ------------------------------------------------------------
class AccessHost {
  constructor() {
    this.pending = new Map();
    this.seq = 0;
    this.buf = Buffer.alloc(0);
    this.dead = false;
    const env = Object.assign({}, process.env, { ACCESSVSC_SESSION_IDLE: String(cfg().get('keepAccessOpenSec') || 300) });
    this.child = cp.spawn(powershell(), ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', cliPath(), 'serve'], { windowsHide: true, env });
    this.child.stdout.on('data', (d) => this.onData(d));
    this.child.stderr.on('data', (d) => output.append(d.toString('utf8')));
    const fail = (why) => {
      this.dead = true;
      for (const [, p] of this.pending) {
        clearTimeout(p.timer);
        p.resolve({ ok: false, command: p.command, error: why });
      }
      this.pending.clear();
    };
    this.child.on('error', (e) => fail(String(e)));
    this.child.on('exit', () => fail('Access VSC の常駐プロセスが終了しました。もう一度実行してください。'));
  }
  onData(d) {
    this.buf = Buffer.concat([this.buf, d]);
    let i;
    while ((i = this.buf.indexOf(10)) >= 0) {
      const line = this.buf.slice(0, i).toString('utf8').trim();
      this.buf = this.buf.slice(i + 1);
      if (!line) continue;
      let m;
      try {
        m = JSON.parse(line);
      } catch (_) {
        output.appendLine(line);
        continue;
      }
      const p = this.pending.get(m.id);
      if (!p) continue;
      this.pending.delete(m.id);
      clearTimeout(p.timer);
      p.resolve(m.response || { ok: false, command: p.command, error: 'Empty response' });
    }
  }
  request(command, params, timeoutMs) {
    const id = ++this.seq;
    return new Promise((resolve) => {
      // 操作ごとの時間切れは常駐プロセスの中で処理する。これは常駐プロセス自体が固まったときの保険
      const timer = setTimeout(() => {
        this.pending.delete(id);
        resolve({ ok: false, command, error: '応答がないため、常駐プロセスを再起動します。もう一度実行してください。' });
        this.kill();
      }, timeoutMs);
      this.pending.set(id, { resolve, timer, command });
      this.child.stdin.write(JSON.stringify({ id, command, params }) + '\n');
    });
  }
  kill() {
    if (this.dead) return;
    this.dead = true;
    try { this.child.stdin.end(); } catch (_) { /* closed */ }
    const c = this.child;
    setTimeout(() => { try { c.kill(); } catch (_) { /* gone */ } }, 30000).unref();
  }
}
let accessHost;
function getHost() {
  if (!accessHost || accessHost.dead) accessHost = new AccessHost();
  return accessHost;
}
// 開いたままの Access を閉じてもらう（Access の画面で開く前など）
async function releaseAccessSession() {
  if (accessHost && !accessHost.dead) await accessHost.request('release', {}, 60000);
}

function access(command, params = {}, opts = {}) {
  const job = async () => {
    const db = opts.db || (await getDb(true));
    if (!db) return undefined;
    const p = Object.assign({}, params);
    const t = cfg().get('timeoutSec');
    if (t && !p.timeoutSec) p.timeoutSec = t;
    if (cfg().get('keepAccessOpen') !== false) {
      p.db = db;
      if (cfg().get('sourceFolder')) p.src = srcFor(db);
      const viaHost = async () => {
        output.appendLine(`\n[${new Date().toLocaleTimeString()}] ${command} ${Object.keys(params).length ? JSON.stringify(params) : ''}`);
        const res = await getHost().request(command, p, 60 * 60 * 1000);
        output.appendLine(JSON.stringify(res, null, 2));
        return res;
      };
      return vscode.window.withProgress(
        { location: opts.quiet ? vscode.ProgressLocation.Window : vscode.ProgressLocation.Notification, title: opts.title || `Access: ${command} を実行中...` },
        viaHost
      );
    }
    const tmp = path.join(os.tmpdir(), `accessvsc_${process.pid}_${Date.now()}_${Math.random().toString(36).slice(2)}.json`);
    fs.writeFileSync(tmp, JSON.stringify(p), 'utf8');
    const args = ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', cliPath(), command, '-DbPath', db, '-ParamsFile', tmp];
    if (cfg().get('sourceFolder')) args.push('-Src', srcFor(db));

    const exec = () =>
      new Promise((resolve) => {
        output.appendLine(`\n[${new Date().toLocaleTimeString()}] ${command} ${Object.keys(params).length ? JSON.stringify(params) : ''}`);
        const child = cp.spawn(powershell(), args, { cwd: path.dirname(db), windowsHide: true });
        const out = [];
        const err = [];
        child.stdout.on('data', (d) => out.push(d));
        child.stderr.on('data', (d) => err.push(d));
        child.on('error', (e) => resolve({ ok: false, command, error: String(e) }));
        child.on('close', (code) => {
          try { fs.unlinkSync(tmp); } catch (_) { /* ignore */ }
          const text = Buffer.concat(out).toString('utf8');
          const etext = Buffer.concat(err).toString('utf8').trim();
          let res;
          try {
            res = JSON.parse(text.slice(Math.max(0, text.indexOf('{'))));
          } catch (_) {
            res = { ok: false, command, error: etext || text.trim() || `PowerShell exited with code ${code}` };
          }
          output.appendLine(JSON.stringify(res, null, 2));
          resolve(res);
        });
      });

    return vscode.window.withProgress(
      { location: opts.quiet ? vscode.ProgressLocation.Window : vscode.ProgressLocation.Notification, title: opts.title || `Access: ${command} を実行中...` },
      exec
    );
  };
  const run = queue.then(job, job);
  queue = run.catch(() => undefined);
  return run.finally(() => scheduleStatusRefresh());
}

// ------------------------------------------------------------
// 結果の表示・診断（コンパイルエラー / 実行時エラーをエディタに表示）
// ------------------------------------------------------------
function applyDiagnostics(res) {
  diagnostics.clear();
  if (!res) return undefined;
  const byFile = new Map();
  let first;
  const add = (file, line, message) => {
    if (!file) return;
    const ln = Math.max(0, (line || 1) - 1);
    const d = new vscode.Diagnostic(new vscode.Range(ln, 0, ln, 10000), message, vscode.DiagnosticSeverity.Error);
    d.source = 'Access';
    const key = file.toLowerCase();
    if (!byFile.has(key)) byFile.set(key, { uri: vscode.Uri.file(file), list: [] });
    byFile.get(key).list.push(d);
    if (!first) first = { file, line: ln };
  };
  const r = res.result || {};
  const ce = res.command === 'compile' ? r.error : r.compile && r.compile.error;
  if (ce) add(ce.file, ce.fileLine, `${ce.message}${ce.code ? `  ← ${ce.code.trim()}` : ''}`);
  if (res.runtimeError) {
    const e = res.runtimeError;
    add(e.file, e.fileLine, `${e.message}${e.code ? `  ← ${e.code.trim()}` : ''}`);
  }
  if (Array.isArray(r.errors) && res.source) {
    for (const e of r.errors) add(path.join(res.source, ...String(e.file).split('/')), 1, e.message);
  }
  for (const { uri, list } of byFile.values()) diagnostics.set(uri, list);
  return first;
}

async function reveal(loc) {
  if (!loc) return;
  try {
    const doc = await vscode.workspace.openTextDocument(loc.file);
    const ed = await vscode.window.showTextDocument(doc, { preview: false });
    const pos = new vscode.Position(loc.line, 0);
    ed.selection = new vscode.Selection(pos, pos);
    ed.revealRange(new vscode.Range(pos, pos), vscode.TextEditorRevealType.InCenter);
  } catch (_) { /* file may not exist */ }
}

// よく出るエラーを、初めての人にもわかる日本語にする
function errorText(res) {
  const r = res.result || {};
  const ce = res.command === 'compile' ? r.error : r.compile && r.compile.error;
  const where = (e) => (e && e.file ? `（${path.basename(e.file)} の ${e.fileLine || e.line} 行目）` : e && e.module ? `（${e.module} の ${e.line} 行目）` : '');
  if (ce) return `${ce.message || 'コンパイル エラー'}${where(ce)}。赤い波線の行を直して、もう一度反映してください。`;
  if (res.runtimeError) {
    const msg = String(res.runtimeError.message || '');
    return `${/^実行時エラー/.test(msg) ? msg : `実行時エラー: ${msg}`}${where(res.runtimeError)}`;
  }
  const e = String(res.error || '');
  if (/open in another process/i.test(e)) return 'データベースが Access などで開かれています。Access を閉じてから、もう一度実行してください。';
  if (/did not finish within/i.test(e)) return '時間内に終わらなかったため、Access を終了しました。入力待ちの画面（InputBox やモーダルフォーム）が原因のことがあります。';
  if (/did not close/i.test(e)) return 'Access が正しく終了しなかったため、強制終了しました。もう一度実行してください。';
  const m = /^(\d+) file\(s\) failed to import/.exec(e);
  if (m) {
    const first = (r.errors || [])[0];
    return `${m[1]} 件のファイルを取り込めませんでした${first ? `（${first.file}: ${first.message}）` : ''}。`;
  }
  if (/No database specified|Multiple databases found/i.test(e)) return '使うデータベースが決まっていません。「データベースを選択」で選んでください。';
  if (/Database not found/i.test(e)) return 'データベースのファイルが見つかりません。「データベースを選択」で選び直してください。';
  if (/action query/i.test(e)) return 'この SQL はデータを書き換えるため、ここでは実行できません（読み取り専用）。';
  return e || '失敗しました';
}

async function report(res, okText) {
  if (!res) return res;
  const loc = applyDiagnostics(res);
  const dialogs = Array.isArray(res.dialogs) ? res.dialogs : [];
  const r = res.result || {};
  const hasCompileError = !!(res.command === 'compile' ? r.error : r.compile && r.compile.error);
  const dlgNote = dialogs.length && !hasCompileError && !res.runtimeError ? `（ダイアログ ${dialogs.length} 件を自動応答: ${dialogs.map((d) => d.message).join(' / ')}）` : '';
  if (res.ok) {
    vscode.window.showInformationMessage(`Access: ${okText(res)}${dlgNote}`);
  } else {
    if (loc) reveal(loc);
    vscode.window.showErrorMessage(`Access: ${errorText(res)}${dlgNote}`, '出力を表示').then((c) => c && output.show(true));
  }
  if (Array.isArray(res.warnings) && res.warnings.length) {
    vscode.window.showWarningMessage(`Access: ${res.warnings.join(' / ')}`);
  }
  return res;
}

// ------------------------------------------------------------
// コマンド
// ------------------------------------------------------------
async function cmdSelectDatabase(uri) {
  let db = uri && uri.fsPath;
  if (!db) {
    const found = await findDatabases();
    const items = found.map((f) => ({ label: path.basename(f), description: vscode.workspace.asRelativePath(f), f }));
    items.push({ label: '$(folder-opened) ファイルを選択...', f: '' });
    const item = await vscode.window.showQuickPick(items, { placeHolder: '使用する Access データベース' });
    if (!item) return;
    db = item.f;
    if (!db) {
      const pick = await vscode.window.showOpenDialog({ filters: { 'Access データベース': ['accdb', 'mdb'] } });
      if (!pick) return;
      db = pick[0].fsPath;
    }
  }
  await saveDbSetting(db);
  vscode.window.showInformationMessage(`Access: ${path.basename(db)} を使用します。ソース: ${path.basename(srcFor(db))}`);
  if (!fs.existsSync(path.join(srcFor(db), MANIFEST))) {
    const c = await vscode.window.showInformationMessage('まだソースフォルダがありません。今エクスポートしますか？', 'エクスポート');
    if (c) await cmdExport();
  }
}

async function cmdExport() {
  const title = 'Access → ソースへエクスポート中...';
  let res = await access('export', {}, { title });
  if (res && !res.ok && res.result && Array.isArray(res.result.unimported)) {
    const n = res.result.unimported.length;
    const c = await vscode.window.showWarningMessage(
      `ソースフォルダに未インポートの変更が ${n} 件あります。エクスポートすると上書きされます。`,
      { modal: true, detail: res.result.unimported.join('\n') },
      '先にインポートする',
      '上書きしてエクスポート'
    );
    if (c === '先にインポートする') {
      const r = await cmdImport();
      if (!r || !r.ok) return;
      res = await access('export', {}, { title });
    } else if (c === '上書きしてエクスポート') {
      res = await access('export', { force: true }, { title });
    } else return;
  }
  return report(res, (r) => {
    const x = r.result || {};
    return `エクスポート完了（${x.objects} オブジェクト / 更新 ${(x.written || []).length} / 削除 ${(x.deleted || []).length}）`;
  });
}

function importSummary(r) {
  const x = r.result || {};
  if (r.message) return r.message;
  const parts = [`インポート ${(x.imported || []).length} 件`];
  if ((x.deleted || []).length) parts.push(`削除 ${x.deleted.length} 件`);
  if (x.compile) parts.push(x.compile.compiled ? 'コンパイル OK' : 'コンパイル エラー');
  return parts.join(' / ');
}

async function cmdImport(params = {}) {
  const db = await getDb(true);
  if (!db) return undefined;
  const p = Object.assign({}, params);
  if (!p.files && !p.all) {
    const st = syncStatus(srcFor(db));
    if (st.missing.length) {
      const c = await vscode.window.showWarningMessage(
        `ソースから削除されたファイルが ${st.missing.length} 件あります。対応するオブジェクトを Access からも削除しますか？`,
        { modal: true, detail: st.missing.join('\n') },
        '削除する',
        '削除しない'
      );
      if (!c) return undefined;
      p.deleteMissing = c === '削除する';
    }
  }
  const res = await access('import', p, { db, title: 'ソース → Access へインポート中...' });
  return report(res, importSummary);
}

async function cmdImportFile(uri, uris) {
  const toUri = (u) => (u && u.resourceUri ? u.resourceUri : u);
  const targets = (uris && uris.length ? uris.map(toUri) : [toUri(uri) || (vscode.window.activeTextEditor && vscode.window.activeTextEditor.document.uri)])
    .filter((u) => u && u.fsPath)
    .map((u) => u.fsPath);
  if (!targets.length) return;
  await vscode.workspace.saveAll(false);
  return cmdImport({ files: targets });
}

async function cmdCompile() {
  const res = await access('compile', {}, { title: 'VBA をコンパイル中...' });
  return report(res, () => 'コンパイル OK');
}

// .bas の中のプロシージャを探す
const PROC_RE = /^\s*(?:(Public|Private|Friend)\s+)?(?:Static\s+)?(Function|Sub)\s+([A-Za-z_\u0080-￿][\w\u0080-￿]*)\s*\(([^)]*)\)/i;

function parseProceduresText(text) {
  const procs = [];
  const lines = text.split(/\r?\n/);
  for (let i = 0; i < lines.length; i++) {
    const m = PROC_RE.exec(lines[i]);
    if (m && (m[1] || '').toLowerCase() !== 'private') procs.push({ line: i, kind: m[2], name: m[3], params: m[4].trim() });
  }
  return procs;
}

function parseProcedures(doc) {
  return parseProceduresText(doc.getText());
}

// サイドバーの項目・エクスプローラーの Uri・開いているエディタから、対象オブジェクトを割り出す
function objectFromArg(arg) {
  if (arg && arg.objName) return { name: arg.objName, kind: arg.kind, file: arg.resourceUri && arg.resourceUri.fsPath };
  let file;
  if (arg instanceof vscode.Uri) file = arg.fsPath;
  else if (vscode.window.activeTextEditor) file = vscode.window.activeTextEditor.document.fileName;
  if (!file) return {};
  const kind = path.basename(path.dirname(file)).toLowerCase();
  if (!MANAGED[kind]) return {};
  return { name: decodeName(path.basename(file, path.extname(file))), kind, file };
}

function isStandardModule(doc) {
  const db = configuredDb();
  if (path.extname(doc.fileName).toLowerCase() !== '.bas') return false;
  const dir = path.basename(path.dirname(doc.fileName)).toLowerCase();
  if (dir !== 'modules') return false;
  return !db || !!relOf(srcFor(db), doc.fileName) || /\.src$/i.test(path.dirname(path.dirname(doc.fileName)));
}

async function ensureImported(db) {
  const st = syncStatus(srcFor(db));
  const dirty = st.changed.length + st.added.length;
  if (!dirty) return true;
  const c = await vscode.window.showWarningMessage(
    `未インポートの変更が ${dirty} 件あります（${[...st.changed, ...st.added].join(', ')}）。`,
    'インポートしてから実行',
    'そのまま実行'
  );
  if (!c) return false;
  if (c === 'インポートしてから実行') {
    await vscode.workspace.saveAll(false);
    const r = await access('import', {}, { db, title: 'インポート中...' });
    await report(r, importSummary);
    return !!(r && r.ok);
  }
  return true;
}

function parseArgs(text) {
  const t = (text || '').trim();
  if (!t) return [];
  const v = JSON.parse(t.startsWith('[') ? t : `[${t}]`);
  return Array.isArray(v) ? v : [v];
}

async function cmdRunProcedure(name, params, presetArgs) {
  const db = await getDb(true);
  if (!db) return;
  if (typeof name === 'string' && Array.isArray(presetArgs)) {
    // 引数を指定して実行（自動化・撮影用。入力欄を出さない）
    if (!(await ensureImported(db))) return;
    const res = await access('run', { function: name, args: presetArgs }, { db, title: `${name} を実行中...` });
    await report(res, (r) => `${name} → ${JSON.stringify(r.result && r.result.value)}`);
    output.show(true);
    return res;
  }
  if (name && typeof name === 'object') {
    // サイドバーのモジュールから：中のプロシージャを一覧から選ぶ
    const o = objectFromArg(name);
    name = undefined;
    if (o.file && fs.existsSync(o.file)) {
      const procs = parseProceduresText(fs.readFileSync(o.file, 'utf8'));
      if (!procs.length) {
        vscode.window.showInformationMessage(`${o.name} に実行できる Public Function / Sub がありません。`);
        return;
      }
      const pick = await vscode.window.showQuickPick(
        procs.map((p) => ({ label: `$(symbol-method) ${p.name}`, description: `${p.kind}(${p.params})`, p })),
        { placeHolder: '実行するプロシージャを選んでください' }
      );
      if (!pick) return;
      name = pick.p.name;
      params = pick.p.params;
    }
  }
  if (!name) {
    const ed = vscode.window.activeTextEditor;
    let def = '';
    if (ed && ed.document.languageId === 'vba') {
      const procs = parseProcedures(ed.document).filter((p) => p.line <= ed.selection.active.line);
      if (procs.length) {
        def = procs[procs.length - 1].name;
        params = procs[procs.length - 1].params;
      }
    }
    name = await vscode.window.showInputBox({ prompt: '実行するプロシージャ名（標準モジュールの Public Function / Sub）', value: def });
    if (!name) return;
    if (name !== def) params = undefined;
  }
  let args = [];
  if (params === undefined || params) {
    const text = await vscode.window.showInputBox({
      prompt: `引数を JSON で入力（例: "abc", 123）${params ? `  — (${params})` : '（なければ空欄）'}`,
      placeHolder: '"文字列", 123, true',
    });
    if (text === undefined) return;
    try {
      args = parseArgs(text);
    } catch (e) {
      vscode.window.showErrorMessage(`引数が JSON として読めません: ${e.message}`);
      return;
    }
  }
  if (!(await ensureImported(db))) return;
  const res = await access('run', { function: name, args }, { db, title: `${name} を実行中...` });
  await report(res, (r) => `${name} → ${JSON.stringify(r.result && r.result.value)}`);
  output.show(true);
}

function mdCell(v) {
  if (v === null || v === undefined) return '*Null*';
  return String(typeof v === 'object' ? JSON.stringify(v) : v).replace(/\|/g, '\\|').replace(/\r?\n/g, '<br>');
}

async function cmdRunSql(sqlArg) {
  const ed = vscode.window.activeTextEditor;
  let sql;
  if (sqlArg && sqlArg.objName) {
    // サイドバーから：クエリはそのまま実行、テーブルは先頭 100 行
    sql = sqlArg.kind === 'tables' ? `SELECT TOP 100 * FROM [${sqlArg.objName}]` : sqlArg.objName;
  } else {
    let def = typeof sqlArg === 'string' ? sqlArg : '';
    if (!def && ed) {
      if (!ed.selection.isEmpty) def = ed.document.getText(ed.selection);
      else if (path.extname(ed.document.fileName).toLowerCase() === '.sql') def = ed.document.getText();
    }
    def = def.split(/\r?\n/).filter((l) => !/^--\s*accessvsc:/.test(l)).join(' ').replace(/\s+/g, ' ').trim();
    sql = await vscode.window.showInputBox({ prompt: 'SQL（SELECT のみ。保存済みクエリ名も可）', value: def });
  }
  if (!sql) return;
  const limit = cfg().get('sqlRowLimit') || 200;
  const res = await access('sql', { sql, limit }, { title: 'SQL を実行中...' });
  if (!res) return;
  if (!res.ok) return report(res, () => '');
  const r = res.result;
  const lines = [`# SQL 結果`, '', '```sql', sql, '```', '', `${r.rowCount} 行${r.truncated ? `（先頭 ${limit} 行まで表示）` : ''}`, ''];
  lines.push(`| ${r.columns.map(mdCell).join(' | ')} |`, `| ${r.columns.map(() => '---').join(' | ')} |`);
  for (const row of r.rows) lines.push(`| ${r.columns.map((c) => mdCell(row[c])).join(' | ')} |`);
  const doc = await vscode.workspace.openTextDocument({ language: 'markdown', content: lines.join('\n') });
  await vscode.window.showTextDocument(doc, { preview: true, viewColumn: vscode.ViewColumn.Beside });
}

async function cmdStatus() {
  const db = await getDb(true);
  if (!db) return;
  const src = srcFor(db);
  const st = syncStatus(src);
  const items = [
    ...st.changed.map((f) => ({ label: `$(diff-modified) ${f}`, description: '変更（未インポート）', f })),
    ...st.added.map((f) => ({ label: `$(diff-added) ${f}`, description: '追加（未インポート）', f })),
    ...st.missing.map((f) => ({ label: `$(diff-removed) ${f}`, description: '削除（Access には残っている）' })),
    ...st.readOnlyChanged.map((f) => ({ label: `$(lock) ${f}`, description: '読み取り専用ファイルが変更されています（DB には反映されません）', f })),
  ];
  if (!st.hasManifest) items.unshift({ label: '$(info) まだエクスポートされていません', description: '「Access → ソースへエクスポート」を実行してください' });
  if (!items.length) items.push({ label: '$(check) ソースと Access は同期しています' });
  const pick = await vscode.window.showQuickPick(items, { placeHolder: `${path.basename(db)} ⇔ ${path.basename(src)}` });
  if (pick && pick.f) vscode.window.showTextDocument(vscode.Uri.file(path.join(src, ...pick.f.split('/'))));
}

async function cmdOpenInAccess() {
  const db = await getDb(true);
  if (!db) return;
  await releaseAccessSession();   // 裏で開いたままの Access を閉じてから
  vscode.env.openExternal(vscode.Uri.file(db));
}

// ------------------------------------------------------------
// 帳票（レポート）：プレビュー・レイアウト・ページ設定
// ------------------------------------------------------------
const PAPER_SIZES = ['A4', 'A3', 'B4', 'B5', 'A5', 'Letter', 'Legal'];
let previewPanel;

async function pickObject(kinds, placeHolder) {
  const db = await getDb(true);
  if (!db) return {};
  const src = srcFor(db);
  const items = [];
  for (const kind of kinds) {
    const dir = path.join(src, kind);
    if (!fs.existsSync(dir)) continue;
    for (const f of fs.readdirSync(dir).filter((n) => n.toLowerCase().endsWith('.txt'))) {
      const name = decodeName(path.basename(f, '.txt'));
      items.push({ label: `${kind === 'forms' ? '$(window)' : '$(file-text)'} ${name}`, description: kind === 'forms' ? 'フォーム' : 'レポート', o: { name, kind, file: path.join(dir, f) } });
    }
  }
  if (!items.length) {
    vscode.window.showInformationMessage('対象がありません。先に「Access から取り出す（エクスポート）」を実行してください。');
    return {};
  }
  const pick = await vscode.window.showQuickPick(items, { placeHolder });
  return pick ? pick.o : {};
}

async function resolveReportArg(arg, kinds = ['reports', 'forms']) {
  let o = objectFromArg(arg);
  if (!o.name || !kinds.includes(o.kind)) o = await pickObject(kinds, '対象のレポートを選んでください');
  return o;
}

async function ensureFileImported(db, o) {
  const rel = o.file ? relOf(srcFor(db), o.file) : undefined;
  if (!rel) return true;
  const st = syncStatus(srcFor(db));
  if (!st.changed.includes(rel) && !st.added.includes(rel)) return true;
  const c = await vscode.window.showWarningMessage(
    `${o.name} の変更がまだ Access に反映されていません。`,
    '反映してから続ける',
    '反映せずに続ける'
  );
  if (!c) return false;
  if (c === '反映してから続ける') {
    await vscode.workspace.saveAll(false);
    const r = await access('import', { files: [o.file] }, { db, title: `${o.name} を反映中...` });
    await report(r, importSummary);
    return !!(r && r.ok);
  }
  return true;
}

async function cmdPreviewReport(arg) {
  const db = await getDb(true);
  if (!db) return;
  const o = await resolveReportArg(arg);
  if (!o.name) return;
  if (!(await ensureFileImported(db, o))) return;
  const res = await access('preview', { report: o.name, objectType: o.kind === 'forms' ? 'form' : 'report', pages: 5 }, { db, title: `${o.name} のプレビューを作成中...` });
  if (!res) return;
  if (!res.ok) return report(res, () => '');
  showPreviewPanel(o, res);
}

function escHtml(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function cmdLink(command, args) {
  return `command:${command}?${encodeURIComponent(JSON.stringify(args || []))}`;
}

function showPreviewPanel(o, res) {
  const r = res.result || {};
  const dir = r.pdf ? path.dirname(r.pdf) : os.tmpdir();
  if (!previewPanel) {
    previewPanel = vscode.window.createWebviewPanel('accessVsc.preview', 'プレビュー', vscode.ViewColumn.Beside, {
      enableCommandUris: true,
      localResourceRoots: [vscode.Uri.file(dir)],
    });
    previewPanel.onDidDispose(() => (previewPanel = undefined));
  }
  const wv = previewPanel.webview;
  previewPanel.title = `プレビュー: ${o.name}`;
  const stamp = Date.now();
  const imgs = (r.images || []).map((p, i) => `<figure><img src="${wv.asWebviewUri(vscode.Uri.file(p))}?t=${stamp}" alt="${i + 1} ページ"><figcaption>${i + 1} ページ</figcaption></figure>`).join('');
  const issues = r.issues || [];
  const pr = r.printer;
  const target = { objName: o.name, kind: o.kind };
  const more = r.pageCount && (r.images || []).length < r.pageCount ? `（全 ${r.pageCount} ページ中 ${(r.images || []).length} ページを表示。全ページは PDF で確認）` : '';
  previewPanel.webview.html = `<!DOCTYPE html><html lang="ja"><head><meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src ${wv.cspSource}; style-src 'unsafe-inline';">
<style>
 body { font-family: var(--vscode-font-family); color: var(--vscode-foreground); padding: 0 16px 24px; }
 .bar { display: flex; flex-wrap: wrap; gap: 8px; margin: 12px 0; }
 .bar a { padding: 4px 12px; border-radius: 2px; background: var(--vscode-button-secondaryBackground); color: var(--vscode-button-secondaryForeground); text-decoration: none; }
 .bar a.primary { background: var(--vscode-button-background); color: var(--vscode-button-foreground); }
 .box { padding: 8px 12px; border-radius: 4px; margin: 8px 0; border-left: 4px solid; }
 .warn { border-color: var(--vscode-editorWarning-foreground); background: var(--vscode-inputValidation-warningBackground); }
 .ok { border-color: var(--vscode-testing-iconPassed, #3c3); background: var(--vscode-editor-inactiveSelectionBackground); }
 .meta { opacity: .8; font-size: 90%; }
 figure { margin: 16px 0; }
 img { max-width: 100%; background: #fff; box-shadow: 0 2px 8px rgba(0,0,0,.35); }
 figcaption { opacity: .7; font-size: 85%; margin-top: 4px; }
 ul { margin: 4px 0; padding-left: 20px; }
</style></head><body>
<h2>${escHtml(o.name)}</h2>
<div class="meta">${pr ? `用紙: ${escHtml(pr.paperSize)} ${pr.orientation === 'landscape' ? '横' : '縦'} / 余白(mm) 上${pr.marginsMm.top} 下${pr.marginsMm.bottom} 左${pr.marginsMm.left} 右${pr.marginsMm.right} / ` : ''}幅 ${escHtml(r.widthCm)} cm${pr && pr.printableWidthCm ? `（印刷可能幅 ${pr.printableWidthCm} cm）` : ''} ${escHtml(more)}</div>
<div class="bar">
 <a class="primary" href="${cmdLink('accessVsc.previewReport', [target])}">再プレビュー</a>
 <a href="${cmdLink('accessVsc.reportLayout', [target])}">レイアウト詳細</a>
 ${o.kind === 'reports' ? `<a href="${cmdLink('accessVsc.pageSetup', [target])}">ページ設定</a>` : ''}
 ${o.file ? `<a href="${cmdLink('vscode.open', [vscode.Uri.file(o.file)])}">ソースを開く</a>` : ''}
 ${r.pdf ? `<a href="${cmdLink('accessVsc.openExternalFile', [r.pdf])}">PDF を開く</a>` : ''}
</div>
${issues.length ? `<div class="box warn"><b>レイアウトの問題 ${issues.length} 件</b><ul>${issues.map((i) => `<li>${escHtml(i)}</li>`).join('')}</ul></div>` : '<div class="box ok">レイアウトの問題は見つかりませんでした。</div>'}
${(res.warnings || []).map((w) => `<div class="box warn">${escHtml(w)}</div>`).join('')}
${imgs || '<p>画像を作成できませんでした。「PDF を開く」で確認してください。</p>'}
</body></html>`;
  previewPanel.reveal(vscode.ViewColumn.Beside, true);
}

async function cmdReportLayout(arg) {
  const db = await getDb(true);
  if (!db) return;
  const o = await resolveReportArg(arg);
  if (!o.name) return;
  const res = await access('layout', { report: o.name, objectType: o.kind === 'forms' ? 'form' : 'report' }, { db, title: `${o.name} のレイアウトを取得中...` });
  if (!res) return;
  if (!res.ok) return report(res, () => '');
  const r = res.result;
  const secName = {};
  for (const s of r.sections) secName[s.index] = s.name;
  const lines = [`# レイアウト: ${r.object}`, ''];
  if (r.printer) {
    const p = r.printer;
    lines.push(`- 用紙: **${p.paperSize} ${p.orientation === 'landscape' ? '横' : '縦'}**、余白(mm) 上 ${p.marginsMm.top} / 下 ${p.marginsMm.bottom} / 左 ${p.marginsMm.left} / 右 ${p.marginsMm.right}`);
    if (p.printableWidthCm) lines.push(`- 印刷可能幅: ${p.printableWidthCm} cm`);
  }
  lines.push(`- 幅: **${r.widthCm} cm**（${r.width} twip。1 cm = 567 twip）`, '');
  lines.push('## 問題', '', ...(r.issues.length ? r.issues.map((i) => `- ⚠ ${i}`) : ['- なし']), '');
  lines.push('## セクション', '', '| セクション | 高さ (cm) |', '|---|---|', ...r.sections.map((s) => `| ${mdCell(s.name)} | ${s.heightCm} |`), '');
  lines.push('## コントロール', '', '| 名前 | 種類 | セクション | 左 | 上 | 幅 | 高さ (cm) | 表示内容 | フォント |', '|---|---|---|---|---|---|---|---|---|');
  for (const c of r.controls) {
    const cm = (v) => (v / 567).toFixed(2);
    lines.push(`| ${mdCell(c.name)} | ${c.type} | ${mdCell(secName[c.section])} | ${cm(c.left)} | ${cm(c.top)} | ${cm(c.width)} | ${cm(c.height)} | ${mdCell(c.caption || c.controlSource || '')} | ${mdCell(c.fontName ? `${c.fontName} ${c.fontSize}pt` : '')} |`);
  }
  const doc = await vscode.workspace.openTextDocument({ language: 'markdown', content: lines.join('\n') });
  await vscode.window.showTextDocument(doc, { preview: true, viewColumn: vscode.ViewColumn.Beside });
}

async function cmdPageSetup(arg) {
  const db = await getDb(true);
  if (!db) return;
  const o = await resolveReportArg(arg, ['reports']);
  if (!o.name) return;
  if (!(await ensureFileImported(db, o))) return;
  const keep = '（変更しない）';
  const paper = await vscode.window.showQuickPick([keep, ...PAPER_SIZES], { placeHolder: `${o.name}: 用紙サイズ` });
  if (!paper) return;
  const orient = await vscode.window.showQuickPick([keep, '縦', '横'], { placeHolder: `${o.name}: 用紙の向き` });
  if (!orient) return;
  const margins = await vscode.window.showInputBox({
    prompt: '余白 (mm)。「上,下,左,右」または 1 つの数値（全部同じ）。空欄なら変更しない',
    placeHolder: '例: 10,10,15,15',
    validateInput: (v) => (!v.trim() || /^\s*\d+(\.\d+)?(\s*,\s*\d+(\.\d+)?){0}\s*$|^\s*\d+(\.\d+)?(\s*,\s*\d+(\.\d+)?){3}\s*$/.test(v) ? undefined : '「10」または「10,10,15,15」の形で入力してください'),
  });
  if (margins === undefined) return;
  const p = { report: o.name };
  if (paper !== keep) p.paperSize = paper;
  if (orient !== keep) p.orientation = orient === '横' ? 'landscape' : 'portrait';
  if (margins.trim()) {
    let m = margins.split(',').map((x) => Number(x.trim()));
    if (m.length === 1) m = [m[0], m[0], m[0], m[0]];
    [p.marginTop, p.marginBottom, p.marginLeft, p.marginRight] = m;
  }
  if (Object.keys(p).length === 1) return;
  const res = await access('pagesetup', p, { db, title: `${o.name} のページ設定を変更中...` });
  await report(res, (r) => {
    const pr = r.result.printer;
    return `ページ設定を変更しました（${pr.paperSize} ${pr.orientation === 'landscape' ? '横' : '縦'}）${r.result.issues.length ? ` — レイアウトの問題 ${r.result.issues.length} 件` : ''}`;
  });
  if (res && res.ok) {
    const c = await vscode.window.showInformationMessage('プレビューで確認しますか？', 'プレビュー');
    if (c) cmdPreviewReport({ objName: o.name, kind: 'reports', resourceUri: o.file && vscode.Uri.file(o.file) });
  }
}

async function cmdMenu() {
  const items = [
    { label: '$(cloud-upload) 変更をインポート', cmd: 'accessVsc.import' },
    { label: '$(cloud-download) Access → ソースへエクスポート', cmd: 'accessVsc.export' },
    { label: '$(check-all) VBA をコンパイル', cmd: 'accessVsc.compile' },
    { label: '$(play) VBA プロシージャを実行', cmd: 'accessVsc.runProcedure' },
    { label: '$(table) SQL を実行', cmd: 'accessVsc.runSql' },
    { label: '$(open-preview) 帳票をプレビュー', cmd: 'accessVsc.previewReport' },
    { label: '$(symbol-ruler) 帳票のレイアウト詳細', cmd: 'accessVsc.reportLayout' },
    { label: '$(settings) 帳票のページ設定', cmd: 'accessVsc.pageSetup' },
    { label: '$(git-compare) 同期状態を表示', cmd: 'accessVsc.status' },
    { label: '$(link-external) Access で開く', cmd: 'accessVsc.openInAccess' },
    { label: '$(database) データベースを選択', cmd: 'accessVsc.selectDatabase' },
    { label: '$(sparkle) AI エージェント連携を設定', cmd: 'accessVsc.setupAgent' },
    { label: '$(output) 出力ログを表示', cmd: 'accessVsc.showOutput' },
  ];
  const pick = await vscode.window.showQuickPick(items, { placeHolder: 'Access VSC' });
  if (pick) vscode.commands.executeCommand(pick.cmd);
}

// ------------------------------------------------------------
// AI エージェント連携
// ------------------------------------------------------------
function agentDoc(db, root) {
  const rel = (p) => {
    const r = root ? path.relative(root, p) : p;
    return (r && !r.startsWith('..') ? r : p).split(path.sep).join('/');
  };
  const src = srcFor(db);
  // リポジトリに入るファイルなので、PC 固有のパス（ユーザー名を含む）は書かない
  const run = 'powershell -NoProfile -ExecutionPolicy Bypass -File "%LOCALAPPDATA%\\access-vscode\\bin\\access.ps1"';
  const srcArg = cfg().get('sourceFolder') ? ` -Src "${rel(src)}"` : '';
  const dbRel = rel(db);
  return `${AGENT_BEGIN}
## Access データベース（Access VSC）

Access データベース \`${rel(db)}\` は \`${rel(src)}/\` にテキストとして展開されている。
**.accdb を直接いじらず、必ずソースを編集してからインポートすること。**

| パス | 内容 | 編集 |
|---|---|---|
| \`modules/*.bas\` | 標準モジュール | 可 |
| \`modules/*.cls\` | クラスモジュール | 可 |
| \`forms/*.txt\` / \`reports/*.txt\` | フォーム / レポート（SaveAsText 形式。末尾の \`CodeBehindForm\` 以降が VBA） | 可（慎重に） |
| \`macros/*.txt\` | マクロ | 可 |
| \`queries/*.sql\` | クエリ（Access SQL。1 ファイル 1 クエリ） | 可 |
| \`tables/*.json\`, \`relationships.json\`, \`references.json\` | テーブル定義・リレーション・VBA 参照設定のスナップショット | **読み取り専用** |

### 手順
1. \`access_status\` で同期状態を確認（Access を起動しないので速い）
2. ソースファイルを編集（新しいファイルを置けば新しいオブジェクトになる）。VBA を書いたら \`access_format\` でインデントを整える
3. \`access_import\` で取り込み＋コンパイル。コンパイルエラーは \`file\` / \`fileLine\` / \`code\` が返るので、そこを直して再インポート。\`result.verify\` は取り込んだオブジェクトと Public な手続きが DB に本当にあるかの確認結果
4. \`access_run\`（VBA の Public Function/Sub を実行）、\`access_sql\`（データ確認。既定は読み取り専用）、\`access_eval\` で動作確認
5. ユーザーが Access 側で変更した場合は \`access_export\` で取り込む（未インポートの編集があると拒否される）
6. テーブル定義の変更は \`access_sql\`（write=true で DDL）→ \`access_export\`（kinds=["tables"]）

### 帳票（レポート）のレイアウト修正
1. \`access_report_layout\` で現状を確認（各コントロールの位置・大きさ、用紙設定、問題点）。単位は twip（**1 cm = 567 twip**）
2. \`reports/<名前>.txt\` の該当コントロールの \`Left\` / \`Top\` / \`Width\` / \`Height\` / \`Caption\` / \`FontSize\` などを書き換える
   - コントロールを増やすときは、既存の \`Begin TextBox ... End\` ブロックを丸ごと複製して \`Name\` を変える
   - \`LayoutCached*\` 行は Access が計算し直すので気にしなくてよい
3. \`access_import\` で反映 → \`access_report_preview\` で **画像を見て** 確認。\`issues\`（はみ出し・重なり・用紙に入らない）が空になり、見た目が意図どおりになるまで繰り返す
4. 用紙サイズ・向き・余白は .txt ではなく \`access_report_pagesetup\` で変える（PrtDevMode / PrtMip のバイナリ行は編集しない）

### テスト・品質
- テストは標準モジュール（例 \`modules/Test_Tax.bas\`）に **\`Public Sub Test_○○()\`（引数なし）** で書く。同じモジュールの \`TestSetup\` / \`TestTeardown\` は各テストの前後に呼ばれる
- 確認は \`Assert\` モジュール: \`Assert.AreEqual 期待値, 実際\` / \`Assert.IsTrue 条件\` / \`Assert.AreClose 期待, 実際, 誤差\` / \`Assert.Contains 文字列, 部分\` / \`Assert.Fail "理由"\`
- フォームの動き（入力 → ボタン → 結果）は \`UI\` モジュールで確かめる: \`UI.OpenForm "frm"\` / \`UI.TypeText "txt", "値"\`（Tab で確定し AfterUpdate が動く）/ \`UI.Click "cmd"\` / \`UI.Value("txt")\` / \`UI.CloseForm\`（TestTeardown で閉じる）。値を直接代入するとイベントが動かないので、操作のテストでは UI を使う
- \`access_import\` してから \`access_test\`（filter で絞り込み可）。テストは DB のコピーで動くのでデータを書き換えてよい（ただしリンクテーブルの先は本物）
- 機能を直したら、対応するテストを追加・更新し、\`access_test\` と \`access_lint\` が通ってから完了とする
- \`access_lint\` の警告（型の書き忘れ \`implicit-variant\`・使っていない変数 \`unused-variable\` など）は、触ったコードの分は直す
- 帳票や、帳票が使うクエリ・コードを変えたら \`access_report_snapshot\` で見た目を基準画像と比べる。変化があれば差分画像を見てユーザーに報告し、**意図した変化だとユーザーが認めた場合だけ** update=true で基準を更新する

### 完了の条件（すべて満たしてから「できました」と報告する）
- \`access_import\` が \`ok: true\`（取り込みエラーなし・コンパイル成功・\`verify.missing\` が空）
- \`access_test\` の失敗が 0 件、\`access_lint\` のエラーが 0 件
- 帳票を変えたら \`access_report_preview\` の画像を見て確認し、\`issues\` が空
- 報告には、実行したツールとその結果（テストの件数・失敗数など）を書く。ファイルの大きさや「保存できた」ことを完了の根拠にしない
- 失敗したら原因を調べて直す。テストを書き換えて通すのではなく、コードを直す（仕様が変わった場合はユーザーに確認）

### リリース
- \`access_build\` で .accde を作れる（版番号の書き込み・本番バックエンドへの張り替え込み）
- **配布（deploy）はユーザーが明示的に頼んだときだけ**。配布は VS Code の「配布する」か CLI の \`deploy -Yes\` で行う

MCP ツールが使えないときは CLI（結果は JSON。プロジェクトのフォルダで実行する。\`%LOCALAPPDATA%\` は cmd の書き方で、PowerShell では \`$env:LOCALAPPDATA\`、bash では \`$LOCALAPPDATA\` に置き換える）:
\`\`\`
${run} status  -DbPath "${dbRel}"${srcArg}
${run} import  -DbPath "${dbRel}"${srcArg}
${run} import  -DbPath "${dbRel}"${srcArg} -Files modules/Module1.bas
${run} compile -DbPath "${dbRel}"${srcArg}
${run} run     -DbPath "${dbRel}"${srcArg} -Function MyFunc -ArgsJson "[1, 'abc']"
${run} sql     -DbPath "${dbRel}"${srcArg} -Sql "SELECT TOP 10 * FROM MyTable"
${run} eval    -DbPath "${dbRel}"${srcArg} -Expr "DCount('*','MyTable')"
${run} export  -DbPath "${dbRel}"${srcArg}
${run} test    -DbPath "${dbRel}"${srcArg} -Text                 # -Filter "Test_Tax*"
${run} lint    -DbPath "${dbRel}"${srcArg} -Text
${run} format  -DbPath "${dbRel}"${srcArg} -Text                 # VBA のインデントを整える（-Check で確認だけ）
${run} layout  -DbPath "${dbRel}"${srcArg} -Report rptSales
${run} preview -DbPath "${dbRel}"${srcArg} -Report rptSales      # result.images の PNG を読んで確認
${run} pagesetup -DbPath "${dbRel}"${srcArg} -Report rptSales -PaperSize A4 -Orientation landscape -Margins "10,10,15,15"
\`\`\`

### 注意
- VBA は Access 内で Shift-JIS で保存される。絵文字など Shift-JIS にない文字は使わない
- MsgBox などのダイアログは自動で答える（OK だけなら OK、確認の質問は安全側の「キャンセル」／「いいえ」）。押したボタンは結果の \`dialogs\` に記録される。InputBox やモーダルフォームで止まるとタイムアウトで Access が強制終了される
- VBA の実行時エラーは \`runtimeError\`（モジュール・行・ソースの行番号）で返る
- インポートと書き込み SQL の前に \`${rel(path.join(path.dirname(db), '.accessvsc', 'backups'))}/\` へ DB のバックアップが作られる
- ユーザーが Access でこの DB を開いていると書き込み系の操作は失敗する（閉じてもらう）
- フォームの .txt はプロパティの書式が厳密。既存の行の書き方に合わせ、GUID や Checksum などのバイナリ行は触らない
${AGENT_END}`;
}

function upsertBlock(file, block, header) {
  let text = '';
  try {
    text = fs.readFileSync(file, 'utf8');
  } catch (_) { /* new file */ }
  const s = text.indexOf(AGENT_BEGIN);
  const e = text.indexOf(AGENT_END);
  if (s >= 0 && e > s) text = text.slice(0, s) + block + text.slice(e + AGENT_END.length);
  else text = (text ? text.replace(/\s*$/, '\n\n') : header || '') + block + '\n';
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, text, 'utf8');
}

// ワークスペースの中なら相対パスにする（リポジトリに入るファイルに、PC 固有のパスやユーザー名を書かないため）
function relToWorkspace(p) {
  const root = workspaceRoot();
  if (!root) return p;
  const r = path.relative(root, p);
  return r && !r.startsWith('..') && !path.isAbsolute(r) ? r : p;
}

function writeMcpJson(file, db, withType = true) {
  let json = {};
  if (fs.existsSync(file)) {
    try {
      json = JSON.parse(fs.readFileSync(file, 'utf8').replace(/^﻿/, ''));
    } catch (_) {
      // コメント入りなど、読めないファイルを空にして上書きしない
      throw new Error(`${relToWorkspace(file)} を JSON として読めないため、書き換えませんでした。MCP サーバー「access」は手で追加してください（リファレンス「AI 連携・MCP ツール」参照）。`);
    }
  }
  json.mcpServers = json.mcpServers || {};
  // スクリプトの場所は、ユーザー名を含む絶対パスではなく、実行時に LocalAppData から組み立てる。
  // "$" を使わないのは、MCP クライアントによっては $VAR を環境変数として置き換えてしまうため
  const q = (s) => `'${String(s).replace(/'/g, "''")}'`;
  let command = `& (Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'access-vscode\\bin\\mcp-server.ps1') -DbPath ${q(relToWorkspace(db))}`;
  if (cfg().get('sourceFolder')) command += ` -Src ${q(relToWorkspace(srcFor(db)))}`;
  const args = ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', command];
  json.mcpServers.access = withType ? { type: 'stdio', command: 'powershell.exe', args } : { command: 'powershell.exe', args };
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, JSON.stringify(json, null, 2) + '\n', 'utf8');
}

async function cmdSetupAgent() {
  const root = workspaceRoot();
  if (!root) {
    vscode.window.showErrorMessage('フォルダ（ワークスペース）を開いてから実行してください。');
    return;
  }
  const db = await getDb(true);
  if (!db) return;
  const picks = await vscode.window.showQuickPick(
    [
      { label: 'Claude Code', description: '.mcp.json（MCP サーバー登録）と CLAUDE.md（作業手順）', id: 'claude', picked: true },
      { label: 'GitHub Copilot', description: '.github/copilot-instructions.md（MCP サーバーはこの拡張機能が自動登録）', id: 'copilot' },
      { label: 'Cursor', description: '.cursor/mcp.json（MCP サーバー登録）と .cursor/rules/access-vsc.mdc（作業手順）', id: 'cursor' },
      { label: 'Gemini CLI', description: '.gemini/settings.json（MCP サーバー登録）と GEMINI.md（作業手順）', id: 'gemini' },
      { label: 'Codex・その他のエージェント', description: 'AGENTS.md（MCP サーバーは各ツールの設定で登録。CLI でも操作可）', id: 'agents' },
    ],
    { canPickMany: true, placeHolder: '設定する AI エージェントを選んでください' }
  );
  if (!picks || !picks.length) return;
  const written = [];
  const block = agentDoc(db, root);
  for (const p of picks) {
    try {
      setupOneAgent(p);
    } catch (e) {
      vscode.window.showErrorMessage(`Access VSC: ${p.label} の設定: ${e.message}`);
    }
  }
  function setupOneAgent(p) {
    if (p.id === 'claude') {
      writeMcpJson(path.join(root, '.mcp.json'), db);
      upsertBlock(path.join(root, 'CLAUDE.md'), block, '# CLAUDE.md\n\n');
      written.push('.mcp.json', 'CLAUDE.md');
    } else if (p.id === 'copilot') {
      upsertBlock(path.join(root, '.github', 'copilot-instructions.md'), block, '');
      written.push('.github/copilot-instructions.md');
    } else if (p.id === 'cursor') {
      writeMcpJson(path.join(root, '.cursor', 'mcp.json'), db, false);
      upsertBlock(path.join(root, '.cursor', 'rules', 'access-vsc.mdc'), block, '---\ndescription: Access VSC（Microsoft Access を VS Code で開発）の作業手順\nalwaysApply: true\n---\n\n');
      written.push('.cursor/mcp.json', '.cursor/rules/access-vsc.mdc');
    } else if (p.id === 'gemini') {
      writeMcpJson(path.join(root, '.gemini', 'settings.json'), db, false);
      upsertBlock(path.join(root, 'GEMINI.md'), block, '# GEMINI.md\n\n');
      written.push('.gemini/settings.json', 'GEMINI.md');
    } else if (p.id === 'agents') {
      upsertBlock(path.join(root, 'AGENTS.md'), block, '# AGENTS.md\n\n');
      written.push('AGENTS.md');
    }
  }
  ensureGitignore(root);
  const extra = picks.some((p) => ['claude', 'cursor', 'gemini'].includes(p.id)) ? ' エージェントを再起動すると MCP サーバー「access」が使えるようになります（利用許可を求められたら許可してください）。' : '';
  const c = await vscode.window.showInformationMessage(`AI エージェント連携を設定しました: ${written.join(', ')}。${extra}`, '開く');
  if (c) for (const f of written) vscode.window.showTextDocument(vscode.Uri.file(path.join(root, f)), { preview: false });
}

function ensureGitignore(root) {
  const file = path.join(root, '.gitignore');
  let text = '';
  try {
    text = fs.readFileSync(file, 'utf8');
  } catch (_) { /* none */ }
  const need = ['.accessvsc/', '*.laccdb'].filter((l) => !text.split(/\r?\n/).includes(l));
  if (!need.length) return;
  fs.writeFileSync(file, (text ? text.replace(/\s*$/, '\n') : '') + '# Access VSC\n' + need.join('\n') + '\n', 'utf8');
}

// ------------------------------------------------------------
// ユニットテスト（VS Code の「テスト」ビューに表示）
//   modules/*.bas の Public Sub Test_*() をテストとして扱う
// ------------------------------------------------------------
let testController;
let lintDiagnostics;
let lintTimer;
const TEST_RE = /^\s*(?:Public\s+)?(?:Sub|Function)\s+(Test_[\w\u0080-￿]*)\s*\(\s*\)/i;

async function discoverTests() {
  if (!testController) return;
  const db = await getDb(false);
  const items = [];
  const dir = db ? path.join(srcFor(db), 'modules') : undefined;
  if (dir && fs.existsSync(dir)) {
    for (const f of fs.readdirSync(dir).filter((n) => n.toLowerCase().endsWith('.bas')).sort()) {
      const file = path.join(dir, f);
      const module = decodeName(path.basename(f, path.extname(f)));
      const tests = [];
      fs.readFileSync(file, 'utf8')
        .split(/\r?\n/)
        .forEach((l, i) => {
          const m = TEST_RE.exec(l);
          if (m) tests.push({ name: m[1], line: i });
        });
      if (!tests.length) continue;
      const uri = vscode.Uri.file(file);
      const mod = testController.createTestItem(`mod:${module}`, module, uri);
      for (const t of tests) {
        const it = testController.createTestItem(`${module}.${t.name}`, t.name, uri);
        it.range = new vscode.Range(t.line, 0, t.line, 0);
        mod.children.add(it);
      }
      items.push(mod);
    }
  }
  testController.items.replace(items);
}

async function runTestsHandler(request, token) {
  const run = testController.createTestRun(request, 'Access VBA テスト');
  const leaves = [];
  const collect = (item) => (item.children.size ? item.children.forEach(collect) : leaves.push(item));
  if (request.include) request.include.forEach(collect);
  else testController.items.forEach(collect);
  const excluded = new Set((request.exclude || []).map((i) => i.id));
  const targets = leaves.filter((i) => !excluded.has(i.id) && !(i.parent && excluded.has(i.parent.id)));
  try {
    targets.forEach((t) => run.enqueued(t));
    const db = await getDb(true);
    if (!db || !targets.length || !(await ensureImported(db)) || token.isCancellationRequested) {
      targets.forEach((t) => run.skipped(t));
      return;
    }
    targets.forEach((t) => run.started(t));
    const res = await access('test', request.include ? { filter: targets.map((t) => t.id) } : {}, { db, title: 'VBA テストを実行中...' });
    if (!res) {
      targets.forEach((t) => run.skipped(t));
      return;
    }
    const byId = new Map(targets.map((t) => [t.id, t]));
    const tests = (res.result && res.result.tests) || [];
    for (const r of tests) {
      const item = byId.get(r.id);
      if (!item) continue;
      byId.delete(r.id);
      if (r.status === 'passed') {
        run.passed(item, r.ms);
        continue;
      }
      const loc = r.location;
      const msg = new vscode.TestMessage(`${r.message}${loc && loc.code ? `\n  場所: ${loc.module} ${loc.line} 行目  ${loc.code.trim()}` : ''}`);
      if (loc && loc.file && loc.fileLine) msg.location = new vscode.Location(vscode.Uri.file(loc.file), new vscode.Position(loc.fileLine - 1, 0));
      else if (item.uri && item.range) msg.location = new vscode.Location(item.uri, item.range);
      if (r.status === 'failed') run.failed(item, msg, r.ms);
      else run.errored(item, msg, r.ms);
    }
    const ce = res.result && res.result.compile && res.result.compile.error;
    if (ce) applyDiagnostics({ command: 'compile', result: { error: ce } });
    for (const item of byId.values()) {
      const m = new vscode.TestMessage(ce ? `コンパイル エラーのため実行できません: ${ce.message}` : res.error || res.message || '実行されませんでした');
      if (ce && ce.file && ce.fileLine) m.location = new vscode.Location(vscode.Uri.file(ce.file), new vscode.Position(ce.fileLine - 1, 0));
      run.errored(item, m);
    }
    const s = res.result && res.result.summary;
    if (s) run.appendOutput(`成功 ${s.passed} / 失敗 ${s.failed} / エラー ${s.errors}（全 ${s.total} 件, ${s.ms} ms）\r\n`);
    for (const w of res.warnings || []) run.appendOutput(`注意: ${w}\r\n`);
    if (res.result && res.result.junit) run.appendOutput(`JUnit: ${res.result.junit}\r\n`);
  } finally {
    run.end();
  }
}

async function cmdRunTests() {
  await discoverTests();
  if (!testController.items.size) {
    const c = await vscode.window.showInformationMessage(
      'テストがまだありません。テストは標準モジュールの「Public Sub Test_○○()」です。テストの例とアサーション（Assert）を追加しますか？',
      'テストの準備をする'
    );
    if (c) cmdSetupCi();
    return;
  }
  vscode.commands.executeCommand('workbench.view.testing.focus');
  const cts = new vscode.CancellationTokenSource();
  await runTestsHandler(new vscode.TestRunRequest(), cts.token);
}

// ------------------------------------------------------------
// コードチェック（lint。Access 不要）
// ------------------------------------------------------------
function applyLint(res) {
  lintDiagnostics.clear();
  if (!res || !res.result || !Array.isArray(res.result.problems)) return;
  const byFile = new Map();
  for (const p of res.result.problems) {
    const file = path.join(res.source, ...String(p.file).split('/'));
    const ln = Math.max(0, (p.line || 1) - 1);
    const d = new vscode.Diagnostic(
      new vscode.Range(ln, 0, ln, 10000),
      p.message,
      p.severity === 'error' ? vscode.DiagnosticSeverity.Error : p.severity === 'info' ? vscode.DiagnosticSeverity.Information : vscode.DiagnosticSeverity.Warning
    );
    d.source = 'Access lint';
    d.code = p.rule;
    if (!byFile.has(file)) byFile.set(file, []);
    byFile.get(file).push(d);
  }
  for (const [file, list] of byFile) lintDiagnostics.set(vscode.Uri.file(file), list);
}

async function cmdLint(quiet) {
  const db = quiet ? await getDb(false) : await getDb(true);
  if (!db) return;
  const res = await access('lint', {}, { db, title: 'コードチェック中...', quiet: quiet === true });
  applyLint(res);
  if (quiet === true || !res) return res;
  const r = res.result || {};
  if (!r.problems) return report(res, () => '');
  const msg = `コードチェック: エラー ${r.errorCount} 件 / 警告 ${r.warningCount} 件 / 情報 ${r.infoCount || 0} 件（${r.files} ファイル）`;
  if (r.errorCount || r.warningCount) {
    const c = await vscode.window.showWarningMessage(msg, '問題パネルを開く');
    if (c) vscode.commands.executeCommand('workbench.actions.view.problems');
  } else vscode.window.showInformationMessage(`${msg} — 問題はありません`);
  return res;
}

// ------------------------------------------------------------
// テスト・CI/CD の準備
// ------------------------------------------------------------
async function cmdSetupCi() {
  const db = await getDb(true);
  if (!db) return;
  if (!fs.existsSync(path.join(srcFor(db), MANIFEST))) {
    vscode.window.showWarningMessage('先に「Access から取り出す（エクスポート）」を実行してください。');
    return;
  }
  const root = workspaceRoot() || path.dirname(db);
  let gitInit = false;
  if (!fs.existsSync(path.join(root, '.git'))) {
    const c = await vscode.window.showInformationMessage(
      'このフォルダはまだ git で管理されていません。git を始めますか？',
      { modal: true, detail: 'git で変更履歴を管理すると、誰がいつ何を変えたかが残り、コミット前の自動チェックも使えます。' },
      'git を始める',
      'git なしで続ける'
    );
    if (!c) return;
    gitInit = c === 'git を始める';
  }
  const res = await access('setup-ci', { root, gitInit }, { db, title: 'テスト・CI/CD を準備中...' });
  if (!res || !res.ok) return report(res, () => '');
  const r = res.result;
  const rel = (p) => p;
  const lines = [
    '# テスト・CI/CD の準備ができました',
    '',
    '## 作成・更新したもの',
    '',
    ...[...r.created, ...r.updated].map((f) => `- \`${rel(f)}\``),
    ...(r.skipped.length ? ['', '既にあったので変更していないもの:', '', ...r.skipped.map((f) => `- \`${f}\``)] : []),
    '',
    '## メモ',
    '',
    ...r.notes.map((n) => `- ${n}`),
    '',
    '## 次にやること',
    '',
    '1. 左の Access →「③ Access へ反映する」で `Assert` と `Test_Example` を DB に入れる',
    '2. 「テストを実行」（または VS Code の「テスト」ビュー）でテストが通ることを確認',
    '3. 自分のテストを書く：標準モジュールに `Public Sub Test_○○()` を作り、`Assert.AreEqual 期待値, 実際の値` で確認',
    '4. GitHub に push すると `.github/workflows/access-ci.yml` が動きます',
    '   - **lint**（コードチェック）は GitHub のクラウド環境でそのまま動きます',
    '   - **test** は Access の入った Windows PC を「セルフホストランナー」として登録する必要があります',
    '     （リポジトリ → Settings → Actions → Runners → New self-hosted runner、ラベル `access` を追加。',
    '     サービスではなくログオン中のユーザーで `run.cmd` を起動しておく）',
  ];
  const doc = await vscode.workspace.openTextDocument({ language: 'markdown', content: lines.join('\n') });
  await vscode.window.showTextDocument(doc, { preview: true });
  scheduleStatusRefresh();
  discoverTests();
  const c = await vscode.window.showInformationMessage('追加したテスト用モジュールを Access へ反映して、テストを実行しますか？', '反映してテスト');
  if (c) {
    const ir = await cmdImport();
    if (ir && ir.ok) cmdRunTests();
  }
}

// ------------------------------------------------------------
// 帳票スナップショット（見た目の回帰テスト）
// ------------------------------------------------------------
let snapshotPanel;
const SNAP_LABEL = { same: '同じ', changed: '変化あり', new: '基準なし', updated: '基準を更新', error: 'エラー', added: 'ページ追加', removed: 'ページ削除' };

async function cmdSnapshot(arg) {
  const db = await getDb(true);
  if (!db) return;
  const o = arg && arg.objName ? { name: arg.objName } : {};
  const res = await access('snapshot', o.name ? { reports: [o.name] } : {}, { db, title: '帳票の見た目を比較中...' });
  if (!res) return;
  if (!res.result || !res.result.reports) return report(res, () => '');
  showSnapshotPanel(db, res);
  const s = res.result.summary;
  if (s.changed || s.errors) vscode.window.showWarningMessage(`帳票の見た目: ${s.changed} 件変化 / ${s.errors} 件エラー（同じ ${s.same} 件）`);
  else if (s.new) vscode.window.showInformationMessage(`帳票の見た目: 基準画像のない帳票が ${s.new} 件あります。内容を確認して「この見た目を基準にする」を押してください。`);
  else vscode.window.showInformationMessage(`帳票の見た目: ${s.same} 件すべて基準どおりです`);
}

async function cmdSnapshotAccept(arg) {
  const db = await getDb(true);
  if (!db) return;
  let names;
  if (arg && arg.objName) names = [arg.objName];
  else if (typeof arg === 'string') names = [arg];
  else {
    const c = await vscode.window.showWarningMessage('すべての帳票の今の見た目を基準にします。よろしいですか？', { modal: true }, '基準にする');
    if (!c) return;
  }
  const res = await access('snapshot', names ? { reports: names, update: true } : { update: true }, { db, title: '基準画像を更新中...' });
  if (!res) return;
  await report(res, (r) => `基準画像を更新しました（${r.result.summary.updated} 件）。snapshots フォルダを git にコミットしてください。`);
  if (snapshotPanel && res.ok) showSnapshotPanel(db, res);
}

function showSnapshotPanel(db, res) {
  const r = res.result;
  const roots = [vscode.Uri.file(path.join(path.dirname(db), '.accessvsc')), vscode.Uri.file(r.baselineDir || srcFor(db))];
  if (!snapshotPanel) {
    snapshotPanel = vscode.window.createWebviewPanel('accessVsc.snapshot', '帳票の見た目', vscode.ViewColumn.Active, { enableCommandUris: true, localResourceRoots: roots });
    snapshotPanel.onDidDispose(() => (snapshotPanel = undefined));
  }
  const wv = snapshotPanel.webview;
  const stamp = Date.now();
  const img = (p, cap) => (p ? `<figure><img src="${wv.asWebviewUri(vscode.Uri.file(p))}?t=${stamp}"><figcaption>${cap}</figcaption></figure>` : `<figure class="none"><div>なし</div><figcaption>${cap}</figcaption></figure>`);
  const order = { changed: 0, error: 1, new: 2, updated: 3, same: 4 };
  const reports = [...r.reports].sort((a, b) => order[a.status] - order[b.status]);
  const body = reports
    .map((x) => {
      const pages = (x.pages || [])
        .filter((p) => x.status !== 'same' && p.status !== 'same')
        .map(
          (p) => `<div class="page"><h4>${p.page} ページ — ${SNAP_LABEL[p.status] || p.status}${p.diffPercent ? `（${p.diffPercent}% のピクセルが変化）` : ''}</h4>
          <div class="row">${p.status === 'new' || p.status === 'updated' ? img(p.current, '今の見た目') : img(p.baseline, '基準') + img(p.current, '今') + img(p.diff, '差分（赤＝違う所）')}</div></div>`
        )
        .join('');
      const accept = ['changed', 'new'].includes(x.status) ? `<a class="btn primary" href="${cmdLink('accessVsc.snapshotAccept', [x.name])}">この見た目を基準にする</a>` : '';
      const preview = `<a class="btn" href="${cmdLink('accessVsc.previewReport', [{ objName: x.name, kind: 'reports' }])}">プレビュー</a>`;
      return `<section class="${x.status}"><h3><span class="chip ${x.status}">${SNAP_LABEL[x.status] || x.status}</span> ${escHtml(x.name)} ${accept} ${preview}</h3>${x.message ? `<p class="msg">${escHtml(x.message)}</p>` : ''}${pages}</section>`;
    })
    .join('');
  const s = r.summary;
  wv.html = `<!DOCTYPE html><html lang="ja"><head><meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src ${wv.cspSource}; style-src 'unsafe-inline';">
<style>
 body { font-family: var(--vscode-font-family); color: var(--vscode-foreground); padding: 0 16px 24px; }
 .btn { display: inline-block; margin-left: 8px; padding: 2px 10px; border-radius: 2px; font-size: 12px; font-weight: normal; text-decoration: none; background: var(--vscode-button-secondaryBackground); color: var(--vscode-button-secondaryForeground); }
 .btn.primary { background: var(--vscode-button-background); color: var(--vscode-button-foreground); }
 .chip { font-size: 11px; padding: 1px 8px; border-radius: 10px; margin-right: 6px; border: 1px solid; }
 .chip.changed, .chip.error { color: var(--vscode-errorForeground); }
 .chip.new { color: var(--vscode-editorWarning-foreground); }
 .chip.same, .chip.updated { color: var(--vscode-testing-iconPassed, #3c3); }
 section { border-top: 1px solid var(--vscode-panel-border); padding: 8px 0; }
 .row { display: flex; gap: 12px; flex-wrap: wrap; }
 figure { margin: 0; flex: 1 1 30%; min-width: 220px; }
 figure img { width: 100%; background: #fff; box-shadow: 0 1px 6px rgba(0,0,0,.35); }
 figure.none div { height: 120px; display: flex; align-items: center; justify-content: center; border: 1px dashed var(--vscode-panel-border); opacity: .6; }
 figcaption { font-size: 85%; opacity: .75; margin-top: 4px; }
 .msg { opacity: .8; }
 .top { margin: 12px 0; }
</style></head><body>
<h2>帳票の見た目（スナップショット）</h2>
<div class="top">同じ ${s.same} / 変化あり ${s.changed} / 基準なし ${s.new} / 更新 ${s.updated} / エラー ${s.errors}
 <a class="btn" href="${cmdLink('accessVsc.snapshot', [])}">もう一度比較</a>
 ${s.changed || s.new ? `<a class="btn" href="${cmdLink('accessVsc.snapshotAccept', [])}">すべて基準にする</a>` : ''}</div>
${(res.warnings || []).map((w) => `<p class="msg">⚠ ${escHtml(w)}</p>`).join('')}
${body}
<p class="msg">基準画像は <code>${escHtml(r.baselineDir || '')}</code> にあります。git にコミットしておくと、CI で見た目の変化を検出できます。</p>
</body></html>`;
  snapshotPanel.reveal();
}

// ------------------------------------------------------------
// 配布（リリース）
// ------------------------------------------------------------
function deployConfigPath(db) {
  return path.join(path.dirname(db), 'access-deploy.json');
}

function readDeployConfig(db) {
  try {
    return JSON.parse(fs.readFileSync(deployConfigPath(db), 'utf8').replace(/^﻿/, ''));
  } catch (_) {
    return undefined;
  }
}

async function cmdDeployInit() {
  const db = await getDb(true);
  if (!db) return;
  const cur = readDeployConfig(db) || {};
  const appName = await vscode.window.showInputBox({ prompt: 'アプリ名（配布ファイル名・利用者 PC のフォルダ名になります）', value: cur.appName || path.basename(db, path.extname(db)), validateInput: (v) => (/^[^\\/:*?"<>|]+$/.test(v) ? undefined : 'ファイル名に使えない文字があります') });
  if (!appName) return;
  let deployTo = await vscode.window.showInputBox({ prompt: '配布先の共有フォルダ（利用者が起動する場所）。例: \\\\server\\share\\Shop', value: cur.deployTo || '', placeHolder: '空欄ならフォルダを選ぶ画面を開きます' });
  if (deployTo === undefined) return;
  if (!deployTo) {
    const pick = await vscode.window.showOpenDialog({ canSelectFolders: true, canSelectFiles: false, openLabel: 'ここに配布する' });
    if (!pick) return;
    deployTo = pick[0].fsPath;
  }
  const be = await vscode.window.showQuickPick(
    [
      { label: 'リンクテーブルの接続先はそのまま', id: 'keep' },
      { label: '本番のバックエンド DB に張り替える', description: 'ビルド時に、リンクテーブルの接続先を本番用に変更します', id: 'set' },
    ],
    { placeHolder: 'リンクテーブル（バックエンド）の扱い' }
  );
  if (!be) return;
  let backendPath = '';
  if (be.id === 'set') {
    backendPath = await vscode.window.showInputBox({ prompt: '本番のバックエンド DB のパス。例: \\\\server\\data\\Shop_be.accdb', value: (cur.backend && cur.backend.path) || '' });
    if (!backendPath) return;
  }
  const res = await access('deploy-init', { appName, deployTo, backendPath, force: true, root: workspaceRoot() }, { db, title: '配布の設定を作成中...' });
  if (!res || !res.ok) return report(res, () => '');
  vscode.window.showTextDocument(vscode.Uri.file(res.result.config), { preview: false });
  vscode.window.showInformationMessage(`配布の設定を作成しました（${path.basename(res.result.config)}、リリース用ワークフロー）。`);
}

async function cmdBuild() {
  const db = await getDb(true);
  if (!db) return;
  if (!readDeployConfig(db)) {
    const c = await vscode.window.showInformationMessage('配布の設定（access-deploy.json）がまだありません。先に設定しますか？', '設定する', '設定せずにビルド');
    if (!c) return;
    if (c === '設定する') return cmdDeployInit();
  }
  if (!(await ensureImported(db))) return;
  const version = await vscode.window.showInputBox({ prompt: 'バージョン（空欄なら日時）', placeHolder: '例: 1.2.0' });
  if (version === undefined) return;
  const res = await access('build', version ? { version } : {}, { db, title: '.accde を作成中（Access のウィンドウが一瞬表示されます）...' });
  await report(res, (r) => `リリースを作成しました: ${path.basename(r.result.accde)}${(r.result.relinked || []).length ? `（リンクテーブル ${r.result.relinked.length} 件を本番に張り替え）` : ''}`);
  return res;
}

async function cmdDeploy() {
  const db = await getDb(true);
  if (!db) return;
  const cfg = readDeployConfig(db);
  if (!cfg || !cfg.deployTo) {
    const c = await vscode.window.showInformationMessage('配布先が設定されていません。設定しますか？', '設定する');
    if (c) cmdDeployInit();
    return;
  }
  const outDir = path.resolve(path.dirname(db), cfg.output || 'dist');
  const builds = fs.existsSync(outDir)
    ? fs
        .readdirSync(outDir)
        .filter((f) => f.toLowerCase().endsWith('.accde'))
        .map((f) => ({ f, t: fs.statSync(path.join(outDir, f)).mtimeMs }))
        .sort((a, b) => b.t - a.t)
    : [];
  const items = [{ label: '$(tools) 新しくビルドして配布', id: 'build' }, ...builds.map((b) => ({ label: `$(package) ${b.f}`, description: new Date(b.t).toLocaleString(), id: path.join(outDir, b.f) }))];
  const pick = await vscode.window.showQuickPick(items, { placeHolder: `配布するもの（配布先: ${cfg.deployTo}）` });
  if (!pick) return;
  let version;
  if (pick.id === 'build') {
    if (!(await ensureImported(db))) return;
    version = await vscode.window.showInputBox({ prompt: 'バージョン（空欄なら日時）', placeHolder: '例: 1.2.0' });
    if (version === undefined) return;
  }
  const target = pick.id === 'build' ? `新しいビルド（${version || '日時の版'}）` : path.basename(pick.id);
  const ok = await vscode.window.showWarningMessage(
    `${target} を配布します。`,
    { modal: true, detail: `配布先: ${cfg.deployTo}\n\n利用者が次に起動したときに、この版に置き換わります。以前の版は releases フォルダに残ります。` },
    '配布する'
  );
  if (ok !== '配布する') return;
  const p = { yes: true };
  if (version) p.version = version;
  if (pick.id !== 'build') p.from = pick.id;
  const res = await access('deploy', p, { db, title: '配布中...' });
  await report(res, (r) => `バージョン ${r.result.version} を配布しました。利用者は ${path.basename(r.result.launcher)} から起動します。`);
}

async function cmdOpenDeployFolder() {
  const db = await getDb(true);
  if (!db) return;
  const cfg = readDeployConfig(db);
  if (!cfg || !cfg.deployTo) return vscode.window.showInformationMessage('配布先が設定されていません。');
  vscode.env.openExternal(vscode.Uri.file(cfg.deployTo));
}

// ------------------------------------------------------------
// 練習用サンプル（チュートリアル用）
// ------------------------------------------------------------
async function cmdCreateSample() {
  const pick = await vscode.window.showOpenDialog({
    canSelectFolders: true,
    canSelectFiles: false,
    openLabel: 'ここに作る',
    title: '練習用サンプルを作る場所を選んでください（中に「AccessVSC-練習」フォルダができます）',
  });
  if (!pick) return;
  let dir = path.join(pick[0].fsPath, 'AccessVSC-練習');
  if (fs.existsSync(dir)) {
    // 中級編・上級編は新しいサンプルから始められるので、別のフォルダに作れるようにする
    let n = 2;
    while (fs.existsSync(`${dir}${n}`)) n++;
    const c = await vscode.window.showQuickPick(
      [
        { label: '$(folder-opened) 今あるサンプルを開く', id: 'open' },
        { label: '$(add) 新しく作る', description: `AccessVSC-練習${n}`, id: 'new' },
      ],
      { placeHolder: '練習用サンプルは既にあります' }
    );
    if (!c) return;
    if (c.id === 'open') {
      vscode.commands.executeCommand('vscode.openFolder', vscode.Uri.file(dir), { forceNewWindow: true });
      return;
    }
    dir = `${dir}${n}`;
  }
  const db = path.join(dir, 'cafe.accdb');
  const out = await vscode.window.withProgress({ location: vscode.ProgressLocation.Notification, title: '練習用サンプルを作成中...' }, () =>
    new Promise((resolve) => {
      const p = cp.spawn(powershell(), ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', path.join(binDir(), 'New-SampleDatabase.ps1'), '-Path', db], { windowsHide: true });
      let text = '';
      p.stdout.on('data', (d) => (text += d));
      p.stderr.on('data', (d) => (text += d));
      p.on('close', (code) => resolve({ code, text }));
      p.on('error', (e) => resolve({ code: 1, text: String(e) }));
    })
  );
  output.appendLine(out.text);
  if (out.code !== 0 || !fs.existsSync(db)) {
    vscode.window.showErrorMessage('サンプルを作成できませんでした。「ログを表示」で詳細を確認してください。', 'ログを表示').then((c) => c && output.show(true));
    return;
  }
  const c = await vscode.window.showInformationMessage(`練習用サンプルを作りました（${db}）。新しいウィンドウで開きますか？`, '開く');
  if (c) vscode.commands.executeCommand('vscode.openFolder', vscode.Uri.file(dir), { forceNewWindow: true });
}

// ------------------------------------------------------------
// サイドバー（左のアクティビティバーの「Access」）
// ------------------------------------------------------------
const KIND_INFO = {
  modules: { label: 'モジュール（VBA）', icon: 'symbol-method' },
  forms: { label: 'フォーム', icon: 'window' },
  reports: { label: 'レポート（帳票）', icon: 'file-text' },
  queries: { label: 'クエリ', icon: 'filter' },
  macros: { label: 'マクロ', icon: 'zap' },
  tables: { label: 'テーブル（定義・読み取り専用）', icon: 'table' },
};

class ActionsProvider {
  constructor() {
    this._em = new vscode.EventEmitter();
    this.onDidChangeTreeData = this._em.event;
  }
  refresh() {
    this._em.fire();
  }
  getTreeItem(el) {
    return el;
  }
  async getChildren(el) {
    if (el) return el.children || [];
    const db = await getDb(false);
    if (!db) return []; // → ようこそ表示（データベースを選択）
    const st = syncStatus(srcFor(db));
    const pending = st.changed.length + st.added.length + st.missing.length;
    const action = (label, description, icon, command, args, tooltip) => {
      const t = new vscode.TreeItem(label, vscode.TreeItemCollapsibleState.None);
      t.description = description;
      t.iconPath = new vscode.ThemeIcon(icon);
      t.command = { command, title: label, arguments: args || [] };
      t.tooltip = tooltip || description;
      return t;
    };
    const group = (label, children) => {
      const t = new vscode.TreeItem(label, vscode.TreeItemCollapsibleState.Expanded);
      t.children = children;
      t.contextValue = 'group';
      return t;
    };
    const dbItem = action(path.basename(db), st.hasManifest ? (pending ? `未反映 ${pending} 件` : '同期済み') : '未取り出し', 'database', 'accessVsc.status', [], `${db}\nクリックで同期状態を表示`);
    return [
      dbItem,
      group('基本の流れ', [
        action('① Access から取り出す', 'エクスポート', 'cloud-download', 'accessVsc.export', [], 'Access の中身（VBA・クエリ・フォーム・レポート）をファイルに書き出します。最初に 1 回、または Access 側で直接変更したあとに実行します。'),
        action('② ファイルを編集する', '左の「オブジェクト」から開く', 'edit', 'accessVsc.focusObjects', [], 'オブジェクト一覧から編集したいものを開きます。AI に頼む場合は、チャットで指示してください。'),
        action('③ Access へ反映する', pending ? `インポート（${pending} 件）` : 'インポート', 'cloud-upload', 'accessVsc.import', [], '変更したファイルだけを Access に取り込み、VBA をコンパイルします。エラーがあれば該当行に赤線が出ます。'),
        action('④ 動かして確かめる', 'VBA を実行', 'play', 'accessVsc.runProcedure', [], 'VBA の Public Function / Sub を実行して結果を表示します。'),
      ]),
      group('確認する', [
        action('帳票をプレビュー', 'レポートを画像で表示', 'open-preview', 'accessVsc.previewReport', [], 'レポートを PDF / 画像にして、印刷イメージとレイアウトの問題を表示します。'),
        action('データを見る', 'SQL を実行', 'table', 'accessVsc.runSql'),
        action('コンパイルする', 'VBA の文法チェック', 'check-all', 'accessVsc.compile'),
        action('Access で開く', '', 'link-external', 'accessVsc.openInAccess', [], 'Access の画面で開きます。開いている間は「反映する」ができないので、終わったら閉じてください。'),
      ]),
      group('テスト・品質', [
        action('テストを実行', 'VBA のユニットテスト', 'beaker', 'accessVsc.runTests', [], 'Test_ で始まる Public Sub をすべて実行します。DB のコピーで実行するので、元のデータは変わりません。'),
        action('コードチェック', 'Access 不要・数秒', 'checklist', 'accessVsc.lint', [], 'Shift-JIS にない文字、Option Explicit の書き忘れ、帳票のはみ出しなどを調べます。'),
        action('帳票の見た目を比較', 'スナップショット', 'eye', 'accessVsc.snapshot', [], '全レポートを画像にして、基準画像と比べます。変わった所は赤で表示します。'),
        action('テスト・CI/CD を準備', 'Assert・GitHub Actions など', 'rocket', 'accessVsc.setupCi', [], 'テスト用の Assert モジュールとテスト例、GitHub Actions の設定、コミット前の自動チェックを追加します。'),
      ]),
      group('配布（リリース）', [
        action('配布の設定', 'access-deploy.json', 'gear', 'accessVsc.deployInit', [], '配布先の共有フォルダ、本番のバックエンド DB などを設定します。'),
        action('リリースを作成', '.accde を作る', 'package', 'accessVsc.build', [], '版番号を書き込み、リンクテーブルを本番に張り替えて .accde を作ります（まだ配布はしません）。'),
        action('配布する', '共有フォルダへ公開', 'cloud-upload', 'accessVsc.deploy', [], '作成したリリースを共有フォルダへ公開します。利用者は Start-○○.cmd から起動すると最新版になります。'),
        action('配布先を開く', '', 'folder-opened', 'accessVsc.openDeployFolder'),
      ]),
      group('AI と使う', [
        action('AI エージェント連携を設定', 'Claude Code / Copilot', 'sparkle', 'accessVsc.setupAgent', [], 'AI が Access を操作できるように .mcp.json や CLAUDE.md を作ります。'),
      ]),
      group('その他', [
        action('ドキュメント', 'チュートリアル・学習コース・リファレンス', 'book', 'accessVsc.openManual', [], 'ブラウザでドキュメントを開きます。'),
        action('練習用サンプルを作る', 'チュートリアル用', 'beaker', 'accessVsc.createSample', [], '練習用の小さなデータベース（受注管理）を作ります。自分のデータベースを使わずに練習できます。'),
        action('使い方ガイド', 'VS Code 内の簡単なガイド', 'list-ordered', 'workbench.action.openWalkthrough', [`${ctx.extension.id}#accessVsc.gettingStarted`, false]),
        action('すべてを反映', '全ファイルをインポート', 'repo-push', 'accessVsc.importAll'),
        action('データベースを切り替える', '', 'arrow-swap', 'accessVsc.selectDatabase'),
        action('ログを表示', '', 'output', 'accessVsc.showOutput'),
        action('設定', '', 'gear', 'workbench.action.openSettings', ['accessVsc']),
      ]),
    ];
  }
}

class ObjectsProvider {
  constructor() {
    this._em = new vscode.EventEmitter();
    this.onDidChangeTreeData = this._em.event;
  }
  refresh() {
    this._em.fire();
  }
  getTreeItem(el) {
    return el;
  }
  async getChildren(el) {
    const db = await getDb(false);
    if (!db) return [];
    const src = srcFor(db);
    if (!fs.existsSync(path.join(src, MANIFEST))) return []; // → ようこそ表示（取り出す）
    const st = syncStatus(src);
    const state = (rel) => (st.changed.includes(rel) ? 'changed' : st.added.includes(rel) ? 'added' : st.readOnlyChanged.includes(rel) ? 'readonly' : '');
    if (!el) {
      return Object.keys(KIND_INFO).map((kind) => {
        const dir = path.join(src, kind);
        const files = fs.existsSync(dir) ? fs.readdirSync(dir).filter((n) => MANAGED[kind].includes(path.extname(n).toLowerCase())) : [];
        const pending = files.filter((f) => state(`${kind}/${f}`)).length + st.missing.filter((m) => m.startsWith(`${kind}/`)).length;
        // 少ないものは最初から開いておく（初めての人に中身が見えるように）
        const collapse = !files.length ? vscode.TreeItemCollapsibleState.None : files.length <= 10 && kind !== 'tables' ? vscode.TreeItemCollapsibleState.Expanded : vscode.TreeItemCollapsibleState.Collapsed;
        const t = new vscode.TreeItem(KIND_INFO[kind].label, collapse);
        t.description = `${files.length}${pending ? `（未反映 ${pending}）` : ''}`;
        t.iconPath = new vscode.ThemeIcon(KIND_INFO[kind].icon);
        t.kindKey = kind;
        t.contextValue = `accessKind.${kind}`;
        return t;
      });
    }
    if (!el.kindKey) return [];
    const kind = el.kindKey;
    const dir = path.join(src, kind);
    const items = [];
    const files = fs.existsSync(dir) ? fs.readdirSync(dir).filter((n) => MANAGED[kind].includes(path.extname(n).toLowerCase())).sort() : [];
    for (const f of files) {
      const rel = `${kind}/${f}`;
      const name = decodeName(path.basename(f, path.extname(f)));
      const s = state(rel);
      const t = new vscode.TreeItem(name, vscode.TreeItemCollapsibleState.None);
      t.objName = name;
      t.kind = kind;
      t.resourceUri = vscode.Uri.file(path.join(dir, f));
      t.description = s === 'changed' ? '● 変更あり（未反映）' : s === 'added' ? '＋ 新規（未反映）' : s === 'readonly' ? '⚠ 編集しても反映されません' : kind === 'modules' ? (f.endsWith('.cls') ? 'クラス' : '') : '';
      t.command = { command: 'vscode.open', title: '開く', arguments: [t.resourceUri] };
      const std = kind === 'modules' && f.toLowerCase().endsWith('.bas') ? '.std' : '';
      t.contextValue = `accessObj.${kind}${std}${s === 'changed' || s === 'added' ? '.pending' : ''}`;
      const hints = {
        modules: std ? '右クリック →「プロシージャを実行」' : '',
        reports: '右クリック →「プレビュー」「ページ設定」',
        forms: '右クリック →「プレビュー」',
        queries: '▶ でクエリを実行',
        tables: '▶ で先頭 100 行を表示',
      };
      t.tooltip = new vscode.MarkdownString(`**${name}**\n\n${rel}${hints[kind] ? `\n\n${hints[kind]}` : ''}`);
      items.push(t);
    }
    for (const m of st.missing.filter((x) => x.startsWith(`${kind}/`))) {
      const t = new vscode.TreeItem(decodeName(path.basename(m, path.extname(m))), vscode.TreeItemCollapsibleState.None);
      t.description = '✕ ファイル削除（Access には残っています）';
      t.iconPath = new vscode.ThemeIcon('trash');
      t.tooltip = '「Access へ反映する」で Access からも削除するか選べます。';
      items.push(t);
    }
    return items;
  }
}

let actionsProvider;
let objectsProvider;

function refreshTrees() {
  if (actionsProvider) actionsProvider.refresh();
  if (objectsProvider) objectsProvider.refresh();
}

// ------------------------------------------------------------
// ステータスバー・監視・CodeLens
// ------------------------------------------------------------
function scheduleStatusRefresh() {
  clearTimeout(statusTimer);
  statusTimer = setTimeout(refreshStatus, 300);
}

let currentDb; // 自動検出したデータベースも含め、今使っているもの（保存時の処理で使う）

async function refreshStatus() {
  const db = await getDb(false);
  currentDb = db;
  vscode.commands.executeCommand('setContext', 'accessVsc.hasDb', !!db);
  vscode.commands.executeCommand('setContext', 'accessVsc.exported', !!db && fs.existsSync(path.join(srcFor(db), MANIFEST)));
  refreshTrees();
  discoverTests().catch(() => undefined);
  if (!db) {
    statusItem.hide();
    return;
  }
  let st;
  try {
    st = syncStatus(srcFor(db));
  } catch (_) {
    st = { changed: [], added: [], missing: [], readOnlyChanged: [], hasManifest: false };
  }
  const n = st.changed.length + st.added.length + st.missing.length;
  statusItem.text = `$(database) ${path.basename(db)}${n ? ` $(circle-filled) ${n}` : ''}`;
  statusItem.tooltip = !st.hasManifest
    ? '未エクスポート。クリックしてメニューを開く'
    : n
      ? `未インポートの変更 ${n} 件\n${[...st.changed, ...st.added, ...st.missing].join('\n')}\n\nクリックしてメニューを開く`
      : 'ソースと Access は同期しています。クリックしてメニューを開く';
  statusItem.show();
  setupWatcher(srcFor(db));
}

let watchedSrc;
function setupWatcher(src) {
  if (watchedSrc === src) return;
  if (watcher) watcher.dispose();
  watchedSrc = src;
  watcher = vscode.workspace.createFileSystemWatcher(new vscode.RelativePattern(vscode.Uri.file(src), '**/*'));
  watcher.onDidChange(scheduleStatusRefresh);
  watcher.onDidCreate(scheduleStatusRefresh);
  watcher.onDidDelete(scheduleStatusRefresh);
}

function onSave(doc) {
  scheduleStatusRefresh();
  const dbForLint = configuredDb() || currentDb;
  if (cfg().get('lintOnSave') !== false && dbForLint && kindOf(relOf(srcFor(dbForLint), doc.fileName))) {
    clearTimeout(lintTimer);
    lintTimer = setTimeout(() => cmdLint(true), 800);
  }
  if (doc.languageId === 'vba') discoverTests();
  if (!cfg().get('importOnSave')) return;
  const db = configuredDb() || currentDb;
  if (!db) return;
  const rel = relOf(srcFor(db), doc.fileName);
  const k = kindOf(rel);
  if (!k || k.readOnly) return;
  pendingSaves.add(doc.fileName);
  clearTimeout(saveTimer);
  saveTimer = setTimeout(() => {
    const files = [...pendingSaves];
    pendingSaves.clear();
    access('import', { files }, { db, title: '保存したファイルをインポート中...' }).then((r) => report(r, importSummary));
  }, 500);
}

const codeLensProvider = {
  onDidChangeCodeLenses: lensChanged.event,
  provideCodeLenses(doc) {
    const lenses = [];
    const ext = path.extname(doc.fileName).toLowerCase();
    if (doc.languageId === 'vba' && isStandardModule(doc)) {
      for (const p of parseProcedures(doc)) {
        const range = new vscode.Range(p.line, 0, p.line, 0);
        lenses.push(new vscode.CodeLens(range, { title: '▶ Access で実行', command: 'accessVsc.runProcedure', arguments: [p.name, p.params] }));
      }
    } else if (ext === '.sql' && path.basename(path.dirname(doc.fileName)).toLowerCase() === 'queries') {
      const name = decodeName(path.basename(doc.fileName, ext));
      lenses.push(new vscode.CodeLens(new vscode.Range(0, 0, 0, 0), { title: '▶ 保存済みクエリを実行', command: 'accessVsc.runSql', arguments: [name] }));
      lenses.push(new vscode.CodeLens(new vscode.Range(0, 0, 0, 0), { title: '▶ この SQL を実行', command: 'accessVsc.runSql', arguments: [doc.getText()] }));
    } else if (ext === '.txt' && ['reports', 'forms'].includes(path.basename(path.dirname(doc.fileName)).toLowerCase())) {
      const kind = path.basename(path.dirname(doc.fileName)).toLowerCase();
      const target = { objName: decodeName(path.basename(doc.fileName, ext)), kind, resourceUri: doc.uri };
      const top = new vscode.Range(0, 0, 0, 0);
      lenses.push(new vscode.CodeLens(top, { title: '$(open-preview) プレビュー', command: 'accessVsc.previewReport', arguments: [target] }));
      lenses.push(new vscode.CodeLens(top, { title: '$(symbol-ruler) レイアウト詳細', command: 'accessVsc.reportLayout', arguments: [target] }));
      if (kind === 'reports') lenses.push(new vscode.CodeLens(top, { title: '$(settings) ページ設定', command: 'accessVsc.pageSetup', arguments: [target] }));
    }
    return lenses;
  },
};

// ------------------------------------------------------------
// VS Code（Copilot など）への MCP サーバー登録
// ------------------------------------------------------------
function registerMcpProvider() {
  if (!vscode.lm || typeof vscode.lm.registerMcpServerDefinitionProvider !== 'function' || !vscode.McpStdioServerDefinition) return;
  ctx.subscriptions.push(
    vscode.lm.registerMcpServerDefinitionProvider('accessVsc.mcp', {
      onDidChangeMcpServerDefinitions: mcpChanged.event,
      provideMcpServerDefinitions: async () => {
        const db = await getDb(false);
        if (!db) return [];
        const args = ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', mcpServerPath(), '-DbPath', db];
        if (cfg().get('sourceFolder')) args.push('-Src', srcFor(db));
        return [new vscode.McpStdioServerDefinition(`Access (${path.basename(db)})`, powershell(), args, {}, ctx.extension.packageJSON.version)];
      },
      resolveMcpServerDefinition: async (server) => server,
    })
  );
}

// ------------------------------------------------------------
// activate / deactivate
// ------------------------------------------------------------
function activate(context) {
  ctx = context;
  output = vscode.window.createOutputChannel('Access VSC');
  diagnostics = vscode.languages.createDiagnosticCollection('access');
  statusItem = vscode.window.createStatusBarItem(vscode.StatusBarAlignment.Left, 50);
  statusItem.command = 'accessVsc.menu';
  context.subscriptions.push(output, diagnostics, statusItem, mcpChanged, lensChanged);

  if (process.platform !== 'win32') {
    output.appendLine('Access VSC は Windows（Microsoft Access がインストールされた環境）専用です。');
    return;
  }
  try {
    installScripts();
  } catch (e) {
    output.appendLine(`スクリプトのコピーに失敗: ${e}`);
  }

  const reg = (id, fn) => context.subscriptions.push(vscode.commands.registerCommand(id, fn));
  reg('accessVsc.menu', cmdMenu);
  reg('accessVsc.selectDatabase', cmdSelectDatabase);
  reg('accessVsc.export', cmdExport);
  reg('accessVsc.import', () => cmdImport());
  reg('accessVsc.importAll', () => cmdImport({ all: true }));
  reg('accessVsc.importFile', cmdImportFile);
  reg('accessVsc.compile', cmdCompile);
  reg('accessVsc.runProcedure', cmdRunProcedure);
  reg('accessVsc.runSql', cmdRunSql);
  reg('accessVsc.status', cmdStatus);
  reg('accessVsc.openInAccess', cmdOpenInAccess);
  reg('accessVsc.setupAgent', cmdSetupAgent);
  reg('accessVsc.showOutput', () => output.show(true));
  reg('accessVsc.previewReport', cmdPreviewReport);
  reg('accessVsc.reportLayout', cmdReportLayout);
  reg('accessVsc.pageSetup', cmdPageSetup);
  reg('accessVsc.refresh', () => scheduleStatusRefresh());
  reg('accessVsc.focusObjects', () => vscode.commands.executeCommand('accessVsc.objects.focus'));
  // 画面内のリンクから呼ばれる。実行ファイルなどを開かないよう、PDF だけに限る
  reg('accessVsc.openExternalFile', (p) => {
    if (typeof p === 'string' && /\.pdf$/i.test(p) && fs.existsSync(p)) vscode.env.openExternal(vscode.Uri.file(p));
  });

  reg('accessVsc.runTests', cmdRunTests);
  reg('accessVsc.lint', () => cmdLint(false));
  reg('accessVsc.setupCi', cmdSetupCi);
  reg('accessVsc.snapshot', cmdSnapshot);
  reg('accessVsc.snapshotAccept', cmdSnapshotAccept);
  reg('accessVsc.deployInit', cmdDeployInit);
  reg('accessVsc.build', cmdBuild);
  reg('accessVsc.deploy', cmdDeploy);
  reg('accessVsc.openDeployFolder', cmdOpenDeployFolder);
  reg('accessVsc.openManual', (page) => {
    // docs フォルダの HTML だけ（パスをたどって別のファイルを開かないように）
    const name = typeof page === 'string' && /^[a-z][a-z0-9-]*\.html$/.test(page) ? page : 'index.html';
    vscode.env.openExternal(vscode.Uri.file(path.join(context.extensionPath, 'docs', name)));
  });
  reg('accessVsc.createSample', cmdCreateSample);

  lintDiagnostics = vscode.languages.createDiagnosticCollection('access-lint');
  testController = vscode.tests.createTestController('accessVsc.tests', 'Access VBA');
  testController.resolveHandler = () => discoverTests();
  testController.refreshHandler = () => discoverTests();
  testController.createRunProfile('実行', vscode.TestRunProfileKind.Run, runTestsHandler, true);
  context.subscriptions.push(lintDiagnostics, testController);

  // VBA の編集支援（アウトライン・定義へ移動・参照・ホバー・補完・引数のヒント・折りたたみ）
  require('./vba-language').register(context, () => {
    const db = configuredDb() || currentDb;
    return db ? srcFor(db) : undefined;
  });

  actionsProvider = new ActionsProvider();
  objectsProvider = new ObjectsProvider();
  context.subscriptions.push(
    vscode.window.registerTreeDataProvider('accessVsc.actions', actionsProvider),
    vscode.window.registerTreeDataProvider('accessVsc.objects', objectsProvider)
  );

  context.subscriptions.push(
    vscode.languages.registerCodeLensProvider(
      [{ language: 'vba' }, { pattern: '**/queries/*.sql' }, { pattern: '**/reports/*.txt' }, { pattern: '**/forms/*.txt' }],
      codeLensProvider
    ),
    vscode.workspace.onDidSaveTextDocument(onSave),
    vscode.workspace.onDidChangeConfiguration((e) => {
      if (e.affectsConfiguration('accessVsc')) {
        mcpChanged.fire();
        lensChanged.fire();
        scheduleStatusRefresh();
      }
    }),
    { dispose: () => watcher && watcher.dispose() }
  );
  registerMcpProvider();
  scheduleStatusRefresh();
}

function deactivate() {
  // 常駐プロセスを止める（開いたままの Access も閉じる）
  if (accessHost) accessHost.kill();
}

module.exports = { activate, deactivate };
