// VBA の編集支援（言語サーバーを使わない軽量版）
//   アウトライン / ワークスペースのシンボル / 定義へ移動（F12）/ 参照の検索（Shift+F12）
//   ホバー / 入力補完 / 引数のヒント / 折りたたみ / 整形（Shift+Alt+F）/ 名前の変更（F2・安全な範囲だけ）
// 対象：modules/*.bas, *.cls と、forms・reports の .txt の CodeBehindForm 以降
'use strict';

const vscode = require('vscode');
const fs = require('fs');
const path = require('path');

const ID = '[A-Za-z_\\u0080-\\uFFFF][\\w\\u0080-\\uFFFF]*';
const PROC_START = new RegExp(`^\\s*(?:(Public|Private|Friend|Global)\\s+)?(?:(Static)\\s+)?(Sub|Function|Property\\s+(?:Get|Let|Set))\\s+(${ID})\\s*\\(`, 'i');
const PROC_END = /^\s*End\s+(Sub|Function|Property)\b/i;
const DECLARE = new RegExp(`^\\s*(?:(Public|Private)\\s+)?Declare\\s+(?:PtrSafe\\s+)?(Sub|Function)\\s+(${ID})\\s+Lib\\b`, 'i');
const TYPE_START = new RegExp(`^\\s*(?:(Public|Private)\\s+)?(Type|Enum)\\s+(${ID})\\s*$`, 'i');
const TYPE_END = /^\s*End\s+(Type|Enum)\b/i;
const MODULE_VAR = new RegExp(`^\\s*(Public|Private|Global|Dim)\\s+(?:(Const)\\s+)?(?!Sub\\b|Function\\b|Property\\b|Declare\\b|Type\\b|Enum\\b|Event\\b)(${ID})`, 'i');
const MODULE_CONST = new RegExp(`^\\s*(?:(Public|Private|Global)\\s+)?Const\\s+(${ID})`, 'i');

// ------------------------------------------------------------
// 解析
// ------------------------------------------------------------
// 文字列リテラルとコメントを空白に置き換える（位置はそのまま）。vba-format.js と共通
const { stripLine, formatVba } = require('./vba-format');

function codeStartLine(lines, file) {
  if (!/\.txt$/i.test(file)) return 0;
  for (let i = 0; i < lines.length; i++) if (/^CodeBehind(Form|Report)\s*$/.test(lines[i].trim())) return i + 1;
  return lines.length; // コードなし
}

function moduleNameOf(file) {
  const dir = path.basename(path.dirname(file)).toLowerCase();
  const base = path.basename(file, path.extname(file)).replace(/%([0-9A-Fa-f]{2})/g, (_, h) => String.fromCharCode(parseInt(h, 16)));
  if (dir === 'forms') return `Form_${base}`;
  if (dir === 'reports') return `Report_${base}`;
  return base;
}

// 行継続（ _）をつないだ論理行を返す
function logicalLine(lines, i) {
  let text = lines[i];
  let end = i;
  while (/\s_\s*$/.test(stripLine(lines[end])) && end + 1 < lines.length) {
    text = text.replace(/\s_\s*$/, ' ') + lines[end + 1].trim();
    end++;
  }
  return { text, end };
}

function parseParams(sig) {
  const open = sig.indexOf('(');
  if (open < 0) return { params: '', returnType: '' };
  let depth = 0;
  let close = -1;
  for (let i = open; i < sig.length; i++) {
    if (sig[i] === '(') depth++;
    else if (sig[i] === ')') { depth--; if (depth === 0) { close = i; break; } }
  }
  const params = close > 0 ? sig.slice(open + 1, close).trim() : sig.slice(open + 1).trim();
  const rest = close > 0 ? sig.slice(close + 1) : '';
  const m = new RegExp(`^\\s*As\\s+(${ID}(?:\\.${ID})?(?:\\(\\))?)`, 'i').exec(rest);
  return { params, returnType: m ? m[1] : '' };
}

function splitParams(params) {
  if (!params) return [];
  const out = [];
  let depth = 0;
  let cur = '';
  for (const ch of params) {
    if (ch === '(') depth++;
    if (ch === ')') depth--;
    if (ch === ',' && depth === 0) { out.push(cur.trim()); cur = ''; continue; }
    cur += ch;
  }
  if (cur.trim()) out.push(cur.trim());
  return out;
}

function paramName(p) {
  const m = new RegExp(`^(?:Optional\\s+)?(?:ByVal\\s+|ByRef\\s+)?(?:ParamArray\\s+)?(${ID})`, 'i').exec(p);
  return m ? m[1] : p;
}

function docCommentAbove(lines, i, stop) {
  const out = [];
  for (let j = i - 1; j >= stop; j--) {
    const t = lines[j].trim();
    if (/^'/.test(t)) out.unshift(t.replace(/^'+\s?/, ''));
    else break;
  }
  return out.join('\n').trim();
}

function parseText(text, file) {
  const lines = text.split(/\r?\n/);
  const start = codeStartLine(lines, file);
  const procs = [];
  const decls = [];
  let cur = null;
  let typeBlock = null;
  for (let i = start; i < lines.length; i++) {
    const raw = lines[i];
    const s = stripLine(raw);
    if (!s.trim()) continue;
    if (/^\s*Attribute\s/i.test(s)) continue;
    if (cur) {
      if (PROC_END.test(s)) { cur.endLine = i; procs.push(cur); cur = null; }
      continue;
    }
    if (typeBlock) {
      if (TYPE_END.test(s)) { typeBlock.endLine = i; decls.push(typeBlock); typeBlock = null; }
      else {
        const m = new RegExp(`^\\s*(${ID})`).exec(s);
        if (m) typeBlock.members.push({ name: m[1], line: i });
      }
      continue;
    }
    let m = PROC_START.exec(s);
    if (m) {
      const ll = logicalLine(lines, i);
      const pr = parseParams(stripLine(ll.text));
      cur = {
        kind: m[3].replace(/\s+/g, ' ').replace(/^property/i, 'Property'),
        scope: m[1] || 'Public',
        name: m[4],
        params: pr.params,
        returnType: pr.returnType,
        startLine: i,
        endLine: i,
        nameCol: raw.indexOf(m[4]),
        signature: ll.text.trim().replace(/\s+/g, ' '),
        doc: docCommentAbove(lines, i, start),
      };
      i = Math.max(i, ll.end);
      continue;
    }
    m = DECLARE.exec(s);
    if (m) {
      const ll = logicalLine(lines, i);
      const pr = parseParams(stripLine(ll.text));
      procs.push({ kind: `Declare ${m[2]}`, scope: m[1] || 'Public', name: m[3], params: pr.params, returnType: pr.returnType, startLine: i, endLine: ll.end, nameCol: raw.indexOf(m[3]), signature: ll.text.trim().replace(/\s+/g, ' '), doc: docCommentAbove(lines, i, start) });
      i = ll.end;
      continue;
    }
    m = TYPE_START.exec(s);
    if (m) { typeBlock = { kind: m[2].replace(/^./, (c) => c.toUpperCase()), scope: m[1] || 'Public', name: m[3], line: i, nameCol: raw.indexOf(m[3]), members: [] }; continue; }
    m = MODULE_CONST.exec(s);
    if (m) { decls.push({ kind: 'Const', scope: m[1] || 'Private', name: m[2], line: i, nameCol: raw.indexOf(m[2]), text: raw.trim() }); continue; }
    m = MODULE_VAR.exec(s);
    if (m && !/^\s*(Public|Private)\s+(Sub|Function|Property)/i.test(s)) {
      decls.push({ kind: m[2] ? 'Const' : 'Variable', scope: /^(Public|Global)$/i.test(m[1]) ? 'Public' : 'Private', name: m[3], line: i, nameCol: raw.indexOf(m[3]), text: raw.trim() });
    }
  }
  return { module: moduleNameOf(file), file, procs, decls, codeStart: start, lines };
}

// ------------------------------------------------------------
// プロジェクト全体の索引（ファイルの更新時刻でキャッシュ）
// ------------------------------------------------------------
class Index {
  constructor(getSrc) {
    this.getSrc = getSrc;
    this.cache = new Map();
  }
  files() {
    const src = this.getSrc();
    if (!src || !fs.existsSync(src)) return [];
    const out = [];
    for (const [dir, re] of [['modules', /\.(bas|cls)$/i], ['forms', /\.txt$/i], ['reports', /\.txt$/i]]) {
      const d = path.join(src, dir);
      if (!fs.existsSync(d)) continue;
      for (const f of fs.readdirSync(d)) if (re.test(f)) out.push(path.join(d, f));
    }
    return out;
  }
  get(file, text) {
    const key = file.toLowerCase();
    if (text !== undefined) {
      // 開いているエディタの内容（未保存も含む）
      return parseText(text, file);
    }
    let st;
    try { st = fs.statSync(file); } catch (_) { this.cache.delete(key); return null; }
    const c = this.cache.get(key);
    if (c && c.mtime === st.mtimeMs) return c.parsed;
    const parsed = parseText(fs.readFileSync(file, 'utf8').replace(/^﻿/, ''), file);
    this.cache.set(key, { mtime: st.mtimeMs, parsed });
    return parsed;
  }
  all() {
    const open = new Map(vscode.workspace.textDocuments.map((d) => [d.fileName.toLowerCase(), d]));
    return this.files().map((f) => {
      const d = open.get(f.toLowerCase());
      return d && d.isDirty ? this.get(f, d.getText()) : this.get(f);
    }).filter(Boolean);
  }
}

// ------------------------------------------------------------
// よく使う組み込み関数・オブジェクト（ホバーと補完用）
// ------------------------------------------------------------
const BUILTINS = {
  DoCmd: 'Access の操作（フォームやレポートを開く、SQL を実行する など）をまとめたオブジェクト。例: DoCmd.OpenForm "frmCustomers"',
  CurrentDb: '今開いているデータベース（DAO.Database）を返す。例: CurrentDb.Execute "UPDATE ...", dbFailOnError',
  CurrentProject: '今のプロジェクト。AllForms / AllReports / Path など',
  Me: 'このコードがあるフォーム／レポート自身',
  Forms: '開いているフォームの集まり。例: Forms!frmCustomers!txtName',
  Reports: '開いているレポートの集まり',
  Nz: 'Nz(値, 代わりの値) … 値が Null なら代わりの値を返す。例: Nz(Me.txtQty, 0)',
  DLookup: 'DLookup(項目, テーブル/クエリ, 条件) … 条件に合う最初の値を返す',
  DCount: 'DCount(項目, テーブル/クエリ, 条件) … 件数を返す',
  DSum: 'DSum(項目, テーブル/クエリ, 条件) … 合計を返す',
  DMax: 'DMax(項目, テーブル/クエリ, 条件) … 最大値を返す',
  DMin: 'DMin(項目, テーブル/クエリ, 条件) … 最小値を返す',
  IIf: 'IIf(条件, 真のとき, 偽のとき) … 両方とも評価される点に注意',
  Format: 'Format(値, 書式) … 例: Format(Date, "yyyy/mm/dd")、Format(1234.5, "#,##0")',
  MsgBox: 'MsgBox(メッセージ, ボタン, タイトル) … Access VSC から実行すると自動で答える（OK だけなら OK、確認の質問は「キャンセル」／「いいえ」）',
  InputBox: 'InputBox(メッセージ) … Access VSC からの自動実行では入力できず時間切れになる。テストしたい処理では使わない',
  CStr: '文字列に変換', CLng: '長整数型に変換（小数は偶数丸め）', CCur: '通貨型に変換', CDbl: '倍精度浮動小数点型に変換', CDate: '日付に変換',
  Round: 'Round(数値, 桁) … 偶数丸め（銀行型）。2.5 → 2',
  Int: 'Int(数値) … 小数を切り捨て（負の数は小さい方へ。-1.5 → -2）',
  Fix: 'Fix(数値) … 小数を切り捨て（0 の方へ。-1.5 → -1）',
  DateAdd: 'DateAdd(単位, 数, 日付) … 例: DateAdd("m", 1, Date)',
  DateDiff: 'DateDiff(単位, 日付1, 日付2) … 例: DateDiff("d", 開始, 終了)',
  DateSerial: 'DateSerial(年, 月, 日) … 日付を作る。月末: DateSerial(y, m + 1, 0)',
  Len: '文字数', Left: 'Left(文字列, 文字数)', Right: 'Right(文字列, 文字数)', Mid: 'Mid(文字列, 開始, 文字数)', Trim: '前後の空白を削除',
  Replace: 'Replace(文字列, 探す文字, 置き換える文字)', InStr: 'InStr(開始, 文字列, 探す文字) … 見つかった位置（なければ 0）',
  Split: 'Split(文字列, 区切り) … 配列に分ける', Join: 'Join(配列, 区切り) … 配列をつなぐ',
  IsNull: 'Null かどうか', IsEmpty: '未初期化の Variant かどうか', IsNumeric: '数値として扱えるかどうか', IsDate: '日付として扱えるかどうか',
  Now: '現在の日時', Date: '今日の日付', Time: '現在の時刻',
  Assert: 'Access VSC のテスト用モジュール。Assert.AreEqual 期待値, 実際 など',
  UI: 'Access VSC のフォーム操作テスト用モジュール。UI.OpenForm \"frm\" / UI.TypeText \"txt\", \"値\" / UI.Click \"cmd\" / UI.Value(\"txt\")。キーボードで操作するのでイベントも動く',
};
const KEYWORDS = ['Dim', 'As', 'If', 'Then', 'Else', 'ElseIf', 'End If', 'For', 'To', 'Step', 'Next', 'For Each', 'In', 'Do', 'Loop', 'While', 'Wend', 'Until', 'Select Case', 'Case', 'End Select', 'With', 'End With', 'Exit Sub', 'Exit Function', 'On Error GoTo', 'On Error Resume Next', 'Resume', 'Set', 'Nothing', 'Not', 'And', 'Or', 'True', 'False', 'Null', 'ByVal', 'ByRef', 'Optional', 'Public', 'Private', 'Const', 'Long', 'Integer', 'String', 'Currency', 'Double', 'Boolean', 'Date', 'Variant', 'Object'];

// ------------------------------------------------------------
// プロバイダー
// ------------------------------------------------------------
const SELECTOR = [{ language: 'vba' }, { language: 'access-object' }];

function wordAt(doc, pos) {
  const r = doc.getWordRangeAtPosition(pos, new RegExp(ID));
  return r ? { range: r, word: doc.getText(r) } : null;
}

function qualifierBefore(doc, range) {
  const line = doc.lineAt(range.start.line).text.slice(0, range.start.character);
  const m = new RegExp(`(${ID})\\s*\\.\\s*$`).exec(line);
  return m ? m[1] : null;
}

// 日本語の識別子にも効く「単語として一致」の正規表現（JavaScript の \b は ASCII だけなので使わない）
function wordRe(word, flags) {
  const w = word.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return new RegExp(`(?<![\\w\\u0080-\\uFFFF])${w}(?![\\w\\u0080-\\uFFFF])`, flags);
}

// カーソルがコメントや文字列の中でなければ true
function inCode(doc, pos) {
  const full = doc.lineAt(pos.line).text;
  const s = stripLine(full);
  const c = Math.max(0, pos.character - 1);
  if (!full.length) return true;
  if (s[c] === ' ' && full[c] !== ' ') return false;
  return !(/"[^"]*$/.test(s.slice(0, pos.character)) && (s.slice(0, pos.character).match(/"/g) || []).length % 2 === 1);
}

function procAtLine(parsed, line) {
  return parsed.procs.find((p) => line >= p.startLine && line <= p.endLine);
}

function symbolKind(kind) {
  if (/^Function|Declare Function/i.test(kind)) return vscode.SymbolKind.Function;
  if (/^Property/i.test(kind)) return vscode.SymbolKind.Property;
  if (/^Sub|Declare Sub/i.test(kind)) return vscode.SymbolKind.Method;
  if (kind === 'Const') return vscode.SymbolKind.Constant;
  if (kind === 'Type') return vscode.SymbolKind.Struct;
  if (kind === 'Enum') return vscode.SymbolKind.Enum;
  return vscode.SymbolKind.Variable;
}

function procMarkdown(p, module) {
  const md = new vscode.MarkdownString();
  md.appendCodeblock(p.signature, 'vba');
  if (p.doc) md.appendMarkdown(`\n${p.doc.replace(/\n/g, '  \n')}\n`);
  md.appendMarkdown(`\n\n*${module}* ・ ${p.scope}`);
  return md;
}

function register(context, getSrc) {
  const index = new Index(getSrc);
  const parsedDoc = (doc) => index.get(doc.fileName, doc.getText());

  // 見つけた候補のうち、同じファイル → Public の順で探す
  function findProcs(word, docFile, qualifier) {
    const res = [];
    for (const m of index.all()) {
      if (qualifier && m.module.toLowerCase() !== qualifier.toLowerCase()) continue;
      for (const p of m.procs) {
        if (p.name.toLowerCase() !== word.toLowerCase()) continue;
        const same = m.file.toLowerCase() === docFile.toLowerCase();
        if (same || p.scope !== 'Private') res.push({ m, p, same });
      }
      for (const d of m.decls) {
        if (d.name.toLowerCase() !== word.toLowerCase()) continue;
        const same = m.file.toLowerCase() === docFile.toLowerCase();
        if (same || d.scope === 'Public') res.push({ m, d, same });
      }
    }
    res.sort((a, b) => Number(b.same) - Number(a.same));
    return res;
  }

  const subs = [];

  // アウトライン（Ctrl+Shift+O・パンくず）
  subs.push(vscode.languages.registerDocumentSymbolProvider(SELECTOR, {
    provideDocumentSymbols(doc) {
      const p = parsedDoc(doc);
      const out = [];
      for (const d of p.decls) {
        const line = d.line;
        const r = new vscode.Range(line, 0, d.endLine !== undefined ? d.endLine : line, doc.lineAt(d.endLine !== undefined ? d.endLine : line).text.length);
        const sel = new vscode.Range(line, Math.max(0, d.nameCol), line, Math.max(0, d.nameCol) + d.name.length);
        const s = new vscode.DocumentSymbol(d.name, `${d.scope} ${d.kind}`, symbolKind(d.kind), r, sel);
        for (const mem of d.members || []) {
          const mr = doc.lineAt(mem.line).range;
          s.children.push(new vscode.DocumentSymbol(mem.name, '', vscode.SymbolKind.Field, mr, mr));
        }
        out.push(s);
      }
      for (const pr of p.procs) {
        const r = new vscode.Range(pr.startLine, 0, pr.endLine, doc.lineAt(pr.endLine).text.length);
        const sel = new vscode.Range(pr.startLine, Math.max(0, pr.nameCol), pr.startLine, Math.max(0, pr.nameCol) + pr.name.length);
        out.push(new vscode.DocumentSymbol(pr.name, `${pr.scope} ${pr.kind}${pr.returnType ? ` As ${pr.returnType}` : ''}`, symbolKind(pr.kind), r, sel));
      }
      return out;
    },
  }));

  // プロジェクト全体のシンボル（Ctrl+T）
  subs.push(vscode.languages.registerWorkspaceSymbolProvider({
    provideWorkspaceSymbols(query) {
      const q = (query || '').toLowerCase();
      const out = [];
      for (const m of index.all()) {
        const uri = vscode.Uri.file(m.file);
        for (const p of m.procs) {
          if (q && !p.name.toLowerCase().includes(q)) continue;
          out.push(new vscode.SymbolInformation(p.name, symbolKind(p.kind), m.module, new vscode.Location(uri, new vscode.Position(p.startLine, Math.max(0, p.nameCol)))));
        }
        for (const d of m.decls) {
          if (q && !d.name.toLowerCase().includes(q)) continue;
          out.push(new vscode.SymbolInformation(d.name, symbolKind(d.kind), m.module, new vscode.Location(uri, new vscode.Position(d.line, Math.max(0, d.nameCol)))));
        }
      }
      return out.slice(0, 500);
    },
  }));

  // 定義へ移動（F12）
  subs.push(vscode.languages.registerDefinitionProvider(SELECTOR, {
    provideDefinition(doc, pos) {
      const w = wordAt(doc, pos);
      if (!w || !inCode(doc, pos)) return undefined;
      const q = qualifierBefore(doc, w.range);
      // 手続きの中のローカル変数・引数
      if (!q) {
        const here = parsedDoc(doc);
        const pr = procAtLine(here, pos.line);
        if (pr) {
          const re = wordRe(w.word, 'i');
          for (let i = pr.startLine; i <= pr.endLine; i++) {
            const s = stripLine(doc.lineAt(i).text);
            let col = -1;
            if (i === pr.startLine) {
              const open = s.indexOf('(');
              const at = open >= 0 ? s.slice(open).search(re) : -1;
              if (at >= 0) col = open + at;
            } else if (/^\s*(Dim|Static|Const)\b/i.test(s)) {
              col = s.search(re);
            }
            if (col >= 0 && !(i === pos.line && col === w.range.start.character)) return new vscode.Location(doc.uri, new vscode.Position(i, col));
          }
        }
      }
      // モジュール名そのもの
      const mod = index.all().find((m) => m.module.toLowerCase() === w.word.toLowerCase());
      if (mod && !q) return new vscode.Location(vscode.Uri.file(mod.file), new vscode.Position(mod.codeStart, 0));
      const hits = findProcs(w.word, doc.fileName, q);
      return hits.map((h) => new vscode.Location(vscode.Uri.file(h.m.file), new vscode.Position(h.p ? h.p.startLine : h.d.line, Math.max(0, h.p ? h.p.nameCol : h.d.nameCol))));
    },
  }));

  // 参照の検索（Shift+F12）。大文字小文字を区別せず、コメントと文字列は除く
  subs.push(vscode.languages.registerReferenceProvider(SELECTOR, {
    provideReferences(doc, pos) {
      const w = wordAt(doc, pos);
      if (!w) return [];
      const re = wordRe(w.word, 'gi');
      const out = [];
      for (const m of index.all()) {
        const uri = vscode.Uri.file(m.file);
        const isObj = /\.txt$/i.test(m.file);
        for (let i = 0; i < m.lines.length; i++) {
          // フォーム・レポートは、コードと ControlSource / RowSource などの式も対象
          if (isObj && i < m.codeStart && !/^\s*(ControlSource|RowSource|RecordSource|DefaultValue|ValidationRule|On\w+)\s*=/.test(m.lines[i])) continue;
          const s = i < m.codeStart ? m.lines[i] : stripLine(m.lines[i]);
          let mm;
          re.lastIndex = 0;
          while ((mm = re.exec(s))) out.push(new vscode.Location(uri, new vscode.Range(i, mm.index, i, mm.index + w.word.length)));
        }
      }
      return out;
    },
  }));

  // ホバー
  subs.push(vscode.languages.registerHoverProvider(SELECTOR, {
    provideHover(doc, pos) {
      const w = wordAt(doc, pos);
      if (!w || !inCode(doc, pos)) return undefined;
      const hits = findProcs(w.word, doc.fileName, qualifierBefore(doc, w.range));
      if (hits.length) {
        const h = hits[0];
        if (h.p) return new vscode.Hover(procMarkdown(h.p, h.m.module), w.range);
        const md = new vscode.MarkdownString();
        md.appendCodeblock(h.d.text || `${h.d.scope} ${h.d.kind} ${h.d.name}`, 'vba');
        md.appendMarkdown(`\n*${h.m.module}*`);
        return new vscode.Hover(md, w.range);
      }
      const key = Object.keys(BUILTINS).find((k) => k.toLowerCase() === w.word.toLowerCase());
      if (key) return new vscode.Hover(new vscode.MarkdownString(`**${key}**  \n${BUILTINS[key]}`), w.range);
      return undefined;
    },
  }));

  // 入力補完
  subs.push(vscode.languages.registerCompletionItemProvider(SELECTOR, {
    provideCompletionItems(doc, pos) {
      const before = doc.lineAt(pos.line).text.slice(0, pos.character);
      if (!inCode(doc, pos)) return undefined; // コメント・文字列の中では出さない
      const items = [];
      const all = index.all();
      const qm = new RegExp(`(${ID})\\.(${ID})?$`).exec(before);
      if (qm) {
        const mod = all.find((m) => m.module.toLowerCase() === qm[1].toLowerCase());
        if (!mod) return undefined; // Me. や DoCmd. などは VS Code の既定に任せる
        for (const p of mod.procs.filter((p) => p.scope !== 'Private')) {
          const it = new vscode.CompletionItem(p.name, symbolKind(p.kind) === vscode.SymbolKind.Function ? vscode.CompletionItemKind.Function : vscode.CompletionItemKind.Method);
          it.detail = p.signature;
          it.documentation = p.doc || undefined;
          items.push(it);
        }
        return items;
      }
      for (const m of all) {
        const same = m.file.toLowerCase() === doc.fileName.toLowerCase();
        for (const p of m.procs) {
          if (!same && p.scope === 'Private') continue;
          const it = new vscode.CompletionItem(p.name, /Function/i.test(p.kind) ? vscode.CompletionItemKind.Function : vscode.CompletionItemKind.Method);
          it.detail = `${m.module} ・ ${p.signature}`;
          it.documentation = p.doc || undefined;
          it.sortText = `0${p.name}`;
          items.push(it);
        }
        for (const d of m.decls) {
          if (!same && d.scope !== 'Public') continue;
          const it = new vscode.CompletionItem(d.name, d.kind === 'Const' ? vscode.CompletionItemKind.Constant : d.kind === 'Variable' ? vscode.CompletionItemKind.Variable : vscode.CompletionItemKind.Struct);
          it.detail = `${m.module} ・ ${d.kind}`;
          it.sortText = `1${d.name}`;
          items.push(it);
        }
        const mi = new vscode.CompletionItem(m.module, vscode.CompletionItemKind.Module);
        mi.detail = 'モジュール';
        mi.sortText = `2${m.module}`;
        items.push(mi);
      }
      // 手続きの中のローカル変数と引数
      const here = parsedDoc(doc);
      const pr = procAtLine(here, pos.line);
      if (pr) {
        for (const pa of splitParams(pr.params)) {
          const it = new vscode.CompletionItem(paramName(pa), vscode.CompletionItemKind.Variable);
          it.detail = `引数 ・ ${pa}`;
          it.sortText = `00${paramName(pa)}`;
          items.push(it);
        }
        for (let i = pr.startLine + 1; i < pos.line; i++) {
          const s = stripLine(doc.lineAt(i).text);
          const dm = /^\s*(?:Dim|Static)\s+(.*)$/i.exec(s);
          if (!dm) continue;
          for (const part of splitParams(dm[1])) {
            const n = new RegExp(`^(${ID})`).exec(part);
            if (!n) continue;
            const it = new vscode.CompletionItem(n[1], vscode.CompletionItemKind.Variable);
            it.detail = `変数 ・ ${part}`;
            it.sortText = `00${n[1]}`;
            items.push(it);
          }
        }
      }
      for (const k of Object.keys(BUILTINS)) {
        const it = new vscode.CompletionItem(k, vscode.CompletionItemKind.Function);
        it.documentation = BUILTINS[k];
        it.sortText = `3${k}`;
        items.push(it);
      }
      for (const k of KEYWORDS) {
        const it = new vscode.CompletionItem(k, vscode.CompletionItemKind.Keyword);
        it.sortText = `4${k}`;
        items.push(it);
      }
      return items;
    },
  }, '.'));

  // 引数のヒント
  subs.push(vscode.languages.registerSignatureHelpProvider(SELECTOR, {
    provideSignatureHelp(doc, pos) {
      const text = stripLine(doc.lineAt(pos.line).text.slice(0, pos.character));
      // 閉じていない ( を探す。なければ「Call なしの Sub 呼び出し」（名前 引数, 引数）
      let depth = 0;
      let open = -1;
      let commas = 0;
      for (let i = text.length - 1; i >= 0; i--) {
        const ch = text[i];
        if (ch === ')') depth++;
        else if (ch === '(') { if (depth === 0) { open = i; break; } depth--; }
        else if (ch === ',' && depth === 0) commas++;
      }
      let name;
      let qualifier;
      if (open >= 0) {
        const m = new RegExp(`(?:(${ID})\\.)?(${ID})\\s*$`).exec(text.slice(0, open));
        if (!m) return undefined;
        qualifier = m[1];
        name = m[2];
      } else {
        const m = new RegExp(`^\\s*(?:Call\\s+)?(?:(${ID})\\.)?(${ID})\\s+(.*)$`, 'i').exec(text);
        if (!m || /^(If|ElseIf|While|Until|Case|Dim|Set|Let|Return|Exit|End|Select|For|Do|With)$/i.test(m[2])) return undefined;
        qualifier = m[1];
        name = m[2];
        commas = (m[3].match(/,/g) || []).length;
      }
      const hits = findProcs(name, doc.fileName, qualifier).filter((h) => h.p);
      if (!hits.length) return undefined;
      const p = hits[0].p;
      const params = splitParams(p.params);
      const sig = new vscode.SignatureInformation(`${p.name}(${params.join(', ')})${p.returnType ? ` As ${p.returnType}` : ''}`, p.doc || undefined);
      sig.parameters = params.map((x) => new vscode.ParameterInformation(x));
      const help = new vscode.SignatureHelp();
      help.signatures = [sig];
      help.activeSignature = 0;
      help.activeParameter = Math.min(commas, Math.max(0, params.length - 1));
      return help;
    },
  }, '(', ',', ' '));

  // 折りたたみ（手続き・Type・Enum、フォームのテキストの Begin...End）
  subs.push(vscode.languages.registerFoldingRangeProvider(SELECTOR, {
    provideFoldingRanges(doc) {
      const p = parsedDoc(doc);
      const out = p.procs.filter((x) => x.endLine > x.startLine).map((x) => new vscode.FoldingRange(x.startLine, x.endLine));
      for (const d of p.decls) if (d.endLine > d.line) out.push(new vscode.FoldingRange(d.line, d.endLine));
      if (doc.languageId === 'access-object') {
        const stack = [];
        for (let i = 0; i < Math.min(doc.lineCount, p.codeStart || doc.lineCount); i++) {
          const t = doc.lineAt(i).text.trim();
          if (/^Begin(\s+\w+)?$/.test(t) || /=\s*Begin$/.test(t)) stack.push(i);
          else if (t === 'End' && stack.length) {
            const s = stack.pop();
            if (i > s) out.push(new vscode.FoldingRange(s, i));
          }
        }
      }
      return out;
    },
  }));

  // 整形（Shift+Alt+F）。インデントと行末の空白だけ。ブロックの対応が取れないときは変えない
  subs.push(vscode.languages.registerDocumentFormattingEditProvider(SELECTOR, {
    provideDocumentFormattingEdits(doc, options) {
      const indent = options.insertSpaces ? ' '.repeat(options.tabSize || 4) : '\t';
      const text = doc.getText();
      const out = formatVba(text, { indent, file: doc.fileName });
      if (out === null) {
        vscode.window.showWarningMessage('Access VSC: ブロックの対応が取れないため整形しませんでした（End If・Next・End Sub などの書き忘れがないか確認してください）。');
        return [];
      }
      if (out === text) return [];
      return [vscode.TextEdit.replace(new vscode.Range(doc.positionAt(0), doc.positionAt(text.length)), out)];
    },
  }));

  // 名前の変更（F2）。誤変更を避けるため、ローカル変数・引数と、モジュール内の Private な名前だけ
  const RESERVED = new Set(KEYWORDS.concat(['Call', 'Let', 'Get', 'Property', 'Sub', 'Function', 'End', 'Exit', 'GoTo', 'Me', 'Is', 'Like', 'Mod', 'Xor', 'Eqv', 'Imp', 'New', 'Type', 'Enum', 'Static', 'Global', 'Friend', 'Event', 'RaiseEvent', 'Declare', 'Option', 'Byte', 'Single', 'LongPtr', 'LongLong', 'Empty', 'Error', 'Print', 'Stop'])
    .map((k) => k.toLowerCase()));

  // 手続きの中で宣言されている名前（引数・Dim・Static・Const）
  function localNames(doc, pr) {
    const names = new Map();
    for (const pa of splitParams(pr.params)) names.set(paramName(pa).toLowerCase(), paramName(pa));
    for (let i = pr.startLine + 1; i < pr.endLine; i++) {
      const s = stripLine(doc.lineAt(i).text);
      const dm = /^\s*(?:Dim|Static|Const)\s+(.*)$/i.exec(s);
      if (!dm) continue;
      for (const part of splitParams(dm[1])) {
        const n = new RegExp(`^(${ID})`).exec(part);
        if (n) names.set(n[1].toLowerCase(), n[1]);
      }
    }
    return names;
  }

  function renameTarget(doc, pos) {
    const w = wordAt(doc, pos);
    if (!w || !inCode(doc, pos)) throw new Error('ここは名前を変更できません。');
    const before = doc.lineAt(pos.line).text.slice(0, w.range.start.character);
    if (/[.!]\s*$/.test(before)) throw new Error('「.」や「!」の後ろの名前（オブジェクトのメンバー・フィールド・コントロール）は変更できません。');
    const here = parsedDoc(doc);
    if (pos.line < here.codeStart) throw new Error('フォーム・レポートのプロパティ部分では名前を変更できません。');
    const key = w.word.toLowerCase();
    const pr = procAtLine(here, pos.line);
    if (pr && localNames(doc, pr).has(key)) return { w, scope: 'local', from: pr.startLine, to: pr.endLine, here, pr };
    const isPrivate = here.decls.some((d) => d.name.toLowerCase() === key && d.scope === 'Private' && d.kind !== 'Type' && d.kind !== 'Enum')
      || here.procs.some((p) => p.name.toLowerCase() === key && p.scope === 'Private' && !/^Declare/.test(p.kind));
    if (!isPrivate) {
      throw new Error('名前を変更できるのは、ローカル変数・引数と、モジュール内の Private な名前だけです（Public な名前は、他のモジュール・クエリ・フォームの式から使われている可能性があるため）。');
    }
    if (w.word.includes('_')) throw new Error('「_」を含む Private な手続きは、イベント（Click など）の可能性があるため変更できません。');
    for (const p of here.procs) {
      if (p.endLine > p.startLine && localNames(doc, p).has(key)) throw new Error(`手続き ${p.name} に同じ名前のローカル変数・引数があるため変更できません。`);
    }
    if (/\.txt$/i.test(doc.fileName)) {
      const re = wordRe(w.word, 'i');
      for (let i = 0; i < here.codeStart; i++) if (re.test(doc.lineAt(i).text)) throw new Error('フォーム・レポートのプロパティ（式）からも使われているため変更できません。');
    }
    return { w, scope: 'module', from: here.codeStart, to: doc.lineCount - 1, here };
  }

  subs.push(vscode.languages.registerRenameProvider(SELECTOR, {
    prepareRename(doc, pos) {
      return renameTarget(doc, pos).w.range;
    },
    provideRenameEdits(doc, pos, newName) {
      const t = renameTarget(doc, pos);
      const nn = newName.trim();
      if (!new RegExp(`^${ID}$`).test(nn) || nn.length > 255) throw new Error(`「${newName}」は VBA の名前として使えません。`);
      if (RESERVED.has(nn.toLowerCase())) throw new Error(`「${nn}」は VBA の予約語です。`);
      const lower = nn.toLowerCase();
      if (lower !== t.w.word.toLowerCase()) {
        const taken = t.scope === 'local'
          ? localNames(doc, t.pr).has(lower) || t.here.procs.some((p) => p.name.toLowerCase() === lower) || t.here.decls.some((d) => d.name.toLowerCase() === lower)
          : t.here.procs.some((p) => p.name.toLowerCase() === lower || (p.endLine > p.startLine && localNames(doc, p).has(lower)))
            || t.here.decls.some((d) => d.name.toLowerCase() === lower)
            || findProcs(nn, doc.fileName).length > 0;
        if (taken) throw new Error(`「${nn}」はすでに使われている名前です。`);
      }
      const edit = new vscode.WorkspaceEdit();
      const re = wordRe(t.w.word, 'gi');
      for (let i = t.from; i <= t.to; i++) {
        const s = stripLine(doc.lineAt(i).text);
        let m;
        re.lastIndex = 0;
        while ((m = re.exec(s))) {
          if (/[.!]\s*$/.test(s.slice(0, m.index))) continue; // obj.Name / rs!Name は別物
          edit.replace(doc.uri, new vscode.Range(i, m.index, i, m.index + t.w.word.length), nn);
        }
      }
      return edit;
    },
  }));

  context.subscriptions.push(...subs);
  return index;
}

module.exports = { register, parseText, stripLine, splitParams, wordRe };
