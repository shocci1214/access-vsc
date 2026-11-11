// VBA の整形（インデントと行末の空白だけを直す、安全側の整形）
//   ・ブロック（Sub/Function/Property, If, For, Do, While, With, Select Case, Type, Enum）でインデント
//   ・ラベル（ErrHandler: など）は行頭、#If などの条件付きコンパイルと Attribute 行は触らない
//   ・行継続（ _）の 2 行目以降は、1 行目と同じだけずらす（揃え方はそのまま）
//   ・ブロックの対応が取れないとき（End If の書き忘れなど）は何も変えない（null を返す）
// VS Code に依存しないので、テストや CLI からも使える。scripts/AccessVsc.Format.ps1 は同じ規則の PowerShell 版。
'use strict';

// 文字列リテラルとコメントを空白に置き換える（位置はそのまま）
function stripLine(line) {
  let out = '';
  let inStr = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (inStr) {
      if (ch === '"') {
        if (line[i + 1] === '"') { out += '  '; i++; continue; }
        inStr = false;
        out += '"';
        continue;
      }
      out += ' ';
      continue;
    }
    if (ch === '"') { inStr = true; out += '"'; continue; }
    if (ch === "'") { out += ' '.repeat(line.length - i); break; }
    out += ch;
  }
  if (/^\s*Rem\b/i.test(out)) return ' '.repeat(line.length);
  return out;
}

const ID = '[A-Za-z_\\u0080-\\uFFFF][\\w\\u0080-\\uFFFF]*';
const PROC_START = new RegExp(`^(?:(?:Public|Private|Friend|Global)\\s+)?(?:Static\\s+)?(?:Sub|Function|Property\\s+(?:Get|Let|Set))\\s+${ID}\\s*\\(`, 'i');
const TYPE_START = new RegExp(`^(?:(?:Public|Private)\\s+)?(?:Type|Enum)\\s+${ID}$`, 'i');
const LABEL = new RegExp(`^(${ID}|\\d+):$`);
const NOT_LABEL = /^(Else|Case|Resume|GoTo|Exit|End|Next|Loop|Wend)$/i;

// コードの始まり（フォーム／レポートは CodeBehindForm の次、クラスはヘッダーと Attribute VB_ の次）
function codeStart(lines, file) {
  if (file && /\.txt$/i.test(file)) {
    for (let i = 0; i < lines.length; i++) if (/^CodeBehind(Form|Report)\s*$/.test(lines[i].trim())) return i + 1;
    return lines.length;
  }
  let start = 0;
  for (let i = 0; i < Math.min(lines.length, 40); i++) if (/^Attribute VB_/i.test(lines[i])) start = i + 1;
  return start;
}

function classify(s) {
  // s: コメントと文字列を除き、前後の空白を取った論理行
  if (/^End\s+(Sub|Function|Property)\b/i.test(s)) return { close: 'proc' };
  if (/^End\s+If\b/i.test(s)) return { close: 'if' };
  if (/^End\s+Select\b/i.test(s)) return { close: 'select' };
  if (/^End\s+With\b/i.test(s)) return { close: 'with' };
  if (/^End\s+(Type|Enum)\b/i.test(s)) return { close: 'type' };
  if (/^Next(\s|$)/i.test(s)) {
    const rest = s.replace(/^Next\s*/i, '').trim();
    return { close: 'for', count: rest ? rest.split(',').length : 1 };
  }
  if (/^Loop(\s|$)/i.test(s)) return { close: 'do' };
  if (/^Wend$/i.test(s)) return { close: 'while' };
  if (/^Else$/i.test(s)) return { middle: 'if' };
  if (/^ElseIf\b.*\bThen$/i.test(s)) return { middle: 'if' };
  if (/^Case(\s|$)/i.test(s)) return { kase: true };
  if (PROC_START.test(s)) return { open: 'proc' };
  if (TYPE_START.test(s)) return { open: 'type' };
  if (/^If\b.*\bThen$/i.test(s)) return { open: 'if' };
  if (/^Select\s+Case\b/i.test(s)) return { open: 'select' };
  if (/^For(\s|$)/i.test(s)) return { open: 'for' };
  if (/^Do(\s|$)/i.test(s)) return { open: 'do' };
  if (/^While\s/i.test(s)) return { open: 'while' };
  if (/^With\s/i.test(s)) return { open: 'with' };
  return {};
}

function leading(s) {
  return /^[ \t]*/.exec(s)[0];
}

// 整形した全文を返す。変える必要がなければ同じ文字列、ブロックの対応が取れなければ null
function formatVba(text, opts) {
  const o = Object.assign({ indent: '    ', file: '' }, opts || {});
  const eol = /\r\n/.test(text) ? '\r\n' : '\n';
  const lines = text.split(/\r?\n/);
  const start = codeStart(lines, o.file);
  const out = lines.slice(0, start);
  const stack = [];
  const width = (ws) => ws.replace(/\t/g, '    ').length;
  for (let i = start; i < lines.length; i++) {
    const raw = lines[i].replace(/[ \t]+$/, '');
    const stripped = stripLine(raw);
    const s0 = stripped.trim();
    if (!raw.trim()) { out.push(''); continue; }
    // 行継続：論理行の最後まで読む
    let end = i;
    let logical = stripped;
    while (/\s_\s*$/.test(stripLine(lines[end])) && end + 1 < lines.length) {
      end++;
      logical = logical.replace(/\s_\s*$/, ' ') + ' ' + stripLine(lines[end]).trim();
    }
    const s = logical.trim();
    const keepAsIs = /^#/.test(s0) || /^Attribute\s/i.test(s0);
    const lab = LABEL.exec(s);
    let depth;
    if (keepAsIs) {
      depth = null;
    } else if (!s0) {
      depth = stack.length; // コメントだけの行
    } else if (lab && !NOT_LABEL.test(lab[1])) {
      depth = -1; // ラベルは行頭
    } else {
      const c = classify(s);
      if (c.close) {
        if (c.close === 'select' && stack[stack.length - 1] === 'case') stack.pop();
        for (let k = 0; k < (c.count || 1); k++) {
          if (stack[stack.length - 1] !== c.close) return null;
          stack.pop();
        }
        depth = stack.length;
      } else if (c.middle) {
        if (stack[stack.length - 1] !== 'if') return null;
        depth = stack.length - 1;
      } else if (c.kase) {
        if (stack[stack.length - 1] === 'case') stack.pop();
        if (stack[stack.length - 1] !== 'select') return null;
        depth = stack.length;
        stack.push('case');
      } else {
        depth = stack.length;
        if (c.open) stack.push(c.open);
      }
    }
    const oldIndent = leading(raw);
    const newIndent = depth === null ? oldIndent : depth < 0 ? '' : o.indent.repeat(depth);
    out.push(newIndent + raw.slice(oldIndent.length));
    // 行継続の 2 行目以降：1 行目と同じだけずらす
    const delta = width(newIndent) - width(oldIndent);
    for (let j = i + 1; j <= end; j++) {
      const r = lines[j].replace(/[ \t]+$/, '');
      const ws = leading(r);
      const w = Math.max(0, width(ws) + delta);
      out.push((o.indent === '\t' ? '\t'.repeat(Math.floor(w / 4)) + ' '.repeat(w % 4) : ' '.repeat(w)) + r.slice(ws.length));
    }
    i = end;
  }
  if (stack.length) return null;
  return out.join(eol);
}

module.exports = { formatVba, stripLine, codeStart };
