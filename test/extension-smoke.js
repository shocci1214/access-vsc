// Loads extension.js with a minimal fake "vscode" module and checks the pure helpers.
// Run: set ELECTRON_RUN_AS_NODE=1 && Code.exe test/extension-smoke.js <src-folder> <db-path>
'use strict';
const Module = require('module');
const path = require('path');
const fs = require('fs');

const commands = {};
const fakeVscode = new Proxy({}, {
  get(_, prop) {
    if (prop === 'EventEmitter') return class { constructor() { this.event = () => {}; } fire() {} dispose() {} };
    if (prop === 'commands') return { registerCommand: (id, fn) => { commands[id] = fn; return { dispose() {} }; }, executeCommand() {} };
    if (prop === 'workspace') return { getConfiguration: () => ({ get: () => '' }), workspaceFolders: undefined, findFiles: async () => [], onDidSaveTextDocument: () => ({ dispose() {} }), onDidChangeConfiguration: () => ({ dispose() {} }), createFileSystemWatcher: () => ({ onDidChange() {}, onDidCreate() {}, onDidDelete() {}, dispose() {} }) };
    if (prop === 'window') return { createOutputChannel: () => ({ appendLine() {}, show() {}, dispose() {} }), createStatusBarItem: () => ({ show() {}, hide() {}, dispose() {} }) };
    if (prop === 'languages') return { createDiagnosticCollection: () => ({ clear() {}, set() {}, dispose() {} }), registerCodeLensProvider: () => ({ dispose() {} }) };
    if (prop === 'lm') return { registerMcpServerDefinitionProvider: () => ({ dispose() {} }) };
    if (prop === 'McpStdioServerDefinition') return class {};
    if (prop === 'StatusBarAlignment') return { Left: 1 };
    if (prop === 'ConfigurationTarget') return { Workspace: 2, Global: 1 };
    return {};
  },
});
const origLoad = Module._load;
Module._load = function (req, ...rest) {
  if (req === 'vscode') return fakeVscode;
  return origLoad.call(this, req, ...rest);
};

const extPath = path.join(__dirname, '..', 'extension.js');
const src = fs.readFileSync(extPath, 'utf8') + '\nmodule.exports.__test = { syncStatus, kindOf, decodeName, parseArgs, agentDoc, PROC_RE };';
const m = new Module(extPath, module);
m.filename = extPath;
m.paths = Module._nodeModulePaths(path.dirname(extPath));
m._compile(src, extPath);
const t = m.exports.__test;

let failures = 0;
const check = (name, cond) => { console.log(`${cond ? 'PASS' : 'FAIL'} ${name}`); if (!cond) failures++; };

check('decodeName', t.decodeName('a%3Fb%25c') === 'a?b%c');
check('kindOf module', t.kindOf('modules/x.bas').kind === 'modules');
check('kindOf table readonly', t.kindOf('tables/T.json').readOnly === true);
check('kindOf invalid', t.kindOf('modules/x.txt') === undefined);
check('parseArgs bare', JSON.stringify(t.parseArgs('"a", 1')) === '["a",1]');
check('parseArgs array', JSON.stringify(t.parseArgs('[true]')) === '[true]');
check('PROC_RE', t.PROC_RE.exec('Public Function 挨拶(ByVal s As String) As String')[3] === '挨拶');
check('PROC_RE private', (t.PROC_RE.exec('Private Sub Foo()') || [])[1] === 'Private');

const [srcDir, db] = process.argv.slice(2);
if (srcDir) {
  const st = t.syncStatus(srcDir);
  console.log('syncStatus:', JSON.stringify(st));
  check('syncStatus manifest', st.hasManifest === true);
}
if (db) {
  const doc = t.agentDoc(db, path.dirname(db));
  check('agentDoc markers', doc.startsWith('<!-- access-vsc:begin -->') && doc.endsWith('<!-- access-vsc:end -->'));
}
process.exitCode = failures ? 1 : 0;
console.log(failures ? `${failures} failure(s)` : 'all passed');
