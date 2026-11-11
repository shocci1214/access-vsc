# Access VSC

Microsoft Access の VBA・クエリ・フォーム・レポート・マクロを **テキストファイルとして VS Code で編集** し、
**AI エージェント（Claude Code / GitHub Copilot など）が Access を操作できる** ようにする拡張機能です。

```
 app.accdb  ──エクスポート──▶  app.accdb.src/   ◀── VS Code / AI エージェントが編集
            ◀─インポート+コンパイル──              
                                  modules/*.bas, *.cls   VBA
                                  forms/*.txt            フォーム（コード込み）
                                  reports/*.txt          レポート
                                  macros/*.txt           マクロ
                                  queries/*.sql          クエリ（SQL）
                                  tables/*.json          テーブル定義（読み取り専用）
                                  relationships.json     リレーション（読み取り専用）
```

> **📖 ドキュメント： https://shocci1214.github.io/access-vsc/**
> [チュートリアル](https://shocci1214.github.io/access-vsc/tutorial.html)（手を動かして一通り体験）・[学習コース](https://shocci1214.github.io/access-vsc/course.html)（初級→中級→上級）・[リファレンス](https://shocci1214.github.io/access-vsc/reference.html)（一覧・辞書）の 3 つがあります。
> 同じものが [docs/](docs/index.html) にも入っていて、拡張機能をインストールした後はサイドバーの「その他」→「ドキュメント」からも開けます。

## できること

| 分類 | 機能 | 内容 |
|---|---|---|
| 同期 | エクスポート | Access → ソースフォルダ。SaveAsText の Shift-JIS / UTF-16 を UTF-8 に変換。 |
| 同期 | インポート | 変更したファイルだけを Access に取り込み、自動でコンパイル。取り込み前に DB をバックアップ。 |
| 同期 | 取り込みの確認 | 取り込んだオブジェクトと **Public な手続きが DB に本当にあるか** を Access 側から確認（`result.verify`）。一時的な失敗は 3 回まで自動で再試行。 |
| 編集支援 | VBA の言語機能 | アウトライン、定義へ移動（F12）、参照の検索（Shift+F12）、ホバー、入力補完、引数のヒント、折りたたみ、スニペット。日本語の名前にも対応。 |
| 編集支援 | 整形 | **Shift+Alt+F** でインデントと行末の空白を整える。ブロックの対応が取れないときは何も変えない（安全側）。CLI・MCP からも使える。 |
| 編集支援 | 名前の変更 | **F2** で変更。誤変更を避けるため、ローカル変数・引数と、モジュール内の Private な名前だけが対象。 |
| 診断 | コンパイル | エラーの **モジュール・行・コード・エラーメッセージ** を返し、VS Code の「問題」パネルに表示。 |
| 診断 | コードチェック | Access 不要の lint（Shift-JIS にない文字・Option Explicit・Stop・型の書き忘れ・未使用の変数・ブロックの対応・インデント・帳票のはみ出し）。保存時に自動実行。 |
| 実行 | VBA 実行 | `Application.Run` で Public Function/Sub を実行して戻り値を取得。`.bas` に「▶ Access で実行」が表示される。 |
| 実行 | 実行時エラー | エラーダイアログを自動で「デバッグ」→ 行番号を取得 →「リセット」。エラー行をエディタに表示。 |
| 実行 | ダイアログ対策 | MsgBox などは自動で答えて内容を記録（確認の質問には安全側の「キャンセル」／「いいえ」）。固まったらタイムアウトで **自分が起動した Access だけ** 強制終了。 |
| 実行 | SQL | SELECT の結果を表で表示（既定は読み取り専用）。 |
| 帳票 | プレビュー | レポートを PDF / PNG にして表示。AI には画像のまま渡るので、見た目を確認しながら修正できる。 |
| 帳票 | レイアウトチェック | コントロールの位置一覧と、**はみ出し・重なり・用紙に入らない** を自動検出。 |
| 帳票 | ページ設定 | 用紙サイズ・向き・余白（バイナリで保存されていてテキストでは直せない部分）をコマンドで変更。 |
| 帳票 | スナップショット | 全レポートを画像化して基準画像と比較。変わった所を赤で示す差分画像つき（見た目の回帰テスト）。 |
| テスト | ユニットテスト | `Public Sub Test_*()` を **DB のコピー上で** 実行。VS Code の「テスト」ビューに ✔/✘、失敗箇所に移動。JUnit XML 出力。 |
| CI/CD | CI の準備 | Assert モジュール・テスト例・GitHub Actions・コミット前チェック・.gitattributes をまとめて追加。 |
| CI/CD | リリース・配布 | 版番号の書き込み → 本番バックエンドへの張り替え → **.accde** 作成 → 共有フォルダへ公開。利用者は起動用ランチャーで常に最新版。 |
| AI 連携 | MCP サーバー・CLI | Copilot には自動登録。Claude Code・Cursor・Gemini CLI 向けの設定と作業手順（完了の条件つき）を生成。JSON を返す CLI も。 |
| 画面 | サイドバー | 左の Access アイコンから、マウス操作だけで取り出し・反映・実行・プレビューができる。使い方ガイド付き。 |
| 安全 | 安全装置 | 未インポートの編集があるとエクスポートを拒否 / DB を他で開いていると書き込みを拒否 / 同じ DB への同時操作を排他 / 信頼されていないフォルダでは無効 |

### 既存の拡張機能との違い

VBA 向けの VS Code 拡張機能の多くは、色分け・スニペット・補完など **エディタの中** の支援が中心です。
Access VSC は、**Access 本体との往復（取り込み・コンパイル・実行・テスト・帳票の確認・配布）まで** を VS Code と AI エージェントから行えることを目指しています。

- `.accdb` の中身をテキストにして **Git で差分を見られる**（フォーム・レポート・クエリも）
- AI が **自分で確かめられる**：コンパイルエラーの位置、実行結果、テスト結果、帳票の画像が返る
- 「できました」の根拠を **ツールの結果で機械的に判断** できる（取り込みの確認・テスト・lint）

### 苦手なこと

- 解析は軽量（言語サーバーなし）。型の推論はしないため、`obj.` のあとの補完は自作モジュールのみ
- `CallByName` や文字列で組み立てた手続き呼び出し、フォームの式（`=MyFunc()`）からの呼び出しは、参照の検索で見つからないことがある
- 名前の変更は、安全のため Public な名前・イベント手続き（`_` を含む名前）には使えない
- 操作のたびに Access を裏で起動するため、1 回 4〜10 秒程度かかる
## 必要なもの

- Windows + Microsoft Access（Microsoft 365 / 2016 以降。Access Runtime では VBE が使えないため不可）
- Windows PowerShell 5.1（Windows に標準で入っています）
- VS Code 1.101 以降
- Node.js / npm は **不要**

## インストール

### VSIX からインストールする（おすすめ）

1. [Releases](https://github.com/shocci1214/access-vsc/releases) から最新の `access-vscode-x.y.z.vsix` をダウンロード
2. VS Code の左の［拡張機能］（<kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>X</kbd>）を開き、右上の［…］→［**VSIX からのインストール...**］→ ダウンロードした `.vsix` を選ぶ
3. 再読み込みを求められたら［再読み込み］（またはコマンドパレット →「Developer: Reload Window」）
4. 左端に Access のアイコンが出れば完了

コマンドラインで入れる場合：

```powershell
code --install-extension .\access-vscode-x.y.z.vsix
```

### ソースからビルドしてインストールする

Node.js / npm は不要です。PowerShell だけで `.vsix` を作れます。

```powershell
git clone https://github.com/shocci1214/access-vsc.git
cd access-vsc
powershell -NoProfile -ExecutionPolicy Bypass -File .\build-vsix.ps1 -Install
```

### アンインストール

拡張機能ビューで「Access VSC」を右クリック →「アンインストール」。

## 使い方（人が使う場合）

`.accdb` があるフォルダを VS Code で開き、**左端の Access アイコン** を押します。

| ビュー | 内容 |
|---|---|
| かんたん操作 | ① 取り出す → ② 編集 → ③ 反映 → ④ 確かめる、の順に並んだボタン。帳票プレビュー・データ表示・AI 連携設定も |
| オブジェクト | モジュール / フォーム / レポート / クエリ / マクロ / テーブルの一覧。未反映のものに「● 変更あり」。行の ▶ ボタンや右クリックで実行・プレビュー・ページ設定 |

初めての方は「かんたん操作」→「使い方ガイド」をどうぞ。

- `.bas` の関数の上に **▶ Access で実行**、レポートの `.txt` の先頭に **プレビュー / レイアウト詳細 / ページ設定** が出ます
- コンパイルエラー・実行時エラーは該当行に赤線（「問題」パネルにも表示）
- 画面左下のステータスバー（`app.accdb ● 2` のような表示）に未反映の変更数。クリックでメニュー
- すべての操作はコマンドパレット（Ctrl+Shift+P →「Access:」）からも使えます

### 帳票（レポート）を直す

1. オブジェクト → レポート → 対象の ▶（プレビュー）で現状を確認
2. `.txt` の `Left` / `Top` / `Width` / `Height`（単位 twip、**1 cm = 567**）や `Caption`、`FontSize` を編集
3. 「Access へ反映する」→ もう一度プレビュー。はみ出し・重なりがあれば黄色の枠で表示されます
4. 用紙・向き・余白は右クリック →「ページ設定」

## テストと CI/CD

サイドバー →「テスト・CI/CD を準備」で次のものが追加されます（Access は使いません）。

| 追加されるもの | 内容 |
|---|---|
| `<src>/modules/Assert.bas` | アサーション（`AreEqual` `AreNotEqual` `AreClose` `IsTrue` `IsFalse` `IsNullValue` `IsNothing` `Contains` `Fail`） |
| `<src>/modules/Test_Example.bas` | テストの例（テストが 1 つもない場合のみ） |
| `tools/access-vsc/` | CI で使うスクリプト一式（拡張機能がない環境でも動く） |
| `.github/workflows/access-ci.yml` | GitHub Actions：**lint**（クラウドの windows-latest）→ **test**（Access 入りのセルフホストランナー） |
| `.githooks/pre-commit` | コミット前に lint。エラーがあるとコミットを止める（`core.hooksPath` を設定） |
| `.gitattributes` / `.gitignore` | ソースは CRLF のまま、DB はバイナリ、作業フォルダは除外 |

### テストの書き方

```vb
' modules/Test_Tax.bas
Option Compare Database
Option Explicit

Public Sub TestSetup()        ' 各テストの前（任意）
End Sub

Public Sub Test_TaxIncluded()
    Assert.AreEqual 110, TaxIncluded(100)
End Sub

Public Sub Test_TaxRounding()
    Assert.AreEqual 109, TaxIncluded(99), "1 円未満切り捨て"
End Sub
```

- 名前が `Test_` で始まる、引数なしの `Public Sub` がテスト
- テストは `.accessvsc/test/` に作った **DB のコピー** で実行（データを消しても元の DB は無事。ただしリンクテーブルの先＝バックエンドは本物）
- 実行時エラーは、エラーになった行（テスト対象のコード内でも）まで特定
- 結果: VS Code の「テスト」ビュー / `.accessvsc/test-results/junit.xml` / CLI `test -Text`

### GitHub Actions の test ジョブを動かすには

1. Access の入った Windows PC で、リポジトリ → Settings → Actions → Runners → **New self-hosted runner** の手順を実行
2. ランナーにラベル `access` を付ける
3. **サービスとしてではなく**、ログオンしたユーザーで `run.cmd` を起動しておく（Office の自動操作はサービス（セッション 0）では正しく動かないため）
4. `.accdb` もコミットしておく（CI はコミットされた DB にソースを取り込んでからテストする）

## 帳票スナップショット（見た目の回帰テスト）

サイドバー →「帳票の見た目を比較」。

1. 初回は全帳票が「基準なし」。画面で内容を確認して **「この見た目を基準にする」**
   → 基準画像が `<src>/snapshots/<レポート名>_p<ページ>.png` に保存されます（git にコミット）
2. 以降は比較するたびに、基準と違うページを **基準 / 今 / 差分（違う所が赤）** の 3 枚で表示
3. 意図した変更なら「この見た目を基準にする」、意図しない崩れなら直す

- CI（access-ci.yml）でも実行され、変化があるとジョブが失敗し、差分画像が成果物に残ります
- 帳票の中身はデータで変わるので、**テスト用データの入った DB** で比較するのが前提です
- 今日の日付などを印刷する帳票は、`-Threshold 0.1`（変化したピクセルが 0.1% 以下なら同じとみなす）を指定してください

## リリース・配布（CD）

| 手順 | サイドバー | CLI |
|---|---|---|
| 1. 配布の設定 | 配布の設定 | `deploy-init -AppName Shop -DeployTo \\server\share\Shop -BackendPath \\server\data\Shop_be.accdb` |
| 2. リリースを作成 | リリースを作成 | `build -Version 1.2.0` |
| 3. 配布 | 配布する（確認あり） | `deploy -Version 1.2.0 -Yes`（`-Yes` がないと予行だけ） |

**build** で行うこと（元の DB は変更しません）

1. DB を `.accessvsc/build/` にコピー
2. 版番号を DB プロパティ `AppVersion`（と `AppBuiltAt`）に書き込み — アプリ内で `CurrentDb.Properties("AppVersion")` で表示できます
3. `access-deploy.json` の `backend.path` があれば、Access のリンクテーブルを本番のバックエンドに張り替え（`RefreshLink` で確認）
4. コンパイルして `.accde` を作成 → `dist/<アプリ名>_<版>.accde` と、ハッシュ・コミット等を書いた `.json`

**deploy** で共有フォルダにできるもの

```
\\server\share\Shop\
  Start-Shop.cmd        ← 利用者はこれ（のショートカット）から起動
  launcher.ps1          ← 起動のたびに版を確認し、新しければ %LOCALAPPDATA%\AccessApps\Shop へコピーしてから開く
  version.json          ← 現在の版・ハッシュ
  app\Shop.accde
  releases\1.2.0\Shop.accde   ← 過去の版（既定 10 個）。「配布する」で選べば元に戻せる
```

- 利用者の PC ではローカルのコピーを開くので、ネットワーク越しに直接開くより速く、壊れにくく、更新時に「使用中」で失敗しません
- `.accde` を開いたときにセキュリティの警告が出る場合は、`%LOCALAPPDATA%\AccessApps` を信頼できる場所に登録するよう管理者に依頼してください（ランチャーは自動では登録しません）
- `.accde` は作成した Access と同じビット数（64 / 32 bit）の Office でしか動きません
- `.accde` の作成中は Access のウィンドウが一瞬（最小化で）表示されます（非表示だと Access が作成を拒否するため）
- GitHub Actions：`deploy-init` で `.github/workflows/access-release.yml` もできます。`v1.2.0` のようなタグを push すると、取り込み → テスト → ビルド → 配布まで自動で行います
- AI エージェントは `access_build` まではできますが、**配布はできません**（人が VS Code か CLI で行う）

## AI エージェントに使わせる

コマンドパレット → **Access: AI エージェント連携を設定** を実行すると、選んだエージェント向けに以下を作成します。

| エージェント | 作成されるもの |
|---|---|
| Claude Code | `.mcp.json`（MCP サーバー `access` の登録）、`CLAUDE.md`（作業手順） |
| GitHub Copilot | `.github/copilot-instructions.md`（MCP サーバーはこの拡張機能が自動登録） |
| Cursor | `.cursor/mcp.json`（MCP サーバーの登録）、`.cursor/rules/access-vsc.mdc`（作業手順） |
| Gemini CLI | `.gemini/settings.json`（MCP サーバーの登録）、`GEMINI.md`（作業手順） |
| Codex・その他 | `AGENTS.md`（MCP サーバーは各ツールの設定で登録。CLI でも操作可） |

エージェントを再起動すると MCP サーバー `access` が使えるようになります（Claude Code は利用許可を求めてきます）。

作業手順には **完了の条件**（取り込み成功・確認 OK・テスト失敗 0・lint エラー 0・帳票の画像確認）が入っていて、AI はツールの結果を根拠に完了を報告します。

### MCP ツール

| ツール | 内容 |
|---|---|
| `access_status` | ソースの変更状況（Access を起動しないので速い） |
| `access_list` | オブジェクト一覧 |
| `access_export` | Access → ソース |
| `access_import` | ソース → Access + コンパイル（エラー位置を返す） |
| `access_compile` | コンパイル |
| `access_run` | VBA 実行（戻り値・ダイアログ・実行時エラーの位置を返す） |
| `access_eval` | `Application.Eval` で式を評価 |
| `access_sql` | SQL 実行（既定は読み取り専用。`write=true` で更新系） |
| `access_report_layout` | レポート/フォームのコントロール位置・用紙設定・問題点 |
| `access_report_preview` | レポートを PDF/PNG に出力し、**画像を AI に返す** |
| `access_report_pagesetup` | 用紙サイズ・向き・余白・レポート幅を変更（ソースも更新） |
| `access_test` | VBA ユニットテスト（DB のコピーで実行、失敗理由・エラー行を返す） |
| `access_lint` | コードチェック（Access 不要） |
| `access_format` | VBA のインデントを整える（Access 不要。空白だけを変更） |
| `access_setup_ci` | テスト・CI/CD の準備 |
| `access_report_snapshot` | 帳票の見た目を基準画像と比較（差分画像を AI に返す） |
| `access_build` | リリース用 .accde を作成（配布はしない） |

### CLI（MCP を使わないエージェントや手作業向け）

拡張機能を有効にすると、スクリプトが `%LOCALAPPDATA%\access-vscode\bin\` にコピーされます（拡張機能の更新でパスが変わらないように）。

```powershell
$A = "$env:LOCALAPPDATA\access-vscode\bin\access.ps1"
powershell -NoProfile -ExecutionPolicy Bypass -File $A status  -DbPath .\app.accdb
powershell -NoProfile -ExecutionPolicy Bypass -File $A export  -DbPath .\app.accdb
powershell -NoProfile -ExecutionPolicy Bypass -File $A format  -DbPath .\app.accdb -Text
powershell -NoProfile -ExecutionPolicy Bypass -File $A import  -DbPath .\app.accdb
powershell -NoProfile -ExecutionPolicy Bypass -File $A run     -DbPath .\app.accdb -Function Greet -ArgsJson "['世界']"
powershell -NoProfile -ExecutionPolicy Bypass -File $A sql     -DbPath .\app.accdb -Sql "SELECT * FROM Customers"
powershell -NoProfile -ExecutionPolicy Bypass -File $A preview -DbPath .\app.accdb -Report rptCustomers
powershell -NoProfile -ExecutionPolicy Bypass -File $A pagesetup -DbPath .\app.accdb -Report rptCustomers -PaperSize A4 -Orientation landscape -Margins "10,10,15,15"
```

結果はすべて JSON（`ok`, `result`, `error`, `dialogs`, `runtimeError` …）。失敗時は終了コード 1。

結果の例（コンパイルエラー）:

```json
{
  "ok": false,
  "command": "import",
  "result": {
    "imported": ["modules/modMain.bas"],
    "compile": {
      "compiled": false,
      "error": {
        "module": "modMain", "line": 13, "column": 9,
        "code": "    x = undefinedThing + 1",
        "message": "コンパイル エラー: 変数が定義されていません。",
        "file": "C:\\work\\app.accdb.src\\modules\\modMain.bas", "fileLine": 13
      }
    }
  },
  "error": "Imported, but the VBA project does not compile"
}
```

## 設定

| 設定 | 既定 | 内容 |
|---|---|---|
| `accessVsc.database` | （自動検出） | 対象の `.accdb` / `.mdb` |
| `accessVsc.sourceFolder` | `<DB名>.src` | ソースフォルダ |
| `accessVsc.importOnSave` | `false` | 保存したら自動でインポート + コンパイル |
| `accessVsc.timeoutSec` | `0`（既定値） | Access 操作のタイムアウト |
| `accessVsc.sqlRowLimit` | `200` | SQL 結果の最大行数 |
| `accessVsc.powershellPath` | `powershell.exe` | PowerShell の実行ファイル |

## 注意・制限

- **書き込み系の操作（インポート・コンパイル・更新 SQL）は、DB を Access で開いていると拒否されます。** Access を閉じてから実行してください。
- バックアップは DB と同じフォルダの `.accessvsc/backups/` に直近 10 個まで残ります（`.gitignore` に追加されます）。
- VBA は Access 内部で Shift-JIS です。Shift-JIS にない文字（絵文字など）は `?` になるため、インポート時に警告します。
- `InputBox` やモーダルフォーム（`acDialog`）で止まった場合はタイムアウトまで待ってから Access を強制終了します。
- 起動時に AutoExec マクロやスタートアップフォームがあれば、操作のたびに実行されます。
- 標準モジュール ↔ クラスモジュールを変えたい場合は拡張子（`.bas` / `.cls`）を変えてインポートすると作り直されます。
- テーブル定義（`tables/*.json`）はスナップショットです。編集しても DB は変わりません。変更は SQL の DDL で行い、その後エクスポートしてください。
- パススルークエリの接続文字列に含まれるパスワードは `***` に置き換えて出力します（インポート時は既存の接続文字列を維持）。

## セキュリティ

- **信頼できる DB・フォルダだけで使ってください。** 操作のたびに Access が DB を開くため、AutoExec マクロや VBA が動きます。信頼されていないワークスペース（VS Code の「制限モード」）では、この拡張機能は無効になります。
- **AI エージェントに渡す権限に注意してください。** `access_run`・`access_eval`・`access_sql`（write=true）は VBA や SQL を実行できます。MCP クライアントで書き込み系のツールを自動承認しないことをおすすめします（読み取り専用のツールには `readOnlyHint` が付いています）。
- 実行中のダイアログには安全側で答えます（「削除しますか？」には「キャンセル」／「いいえ」）。
- 書き込みの前に自動でバックアップを取ります。終了させるのは、この拡張機能が起動した Access だけです。ユーザーが開いている Access には触れません。
- **GitHub Actions（セルフホストランナー）**：テンプレートの test ジョブは、フォークからのプルリクエストでは動きません（他人のコードを自分の PC で実行しないため）。公開リポジトリでセルフホストランナーを使う場合は、Settings → Actions で「外部コントリビューターのワークフローに承認を必須にする」も設定してください。ワークフローの権限は `contents: read` だけです。
- 配布用ランチャーは、共有フォルダの外のファイルを読み込みません。また、SHA-256 が一致しないファイルは起動しません。
## 構成

```
extension.js                VS Code 拡張機能（コマンド・CodeLens・診断・MCP 登録）
scripts/
  access.ps1                CLI（JSON 出力）
  mcp-server.ps1            MCP サーバー（stdio）
  AccessVsc.Runner.ps1      実行エンジン（別スレッドで Access を操作、ダイアログ見張り、タイムアウト、排他、バックアップ）
  AccessVsc.Ops.ps1         Access 内での操作（エクスポート / インポート / コンパイル / 実行 / SQL）
  AccessVsc.Common.ps1      共通処理（ファイル名・文字コード・同期マニフェスト・行番号対応）
  AccessVsc.Win32.cs        ダイアログ検出・ボタン操作、JSON 整形
syntaxes/                   VBA とフォームテキストの色分け
test/                       テスト用 DB 作成スクリプト・スモークテスト
build-vsix.ps1              .vsix 作成（npm 不要）
```

## ライセンス

[MIT](LICENSE)。商用・非商用を問わず、誰でも自由に使用・改変・再配布できます（無保証）。
