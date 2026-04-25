# Core API Import / Export Notes

この文書は、`globalThis.__mikuprojectCoreApi` の import / export 公開方針について、現時点の判断経緯を簡潔に残すためのメモである。

確定仕様の置き場は `README.md` と `docs/architecture.md` と `docs/spec.md` とし、この文書はそれらの補助メモとして扱う。

利用者向けの導線整理は `docs/import-export-workflows.md` を参照する。

## 前提

`mikuproject` は `MS Project XML` を意味の基軸として扱う。

- 正本: `MS Project XML`
- 内部中立表現: `ProjectModel`
- 補助交換形式: `XLSX`, `mikuproject_workbook_json`
- 派生表示 / AI 向け表現: `WBS XLSX`, `Markdown`, `SVG`, `Mermaid`, `aiViews`

外部再利用向けには、single-file web app 向けの既存 global を維持しつつ、集約入口として `globalThis.__mikuprojectCoreApi` を公開する。

Node / CLI 側では、`scripts/lib/core-api-loader.mjs` が `globalThis.__mikuprojectXmlDom` も初期化する。これにより、`msproject-xml` と `excel-io` の XML parse / serialize はブラウザ DOM 直依存ではなく、CLI では `@xmldom/xmldom` 優先、未導入時は `jsdom` フォールバックで動く。

## 現状整理

### format-specific export

主要交換形式の export は、現時点では format ごとの公開面をそのまま使う整理とする。

- `msProject.exportToXml(model)`
- `xlsx.exportWorkbook(model)`
- `workbookJson.exportDocument(model)`

この段階では、`exportExternal(...)` のような統一 export 入口は持たない。

理由:

- 現状でも利用側の負担が大きくない
- 返り値の型が format ごとに自然に異なる
- `WBS XLSX` や `SVG` など派生出力まで同列に混ぜると責務が広がりやすい

したがって、export 側はまず現状維持で十分と判断する。

### format-specific import

主要交換形式の import も、現時点では format ごとの公開面を持つ。

- `msProject.importFromXml(text)`
- `xlsx.decodeWorkbook(bytes)`
- `xlsx.importAsProjectModel(workbook)`
- `xlsx.importIntoProjectModel(workbook, baseModel)`
- `workbookJson.importAsProjectModel(document)`
- `workbookJson.importIntoProjectModel(document, baseModel)`
- `patchJson.applyToProjectModel(document, baseModel)`
- `importAiJsonDocument(document, options)`

このうち `XLSX` については、`core API` 上でも次を公開する。

- `xlsx.decodeWorkbook(bytes)`
- `xlsx.encodeWorkbook(workbook)`
- `xlsx.exportWorkbook(model)`
- `xlsx.importAsProjectModel(workbook)`
- `xlsx.importIntoProjectModel(workbook, baseModel)`

## import 側の統一入口

`importExternal(...)` を first cut として追加した。

```ts
api.importExternal({
  source: { format: "xlsx", bytes },
  mode: "replace"
});

api.importExternal({
  source: { format: "xlsx", bytes },
  mode: "merge",
  baseModel
});
```

first cut の対象と対応 `mode` は次のとおり。

- `ms_project_xml`: `replace`
- `xlsx`: `replace`, `merge`
- `workbook_json`: `replace`, `merge`
- `project_draft_view`: `replace`
- `patch_json`: `patch`

不正な組み合わせは `core API` 側で明示的に reject する。

reject 例:

- `ms_project_xml + merge` は reject する
- `patch_json + replace` は reject する
- `xlsx + patch` は reject する
- `workbook_json + patch` は reject する
- `merge` で `baseModel` がない場合は reject する
- reject 時の error 文言は `importExternal: format=...` と `mode=...` を含め、期待する mode と `baseModel` 要否が分かる形に寄せる

## import 側の次段候補

import 側は、利用側が format ごとの分岐を毎回持ちやすいため、統一入口の価値が export より高い。

`importExternal(...)` では、`text / bytes / document / workbook` のような物理媒体ではなく、`replace / merge / patch` のような操作意味を前面に出す設計を採る。

### この案を好む理由

- `format` と `operation` を分離できる
- `XLSX / XML / workbook JSON / patch JSON` を同じ入口で扱える
- `baseModel` が必要な場面を `mode` と組み合わせて自然に表現できる
- `importAiJsonDocument()` の責務を壊さずに済む

## `importAiJsonDocument()` を広げない理由

`importAiJsonDocument()` は名前の通り、AI JSON 系の共通入口として残す。

ここへ `XLSX` や `MS Project XML` を直接入れ始めると、

- API 名と責務がずれる
- binary / text / JSON の境界が曖昧になる
- 将来の保守で利用者が混乱しやすい

そのため、統一 import 入口が必要なら `importAiJsonDocument()` を拡張するのではなく、別に `importExternal(...)` のような上位 API を足す方針を優先する。

## 当面の判断

- `xlsx.*` を含む format-specific API は維持する
- import 側は、`importExternal(...)` を上位 API として使えるようにする
- export 側は、まず現状維持で十分とする
- ただし report / presentation outputs については、`report.*` 配下に format-specific export を追加してよい
- `WBS XLSX` / `Markdown` / `SVG` / `Mermaid` / `aiViews` は、主要交換形式の統一 export と無理に同列化しない

## CLI naming の当面案

`CLI` を追加する場合も、`core API` の責務分離をそのまま保つ。

- `ai`: AI との契約や spec を扱う
- `state`: state の生成・更新を扱う
- `export`: `workbook-json` / `xml` / `xlsx` のような主要交換形式を扱う
- `report`: `wbs-xlsx` / `wbs-markdown` / `mermaid` / `daily-svg` / `weekly-svg` / `monthly-calendar-svg` のような派生出力を扱う

first cut の候補は次とする。

- `mikuproject ai spec`
- `mikuproject ai export project-overview`
- `mikuproject ai export task-edit`
- `mikuproject ai export phase-detail`
- `mikuproject ai export bundle`
- `mikuproject ai detect-kind`
- `mikuproject ai validate-patch`
- `mikuproject state from-draft`
- `mikuproject state summarize`
- `mikuproject state diff`
- `mikuproject state apply-patch`
- `mikuproject export workbook-json`
- `mikuproject export xml`
- `mikuproject export xlsx`

`report` 系は first cut から分離し、次段候補として扱う。

補足:

- `mikuproject ai export task-edit` は `mikuproject_workbook_json` を入力とし、`task_edit_view` を stdout または `--out` へ出力する
- `mikuproject ai export phase-detail` は `mikuproject_workbook_json` を入力とし、`phase_detail_view` を stdout または `--out` へ出力する
- `mikuproject ai export bundle` は `mikuproject_workbook_json` を入力とし、`ai_projection_bundle` を stdout または `--out` へ出力する
- `bundle` は `ai export` 配下に残し、主用途は調査 / デバッグ / 比較検証とする
- `task-edit` は `--select auto|first-task|uid`、`phase-detail` は `--select auto|first-phase|uid` を受ける
- `--select` は現時点では `task-edit` / `phase-detail` 専用とし、`project-overview` には広げない
- `bundle` / `detect-kind` / `validate-patch` / `state summarize` / `state diff` は現時点で `--select` を持たない
- `export phase-detail` の主オプションは `--phase-uid` / `--mode` / `--root-task-uid` / `--max-depth` とする
- `mikuproject ai validate-patch` は `--state` と `patch_json` を受け、dry-run apply ベースの validation を返す
- `mikuproject ai detect-kind` は JSON document の kind 判定を返す
- `mikuproject state summarize` / `state diff` は state 確認系の補助コマンドとする
- `--in -` と `--out -` を受け、標準入力 / 標準出力を明示的に指定できる
- 同一コマンドで標準入力を読む入力オプションは 1 つだけとする
- `--in path` が最優先、`--in -` は明示 stdin、`--in` 省略時だけ暗黙 stdin を使う
- `--out path` が最優先、`--out -` または `--out` 省略時は stdout を使う
- `ai export` / `detect-kind` / `validate-patch` / `state summarize` / `state diff` / `state apply-patch` / `export` / `report` は `--diagnostics text|json` を受けられる
- `json` diagnostics は少なくとも `diagnostics_version` / `ok` / `command` / `context` / `status` / `exit_code` / `warning_count` / `error_count` / `io` / `warnings` / `errors` を共通キーとして持ち、追加メタ情報をコマンド別に載せる
- `json` diagnostics は `status` も共通キーとし、少なくとも `success` / `warning` / `noop` / `error` を使う
- `--diagnostics json` 指定時は、usage error や処理失敗でも stderr に JSON diagnostics を返す
- warning-only は `status=warning` / `exit_code=0`、no-op は `status=noop` / `exit_code=0`、usage error は `exit_code=2`、processing error は `exit_code=1` とする
- 異常系 diagnostics では `error_type=usage_error|processing_error` を返せる
- 異常系 diagnostics では `error_code` に安定識別子を返せる
- 異常系 diagnostics では top-level に `error_details` を返せる
- 異常系 diagnostics の `errors[]` 要素にも `code` を返せる
- 異常系 diagnostics の `errors[]` 要素には `details` を返せる
- 主要な usage error は文言推定ではなく、CLI 側で安定した `error_code` を直接付与する
- 主要な processing error も、kind mismatch や detect-kind failure などは安定した `error_code` を直接付与する
- JSON parse failure も `invalid_json_input` として安定した `error_code` / `details.context` を返す
- `io` には、各入力が `file` / `stdin` / `stdin_implicit` のどれだったかと、出力先が `stdout` / `file` のどちらかを載せる
- この `io` は正常系だけでなく、`--diagnostics json` の異常系 diagnostics にも載せる
- 異常系の `io` でも、コマンド形から `stdin_implicit` を推定して返す
- `scripts/cli-ai-workflow-example.mjs` で、CLI を使った局所 projection 取得から patch validate / apply / diff までの最小例を示す
- `scripts/cli-ai-stdio-example.mjs` で、`--in -` / `--out -` を使った stdio ベースの最小例を示す

## Node.js CLI runtime artifact

`npm run build:cli-bundle` は、下流 Agent Skills に渡す Node.js CLI runtime artifact として、単一 `MJS` ファイルと source archive を生成する。

既定の出力先は次のとおり。

```text
bundle/mikuproject.mjs
bundle/mikuproject-sources.tgz
```

この artifact は、`scripts/mikuproject-cli.mjs` 相当の CLI entrypoint、`core API` 実行に必要な `src/js` runtime、XML DOM 実装を 1 ファイルに埋め込む。生成後の実行時には、従来の `bundle/mikuproject-cli-bundle/` ディレクトリや `node_modules` を同梱しない。
`mikuproject-sources.tgz` は実行には不要だが、受け取り側が source / docs / tests を確認し、必要に応じて再ビルドや監査を行うための補助 artifact として扱う。

`mikuproject-skills` 側では、この artifact を `skills/mikuproject/runtime/mikuproject.mjs` のような skill-local runtime path に配置して呼び出す想定である。
