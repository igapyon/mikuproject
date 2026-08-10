# miku-project

GitHub: https://github.com/igapyon/miku-project

Agent Skills 版: https://github.com/igapyon/miku-project-skills

`miku-project` は、`MS Project XML` を意味の基軸に、生成AIとの往復を支えるために設計されたローカル HTML ツールです。WBS の草案作成から再編集・再取込、人向けの可視化・帳票化までを、ひとつの流れとして扱えます。

`miku-project` の強みは、`MS Project XML` を意味の基軸として保ちながら、生成AIと人のあいだを往復できることです。WBS 草案の作成、生成AI が扱いやすい形への表現変換、生成AI から返った内容の再取込、人による確認と修正、そして可視化・帳票化までを、同じプロジェクト情報の流れとして扱えます。`XLSX`、`Markdown`、`JSON`、`Mermaid`、生成AI向け表現、そして必要に応じた `MS Project` への橋渡しは、それぞれの用途に応じた周辺表現として無理なく出し分けられます。

特に、次の 3 つを重視して設計しています。

- `MS Project XML` を意味の基軸として保つこと
- 生成AIと人の往復に適した表現変換 / 再取込 / 介在を支えること
- 人が読むための可視化と、WBS 帳票・SVG を含む成果物出力を提供すること

Web UI と single-file 配布物は [miku-project-web](https://github.com/igapyon/miku-project-web) が canonical repository として管理します。Main Application は UI 非依存の core API、Node.js CLI、browser-compatible runtime bundle を提供します。

`MS Project XML` を意味の基軸として扱い、`.xlsx` と workbook JSON は確認・可視化・限定編集のための周辺表現として扱います。生成AI 連携の編集用 JSON は、workbook JSON と区別するため当面 `.editjson` 拡張子を推奨します。

Agent Skills から `miku-project` の CLI / AI JSON 連携を扱うための関連リポジトリとして、[`miku-project-skills`](https://github.com/igapyon/miku-project-skills) があります。

## 代表的なユースケース

- その1: 生成AI との対話で WBS 草案を作成し、`miku-project` に取り込んで、人と生成AIが確認・修正しながら、帳票や可視化成果物として仕上げる
- その2: 既存の `MS Project XML` を `miku-project` に取り込み、内容を確認しながら、`WBS Excel ブック (.xlsx)` や日次・週次のガント表現や月次カレンダーの `SVG`、`Markdown` などの人向け成果物へ展開する
- その3: `miku-project` で扱う WBS やプロジェクト情報を生成AI向けに表現変換し、生成AIが返した結果を再び取り込みながら、人と生成AIがレビュー・調整・再利用しやすい形へ整える

import / export / 生成AI連携の使い分けを「何をしたいか」から辿りたい場合は、[docs/import-export-workflows.md](docs/import-export-workflows.md) を参照してください。`replace / merge / patch` の違い、`project-overview / task-edit / phase-detail / bundle` の使い分け、既存WBSの安全な局所修正フローをまとめています。

## Web App

ブラウザでの入力、可視化、ダウンロード、screenshots、single-file HTML は [miku-project-web](https://github.com/igapyon/miku-project-web) を参照してください。Web App は固定した Main Application runtime を build 時に検証・内包し、実行時にネットワークから runtime を取得しません。

## できること

- 生成AIに渡すためのプロジェクト概要・工程詳細・一式データの出力（`project_overview_view` / `phase_detail_view` / `full bundle`）
- 生成AIが返した WBS 素案の取込（`project_draft_view`）
- 生成AI向けの task 単位編集ビューの出力（`task_edit_view`）
- 生成AIが返した Patch JSON の取込と反映
- `MS Project XML` の読込
- `ProjectModel` への変換と内容確認
- 日次・週次のガント表現、および月次カレンダー可視化の `SVG` 出力
- `Project / Tasks / Resources / Assignments / Calendars` workbook の構造を保ったまま、`XLSX / JSON` で `Export / Import`
- `CSV + ParentID` のファイル読込とダウンロード
- `MS Project XML` の再生成
- 表示専用の `WBS XLSX Export`
- Mermaid gantt テキスト生成

## 使い始め方

Web App の使い始め方とブラウザ配布物は [miku-project-web](https://github.com/igapyon/miku-project-web) を参照してください。

CLI の使い始め方、AI JSON API、runtime contract はこの repository の [docs/development.md](docs/development.md) と [docs/browser-runtime.md](docs/browser-runtime.md) を参照してください。

## 開発

```bash
npm install
npm run build
npm test
```

`npm run build` には `build:core`、`build:browser-runtime`、`build:cli-bundle`、`test:fast` が含まれる。開発用コマンドの詳細、テスト運用、`local-data/` の扱いは [docs/development.md](docs/development.md) を参照してください。

`workplace/` は外部リポジトリの一時 clone、展開物、検証成果物などを置くローカル作業領域として扱い、`workplace/.gitkeep` 以外は Git 管理しません。

`src/js/` は core runtime の Git 管理する生成物です。手編集はせず、対応する `src/ts/` を更新して `npm run build:core` で再生成します。`node_modules/`、`.npm-cache/`、`coverage/`、`bundle/`、`local-data/`、`workplace/` と個人用の `.vscode/mcp.json` はローカル作業用であり、Git 管理しません。

miku-soft の共有標準と、このリポジトリの追従状況は [docs/miku-soft-reference.md](docs/miku-soft-reference.md) と [docs/migration-worklog.md](docs/migration-worklog.md) を参照してください。

CLI、Java CLI、Agent Skillsを中心にしたゼロベース再設計は、[2026-08-09版仕様](docs/miku-project-zero-base-spec-v20260809.md)、[2026-08-10版実施計画](docs/miku-project-zero-base-implementation-plan-v20260810.md)、[semantic contract v1](docs/miku-project-semantic-contract-v1.md)、[format and loss contract v1](docs/miku-project-format-and-loss-contract-v1.md)、[change contract v1](docs/miku-project-change-contract-v1.md)、[CLI contract v1](docs/miku-project-cli-contract-v1.md)、[CLI result and diagnostics contract v1](docs/miku-project-cli-result-contract-v1.md)、[runtime capability contract v1](docs/miku-project-runtime-capability-contract-v1.md)、[runtime manifest contract v1](docs/miku-project-runtime-manifest-contract-v1.md)、[conformance corpus v1](docs/miku-project-conformance-corpus-v1.md)、[human gate and next action contract v1](docs/miku-project-human-gate-and-next-action-contract-v1.md) を参照してください。WebとMCPは、この初期計画では後続検討です。

## 再利用 API

Web App / Agent Skills / CLI / MCP から使いやすい集約入口として `globalThis.__mikuProjectCoreApi` を公開しています。

- `getAiJsonSpec()` / `getAiJsonSpecText()`: `miku-project-ai-json-spec` の安定取得
- `parseAiJsonText()` / `importAiJsonDocument()` / `importAiJsonText()`: `project_draft_view` / Patch JSON / workbook JSON の UI 非依存な共通入口
- `importExternal()`: `MS Project XML / XLSX / workbook JSON / project_draft_view / patch JSON` の format-aware な共通 import 入口
- `projectModel`, `msProject`, `aiViews`, `workbookJson`, `xlsx`, `patchJson`, `report`: `ProjectModel` 周りの import / export / validate の集約 entrypoint

`xlsx` では次を公開する。

- `decodeWorkbook()` / `encodeWorkbook()`: workbook binary と workbook object の相互変換
- `exportWorkbook()`: `ProjectModel` から構造忠実 workbook を生成
- `importAsProjectModel()` / `importIntoProjectModel()`: workbook object を `ProjectModel` へ replace / merge import

`report` では次を公開する。

- `all.export()`: `report` 成果物一式 ZIP の生成
- `wbsXlsx.exportWorkbook()` / `wbsXlsx.exportBytes()`: `WBS XLSX` workbook と `.xlsx` bytes の生成
- `svg.exportDaily()` / `svg.exportWeekly()` / `svg.exportMonthlyCalendar()`: `Daily / Weekly / Monthly Calendar SVG` の生成
- `wbsMarkdown.export()`: `WBS Markdown` の生成
- `mermaid.exportGantt()`: Mermaid gantt text の生成

`importExternal()` の最小例:

```ts
const api = globalThis.__mikuProjectCoreApi;

const replaceResult = api.importExternal({
  source: { format: "xlsx", bytes },
  mode: "replace"
});

const mergeResult = api.importExternal({
  source: { format: "workbook_json", document },
  mode: "merge",
  baseModel
});
```

first cut の対応は次のとおり。

- `ms_project_xml`: `replace` のみ
- `xlsx`: `replace` / `merge`
- `workbook_json`: `replace` / `merge`
- `project_draft_view`: `replace` のみ
- `patch_json`: `patch` のみ

Node 側から `core API` を起動する最小 helper は [`scripts/lib/core-api-loader.mjs`](scripts/lib/core-api-loader.mjs) に置いている。CLI ではこの loader が `globalThis.__mikuprojectXmlDom` を初期化し、XML 系の `DOMParser` / `XMLSerializer` / XML document 生成を環境非依存に扱う。`importExternal()` の利用例は [`scripts/core-api-import-external-example.mjs`](scripts/core-api-import-external-example.mjs) を参照。

`AI JSON spec` 単体の取得用には `globalThis.__mikuprojectAiJsonSpec` も公開しています。

## CLI first cut

Node 側から `core API` を薄く包む最小 CLI first cut として、次の入口を追加している。

- `miku-project ai spec`
- `miku-project --version`
- `miku-project ai export project-overview`
- `miku-project ai export task-edit`
- `miku-project ai export phase-detail`
- `miku-project ai export bundle`
- `miku-project ai detect-kind`
- `miku-project ai validate-patch`
- `miku-project state from-draft`
- `miku-project state summarize`
- `miku-project state diff`
- `miku-project state apply-patch`
- `miku-project import xlsx`
- `miku-project export workbook-json`
- `miku-project export xml`
- `miku-project export xlsx`
- `miku-project report wbs-xlsx`
- `miku-project report daily-svg`
- `miku-project report weekly-svg`
- `miku-project report monthly-calendar-svg`
- `miku-project report all`
- `miku-project report wbs-markdown`
- `miku-project report mermaid`

text 系の主成果物は `stdout` または `--out <path>`、warning / diagnostics は `stderr` を基本とする。
XLSX / ZIP などの binary artifact は `--out <path>` へ出力する。stream-friendly な binary 入出力が必要な場合は、明示的に `--in-base64 -` / `--out-base64 -` を使う。
`--diagnostics text|json` を受けるコマンドでは、構造化 diagnostics を扱える。

既存WBSの安全な局所修正フローは、まず `ai export project-overview` で全体を見て、`task-edit` または `phase-detail` を AI に渡し、返ってきた `patch_json` を `validate-patch` してから `state apply-patch` / `state diff` へ進む形を基本とする。導線全体は [docs/import-export-workflows.md](docs/import-export-workflows.md) を参照。

例:

```bash
miku-project --version
miku-project ai spec
miku-project ai export project-overview --in workbook.json --out overview.editjson
miku-project ai export task-edit --in workbook.json --task-uid 123 --out task.editjson
miku-project ai export phase-detail --in workbook.json --phase-uid 100 --root-task-uid 123 --max-depth 2 --out phase.editjson
miku-project ai detect-kind --in patch.json
miku-project ai validate-patch --state workbook.json --in patch.json --diagnostics json
miku-project state from-draft --in draft.json --out workbook.json
miku-project state summarize --in workbook.json --diagnostics json
miku-project state diff --before workbook.before.json --after workbook.after.json --diagnostics json
miku-project state apply-patch --state workbook.json --in patch.json --out workbook.next.json
miku-project import xlsx --in project.xlsx --out workbook.json
base64 < project.xlsx | miku-project import xlsx --in-base64 - --out -
miku-project export xml --in workbook.json --out project.xml
miku-project export xlsx --in workbook.json --out project.xlsx
miku-project export xlsx --in workbook.json --out-base64 -
miku-project report wbs-xlsx --in workbook.json --out project-wbs.xlsx
miku-project report daily-svg --in workbook.json --out project-daily.svg
miku-project report weekly-svg --in workbook.json --out project-weekly.svg
miku-project report monthly-calendar-svg --in workbook.json --out project-monthly.zip
miku-project report all --in workbook.json --out project-report-bundle.zip
miku-project report wbs-markdown --in workbook.json --out project-wbs.md
miku-project report mermaid --in workbook.json --out project.mmd
```

`report monthly-calendar-svg` は月別 SVG 一式をまとめた ZIP を出力する。
`report all` は `wbs.xlsx` / `wbs.md` / `mermaid.mmd` / `daily.svg` / `weekly.svg` / `monthly-calendar/YYYY-MM.svg` をまとめた ZIP を出力する。

## Browser runtime artifact

`miku-project-web` などの browser downstream が build 時に取り込む importable runtime を生成できる。

```bash
npm run build:browser-runtime
```

既定の出力先は `bundle/miku-project-runtime.mjs` である。`version`、`embeddedCorePaths`、`loadMikuProjectRuntime(options)` と default loader を公開し、Node.js API、CLI 自動実行、UI event/download 層を含めない。公開契約、smoke、Release asset、SHA-256 固定による downstream 取込は [Browser Runtime Contract](docs/browser-runtime.md) を参照。

## CLI runtime artifact

`miku-project` 側で、Agent Skills など下流から受け取って実行できる単一 `MJS` CLI runtime artifact を生成できる。

```bash
npm run build:cli-bundle
```

既定の出力先は `bundle/miku-project.mjs` と `bundle/miku-project-sources.tgz` である。
この artifact には、CLI entrypoint、`core API` 実行に必要な `src/js` runtime、XML DOM 実装を含めている。
`miku-project-sources.tgz` には、再ビルド・監査・下流確認用の source / docs / tests をまとめて格納する。

生成後は追加の `npm install` なしで、そのまま CLI 実行に使える。たとえば次で動く。

```bash
node bundle/miku-project.mjs ai spec
node bundle/miku-project.mjs export xml --in workbook.json --out project.xml
```

`miku-project-skills` などの下流 Agent Skills では、この Node.js runtime artifact を `skills/miku-project/runtime/miku-project.mjs` のような skill-local runtime path に配置して使う想定である。

生成時は repo root の `node_modules/@xmldom/xmldom` から XML DOM 実装を artifact 内へ埋め込む。そのため、artifact 生成前には一度 `npm install` 済みであることを前提とする。生成後の実行時には、`@xmldom/xmldom` や `jsdom` の `node_modules` は不要である。

## 関連ドキュメント

- [docs/miku-project-zero-base-spec-v20260809.md](docs/miku-project-zero-base-spec-v20260809.md)
- [docs/miku-project-zero-base-implementation-plan-v20260810.md](docs/miku-project-zero-base-implementation-plan-v20260810.md)
- [docs/miku-project-semantic-contract-v1.md](docs/miku-project-semantic-contract-v1.md)
- [docs/miku-project-semantic-fixture-catalog-v1.md](docs/miku-project-semantic-fixture-catalog-v1.md)
- [docs/miku-project-format-and-loss-contract-v1.md](docs/miku-project-format-and-loss-contract-v1.md)
- [docs/miku-project-change-contract-v1.md](docs/miku-project-change-contract-v1.md)
- [docs/miku-project-cli-contract-v1.md](docs/miku-project-cli-contract-v1.md)
- [docs/miku-project-cli-result-contract-v1.md](docs/miku-project-cli-result-contract-v1.md)
- [docs/schemas/miku-project-artifacts-v1.schema.json](docs/schemas/miku-project-artifacts-v1.schema.json)
- [docs/schemas/miku-project-cli-result-v1.schema.json](docs/schemas/miku-project-cli-result-v1.schema.json)
- [docs/schemas/miku-project-cli-diagnostic-v1.schema.json](docs/schemas/miku-project-cli-diagnostic-v1.schema.json)
- [docs/miku-project-runtime-capability-contract-v1.md](docs/miku-project-runtime-capability-contract-v1.md)
- [docs/miku-project-runtime-manifest-contract-v1.md](docs/miku-project-runtime-manifest-contract-v1.md)
- [docs/miku-project-conformance-corpus-v1.md](docs/miku-project-conformance-corpus-v1.md)
- [docs/miku-project-human-gate-and-next-action-contract-v1.md](docs/miku-project-human-gate-and-next-action-contract-v1.md)
- [docs/architecture.md](docs/architecture.md)
- [docs/import-export-workflows.md](docs/import-export-workflows.md)
- [docs/core-api-import-export-notes.md](docs/core-api-import-export-notes.md)
- [docs/development.md](docs/development.md)
- [docs/spec.md](docs/spec.md)
- [docs/gap-notes.md](docs/gap-notes.md)
- [docs/miku-project-ai-json-spec.md](docs/miku-project-ai-json-spec.md)
- [docs/msprojectxml-ai-integration.md](docs/msprojectxml-ai-integration.md)
- [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md)
- [docs/TODO.md](docs/TODO.md)
- [CONTRIBUTING.md](CONTRIBUTING.md)
- [CONTRIBUTORS.md](CONTRIBUTORS.md)
- [CODE_OF_CONDUCT.md](CODE_OF_CONDUCT.md)
- [LICENSE](LICENSE)
