# Core API Import / Export Notes

この文書は、`globalThis.__mikuprojectCoreApi` の import / export 公開方針について、現時点の判断経緯を簡潔に残すためのメモである。

確定仕様の置き場は `README.md` と `docs/architecture.md` と `docs/spec.md` とし、この文書はそれらの補助メモとして扱う。

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
- `mikuproject state from-draft`
- `mikuproject state apply-patch`
- `mikuproject export workbook-json`
- `mikuproject export xml`
- `mikuproject export xlsx`

`report` 系は first cut から分離し、次段候補として扱う。
