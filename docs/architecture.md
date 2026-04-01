# Architecture

## 概要

`mikuproject` は、`MS Project XML` を意味の基軸として扱い、内部モデル `ProjectModel` を経由して複数の表現へ出し分ける構成を採る。

重視しているのは次の 3 点である。

- `MS Project XML` を基軸にした変換・可視化・限定編集
- 生成AI 連携を意識した projection / 再取込
- 人が読むための `WBS Excel ブック (.xlsx)` 帳票出力

## 全体構成

- 正本は `MS Project XML`
- 内部の中立表現は `ProjectModel`
- `.xlsx` と workbook JSON は補助的な入出力
- まずは意味的ラウンドトリップを優先

主な流れは次のとおり。

- `MS Project XML -> ProjectModel -> MS Project XML`
- `MS Project XML -> ProjectModel -> XLSX`
- `MS Project XML -> ProjectModel -> mikuproject_workbook_json`
- `MS Project XML -> ProjectModel -> Mermaid`
- `MS Project XML -> ProjectModel -> 生成AI向け JSON projection`

また、新規生成向けには次の流れを持つ。

- `project_draft_view -> ProjectModel -> MS Project XML / WBS Excel ブック (.xlsx)`

`XLSX Import` と workbook JSON import は、自由編集をそのまま受け入れるのではなく、限定列の部分更新として扱う。

`mikuproject-sample.xlsx` は `MS Project XML` との対応関係を確認するための構造忠実 workbook として扱う。列やシートの対応関係は崩さず、見た目改善は可読性補助に留める。これに対して `mikuproject-wbs-sample.xlsx` は人が読むための表示重視 workbook として扱う。

## Single-file Web App

配布物は `mikuproject.html` ひとつの single-file web app である。Web ブラウザさえあればインストール不要・ネットワーク不要で利用できる。

開発用 source は分割して管理するが、配布物としては single-file web app を維持する。

- HTML source: `mikuproject-src.html`
- TypeScript source: `src/ts/`
- generated JavaScript: `src/js/`
- CSS: `src/css/`

`mikuproject.html` は `mikuproject-src.html` をもとに、ローカル CSS / JS を単一 HTML へインライン展開して生成する。

## リポジトリ構成

- `mikuproject.html`: 生成済みの単一 HTML アプリ
- `mikuproject-src.html`: HTML ソース
- `package.json`: Node.js ベースの開発設定
- `src/ts/`: TypeScript ソース
- `src/js/`: `src/ts/` から生成し、Git 管理も行うブラウザ実行用 JavaScript
- `src/css/`: アプリ用 CSS
- `tests/`: Vitest ベースのテスト
- `testdata/`: XML テストデータ
- `scripts/`: ビルド補助スクリプト
- `docs/spec.md`: 現行仕様メモ
- `docs/gap-notes.md`: 保持項目や互換性のギャップメモ

## 画面構成

### Input

- `Load from file` からの `MS Project XML / XLSX / workbook JSON (.json) / 生成AI向け編集用 JSON (.editjson) / CSV + ParentID` の読込
- `project_draft_view` ベースで生成したサンプル XML の読込
- 生成AIが返した `project_draft_view` の JSON 貼り付け取込

### Overview

- 内部モデルの要約確認
- validation メッセージの確認
- `Daily / Weekly / Monthly Calendar` プレビューの確認
- `Project / Tasks / Resources / Assignments / Calendars` の preview 確認

### Output

- `MS Project XML` の保存
- `XLSX` の保存
- `WBS XLSX` の保存
- `XLSX` 相当 workbook JSON の保存
- `CSV + ParentID` の保存
- Mermaid fenced code block を含む `.md` の保存
- `Daily SVG / Weekly SVG / Monthly Calendar SVG` の保存
- 生成AI向け `project_overview_view` / `phase_detail_view` / `full bundle` の `.editjson` 保存

## 生成AI連携

`mikuproject` は、生成AIとの直接連携に `MS Project XML` ではなく `JSON` を使う。

- 既存 project 向けには `project_overview_view` と `phase_detail_view` を出力する
- 既存 project 向けには `full bundle` も出力できる
- 新規生成向けには、生成AIが返した `project_draft_view` を取り込める
- `mikuproject_workbook_json` は `.json`、生成AI 向け編集用 JSON は `.editjson` を推奨拡張子とする
- 現時点で UI から実装済みなのは `project_overview_view` / `phase_detail_view` / `full bundle` の出力と `project_draft_view` の取込である
- `task_edit_view` と Patch JSON 適用は設計メモ段階で、まだ実装していない

詳細な考え方は `docs/mikuproject-ai-json-spec.md` と `docs/msprojectxml-ai-integration.md` に置く。

## 開発

依存関係の導入:

```bash
npm install
```

TypeScript 由来のブラウザ実行 JavaScript を再生成:

```bash
npm run build:js
```

単一 HTML を再生成:

```bash
npm run build:html
```

サンプル XLSX を再生成:

```bash
npm run build:xlsx-sample
```

テスト実行:

```bash
npm test
```

ビルドとテストをまとめて実行:

```bash
npm run build
```

`npm run build` は `build:app` と `test` を順に実行する。`build:app` は `build:web` と `build:xlsx-sample` を順に実行する。

スクリプトの役割は次のとおり。

- `npm run build:js`: `src/ts/` から `src/js/` を生成する
- `npm run build:html`: `index-src.html` と `mikuproject-src.html` から `index.html` と `mikuproject.html` を生成する
- `npm run build:web`: JavaScript 生成と HTML 生成をまとめて行う
- `npm run build:xlsx-sample`: `local-data/` 配下へサンプル XLSX / Markdown を生成する
- `npm run build:app`: `build:web` と `build:xlsx-sample` を順に実行する

`scripts/build-project.mjs` は `--js-only` と `--html-only` を受け取り、JavaScript 生成と HTML 生成を切り替える。

## ソースと生成物の扱い

`src/ts/` を正本として扱い、`src/js/` はそこから生成する中間生成物として扱う。ただし、現状では `src/js/` も Git 管理する。ブラウザ実行、テスト、`build:xlsx-sample` は `src/js/` を参照する。

運用ルール:

- アプリロジックの修正は原則 `src/ts/` で行う
- `src/js/` は手編集の正本としては扱わない
- `src/ts/` を更新した場合は `npm run build:js` を実行し、`src/js/` の差分もあわせて扱う

## 現在の状態

- `package.json` と `package-lock.json` を持つ単独の Node.js プロジェクトとして扱える
- ソース配置は `src/ts/`, `src/js/`, `src/css/`
- `npm run build:js`、`npm run build:html`、`npm test` は通る
- `local-data/` と `node_modules/` は Git 管理対象外
- `local-data/` は確認用の再生成可能な生成物置き場として扱う
- `local-data/` 配下の sample や検証用出力は、消えてもよく、必要時に再生成できればよい前提とする

## 制約

- `MS Project` 実機は未保有
- 目標は XML の完全一致ではなく、意味的に往復できること
- `XLSX Import` の反映対象は限定列のみ
- workbook JSON import の反映対象も同じ限定列のみ
- `CSV + ParentID` は現在ファイルベースの補助入出力として扱う
- `Calendars` の `WeekDays / Exceptions / WorkWeeks` などは現時点では反映対象外

## ドキュメントの役割

- `README.md`: このリポジトリの入口
- `docs/architecture.md`: 全体構成、配布形態、ビルド、運用ルール
- `docs/spec.md`: 仕様と設計判断の置き場
- `docs/TODO.md`: 未完了作業の管理

## 実装 FAQ

### workbook JSON と `project_draft_view` は同じですか

違う。

- workbook JSON は、通常 workbook (`XLSX`) のテキスト版である
- `Project / Tasks / Resources / Assignments / Calendars / NonWorkingDays` を持つ
- import 時も、基本的には `XLSX Import` と同じ限定列の部分更新として扱う

一方で `project_draft_view` は、生成AI などが返す「新規 project 草案」である。

- 主に `project` と `tasks` を持つ草案用 JSON である
- 既存 project の限定更新ではなく、新しい `ProjectModel` を組み立てる入口である
- `Resources / Assignments / Calendars / NonWorkingDays` を workbook JSON と同じ形では持たない

つまり、`workbook JSON` は「既存 project を workbook 相当で扱うための JSON」、`project_draft_view` は「新規草案を取り込むための JSON」として分離している。

### 画面の `サンプル` は XML を読んでいますか、それとも JSON を読んでいますか

画面の通常 `サンプル` は `SAMPLE_XML` を読み込む。

ただし、その `SAMPLE_XML` 自体は内部では `SAMPLE_PROJECT_DRAFT_VIEW` から生成している。つまり、画面上の導線は XML 読込だが、サンプルの元データは `project_draft_view` 系である。
