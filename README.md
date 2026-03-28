# mikuproject

`mikuproject` は、`MS Project XML` を読み込み、内部の `ProjectModel` へ変換し、再び `MS Project XML` や補助表現へ出力するためのローカル HTML ツールです。

`MS Project XML` を意味の基軸として扱い、`.xlsx` は確認・可視化・限定編集のための周辺表現として扱います。

## できること

- `MS Project XML` の読込
- `ProjectModel` への変換と内容確認
- `MS Project XML` の再生成
- Mermaid gantt テキスト生成
- `CSV + ParentID` 生成と解析
- 構造忠実な `Project / Tasks / Resources / Assignments / Calendars` workbook の `XLSX Export / Import`
- 表示専用の `WBS XLSX Export`

## 現在の考え方

- 正本は `MS Project XML`
- 内部の中立表現は `ProjectModel`
- `.xlsx` は補助的な入出力
- まずは意味的ラウンドトリップを優先

`XLSX Import` は自由編集をそのまま受け入れるのではなく、限定列の部分更新として扱います。

`mikuproject-sample.xlsx` は `MS Project XML` との対応関係を確認するための構造忠実 workbook として扱います。列やシートの対応関係は崩さず、見た目改善は可読性補助に留めます。これに対して `mikuproject-wbs-sample.xlsx` は人が読むための表示重視 workbook として扱います。

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
- `mikuproject-spec.md`: 現行仕様メモ
- `mikuproject-gap-notes.md`: 保持項目や互換性のギャップメモ

## ドキュメントの役割

- `README.md`: このリポジトリの入口です。概要、使い方、ビルド方法、運用ルールを簡潔にまとめます。
- `mikuproject-spec.md`: 仕様と設計判断の置き場です。データモデル、入出力方針、対応範囲、制約を継続的に整理します。
- `TODO.md`: まだ終わっていない作業だけを書きます。方針メモや完了済み事項は原則ここに残しません。

## 使い方

もっとも簡単なのは、生成済みの [mikuproject.html](/Users/igapyon/Documents/git/mikuproject/mikuproject.html) をブラウザで開く方法です。

画面上では主に次の操作を行えます。

- `XML Import`
- `XML を解析`
- `XML を再生成`
- `XLSX Export`
- `XLSX Import`
- `WBS XLSX Export`
- `Mermaid を生成`

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

`mikuproject.html` は `mikuproject-src.html` をもとに、ローカル CSS / JS と `src/vendor/mermaid/mermaid.min.js` を単一 HTML へインライン展開して生成します。

`Mermaid Preview` を single-file WebApp のままオフライン再現するため、`Mermaid` ランタイムは `src/vendor/mermaid/mermaid.min.js` を同梱します。

- 同梱ファイル: `src/vendor/mermaid/mermaid.min.js`
- バージョン: `mermaid@11.12.0`
- 取得元: `https://cdn.jsdelivr.net/npm/mermaid@11.12.0/dist/mermaid.min.js`
- 更新手順: 上記 URL のバージョンを差し替えて `src/vendor/mermaid/mermaid.min.js` を更新し、`npm run build:app` を実行する

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

`npm run build` は `build:app` と `test` を順に実行します。

スクリプトの役割は次のとおりです。

- `npm run build:js`: `src/ts/` から `src/js/` を生成します。
- `npm run build:html`: `index-src.html` と `mikuproject-src.html` から `index.html` と `mikuproject.html` を生成します。
- `npm run build:xlsx-sample`: `local-data/` 配下へサンプル XLSX を生成します。
- `npm run build:app`: `build:js`、`build:html`、`build:xlsx-sample` を順に実行します。

[scripts/build-project.mjs](/Users/igapyon/Documents/git/mikuproject/scripts/build-project.mjs) は `--js-only` と `--html-only` を受け取り、JavaScript 生成と HTML 生成を切り替えます。

`src/ts/` を正本として扱い、`src/js/` はそこから生成する中間生成物として扱います。ただし、現状では `src/js/` も Git 管理します。ブラウザ実行、テスト、`build:xlsx-sample` は `src/js/` を参照します。

運用ルール:

- アプリロジックの修正は原則 `src/ts/` で行います。
- `src/js/` は手編集の正本としては扱いません。
- `src/ts/` を更新した場合は `npm run build:js` を実行し、`src/js/` の差分もあわせて扱います。

## 現在の状態

- `package.json` と `package-lock.json` を持つ単独の Node.js プロジェクトとして扱える
- ソース配置は `src/ts/`, `src/js/`, `src/css/`
- 外部ランタイムの同梱先は `src/vendor/`
- `npm run build:js`、`npm run build:html`、`npm test` は通る
- `local-data/` と `node_modules/` は Git 管理対象外

## 制約

- `MS Project` 実機は未保有
- 目標は XML の完全一致ではなく、意味的に往復できること
- `XLSX Import` の反映対象は限定列のみ
- `Calendars` の `WeekDays / Exceptions / WorkWeeks` などは現時点では反映対象外
- Mermaid の SVG プレビューは `src/vendor/mermaid/mermaid.min.js` を `build:html` で内包した `mikuproject.html` を前提とする

## 関連ドキュメント

- [mikuproject-spec.md](/Users/igapyon/Documents/git/mikuproject/mikuproject-spec.md)
- [mikuproject-gap-notes.md](/Users/igapyon/Documents/git/mikuproject/mikuproject-gap-notes.md)
- [TODO.md](/Users/igapyon/Documents/git/mikuproject/TODO.md)
- [LICENSE](/Users/igapyon/Documents/git/mikuproject/LICENSE)
