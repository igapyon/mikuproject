# Development

## セットアップ

```bash
npm install
```

補足:

- CLI の XML parse / serialize は `@xmldom/xmldom` を優先利用する
- `jsdom` は CLI 上の HTML / Blob / File など、XML 以外の Web API 補完にも使っているため依存として残す
- XML DOM の入口は `globalThis.__mikuprojectXmlDom` で、`msproject-xml` と `excel-io` が参照する

## よく使うコマンド

```bash
npm run build
npm test
```

- `npm run build` は日常開発向けの標準 build で、`build:web`、`build:cli-bundle`、`test:fast` を順に実行する
- `npm run build:app` は `build:web` と `build:xlsx-sample` を順に実行する
- `npm run build:full` は `build:web`、`build:cli-bundle`、`test:full` を順に実行し、日常で見たい core UI smoke suite までを確認する
- `build:xlsx-sample` は必要なときだけ `build:app` か `npm run build:xlsx-sample` で明示実行する

## テストの使い分け

- `npm run test:extended` は validation、`XLSX import`、preview 切替、重い patch/export 系、projection/replace 系を追加で確認する
- `npm test` / `npm run test:all` はそれらも含めた完全実行である

## `local-data/` の扱い

`local-data/` は確認用の再生成可能な生成物置き場として扱う。ここに出す sample や検証用出力は、Git 管理下の永続成果物ではなく、必要時に再生成できればよい前提とする。
