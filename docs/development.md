# Development

## Setup

```bash
npm ci
```

- CLI XML parsing and serialization prefer `@xmldom/xmldom`.
- `jsdom` provides the remaining Web API compatibility used by the Node.js CLI and tests.
- XML DOM is obtained through `globalThis.__mikuprojectXmlDom` by `msproject-xml` and `excel-io`.

## Common commands

```bash
npm run build
npm test
```

- `npm run build:core` transpiles the core TypeScript files to their tracked `src/js/` runtime modules.
- `npm run build:browser-runtime` creates the browser-compatible downstream artifact at `bundle/miku-project-runtime.mjs`.
- `npm run build:cli-bundle` creates the Node.js CLI artifact at `bundle/miku-project.mjs` and a source archive.
- `npm run build:full` runs the core build, both bundles, and the complete Main Application suite.
- `npm run build:xlsx-sample` writes optional sample XLSX and Markdown output under `local-data/`.

テストsuiteは次の役割に固定する。追加した `*.test.js` は必ずいずれかへ明示的に分類し、`tests/mikuproject-test-suite-topology.test.js` が実行漏れと重複を検出する。

- `npm run test:fast`: 日常開発用。core utility / codec / workbook / core APIとlegacy CLI compatibility contractを検証する。時間のかかるCLI統合testとbrowser runtime testは含めない。
- `npm run test:full`: `fast` にCLI統合testとbrowser runtime contractを加えた完全回帰。
- `npm test` / `npm run test:all`: 現在は `full` と同じ全checked-in test fileを実行する、CI向けの安定した完全suite alias。

Web UI development, browser/UI tests, single-file HTML build, and offline browser verification are maintained in [miku-project-web](https://github.com/igapyon/miku-project-web).

## local-data

`local-data/` is a disposable workspace for generated samples and verification outputs. It is not tracked by Git and must be reproducible when needed.
