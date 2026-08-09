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

`npm test`, `npm run test:fast`, `npm run test:full`, and `npm run test:all` validate the Main Application core, CLI, and browser runtime contract.

Web UI development, browser/UI tests, single-file HTML build, and offline browser verification are maintained in [miku-project-web](https://github.com/igapyon/miku-project-web).

## local-data

`local-data/` is a disposable workspace for generated samples and verification outputs. It is not tracked by Git and must be reproducible when needed.
