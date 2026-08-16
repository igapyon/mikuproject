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
- `npm run build:cli-v1-runtime` creates the versioned v1 Node runtime under `runtime/node/` only from a clean source tree whose HEAD has the exact `v<package.version>` tag. It refuses existing output and is not a Release command. P4.10's test/audit-owned consumer pins the raw `runtime-manifest.json` SHA-256 supplied outside the runtime, verifies the canonical executable/source entries, sizes, and SHA-256 digests before launching Node, and validates the complete result/runtime binding.
- `npm run verify:cli-v1-release-candidate -- --runtime-dir <directory> --lock <lock.json>` verifies an already-built internal reference candidate against a Git-tracked external lock, copies only the three locked members to an isolated consumer, and replays `validate` / `inspect` / `plan-change` / `apply-change` / `verify-artifact`. The Gate G4 `v1.0.3` lock is `docs/miku-project-node-reference-runtime-lock-v1.0.3.json`; its local retained candidate is expected at `workplace/gate-g4/v1.0.3/runtime/` and is intentionally ignored by Git.
- The Gate G4 lock is an internal approval trust anchor, not a public Release checksum or Skills lock. The current `.github/workflows/release-runtime-bundles.yml` does not build or upload this v1 three-member runtime; do not publish `v1.0.3` as that public artifact through the current workflow.
- `npm run build:full` runs the core build, both bundles, and the complete Main Application suite.
- `npm run build:xlsx-sample` writes optional sample XLSX and Markdown output under `local-data/`.

テストsuiteは次の役割に固定する。追加した `*.test.js` は必ずいずれかへ明示的に分類し、`tests/mikuproject-test-suite-topology.test.js` が実行漏れと重複を検出する。

- `npm run test:fast`: 日常開発用。core utility / codec / workbook / core APIとlegacy CLI compatibility contractを検証する。時間のかかるCLI統合testとbrowser runtime testは含めない。
- `npm run test:full`: `fast` にCLI統合testとbrowser runtime contractを加えた完全回帰。
- `npm test` / `npm run test:all`: 現在は `full` と同じ全checked-in test fileを実行する、CI向けの安定した完全suite alias。

Web UI development, browser/UI tests, single-file HTML build, and offline browser verification are maintained in [miku-project-web](https://github.com/igapyon/miku-project-web).

## local-data

`local-data/` is a disposable workspace for generated samples and verification outputs. It is not tracked by Git and must be reproducible when needed.
