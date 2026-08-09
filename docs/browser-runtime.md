---
title: miku-project browser runtime contract
description: Public contract, build, verification, and downstream intake rules for the browser-compatible runtime bundle.
topics:
  - miku-project
  - browser-runtime
  - web-app
  - release
category: reference
status: active
audience:
  - developer
  - downstream-maintainer
  - agent
created: 2026-08-09
updated: 2026-08-09
---

# Browser Runtime Contract

`miku-project-runtime-<version>.mjs` は、`miku-project-web` などの browser downstream が build 時に取り込む importable runtime bundle である。Node.js CLI bundle `miku-project-<version>.mjs` とは別成果物として扱う。

## Public exports

```js
import loadMikuProjectRuntime, {
  version,
  embeddedCorePaths,
  loadMikuProjectRuntime
} from "./miku-project-runtime-<version>.mjs";

const api = loadMikuProjectRuntime({ expectedVersion: version });
```

- `version`: build 元の `package.json` version
- `embeddedCorePaths`: runtime に含まれる core module path の read-only 一覧
- `loadMikuProjectRuntime(options)`: core を初期化し、`globalThis.__mikuProjectCoreApi` と同じ API object を返す
- default export: `loadMikuProjectRuntime`

`options.expectedVersion` を指定すると、runtime version が一致しない場合に初期化前に失敗する。既存ページですでに別の core API が初期化されている場合は既定でエラーにする。移行中のページで既存 API を意図的に再利用するときだけ `options.reuseExisting: true` を指定する。同じ runtime loader の2回目以降の呼び出しは、最初に初期化した API を返す。

互換 alias `globalThis.__mikuprojectCoreApi` は現行移行期間中も canonical API と同じ object を指す。

## Boundary

browser runtime は次を含めない。

- `node:` module import または参照
- `process` 参照
- CLI entrypoint と CLI 自動実行
- UI event handler、download adapter、画面初期化

Office ZIP helper の同期 Node zlib fallback は browser runtime 生成時に無効化する。browser runtime では stored ZIP entry、`DecompressionStream` を用いる async read、または明示的に注入した async inflater を使用する。これは Node.js CLI bundle の動作を変更しない。

## Build and verification

```bash
npm run build:web
npm run build:browser-runtime
node scripts/smoke-browser-runtime.mjs \
  bundle/miku-project-runtime.mjs \
  --expected-version "$(node -p "require('./package.json').version")"
```

既定出力は `bundle/miku-project-runtime.mjs` である。`--out <path>` で出力先を変更できる。smoke は禁止された Node/CLI 参照、公開 export、version、core API 初期化を検証する。

## Release and downstream intake

Release workflow は次を別 asset として公開する。

- `miku-project-<release-version>.mjs`: Node.js CLI bundle
- `miku-project-runtime-<release-version>.mjs`: browser runtime bundle
- `miku-project-runtime-<release-version>.json`: release tag、package version、asset 名、runtime SHA-256 を結ぶ machine-readable lock
- `miku-project-sources-<release-version>.tgz`: source archive
- `miku-project-SHA256SUMS-<release-version>.txt`: 上記4成果物の SHA-256

`miku-project-web` は `miku-project-runtime-<release-version>.json` の内容を repository-local lock として固定する。build 時に runtime asset を取得し、lock の tag、asset 名、SHA-256 を検証してから Web 配布物へ組み込む。browser 実行時には GitHub やネットワークから runtime を取得しない。

Main Application と Web App の分離完了条件は、Web repository 単独で single-file Web App の build と browser/UI test を再現でき、ネットワーク遮断状態でも主要導線が動作することである。それを確認する前に Main Application 側の Web 資産を削除しない。

移動対象、残置対象、downstream bootstrap の順序は [Web Separation Inventory](web-separation-inventory.md) を参照する。
