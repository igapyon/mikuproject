---
title: miku-project Web separation inventory
description: Ownership inventory and migration gates for separating the historical combined Web surface into miku-project-web.
topics:
  - miku-project
  - web-app
  - migration
category: workflow
status: draft
audience:
  - maintainer
  - developer
  - agent
created: 2026-08-09
updated: 2026-08-09
---

# Web Separation Inventory

この文書は Issue [#124](https://github.com/igapyon/miku-project/issues/124) の移動対象と残置対象を固定する。実際の削除は `miku-project-web` の独立 build、browser/UI test、オフライン動作が成立した後に行う。

## Ownership boundary

| Owner | Current paths and contracts |
| --- | --- |
| Main Application | `src/ts` / `src/js` の core modules、`globalThis.__mikuProjectCoreApi`、CLI、Node CLI bundle、browser runtime bundle、source archive、domain fixtures、core/CLI tests |
| Web App | HTML source and generated HTML、browser adapter、UI state/event/download modules、`src/css/`、`lht-cmn/`、browser/UI tests、single-file Web App build and publication |
| Transitional | `scripts/build-project.mjs`、`scripts/lib/single-html.mjs`、現行 Web screenshots/docs。Web repository の再現性が確認されるまで Main Application 側に残す |

core と Web surface の module path は [runtime-module-paths.mjs](../scripts/lib/runtime-module-paths.mjs) の `CORE_API_MODULE_RELATIVE_PATHS` と `WEB_SURFACE_MODULE_RELATIVE_PATHS` を判定元にする。browser runtime build と smoke は Web surface module の混入を拒否する。

## Main Application に残すもの

- `scripts/miku-project-cli.mjs` と Node adapter
- `scripts/build-cli-bundle.mjs`
- `scripts/build-browser-runtime.mjs`
- `scripts/smoke-browser-runtime.mjs`
- `scripts/lib/core-api-loader.mjs`
- `scripts/lib/runtime-module-paths.mjs`
- `src/vendor/` と core module の `src/ts/` / tracked `src/js/`
- `tests/miku-project-browser-runtime.test.js`、core API、codec、workbook、report、CLI の tests
- [browser-runtime.md](browser-runtime.md) の公開契約
- `.github/workflows/release-runtime-bundles.yml` の runtime Release 導線

`main-util` は歴史的な名前だが core runtime に含まれる。ファイル名だけを理由に Web App へ移さない。

## miku-project-web へ移すもの

- `miku-project-src.html` と Web App の page markup
- `index-src.html` と Web landing page
- `miku-project.html`、`index.html` の generated Web artifacts
- `mikuproject.html` の compatibility redirect と、その終了条件
- `src/css/`
- `lht-cmn/`
- `WEB_SURFACE_MODULE_RELATIVE_PATHS` に列挙した UI state/event/render/download modules と対応する `src/ts/`
- `tests/helpers/main-*` の UI harness
- `tests/mikuproject-main-*` のうち `main-util` 以外の browser/UI tests
- `tests/mikuproject-single-html.test.js`
- `lht-cmn/components.test.js`
- Web App の screenshots、公開導線、GitHub Pages または Web Release 設定

移動時に test 名だけで機械的に所有者を決めない。core API を直接検証している assertion は Main Application の core test に残し、DOM wiring、画面状態、download、single HTML を検証する assertion を Web repository へ移す。

## Downstream bootstrap order

1. 人手で `igapyon/miku-project-web` repository を作成する。
2. Main Application Release の `miku-project-runtime-<release-version>.json` を repository-local lock として固定する。
3. build 時に runtime を取得し、SHA-256 が一致しない場合は Web build を停止する。
4. Web bootstrap は `loadMikuProjectRuntime({ expectedVersion })` を完了してから UI modules を初期化する。
5. runtime と Web assets を single-file 配布物へ build 時に内包する。browser 実行時の GitHub/runtime download は禁止する。
6. repository 単独で browser/UI test と single-file build を実行する。
7. ネットワーク遮断状態で input、overview、output、主要 import/export の smoke を確認する。
8. 以上が成立してから Main Application 側 Web 専用 path を削除し、core/CLI/browser runtime を再検証する。

## Required evidence before Main Application cleanup

- Web repository の固定 runtime tag / asset / SHA-256 設定
- hash mismatch を拒否する test
- runtime loader が UI 初期化より先に完了する test
- Web repository 単独の build result
- browser/UI suite result
- single-file Web App が追加 network request なしで動く証拠
- Main Application cleanup 後の `npm run build:browser-runtime`、CLI bundle smoke、core/CLI full tests

この証拠が不足している間は、combined layout が残っていても分離未完了として扱う。

## 2026-08-09 checkpoint

`igapyon/miku-project-web` の local checkout に runtime lock、SHA-256 検証、runtime-first bootstrap、single-file build、移行 browser/UI suite、offline smoke、CI と Web Release workflow を構築した。tests は core source/shim を Web source tree に置かず、検証済み runtime を直接起動する。local `miku-project-runtime.mjs` を lock と同じ SHA-256 で注入した通常 checkout と clean copy の検証では、18 test files / 135 tests と、Input / Overview / Output / 15 output entries の offline smoke が成功した。

未完了 gate は `miku-project` `v0.13.0` Release の公開、clean clone の通常 runtime 取得、CI、実ブラウザ smoke である。この gate が閉じるまで Main Application 側の Web 専用 path は削除しない。
