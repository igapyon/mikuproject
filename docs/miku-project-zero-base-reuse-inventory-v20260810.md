---
title: miku-project ゼロベース再利用資産棚卸し v20260810
description: R1/C1を対象に、現行module、fixture、test、artifactをreuse/evidence/rewrite/defer/dropへ仮分類するG0資料。
topics:
  - miku-project
  - migration
  - cli
  - testing
category: reference
status: draft
audience:
  - maintainer
  - developer
created: 2026-08-10
updated: 2026-08-10
---

# miku-project ゼロベース再利用資産棚卸し v20260810

## 位置づけ

この文書は、[v1利用シナリオ案](miku-project-zero-base-scenarios-v1.md) のR1/C1を実証するために、現行資産を `reuse / evidence / rewrite / defer / drop` へ仮分類する `ZB-P0.3` の成果物である。ここでの分類は実装順を決めるためのものであり、旧API・形式・名称を直ちに削除または変更する決定ではない。

`reuse` は「現在のAPIまたは形式をそのまま正本にする」意味ではない。新しいsemantic、artifact、CLI契約に適合できると確認できた場合だけ利用する。

## moduleの仮分類

| 分類 | 現行module群 | R1/C1との関係 | 扱い |
| --- | --- | --- | --- |
| reuse候補 | `types.js`、`msproject-codec.js`、`msproject-xml*.js`、`msproject-validate*.js`、`msproject-calendar.js` | XML fixtureを意味表現へ読み、hierarchy、dependency、日付などを検証する根拠 | G1で不変条件と保持範囲を照合し、必要な部分だけ使う |
| reuse候補 | `msproject-ai-views.js`、`core-api-msproject-ai.js`、`ai-json-spec.js`、`ai-json-util.js` | R1のProjection設計とAI向けJSONの経験を提供する | 現行view schemaを正本にせず、purpose・scope・versionをG2で再定義する |
| reuse候補 | `project-patch-json*.js`、`core-api-ai-json-import.js`、`core-api-external-import.js` | C1のoperation-based change、Patch検証・適用の経験を提供する | G2でselector、precondition、許可operation、loss/provenanceを再定義する |
| reuse候補 | `project-workbook-json*.js`、`core-api-workbook.js` | 現行状態のimport/exportとPatch結果の観測に利用できる | workbook JSONを新IRや正本と決めない。互換adapter候補として扱う |
| reuse候補 | `core-api.js`、`core-api-public.js`、`core-api-registry.js`、`scripts/lib/core-api-loader.mjs` | CLIからcoreを明示的に読み、bundleでも動かす入口 | 新CLIの責務境界に合わせてAPIを分離・整理する |
| evidence | `project-xlsx*.js`、`excel-io*.js`、`core-api-workbook-xlsx.js` | XLSX import/exportの制約と回帰知見 | v1 R1/C1の必須範囲外。G2/G4で選択するまで変更しない |
| defer | `wbs-xlsx*.js`、`wbs-svg*.js`、`wbs-markdown.js`、`msproject-mermaid.js`、report API群 | 帳票、可視化、派生成果物 | v1中核を実証した後のderived outputとして扱う |
| defer | `main-util.js`、browser runtime build/manifest/smoke周辺 | Web App互換とbrowser runtime | Web互換を維持できるが、新CLI契約を支配させない |
| rewrite候補 | `scripts/miku-project-cli.mjs` のcommand dispatch、I/O、diagnostics、output write | 現行の `ai/state/import/export/report` 分類、統一されないresult、無確認上書きを持つ | G3で新CLI契約を定義後、現行動作を回帰固定してから責務別に置換する |
| rewrite候補 | `scripts/run-tests.mjs` | `fast/full/all` が同じlistで、core API testsを明示実行しない | G4開始前にtest topologyとcontract suiteを整理する |
| drop候補 | なし | 互換性と移行は未決定 | G7まで削除・改名を行わない |

## fixtureの仮分類

| 分類 | asset | R1/C1で使う理由 | 次の作業 |
| --- | --- | --- | --- |
| reuse候補 | `testdata/dependency.xml` | task、dependency、resource、assignmentを持ち、R1の観測とC1の局所変更を同時に試せる | G3のconformance corpusへ移植または複製し、期待semantic resultを追加する |
| reuse候補 | `testdata/hierarchy.xml` | summary taskとchild taskを持つ | hierarchy・Projection scope・task identityのfixture候補 |
| reuse候補 | `testdata/minimal.xml` | 最小入力の読み書き境界を示す | valid/invalid/boundary corpusの起点 |
| evidence | `testdata/workbook-import-sample.json` | 現行workbook JSON importの実例 | workbook JSON互換adapterの回帰候補 |
| rewrite候補 | Node/Java/Skills共通のconformance corpus | 現在はruntime横断のgolden artifactがない | G3で `testdata/conformance/` を新設し、result/diagnostics/semantic diffを固定する |

## testの仮分類

| 分類 | test | 役割 | 次の扱い |
| --- | --- | --- | --- |
| reuse候補 | `tests/mikuproject-msproject-xml-roundtrip.test.js` | XML、hierarchy、dependency、calendar等の意味資産を検証 | G1/G2のsemantic fixtureへ分解・接続する |
| reuse候補 | `tests/mikuproject-project-workbook-json.test.js`、`tests/mikuproject-project-xlsx.test.js` | workbook/XLSXの現行互換を検証 | v1では互換回帰の証拠として維持する |
| reuse候補 | `tests/mikuproject-cli.test.js` | CLI、diagnostics、bundle、file/stdio workflowを検証 | 旧CLIのfreezeと新CLI contract testsの出発点にする |
| reuse候補 | `tests/mikuproject-core-api.test.js`、`tests/mikuproject-core-api-loader.test.js` | core APIの公開面とloaderを検証 | test runnerへ明示的に編入し、API境界変更時に実行する |
| defer | `tests/mikuproject-wbs-xlsx.test.js`、`tests/mikuproject-wbs-markdown.test.js` | 派生成果物を検証 | derived outputのworkstreamまで維持する |
| defer | `tests/miku-project-browser-runtime.test.js` | browser runtime互換を検証 | Web移行判断まで維持する |
| rewrite候補 | Node/Java/Skills共通conformance tests | 現在は存在しない | G3でcontract fixtureとともに作る |

## build・配布artifactの仮分類

| 分類 | asset | 現状 | 次の扱い |
| --- | --- | --- | --- |
| reuse候補 | `scripts/build-cli-bundle.mjs`、`bundle/miku-project.mjs`、`bundle/miku-project-sources.tgz` | single-file CLIとsource archiveを生成し、isolated bundle testがある | 新CLIでもisolated bundle smokeを維持する |
| rewrite候補 | CLI artifactのmachine-readable manifest | CLI bundleにはSkillsが固定asset/digestを選べるmanifestがない | G3でproduct contract、runtime、fixture suite、asset、SHA-256、capabilityを記録する |
| evidence / defer | `scripts/build-browser-runtime.mjs`、`bundle/miku-project-runtime.mjs`、browser manifest | browser runtimeのversion/digest検証はある | Web互換資産として維持し、v1 CLI contractとは独立させる |
| reuse候補 | `scripts/cli-ai-workflow-example.mjs`、`scripts/cli-ai-stdio-example.mjs` | 現行のfile/stdio Agent連携例 | 新しいR1/C1 workflowの安全性testへ置換・発展させる |

## 現時点で着手しないもの

- 現行 `mikuproject` 名、bin alias、wire identifierの改名・削除
- browser runtime、Web UI、MCPの削除または新規設計
- XLSX、帳票、SVG、Markdownの再実装・見た目調整
- 実績、Earned Value、baseline、timephased data、extended attributeの拡張
- Java implementationのmoving source追随

これらはG0でR1/C1を承認し、G1〜G3で製品契約が固定された後に、計画で定めた順序に従って扱う。
