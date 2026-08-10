---
title: miku-project 現行 capability matrix v20260810
description: ゼロベース再設計のG0で参照する、現行Node CLIと周辺資産の能力・制約・再利用候補の棚卸し。
topics:
  - miku-project
  - cli
  - migration
  - specification
category: reference
status: current-state
audience:
  - maintainer
  - developer
  - agent
created: 2026-08-10
updated: 2026-08-11
---

# miku-project 現行 capability matrix v20260810

## 目的とauthority

この文書は、[ゼロベース再設計仕様](miku-project-zero-base-spec-v20260809.md) の設計判断と、[実施計画](miku-project-zero-base-implementation-plan-v20260810.md) の `ZB-P0.2` / `ZB-P0.3` に使う現状証拠である。現行実装の能力を正確に残すことを目的とし、v1の製品契約や正本形式を決める文書ではない。

| 文書・資産 | authority |
| --- | --- |
| `docs/miku-project-zero-base-spec-v20260809.md` | 目標とする製品契約の仕様ドラフト |
| `docs/miku-project-zero-base-implementation-plan-v20260810.md` | 実装順、判断ゲート、完了条件 |
| 本書 | 2026-08-10時点のNode CLI・tests・artifactの現状証拠 |
| `docs/spec.md`、AI JSON関連文書、現行コード | 旧設計と実装の詳細証拠。vNextの正解ではない |

調査対象は、このリポジトリのNode CLI、core API、tests、bundle、`package.json`である。Java CLIとAgent Skillsの実装状態は、この文書の対象外とする。

## 現行CLIの入口

現行CLIは `scripts/miku-project-cli.mjs` であり、`miku-project` と旧名 `mikuproject` の両方をbinとして公開している。`package.json` のversionは `1.0.0`、descriptionはWeb App中心の歴史を残している。通常buildはcore、browser runtime、CLI bundle、fast testを一括で実行する。

現行CLIの主な名詞空間は `ai`、`state`、`import`、`export`、`report` である。G3で承認された `inspect`、`validate`、`plan-change`、`apply-change`、`verify-artifact` とは異なる分類である。

## 操作 capability matrix

| 新仕様で必要な能力 | 現行の操作・資産 | 現状 | vNextでの扱い |
| --- | --- | --- | --- |
| 外部artifactの読取り | CLIでは `import xlsx` のみ。coreにはMS Project XMLの読取り資産がある | 部分対応。CLIの入口はXLSXに偏る | R1で外部入力例を決め、read/inspect契約として再設計 |
| purpose別`inspect` | `ai detect-kind`、`state summarize` はある | 部分対応。JSON種別・概要に限定され、外部artifact全般の能力・損失・diagnosticsを返さない | G3のread-only command。Projection purposeとscopeを契約化 |
| whole-project `validate` | `ai validate-patch` とcore validatorがある | 部分対応。CLIではPatch検査中心で、プロジェクト全体の検証入口がない | R1/C1共通のread-only operationとして定義 |
| AI向けProjection | `ai export project-overview / task-edit / phase-detail / bundle` | 実装済み。入力は現行workbook JSONで、局所Projectionもある | 再利用候補。purpose、範囲、schema、返却契約をG2で再評価 |
| 変更要求の検証 | `ai validate-patch` | 実装済み。dry-run applyに基づく | C1の入力契約、precondition、許可operation、diagnosticsへ発展させる |
| `plan-change`内のsemantic diff | `state diff --before workbook.json --after workbook.json` | 部分対応。二つのworkbook JSONが前提で、現行import規則によるsummaryである | C1のrequestとoutput planに束縛したdiffとして契約・fixture化 |
| `apply-change` | `state apply-patch --state ... --in ...` | 部分対応。次のworkbook JSONを出力するがhuman gateとpublication契約はない | 再利用候補。ただしapproval、pre/post validate、commit-marker publicationを新契約で定義 |
| 新規draft生成 | `state from-draft`、`project_draft_view` import | 実装済み | v1のR1/C1には含めず、後続scenario候補として保持 |
| 交換形式への変換 | `export workbook-json / xml / xlsx`、`import xlsx` | 実装済み。ただしCLIのdirect XML importはない | 形式・損失・roundtrip規則を決めた後に選別 |
| 帳票・可視化 | WBS XLSX、SVG、Markdown、Mermaid、ZIP bundle | 実装済み | v1中核から外す。derived output候補としてdefer |
| stable structured result | JSON本文と `--diagnostics json` はある | 部分対応。全コマンド共通のresult envelopeやnext action契約はない | G3で`miku_project_cli_result/v1`へ統一。現行実装はまだ未適合 |
| stable diagnostics / exit | diagnostics version、0/1/2のexit code、明示error codeを一部持つ | 部分対応。messageからcodeを推定する経路が残り、location・loss・normalization・retryabilityが一様でない | G3でclosed code catalogとexit `0/1/2/3`を定義。現行実装はまだ未適合 |
| 明示I/Oと非対話 | `--in`、`--out`、stdin/stdout、binaryのBase64経路がある | 概ね対応。ただし一部のtext inputは暗黙stdinを許す | vNextではartifact role、encoding、出力先、許可条件を明示 |
| safe output | `--out`で指定した既存ファイルを無確認で上書きし、`writeFileSync`で直接書く | 未対応。既定のoverwrite拒否やpublication状態がない | 新規directoryのexclusive create、commit marker、incomplete/corrupt判定を必須化 |
| hidden stateなし | ファイルとstdin/stdoutを中心に処理する | 概ね対応 | 維持。artifact間の遷移を明文化し、Agent会話履歴へ依存しない |
| deterministic artifact build | single-file CLI bundleとsource archiveを生成する | 部分対応。bundleはあるが、CLI artifactをSkillsが検証するmanifestはない | runtime manifest、asset role、SHA-256、provenanceを追加 |
| Node/Java conformance | Node側のunit/E2E testsはある | 未対応。共通fixture/golden/capability matrixがない | Node契約固定後にJavaを適合させる |
| Agent Skills integration | 現行CLIはprojection、Patch、diagnosticsを持つ | 部分対応。新仕様に必要なruntime pinning/human gateはNode CLIだけでは満たさない | G6でCLI-only workflow、manifest/digest検証を実装 |

## 現行形式とartifactの役割

| 形式・artifact | 現行の主な役割 | 現状の注意点 | vNextでの暫定位置 |
| --- | --- | --- | --- |
| MS Project XML | coreで読取り・書出し、CLIではworkbook JSONからのexport | 現行文書ではハブ扱いだが、CLI direct importはない | R1の外部fixture候補。正本かどうかは未決定 |
| `mikuproject_workbook_json` | 現行CLIの主要state入出力 | workbook構造と編集可能列の制約を含み、AI編集JSONと別物 | 現行証拠・互換対象。新IRの採用可否は未決定 |
| Projection JSON / `.editjson` | AIまたは人へ局所的な理解・編集範囲を渡す | `project_overview_view`等のschemaが現行設計へ結び付く | R1/C1の再利用候補。purposeとschemaを再定義 |
| Patch JSON | 現行projectへのoperation-based変更要求 | 現行operationの範囲・precondition・loss規則は新契約で未確定 | C1の再利用候補。全量置換の代替として検討 |
| XLSX | import/exportとWBS帳票 | binary I/Oと限定importを持つ。自由編集roundtripは未証明 | v1の必須から外し、形式gateで再評価 |
| SVG / Markdown / Mermaid / report ZIP | 派生出力 | 表示・見た目に関する機能が多い | v1中核から外し、derived outputとしてdefer |

## 現行CLIの安全性・Agent利用性

| 観点 | 現行状態 | 新仕様との差 |
| --- | --- | --- |
| 読み取りと意味変更の区別 | command namespaceには混在があるが、`state apply-patch`は識別できる | `read-only / artifact生成 / 意味変更`の共通分類がない |
| human gate | CLIに承認段階はない | C1では`diff`後・`apply`前の明示gateが必要 |
| output overwrite | `--out`は既存ファイルを無確認で上書きする | 既定拒否と明示許可が必要 |
| publication | 直接`writeFileSync`する | commit markerを境界にcommitted artifactだけを成功として公開する契約が必要 |
| structured branching | JSON diagnosticsを選択できる | command横断のstatus、severity、retryability、next actionが不足 |
| Agent会話への依存 | CLI自体はローカルartifactを入力にする | artifact roleとschema versionをさらに明示し、会話履歴依存を排除する |

## tests・配布artifactの現状

2026-08-10に `npm test` を実行し、現行runnerが対象にする10 test files・183 testsは成功した。これは現行資産のベースラインであり、新仕様のcontract suiteが十分であることを示すものではない。

| 項目 | 現状 | vNextへの入力 |
| --- | --- | --- |
| CLI tests | `tests/mikuproject-cli.test.js` にCLI、bundle、出力の回帰testsがある | 現行互換を固定する候補 |
| core tests | `tests/mikuproject-core-api.test.js` とloader testがある | semantic fixtureへの再編候補 |
| test runner | `fast`、`full`、`all` が同じtest listを実行し、core API testsを明示listに含めない | G4より前にtest topologyを修正する |
| bundle | `bundle/miku-project.mjs` とsource archiveを生成できる | clean bundle smokeを維持し、CLI manifestを追加 |
| browser runtime | default buildに含まれる | Web互換として維持可否を移行gateで判断。v1 CLI契約を支配させない |

## 初期分類: reuse / evidence / rewrite / defer / drop

これはv1候補R1/C1に限った初期分類である。すべての現行moduleとfixtureの最終処分ではない。

| 分類 | 対象 | 理由 |
| --- | --- | --- |
| reuse候補 | MS Project XML codec、ProjectModel validator、AI Projection exporter、Patch importer/applicator、workbook JSON codec、CLI bundle smoke | R1/C1を実証する技術的根拠がある。ただし新契約への適合が条件 |
| evidence | 現行CLI command、AI JSON仕様、XLSX import/export、reports、Web runtime | 利用経験とformat制約を示す。新v1の正本や語彙を決めない |
| rewrite候補 | command dispatchとI/O、result/diagnostics contract、安全出力、semantic diff、test runner | 新仕様の副作用分類、Agent分岐、commit-marker publication、共通conformanceを満たさない |
| defer | reports、SVG、Markdown、Mermaid、XLSXの自由編集・見た目、Web/MCP | R1/C1を成立させる前提ではない |
| drop候補 | なし | 互換性と移行の判断はG7まで保留する |

## 証拠

- CLI command、stdin/stdout、diagnostics、上書き挙動、exit code: `scripts/miku-project-cli.mjs` の `writeHelp` と `writeOutput`
- command dispatch、Projection、Patch validation / apply、state diff: `scripts/miku-project-cli.mjs` の `runCommand`
- 現行AI JSONの役割とPatch operation: `docs/miku-project-ai-json-spec.md`
- 現行テスト構成: `scripts/run-tests.mjs`、`tests/mikuproject-cli.test.js`
- package version、bin、default build: `package.json`

このmatrixは、G0でR1/C1を承認するための現状情報である。G0通過後、`G1`〜`G3`で意味、artifact、形式、CLI契約を定義する際に更新する。
