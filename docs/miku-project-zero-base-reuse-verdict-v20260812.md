---
title: miku-project v1 core再利用採否 v20260812
description: R1/C1のconformance fixtureを根拠に、現行core資産をv1実行経路、互換性証拠、後続deferへ採否するP4.7の判断記録。
topics:
  - miku-project
  - cli
  - migration
  - testing
category: decision-record
status: approved
audience:
  - maintainer
  - developer
  - agent
created: 2026-08-12
updated: 2026-08-13
sources:
  - type: local-file
    role: preliminary-inventory
    path: docs/miku-project-zero-base-reuse-inventory-v20260810.md
    label: ゼロベース再利用資産棚卸し v20260810
    checked: 2026-08-13
  - type: local-file
    role: implementation-plan
    path: docs/miku-project-zero-base-implementation-plan-v20260810.md
    label: ゼロベース新仕様適合計画 v20260810
    checked: 2026-08-13
---

# miku-project v1 core再利用採否 v20260812

## 結論

R1/C1のv1実行経路は、legacy `ProjectModel`、global core API loader、旧AI view schema、旧Patch schemaをsemantic dependencyとして採用しない。これらは現行CLIと将来の互換移行のための証拠として維持する。

一方、single-MJSとsource archiveを作る`build-cli-bundle.mjs`は、v1 module closureをlegacy helperと分離して内包できているため、配布の非semantic基盤として適応しつつ再利用する。ただし、manifest、asset digest、runtime bindingを満たすまで公開v1 runtimeとしては扱わない。workbook/XLSX、帳票、browser runtimeは選択scope外なのでdeferする。

これは旧資産の削除・改名・価値否定ではない。`ZB-P0.3`の「reuse候補」を、R1/C1に対して何を再利用できるかではなく、**どの契約責務を曖昧にせずに再利用できるか**へ具体化した判断である。

## 判断規則

- `v1実行経路へ採用` は、v1 semantic state、stable diagnostic、raw/provenance binding、human approval、artifact publicationの責務を損なわず、conformance fixtureで直接検証できるものだけを指す。
- `互換性の証拠として保持` は、現行機能と既存入力を回帰保守するが、v1 resultやsemantic stateの正本にはしないことを指す。
- `後続scopeまでdefer` は、v1 R1/C1が選択していない能力であり、不採用や削除を意味しない。
- P4.7はNode参照実装内の採否である。Javaの移植範囲、Skillsのruntime選択、旧APIの削除、wire identifierの改名は決めない。

## conformance上の根拠

| 根拠 | v1が固定すること | P4.7での意味 |
| --- | --- | --- |
| `S-V001` | canonical XMLをexact semantic goldenへdecodeし、raw digestとnormalizationなしを保持する | 「読める」だけでなく、入力bytesとv1 semantic stateの対応を固定する |
| `S-I012` | `percent_complete`が101の状態を`semantic.invalid`、rule IDとsemantic location付きで返す | messageだけのwarningや曖昧な診断をv1 validatorの代わりにしない |
| `S-I020` | actual/baselineなどscope外の外部dataを`semantic.unsupported`としてinvalidと区別する | 広いモデルへの取込みや暗黙の欠落を成功扱いしない |
| `CI-OVERVIEW-001` / `CI-CONTEXT-001` | `project_overview`と対象leafだけの`task_change_context`をsource digest付きで固定する | 汎用AI viewではなく、目的・scope・content bindingを持つProjectionを正本にする |
| `CP-CHANGE-001` とC1 apply / verification | strict request、semantic diff、human approval、committed three-member artifact setを束縛する | 広いPatch operationや暗黙writeをC1の代わりにしない |
| legacy XML round-trip / core API loader回帰 | 旧`ProjectModel`経路とNode側loaderが現在も動く | v1の正本性とは別に、現行互換性を退行させない |

### fixtureの系譜

`testdata/dependency.xml`はv1 fixtureを考える出発点として有用だが、`testdata/conformance/v1/fixtures/project/dependency-canonical.xml`のraw bytesとは一致しない。v1 fixtureではProject childの順序を固定し、dependency lagを`PT0H0M0S`から`0`へ表現し直し、`LagFormat`を明示している。

したがって、旧fixtureは入力領域の互換性証拠、v1 fixtureとgoldenはR1/C1の契約正本である。両者を同一のraw artifact、または片方だけで全契約を証明するfixtureとして扱わない。

## 採否表

| 現行資産 | P0での仮分類 | P4.7の採否 | v1での扱いと理由 | 根拠 | 再評価 |
| --- | --- | --- | --- | --- | --- |
| `types.js`、`msproject-codec.js`、`msproject-xml*.js`、`msproject-validate*.js`、`msproject-calendar.js` | reuse候補 | 互換性の証拠として保持 | v1実行経路には採用しない。旧codecは`ProjectModel`をimport/exportし、export時にnormalizationとdefault calendar補完を行う。旧validatorはwarning/message中心で、v1が必要とするraw provenance、stable rule ID、invalid/unsupported分離を正本として提供しない。v1は`cli-v1-xml-adapter.mjs`、`cli-v1-semantic-validator.mjs`、`cli-v1-xml-encoder.mjs`を使う | `S-V001`、`S-I012`、`S-I020`、legacy XML round-trip | format拡張はP4.8、互換移行はP7 |
| `msproject-ai-views.js`、`core-api-msproject-ai.js`、`ai-json-spec.js`、`ai-json-util.js` | reuse候補 | 互換性の証拠として保持 | 旧viewは広い`allow_patch_ops`を持つ。v1は`project_overview`と対象leafの`task_change_context`だけをpurpose/version/source digest/content binding付きで生成する。旧view schemaをv1 Projectionへ変換せず、設計知見として保つ | `CI-OVERVIEW-001`、`CI-CONTEXT-001`、`RB-012` | 新Projectionの追加はP4.8 |
| `project-patch-json*.js`、`core-api-ai-json-import.js`、`core-api-external-import.js` | reuse候補 | 互換性の証拠として保持 | 旧Patchはtask/resource/calendar/assignmentの追加・更新・削除・移動・link/unlinkを扱う。R1/C1では`set_task_percent_complete`だけをstrict request、base digest、leaf、expected current value、human approval、semantic diff、artifact bindingで許可する。旧Patchをpartial C1 implementationとして呼び出さない | `CP-CHANGE-001`、C1 apply / verification、`RB-001`〜`RB-008` | operation追加はP4.8、互換移行はP7 |
| `core-api.js`、`core-api-public.js`、`core-api-registry.js`、`core-api-msproject.js`、`scripts/lib/core-api-loader.mjs` | reuse候補 | 互換性の証拠として保持 | loaderはJSDOMとglobal APIを組み立てる現行CLIのbootstrapである。fixed v1 harnessはv1 module graphだけを使い、v1 serviceがlegacy loaderを暗黙に起動しない。現行CLIの保守とlegacy regressionのため残す | core API / loader回帰、R1/C1 harness、bundle source包含 | P4.9で公開runtimeの明示bindingを設計、P7でcompatibility判断 |
| `scripts/build-cli-bundle.mjs`、single-MJS、source archive | reuse候補 | v1実行経路へ採用（非semantic基盤、適応あり） | v1 module graphをlegacy helperとclosureで分離してbundle/source archiveへ含める仕組みは採用する。これはsemantic conversionやruntime選択を再利用する判断ではない。P4.9でv1 coreだけを識別するmanifest、asset path、SHA-256、capability / fixture version bindingを追加するまでpublic v1 commandはfail-closedを維持する | R1/C1 bundle/source archive integration、public fail-closed boundary | P4.9 / P4.10 |
| `scripts/cli-ai-workflow-example.mjs`、`scripts/cli-ai-stdio-example.mjs` | reuse候補 | 互換性の証拠として保持 | file / stdinによるAgent連携の経験は残すが、どちらもlegacy workbook JSON、旧`ai validate-patch`、旧`state apply-patch`を使う。v1 Skill/Agent workflowへ流用せず、v1のstructured result、human approval、committed artifact verificationを使うP6 workflowへ置換する | 例2本の実行成功、C1 / artifact conformance | P6（G5後） |
| `testdata/dependency.xml`、`hierarchy.xml`、`minimal.xml`、legacy XML / core API tests | reuse候補 | 互換性の証拠として保持 | v1 corpusの設計起点と回帰資産として残す。v1のraw input、semantic state、Projection、diff、resultの正本は`testdata/conformance/v1/`とそのgoldenである | fixture diff、legacy XML / loader回帰、R1/C1 conformance | P4.8で新scopeごとにfixtureを追加、P7で移行fixtureを追加 |
| `project-workbook-json*.js`、`core-api-workbook.js`、`project-xlsx*.js`、`excel-io*.js`、`core-api-workbook-xlsx.js` | reuse候補 / evidence | 後続scopeまでdefer | R1/C1の直接XML scenarioでは使わない。workbook JSONをv1 IRやCLI exchange formatと決めず、format選択時にloss / round-trip fixtureから再評価する | v1 format scope、現行workbook/XLSX regression | P4.8のformat選択 |
| `scripts/miku-project-cli.mjs`、legacy CLI router / diagnostics / I/O | rewrite候補 | 互換性の証拠として保持 | current command surfaceの回帰は維持するが、manifest未検証のv1 commandはlegacyへfall throughさせずfail-closedにする。v1のpublic entrypointはP4.9でruntime bindingを実装してから別途接続する | legacy CLI compatibility regression、R1/C1 public boundary | P4.9、P7 |
| report、SVG、Markdown、browser runtime | defer | 後続scopeまでdefer | 派生出力とWeb互換であり、v1 semantic / C1 publicationの実装を支配させない | P4.4〜P4.6の非目標、既存browser/report回帰 | P4.8またはWeb workstream |

## 明示的に行わないこと

- legacy coreをv1 moduleへimportして「既にある変換」を近道として使わない。
- legacy fixtureの成功を、v1 raw digest、semantic golden、structured result、artifact bindingの成功と読み替えない。
- P4.7のために旧command、`mikuproject`名称、wire identifier、browser runtimeを変更・削除しない。
- manifest未検証のpublic source CLI / development bundleをv1 runtimeとして公開しない。P4.9までは`runtime.capability-missing`でfail-closedのままとする。

## 実施記録

2026-08-13に、次の回帰を一つの対象群として実行した。

```text
./node_modules/.bin/vitest run \
  tests/mikuproject-cli-v1-xml-adapter.test.js \
  tests/mikuproject-cli-v1-inspect.test.js \
  tests/mikuproject-cli-v1-plan-change.test.js \
  tests/mikuproject-cli-v1-apply-preparation.test.js \
  tests/mikuproject-cli-v1-verify-artifact.test.js \
  tests/mikuproject-cli-v1-r1-integration.test.js \
  tests/mikuproject-msproject-xml-roundtrip.test.js \
  tests/mikuproject-core-api.test.js \
  tests/mikuproject-core-api-loader.test.js --reporter=dot
```

結果は9 test files、78 testsすべて成功した。前半6 fileがv1 adapter / Projection / C1 apply・verification・bundle integrationを、後半3 fileがlegacy XML / public core API / Node loaderの現行回帰を担う。この成功は表の採否を裏づけるものであり、legacy経路をv1のsemantic contractと同一視する根拠にはしない。

続けてlegacy CLIの`tests/mikuproject-cli-compatibility-contract.test.js`と`tests/mikuproject-cli.test.js`を実行し、2 test files、64 testsすべて成功した。さらに`node scripts/cli-ai-workflow-example.mjs`と`node scripts/cli-ai-stdio-example.mjs`はともに成功した。これらは旧workflowが保守可能であることの証拠であり、v1 workflowの適合証明ではない。

defer判断の根拠も最新化するため、`tests/mikuproject-project-workbook-json.test.js`と`tests/mikuproject-project-xlsx.test.js`を実行し、2 test files、33 testsすべて成功した。workbook/XLSXが現行経路で回帰していないことは確認したが、v1 semantic stateやexchange formatへ採用する判断はP4.8のformat選択まで保留する。

## 承認境界と次作業

この記録は2026-08-13に承認された。承認範囲はR1/C1における再利用採否だけであり、Javaの移植範囲、Skillsのruntime選択、旧APIの削除、旧名称の改名、public v1 runtimeの有効化は含まない。

次の`ZB-P4.8`では、承認済み`S-V002`を用いた階層C1 sliceを実装する。旧coreの機能量で選ばず、nested leafのR1 Projectionと既存C1 operationを、v1 fixture、loss policy、Projection、change、publicationの既存契約に照らしてmaterializeする。新format、new Projection purpose、new change operationは導入しない。詳細は[実施計画のP4.8](miku-project-zero-base-implementation-plan-v20260810.md#p48の実行計画)を正本とする。

P4.9はこの採否を根拠に、Node参照実装の実assetとmanifestを固定する。ただしP4.7だけではP4.9、P4.10、Gate G4、Java、Skillsへ進む承認にはならない。
