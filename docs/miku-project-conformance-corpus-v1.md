---
title: miku-project conformance corpus v1
description: Gate G3で承認された、Node/Java共通fixture、golden、比較方式、unknown outcome回復case。
topics:
  - miku-project
  - cli
  - conformance
  - testing
  - specification
category: specification
status: approved
audience:
  - maintainer
  - developer
  - agent
created: 2026-08-10
updated: 2026-08-11
sources:
  - type: local-file
    role: primary
    path: docs/miku-project-semantic-fixture-catalog-v1.md
    label: Gate G1承認済みsemantic fixture catalog
    checked: 2026-08-10
  - type: local-file
    role: primary
    path: docs/miku-project-cli-contract-v1.md
    label: Gate G3承認済みCLI契約
    checked: 2026-08-11
  - type: local-file
    role: primary
    path: docs/miku-project-cli-result-contract-v1.md
    label: Gate G3承認済みresult/diagnostics契約
    checked: 2026-08-11
  - type: local-file
    role: primary
    path: docs/miku-project-runtime-capability-contract-v1.md
    label: Gate G3承認済みruntime capability契約
    checked: 2026-08-11
---

# miku-project conformance corpus v1

## 文書の位置づけ

これは`ZB-P3.9`の成果物であり、Node参照実装とJava適合runtimeが共有するfixture、golden、case ID、比較方式を定義する。checked-in corpusは [`testdata/conformance/v1/`](../testdata/conformance/v1/) に置く。

P3ではtest oracleとseed fixtureを固定する。Node用runnerはP4、Java用runnerはP5で実装する。P3時点でruntime未実装のcaseを「pass」とは記録しない。P4/P5は同じcaseをruntime固有testdataへ複製せず、このdirectoryを直接利用する。

## authority

conformanceの正本は次の組合せである。

1. 承認済みsemantic / format / change / CLI / result / capability契約
2. CLI result、diagnostic、semantic / exchange artifact、runtime manifestのmachine-readable JSON Schema
3. この文書の比較規則
4. `suite-index.json`のcaseとchecked-in fixture / golden

Nodeのlive出力をJavaの唯一のoracleにしない。goldenと契約がNodeの挙動に矛盾する場合、Node defectまたは契約不整合として解決する。goldenを実装出力で無条件に更新してはならない。

## directory contract

```text
testdata/conformance/v1/
├── README.md
├── contract-cases.json
├── suite-index.json
├── fixtures/
│   ├── project/
│   │   ├── dependency-canonical.xml
│   │   ├── dependency-percent-101.xml
│   │   └── dependency-unsupported-actual.xml
│   └── change/
│       └── set-task-2-percent-0-to-50.template.json
└── golden/
    └── semantic/
        ├── dependency.state.json
        └── dependency-percent-50.state.json
```

- pathは`suite-index.json`からの相対pathとし、`..`、absolute path、symlinkを許可しない。
- fixture/goldenはUTF-8 BOMなし、LF、末尾LF一件とする。`COMMITTED`だけは0 byteである。
- `.template.json`は通常のJSONとしてparseでき、文字列値`${NAME}`だけをharnessが置換する。key、数値、boolean、構造を置換しない。
- runtimeがfixture/goldenを変更してはならない。runnerはcaseごとに新しいtemporary directoryへcopyし、終了後の差分検査で原本不変を確認する。
- temporary pathとruntime artifactはrepository外に置けるようにし、source checkoutのmodule解決へ依存させない。

## seed fixtureの意味

`dependency-canonical.xml`は`S-V001`のG3 canonical variantであり、正式な`LinkLag=0 / LagFormat=3`とcanonical child順を使う。現行`testdata/dependency.xml`はlegacy zero-lag normalizationを検証する別caseとして残し、新しいcanonical write goldenと同一視しない。

`dependency.state.json`は入力のruntime非依存semantic state、`dependency-percent-50.state.json`はC1後の期待stateである。stateのscalar JSON表現は次で固定する。

| semantic type | fixture/golden JSON表現 |
| --- | --- |
| identity token、text | JSON string。Unicode normalizationしない |
| local civil datetime | `YYYY-MM-DDTHH:mm:ss` string |
| working duration | total durationの`PT{H}H{M}M{S}S` string |
| units | 不要な先頭・末尾zeroを持たないbase-10 string |
| percent | JSON integer |
| boolean | JSON boolean |
| root parent | JSON null |

semantic stateのcanonical digestは、G2のcollection canonicalizationを先に行い、object keyをUnicode code point昇順に再帰sortし、insignificant whitespaceと末尾LFを含まないUTF-8 JSON byte列へSHA-256を適用する。integerはbase-10、先頭zeroと`-0`なしとする。v1 semantic stateにinteger以外のJSON numberを入れない。この規則はfixture fileのpretty-print byte列ではなく、parse後の値へ適用する。

canonical JSON stringはUnicode scalar sequenceをnormalizationせず保持する。quotation markとreverse solidusは`\"`と`\\`、U+0008/U+0009/U+000A/U+000C/U+000Dはそれぞれ`\b / \t / \n / \f / \r`、その他のU+0000〜U+001Fはlowercase hexの`\u00xx`でescapeする。solidusとそれ以外のUnicode scalarはescapeしない。unpaired surrogateはinvalidである。これによりNode/Javaのserializer既定値へdigestを依存させない。

change request templateの`${BASE_STATE_DIGEST}`は、上記規則で`dependency.state.json`から計算した64桁lowercase hexへ置換する。置換後のartifact digestも同じcanonical JSON規則で計算する。runtime family/version/pathのような環境値をsemantic state digestへ含めない。

## comparison mode

| mode | 比較対象 | 規則 |
| --- | --- | --- |
| `schema` | CLI result / diagnostics | 対応JSON Schemaに適合し、status、exit、code、rule ID、side effect、`next_action`がcase期待値と一致する。`message`は比較しない |
| `exact-json` | Projection、semantic diffなどruntime非依存JSON | object key順とformattingを無視し、全valueとarray順を一致させる。非task collectionは比較前に契約どおりsortする |
| `cross-artifact-binding` | state / Projection / request / diff / plan / approval / provenance / result | `RB-001`〜`RB-012`のcanonical digest、runtime、I/O/effect/path、Projection content、status/next action bindingを検査する |
| `semantic-state` | input/output project | runtime内IRを直接公開させず、test-only adapterまたはoutput XML再decode結果をgolden stateとsemantic比較する |
| `byte-same-runtime` | JSON result、canonical XML、provenance | 同じruntime artifact、同じ入力、同じoption、同じcanonical path tokenで二回実行したbyte列が一致する |
| `semantic-cross-runtime` | Node/JavaのXML、result data | runtime/path/digest bindingを除く契約fieldとsemantic stateが一致する。Node/Javaのbyte一致は契約で明示したfile以外に要求しない |
| `artifact-topology` | C1 destination | member名、type、symlink拒否、marker size、schema、digest、publication stateを比較する |
| `runtime-integrity` | runtime manifest / executable / sources | manifest schema、既知capability、固定path、size/digest、project未読、destination未作成を比較する |

`runtime.family`、runtime version、artifact/manifest digest、capability profile、fixture suite version、canonical absolute temporary pathはcaseごとに期待値を注入して比較し、無視しない。Node/Java間でruntime固有値が違うことを許すが、それぞれの[runtime manifest contract v1](miku-project-runtime-manifest-contract-v1.md)、output plan、approval、provenanceとのbinding一致を要求する。

artifact生成caseの`byte-same-runtime`は、runner自身が所有する明示的なtest directoryをsnapshot後に除去し、同じabsolute pathでfreshな親directoryを再構成して二回目を実行する。CLIへ既存destinationの上書きやcleanupをさせない。read-only caseは入力とresult pathを変えずに反復する。

## suite case

[`suite-index.json`](../testdata/conformance/v1/suite-index.json)は21件のworkflow / harness caseを固定する。schemaとcross-artifact bindingを直接検証する31件は[`contract-cases.json`](../testdata/conformance/v1/contract-cases.json)で別に固定し、workflow件数へ重複加算しない。

各caseの`expected_next_action`は[human gate and next action contract v1](miku-project-human-gate-and-next-action-contract-v1.md)に従う。runnerはactionだけでなく`command / source_retryability`も完全一致させ、複数diagnostic caseでは定義済み優先順から集約したことを検証する。

| case ID | command / flow | seedまたはsetup | 主な期待 |
| --- | --- | --- | --- |
| `CV-VALID-001` | `validate` | canonical S-V001 XML | success、診断0、semantic stateはbase golden |
| `CV-INVALID-001` | `validate` | percent 101 XML | rejected、`semantic.invalid`、rule `S-I012` |
| `CV-UNSUPPORTED-001` | `validate` | ActualStartを含むXML | rejected、`semantic.unsupported`、rule `S-I020` |
| `CI-OVERVIEW-001` | `inspect project_overview` | canonical S-V001 XML | success、runtime非依存Projectionをexact-json比較 |
| `CI-CONTEXT-001` | `inspect task_change_context --task-uid 2` | canonical S-V001 XML | success、target/context/許可operationをexact-json比較 |
| `CP-CHANGE-001` | `plan-change` | canonical XML + materialized request + fresh destination | success、`0 → 50`だけのdiff、loss/unsupportedなし、destination未作成 |
| `CA-CHANGE-001` | `apply-change` | CP result + explicit test approval | committed三member、input不変、outputはafter goldenとsemantic equivalent |
| `CA-DEST-EXISTS-001` | `apply-change` | approval後にdestinationを競合生成 | rejected、`publication.reservation-conflict`、既存entry不変 |
| `CA-BINDING-001` | `apply-change` | request/plan/approvalのdigest一つを変更 | rejected、`change.binding-mismatch`、destinationなし |
| `CVF-ABSENT-001` | `verify-artifact` | path不存在 | rejected、absent、`publication.artifact-absent` |
| `CVF-INCOMPLETE-001` | `verify-artifact` | markerなしdirectory | rejected、incomplete、`publication.artifact-incomplete`、非変更 |
| `CVF-CORRUPT-001` | `verify-artifact` | markerありでmember/digest不一致 | rejected、corrupt、`publication.artifact-corrupt`、非変更 |
| `CVF-COMMITTED-001` | `verify-artifact` | CAの成功destination | success、committed、plan/runtime binding一致 |
| `CVF-EXPECTED-PLAN-MISMATCH-001` | `verify-artifact` | committed set + 別plan | rejected、artifact非変更、`publication.expected-plan-mismatch` |
| `CU-UNKNOWN-OUTCOME-001` | apply → result delivery遮断 → verify | `COMMITTED`作成後・success result受信前にharnessがtransportだけを遮断 | applyを再試行せずverifyでcommittedと前回bindingを回収 |
| `CU-USAGE-001` | invalid option | project未読 | usage-error、exit 2、`cli.unknown-option`、副作用なし |
| `CA-CLEANUP-AGGREGATE-001` | `apply-change` | write failure後にcleanupも失敗 | runtime-error、二diagnostic、最保守`not-retryable`からabortへ集約 |
| `CR-MANIFEST-INVALID-001` | `validate` | manifest path escape | runtime-error、`runtime.manifest-invalid`、project未読 |
| `CR-ASSET-DIGEST-001` | `validate` | executable digest mismatch | runtime-error、`runtime.artifact-digest-mismatch`、project未読 |
| `CR-CAPABILITY-MISSING-001` | `validate` | 既知validate capabilityを除外 | runtime-error、`runtime.capability-missing`、environment修復待ち |
| `CR-SOURCE-MISSING-001` | `validate` | source archive不在 | runtime-error、`runtime.manifest-invalid`、project未読 |

`CR-*`の`runtime_setup`はNode manifest exampleをbaseにし、JSON Pointer mutationとfilesystem mutationを機械可読に記録する。runnerはbaseを専用temporary runtime directoryへmaterializeしてからmutationを適用し、元exampleを変更しない。manifest schema違反と、schema上は既知subsetとしてvalidだがcore profileを満たさないcapability不足を別stageで判定する。全runtime integrity caseでproject inputのopen/readを監視し、`expected_project_input_read = false`を満たすこと、destination entryを一件も作らないことをassertする。

## schema / binding adversarial case

`contract-cases.json`は、base exampleへRFC 6901 pointerの`add / remove / replace` mutationを順に適用し、指定layerのaccept/rejectを比較する。`json-schema` layerはschema registryへresult、diagnostic、artifact schemaを同時登録する。`cross-artifact-binding` layerはschema適合後に[result contractの`RB-001`〜`RB-012`](miku-project-cli-result-contract-v1.md#artifact-schema%E3%81%A8cross-artifact-binding)を検査する。

収録caseは、空semantic diff、不完全output plan、statusとdiagnostic codeの不一致、禁止された件数summary、楽観的next action、expected plan不一致をsuccessにする改変、runtime / request / diff / destination / approval bindingに加え、command入力またはdestinationの欠落、stdin metadata矛盾、read-only commandの虚偽effect、apply/verify path不一致、Projection scope/content/source-state不一致を含む。schema正例に加え、`RB-001`〜`RB-006`および`RB-012`が成立するbinding正例も同じindexへ置き、validatorが全入力を機械的に拒否するだけの偽実装を防ぐ。

JSON Schemaだけでは別objectの値同士やcanonical digest計算結果の同値を表現できない。したがって`json-schema` caseがpassしても`cross-artifact-binding`失敗を成功として扱わず、両layerをGate G4/G5の同じtest commandから実行する。

## unknown outcome injection

`CU-UNKNOWN-OUTCOME-001`は、runtimeへ「途中で壊れたふりをする」非公開optionを追加しない。runnerがresult transportを閉じるか、result fileの受信を意図的に放棄し、destinationに空の`COMMITTED`が現れた後だけ対象processを終了させる。成果物を書き換えず、次に独立した`verify-artifact`を実行する。

受入条件は次のとおりである。

1. callerはapplyのsuccess resultを受け取っていないため、同じapplyを再試行しない。
2. destinationの存在だけでsuccessと判断しない。
3. `verify-artifact`が三member、schema、XML、digest、runtime/request/diff/plan bindingを検査する。
4. committedなら前回operationの成果として回収し、incomplete/corruptなら利用を中止する。
5. verifyはcleanup、repair、marker再作成を行わない。

このcaseはNode/Javaのprocess停止timingを同じAPIで強制する必要はない。inject pointの証拠と最終filesystem/result assertionを共通化する。

## semantic catalogのmaterialization

G1の`S-V001`〜`S-I025`はID、期待status、意味を変更せず、P4で具体的なXML/JSON fixtureへ展開する。一行に複数変種があるIDは`S-I003-a`のようなcase suffixを付けるが、diagnostic `rule_id`は親の`S-I003`を保持する。

- valid/boundaryはvalidate successと期待semantic stateを持つ
- invalidは`semantic.invalid`と該当`rule_id`を持つ
- unsupported domain dataは`semantic.unsupported`、unsupported operationは`change.operation-unsupported`を持つ
- C1 rejectは`change.request-invalid / precondition-failed / no-op / operation-unsupported`のどれかを契約どおり固定する
- 全caseで入力不変、success時diagnostics 0、failure時成功payloadなしを検証する

seed三件だけでsemantic catalog全体を実装済みとは扱わない。G4のNode完了条件は全materialized caseのpass、G5のJava完了条件は同じcaseのpassである。

## result/golden更新規則

- golden変更は、関連契約またはfixture catalogの変更理由と同じreview単位で行う。
- runtime testの`--update-golden`相当をrelease testで使用しない。
- message、stack trace、timestamp、hostnameをgoldenへ入れない。
- pathやruntime bindingを`<ANY>`で無条件に無視せず、runnerが期待manifest/pathから具体値を算出する。
- failure caseでは禁止されたartifactが存在しないこと、cleanup不能caseではmarkerなしincompleteであることまでassertする。
- result JSONはcross-runtimeでsemantic比較し、同一runtimeの反復ではbyte determinismも検証する。

## P3.9 review checklist

- [x] corpusのauthority、directory、case ID、比較modeを定義した
- [x] canonical S-V001、invalid S-I012、unsupported S-I020のexternal XML seedをchecked-inした
- [x] C1前後のruntime非依存semantic goldenをchecked-inした
- [x] digestを含むchange requestのdeterministic materialization規則を定義した
- [x] 五command、usage error、publication競合、binding mismatchをcaseへ対応づけた
- [x] runtime manifest不正、asset digest不一致、capability不足、source欠落をproject未読のruntime-error caseへ固定した
- [x] status/code、複数diagnostic集約、expected plan、artifact間digest/runtime/destinationのadversarial caseをmachine-readable化した
- [x] `COMMITTED`後にresultを受け取れないunknown outcomeの注入・回復条件を定義した
- [x] JSON byte一致、semantic equality、artifact topologyの使い分けを定義した
- [x] G1 catalog全件をP4/P5で同じrule IDの実fixtureへmaterializeする境界を定義した
- [x] Node/Javaのlive比較だけをoracleにしていない
