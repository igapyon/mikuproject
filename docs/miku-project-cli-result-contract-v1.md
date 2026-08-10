---
title: miku-project CLI result and diagnostics contract v1
description: Gate G3で承認された、CLI result envelope、diagnostics、status、exit codeの機械契約。
topics:
  - miku-project
  - cli
  - diagnostics
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
    path: docs/miku-project-cli-contract-v1.md
    label: Gate G3承認済みcommand/I/O/publication契約
    checked: 2026-08-11
  - type: local-file
    role: semantic-rules
    path: docs/miku-project-semantic-fixture-catalog-v1.md
    label: Gate G1承認済みrule/fixture ID
    checked: 2026-08-10
  - type: local-file
    role: current-state
    path: scripts/miku-project-cli.mjs
    label: 現行diagnosticsの再利用証拠
    checked: 2026-08-10
  - type: local-file
    role: runtime-binding
    path: docs/miku-project-runtime-manifest-contract-v1.md
    label: Gate G3承認済みruntime manifest契約
    checked: 2026-08-11
---

# miku-project CLI result and diagnostics contract v1

## 文書の位置づけ

これは`ZB-P3.4`〜`ZB-P3.6`の成果物であり、五つのworkflow commandが返すresult envelope、diagnostics、status、process exit codeを定義する。[CLI contract v1](miku-project-cli-contract-v1.md) のtransport規則、[runtime capability contract v1](miku-project-runtime-capability-contract-v1.md) のruntime適合規則、[runtime manifest contract v1](miku-project-runtime-manifest-contract-v1.md) のruntime binding規則、[human gate and next action contract v1](miku-project-human-gate-and-next-action-contract-v1.md)の続行判断規則と合わせてGate G3の承認済み正本とする。

machine-readable schemaは次の三つである。

- [CLI result JSON Schema v1](schemas/miku-project-cli-result-v1.schema.json)
- [CLI diagnostic JSON Schema v1](schemas/miku-project-cli-diagnostic-v1.schema.json)
- [semantic / exchange artifact JSON Schema v1](schemas/miku-project-artifacts-v1.schema.json)

現行CLIのstructured diagnosticsはI/O metadataやexit codeを持つ点を再利用候補とするが、message文字列からcodeを推測する実装、primary outputとstderr diagnosticsの分離、`warning / error`だけの状態表現は引き継がない。

## result envelope

workflow commandがresult channelを確立できた場合、成功・拒否・usage error・runtime errorのすべてを`miku_project_cli_result/v1`一件で返す。時刻、乱数、session ID、hostname、会話IDは含めず、同じruntime・入力・option・filesystem観測から決定的なresultを作る。

| field | 必須 | 意味 |
| --- | --- | --- |
| `kind / schema_version` | ○ | `miku_project_cli_result / 1` |
| `contract` | ○ | product contract、result/diagnostic schema、diagnostic catalog version |
| `runtime` | ○ | binding status、runtime family/version、artifact/manifest digest、capability profile、fixture suite version |
| `command` | ○ | 五commandのいずれか。commandを確定できないpre-dispatch errorは`cli` |
| `side_effect_class` | ○ | `none / read-only / exchange-artifact-generation / meaning-change-and-project-artifact-generation` |
| `status / exit_code` | ○ | 下表の一対一対応 |
| `io` | ○ | stdin option、入力role/source/path/digest、result target、destination |
| `effects` | ○ | input非変更、project artifactのpath/state/ownership、cleanup |
| `observations` | ○ | normalization、loss、unsupported。diagnosticsとは分離 |
| `next_action` | ○ | successまたは全diagnosticの最保守retryabilityから導出した、次の安全な行動class |
| `diagnostics` | ○ | stable codeを持つ0件以上のdiagnostic |
| `data` | ○ | command固有payload。公開できるpayloadがなければ`null` |

fieldの追加、削除、型変更、status/exit codeの再解釈はschema versionを上げる。未知fieldを無視してv1として処理しない。JSON objectのkey順は意味を持たないが、同一runtimeの出力順は決定的にする。

### runtime binding

`runtime.binding_status`は`verified / unverified`の二値である。`verified`ではartifact digest、manifest digest、`miku-project-cli-core/v1`、fixture suite version `1`をすべてnon-nullで記録する。`succeeded`とdomain/validation上の`rejected`は必ず`verified`であり、未検証runtimeでproject inputを読まない。

`unverified`は、manifest/assetの検証自体に失敗した`runtime-error`、またはcommandを確定できない早期`usage-error`だけに許可する。この場合も推測値を入れず、確定できないdigest/profile/suite fieldは`null`にする。完全な検証順、manifest外部pin、output plan/provenanceとの同一bindingは[runtime manifest contract v1](miku-project-runtime-manifest-contract-v1.md)を正本とする。

## statusとexit code

| status | exit code | 境界 | 例 |
| --- | ---: | --- | --- |
| `succeeded` | `0` | 要求したcommand固有の成功payloadを安全に返した | valid Projection、validation成功、承認可能plan、committed set、committed検証 |
| `rejected` | `1` | invocationは正しくruntimeも契約どおり動いたが、入力・意味・安全条件・承認・artifact stateが成功条件を満たさない | invalid XML、unsupported data、precondition不一致、既存destination、absent/incomplete/corrupt set |
| `usage-error` | `2` | command grammarまたはoption指定が不正で、domain operationを開始していない | 未知command/option、必須option欠落、stdin複数指定 |
| `runtime-error` | `3` | I/O、runtime integrity、予期しない内部errorにより、要求結果を契約どおり完成できない | read/write failure、cleanup failure、runtime digest不一致、internal exception |

追加規則は次のとおりである。

- catalog v1のdiagnosticはすべて`error`であるため、`succeeded`はdiagnostics空を必須とする。normalization等の成功時観測は`observations`へ載せ、`losses`と`unsupported`が非空なら成功にしない。
- `rejected / usage-error / runtime-error`は少なくとも一件の`error` diagnosticを持つ。
- `validate`の`valid = false`と、`verify-artifact`の`absent / incomplete / corrupt`は、検査処理自体が完了していても利用者ジョブが成功していないため`rejected / 1`とする。
- `apply-change`のbinding/precondition/approval不一致は`rejected / 1`、destination予約後のwrite/cleanup failureは`runtime-error / 3`とする。
- marker作成後のresult delivery failureなど、valid envelope自体を返せないfailureはschema外である。可能ならexit `3`とemergency stderrを返し、callerは`verify-artifact`でoutcomeを回復する。
- `--help`、`--version`、`<command> --help`はcontrol operationであり、成功時はplain textとexit `0`を返す。強制終了、signal、launcher failureによるplatform固有exitはworkflow exit code契約外である。

## command固有data

| command / status | `data`の必須内容 | 禁止事項 |
| --- | --- | --- |
| `inspect / succeeded` | `projection` | full internal stateを追加しない |
| `validate / succeeded` | `validation.valid = true`、format profile、state digest | Projectionやplanned stateを返さない |
| `validate / rejected` | `validation.valid = false`。解釈前に失敗した場合はprofile/digestを`null`可 | valid state digestを捏造しない |
| `plan-change / succeeded` | `semantic_diff`と`output_plan` | 片方だけを返さない |
| `plan-change / rejected` | `null`または承認不能なvalidation情報だけ | approvalに使えるdiff/output planの組を返さない |
| `apply-change / succeeded` | committed `artifact_set` descriptor | incomplete/corrupt setを`data.artifact_set`へ入れない |
| `apply-change / rejected/runtime-error` | 原則`null` | 残ったpath/stateは`effects.project_artifact`へ記録する |
| `verify-artifact / succeeded` | `verification.publication_state = committed` | digest検証前にcommittedとしない |
| `verify-artifact / rejected` | 通常は`verification`に`absent / incomplete / corrupt`と`matches_expected_plan / bindings = null`。committedだがexpected plan不一致の場合だけ`committed / false / 実測bindings` | 利用可能artifact descriptorを返さない。不一致をsuccessにしない |
| `verify-artifact / runtime-error` | 判定不能なら`publication_state = null` | 第五のpublication stateを推測しない |
| `usage-error` | `null` | domain payloadを返さない |

`plan-change` result envelope全体は`apply-change --plan-result`の入力である。CLIは`command = plan-change`、`status = succeeded`、schema/contract/runtime binding、diff/output planの存在とdigestをすべて検証する。

### artifact schemaとcross-artifact binding

Projection、change request、semantic diff、output plan、approval、provenance、conformance用semantic stateは、kind名だけで判定せず[artifact JSON Schema v1](schemas/miku-project-artifacts-v1.schema.json)へ全体を適合させる。特にC1のsemantic diffは一件の意味変更とpreservation/provenanceを、output planはruntime binding、canonical destination、三memberのpublication topology、preflight digest、normalization/loss/unsupportedを省略できない。空objectや未知fieldを「後で埋める成功payload」として受理しない。

JSON Schemaで表現できない別object間の同値関係は、次のbinding ruleをNode/Java共通validatorとconformance runnerで検査する。比較するdigest objectは`algorithm`と`value`の両方を一致させる。

| rule ID | 必須binding |
| --- | --- |
| `RB-001` | `semantic_diff.base_state_digest = output_plan.base_state_digest`、かつ両者はrequestの`base.state_digest`と一致する |
| `RB-002` | `semantic_diff.change_request_digest = output_plan.change_request_digest = canonical(change_request)` |
| `RB-003` | `output_plan.semantic_diff_digest = canonical(semantic_diff)`、`output_plan.preflight.proposed_state_digest = semantic_diff.proposed_state_digest` |
| `RB-004` | resultのverified runtime bindingから`binding_status`を除いた六fieldが`output_plan.runtime`と一致する |
| `RB-005` | `io.destination.path = output_plan.output.destination.path`。`requested_path`は比較対象にしない |
| `RB-006` | approvalのbase/request/diff/plan digestが、承認時に提示したrequest、semantic diff、output planのcanonical digestと一致する |
| `RB-007` | provenanceのruntime、input/change/output digest、target/before/afterが、承認済みplan、実際の入力、生成artifact、再decode後semantic stateと一致する |
| `RB-008` | `verify-artifact --expect-plan-result`指定時は、committed artifactのprovenance bindingがexpected planと完全一致した場合だけ`succeeded / matches_expected_plan = true`とする。不一致は`rejected / false / publication.expected-plan-mismatch` |
| `RB-009` | result status、diagnostic code、retryability、集約`next_action`をschemaどおり一致させる。result直下にdiagnostic件数summaryを複製しない |
| `RB-010` | usage error以外ではcommandごとの入力role、option、source、順序、件数とdestination有無をschemaの固定matrixに一致させる。stdinは一入力だけ、`stdin_option`はそのoption、stdinのpathは`null`、file/directory/path入力とfile result/destinationはcanonical absolute pathでなければならない。successで実際に読んだartifact-set path以外の入力digestは非nullとする |
| `RB-011` | plan resultでは`io.destination.path = output_plan.output.destination.path`。apply successではそのpath、承認済みplanのdestination、`effects.project_artifact.path`、`data.artifact_set.path`を一致させる。verifyではartifact-set入力path、`verification.path`、観測した`effects.project_artifact.path`と、verification/effect双方のpublication stateを一致させる。cleanup pathが非nullなら同じartifact pathでなければならない |
| `RB-012` | Projectionの`source_state_digest`はcanonical semantic state digestと一致し、purpose固有scopeとcontentは同じstateから決定的に導出されなければならない。`project_overview`はproject、全taskのoverview field、全dependencyを、`task_change_context`はtarget leaf task、ancestor chain、targetに接続するdependency、target assignmentとそのresourceを過不足なく反映する |

canonical JSONとdigest規則は[conformance corpus v1](miku-project-conformance-corpus-v1.md)を正本とする。外部fileを表す`io.inputs[].digest`は、その入力として読んだraw byte列のSHA-256であり、内部artifactのcanonical digestと混同しない。schema適合後にbinding ruleを検査し、どちらか一方でも失敗すれば成功payloadとして扱わない。[machine-readable contract cases](../testdata/conformance/v1/contract-cases.json)はschema層とbinding層を分けて期待結果を固定する。

## I/O metadata

`io.inputs`はcommand option順で並べ、各entryにlogical role、option、source、canonical path、digestを持たせる。usage error以外の固定matrixは、`inspect / validate = project`、`plan-change = project, change_request`、`apply-change = project, change_request, plan_result, approval`、`verify-artifact = artifact_set, optional expected_plan_result`である。stdinは`path = null`、未読または読取り前failureは`digest = null`である。pathを持つ入力はcanonical absolute pathを記録する。usage errorはoption解析が完了していない可能性があるため、この入力matrixだけを免除する。

`source = file / directory`は`--project`等で実際に分類して読んだentry、`source = stdin`は選択された一入力を表す。`verify-artifact --artifact-set`はpathがabsentまたは壊れたdirectoryでも検査対象になるため、存在・typeの成功を含意しない`source = filesystem-path`へ固定し、raw file digestは持たせない。

`io.result`は実際に選んだ`stdout`または予約済み`file`とcanonical pathを記録する。指定result fileを予約できずstdoutへerror resultをfallbackした場合は、実際のtargetを`stdout`とする。`io.destination`は`plan-change / apply-change`だけが持ち、caller指定値とcanonical absolute pathを区別する。

pathやdigestを取得できなかった場合、推測値を入れず`null`にする。秘密値、file内容、環境変数、会話履歴はI/O metadataへ入れない。

## effectsとpublication

`effects.project_input_modified`はv1で常に`false`である。`effects.project_artifact`は、commandが観測または生成したproject artifact setの状態を報告するdescriptorであり、成功payloadではない。

- `created_by_invocation = true`は、その`apply-change`がdestinationのexclusive reservationに成功した場合だけ許可する。
- committed setだけを`data.artifact_set`へ載せる。incomplete/corrupt pathは`effects`とdiagnosticsだけへ載せる。
- `cleanup.status`は`not-needed / succeeded / failed / prohibited-after-commit`のいずれかである。
- `apply-change`がcommittedになった後はcleanupを行わないため`prohibited-after-commit`とする。destination予約前は`not-needed`、marker前のcleanup結果は`succeeded / failed`を使う。
- cleanup対象がなければ`path = null`、対象があればcanonical absolute pathを使う。
- `verify-artifact`は観測だけなので`created_by_invocation = false`、cleanupは`not-needed`である。
- `inspect / validate`はproject artifact setを観測対象にしていても`effects.project_artifact = null`とする。publication状態を報告する責務は`verify-artifact`だけに置く。
- `apply-change / runtime-error`は`publication_state = committed`を報告してはならない。marker作成後にresult deliveryを失ったケースはvalid envelopeへ推測で書かず、独立した`verify-artifact`で回復する。
- schemaが各pathをabsoluteに制約し、`RB-011`がobject間のpath同値を検査する。片方だけ正しいpathを持つresultは成功payloadとして扱わない。

## observations

normalization、loss、unsupportedは人間向けmessageから独立した構造化配列である。

| array | item | successとの関係 |
| --- | --- | --- |
| `normalizations` | `code / path / before / after` | semantic equivalentならsuccess可 |
| `losses` | `code / path / description` | v1では非空なら`rejected` |
| `unsupported` | `code / path / description` | v1では非空なら`rejected` |

配列は`code`、`path`順に決定的に並べる。同一`code + path`を重複させない。`before / after`はschema化されたJSON valueであり、messageの抜粋ではない。入力XMLのUTF-8 BOMは`text.utf8-bom-removed` normalizationとして扱う。

## diagnostic schema

各diagnosticは次を必須とする。

| field | 意味 |
| --- | --- |
| `kind / schema_version` | `miku_project_cli_diagnostic / 1` |
| `code` | catalogにあるstable code。messageから導出しない |
| `severity` | severity field。closed catalog v1では全codeを`error`へ固定 |
| `category` | usage、I/O、artifact、semantic、change、publication、runtime等 |
| `message` | 人向け説明。文言・言語はmachine contractではない |
| `location` | scope、path、option、artifact role、semantic fixture/rule ID |
| `retryability` | 同じ入力で再実行してよいかを示す固定語彙 |
| `details` | 補助的な構造化値。machine分岐はcode/location/retryabilityを優先する |

`location.scope = semantic`の`path`はsemantic contractのpath、`artifact`はRFC 6901 JSON Pointer、`filesystem`はcanonical absolute pathを使う。該当しないfieldも省略せず`null`にする。G1 fixtureに対応する場合は`rule_id = S-Ixxx`を必須とし、対応しないI/O/usage errorでは`null`にする。

catalog v1では全件`error`なので、diagnosticsは`code`、`location.scope`、`path`、`option`の順で決定的に並べる。将来severityを追加するschemaでは`error / warning / info`を第一sort keyとする。v1 resultまたはProjectionに件数summaryを重複保持せず、必要なcallerがresultの`diagnostics`配列から算出する。

### retryability

| value | 意味 |
| --- | --- |
| `after-input-change` | optionまたは入力artifactを修正してから再実行 |
| `after-environment-change` | permission、容量、runtime等の外部条件を直してから再実行 |
| `after-replan-and-approval` | current state/destination/runtimeを再計画し、人の再承認後に実行 |
| `not-retryable` | 自動再試行しない。人による調査またはruntime修正が必要 |

複数diagnosticがある場合、`not-retryable`、`after-replan-and-approval`、`after-environment-change`、`after-input-change`の順で保守的に集約する。集約値と`next_action.action / command`の一対一対応、human gate、schema外failureは[human gate and next action contract v1](miku-project-human-gate-and-next-action-contract-v1.md)を正本とする。

v1 catalogには、条件を何も変えず自動retryしてよいdiagnosticを置かない。各codeに許されるretryabilityは[diagnostic JSON Schema v1](schemas/miku-project-cli-diagnostic-v1.schema.json)でも固定し、実装が同じcodeへ楽観的な値を付け替えることを許さない。

## diagnostic code catalog v1

catalog v1は次のcodeを閉じた集合とする。同じcodeの意味や既定statusをruntimeごとに変えない。追加・再解釈が必要ならdiagnostic schemaとcatalog versionを上げ、runtime manifestで対応versionを宣言する。

| code | category | 通常status | 主なretryability |
| --- | --- | --- | --- |
| `cli.unknown-command` | usage | usage-error | after-input-change |
| `cli.unknown-option` | usage | usage-error | after-input-change |
| `cli.missing-option` | usage | usage-error | after-input-change |
| `cli.duplicate-option` | usage | usage-error | after-input-change |
| `cli.unexpected-argument` | usage | usage-error | after-input-change |
| `cli.invalid-option-value` | usage | usage-error | after-input-change |
| `cli.multiple-stdin-sources` | usage | usage-error | after-input-change |
| `io.input-not-found` | io | rejected | after-input-change |
| `io.input-type-invalid` | io | rejected | after-input-change |
| `io.input-symlink-rejected` | io | rejected | after-input-change |
| `io.input-read-failed` | io | runtime-error | after-environment-change |
| `io.result-path-exists` | io | rejected | after-input-change |
| `io.result-path-unsafe` | io | rejected | after-input-change |
| `io.result-reservation-failed` | io | runtime-error | after-environment-change |
| `text.invalid-utf8` | encoding | rejected | after-input-change |
| `json.invalid` | json | rejected | after-input-change |
| `json.bom-not-allowed` | json | rejected | after-input-change |
| `json.duplicate-key` | json | rejected | after-input-change |
| `artifact.kind-unsupported` | artifact | rejected | after-input-change |
| `artifact.schema-version-unsupported` | artifact | rejected | after-input-change |
| `xml.invalid` | xml | rejected | after-input-change |
| `xml.encoding-unsupported` | xml | rejected | after-input-change |
| `xml.profile-unsupported` | xml | rejected | after-input-change |
| `semantic.invalid` | semantic | rejected | after-input-change |
| `semantic.unsupported` | semantic | rejected | after-input-change |
| `change.request-invalid` | change | rejected | after-input-change |
| `change.operation-unsupported` | change | rejected | after-input-change |
| `change.precondition-failed` | change | rejected | after-replan-and-approval |
| `change.no-op` | change | rejected | after-input-change |
| `change.binding-mismatch` | change | rejected | after-replan-and-approval |
| `change.approval-invalid` | change | rejected | after-replan-and-approval |
| `publication.destination-exists` | publication | rejected | after-replan-and-approval |
| `publication.destination-unsafe` | publication | rejected | after-input-change |
| `publication.capability-unsupported` | publication | rejected | after-environment-change |
| `publication.reservation-conflict` | publication | rejected | after-replan-and-approval |
| `publication.write-failed` | publication | runtime-error | after-environment-change |
| `publication.postwrite-verification-failed` | publication | runtime-error | not-retryable |
| `publication.cleanup-failed` | publication | runtime-error | not-retryable |
| `publication.artifact-absent` | publication | rejected | not-retryable |
| `publication.artifact-incomplete` | publication | rejected | not-retryable |
| `publication.artifact-corrupt` | publication | rejected | not-retryable |
| `publication.expected-plan-mismatch` | publication | rejected | after-replan-and-approval |
| `runtime.manifest-invalid` | runtime | runtime-error | not-retryable |
| `runtime.artifact-digest-mismatch` | runtime | runtime-error | not-retryable |
| `runtime.capability-missing` | runtime | runtime-error | after-environment-change |
| `internal.unexpected-error` | internal | runtime-error | not-retryable |

全codeのseverityはv1では`error`である。normalizationは`observations`へ載せ、単なるinformational messageのためにdiagnostic codeを増やさない。将来warning/info codeを追加する場合はdiagnostic schemaとcatalog versionを更新する。

## examples

schema validation用の例は次に置く。

- [validate rejected result](examples/cli-v1/validate-rejected.result.json)
- [inspect succeeded result](examples/cli-v1/inspect-succeeded.result.json)
- [plan-change succeeded result](examples/cli-v1/plan-change-succeeded.result.json)
- [apply-change succeeded result](examples/cli-v1/apply-change-succeeded.result.json)
- [apply-change incomplete result](examples/cli-v1/apply-change-incomplete.result.json)
- [verify-artifact committed result](examples/cli-v1/verify-artifact-committed.result.json)
- [verify-artifact expected-plan mismatch result](examples/cli-v1/verify-artifact-expected-plan-mismatch.result.json)
- [usage error result](examples/cli-v1/usage-error.result.json)
- [runtime manifest invalid result](examples/cli-v1/runtime-manifest-invalid.result.json)

exchange artifact単体のschema例は[artifact examples](examples/artifacts-v1/)に置く。公式例に含むsemantic state、change request、semantic diff、output planのdigestはcanonical JSON規則から再計算でき、関連例のbinding値と一致させる。runtime release artifactや実際の出力XMLがまだ存在しない箇所のdigestは例示値であり、P4のmaterialization時に実byte値へ置き換える。

これらはconformance fixtureの代替ではない。実データ、golden semantic result、Node/Java比較は[conformance corpus v1](miku-project-conformance-corpus-v1.md)と`testdata/conformance/v1/`を正本とする。

## P3.4 / P3.5 / P3.6 / P3.10 / P3.12 review checklist

- [x] successと全failure classが同じversioned result envelopeを使う
- [x] statusとexit codeが一対一で、validation rejectionとruntime failureを区別する
- [x] command固有dataと失敗時に公開してよい情報が定義されている
- [x] I/O path/digest、project effect、cleanup、normalization/loss/unsupportedがmessage外にある
- [x] diagnosticがstable code、severity、scope/path、rule ID、retryabilityを持つ
- [x] G1 fixture IDをsemantic diagnosticへ対応づけられる
- [x] incomplete/corrupt setを成功payloadへ載せない
- [x] Agentがstderrやmessage文字列をparseしなくてよい
- [x] result/diagnostic schemaがJSON Schema 2020-12として保存されている
- [x] Projection、semantic diff、output plan、approval、provenanceの全体schemaとcross-artifact binding ruleが固定されている
- [x] successとdomain rejectionを検証済みmanifest/asset/profile/fixture suiteへ束縛する
- [x] manifest検証前のfailureだけを`unverified`として機械判定できる
- [x] successまたはdiagnostic retryabilityから`next_action`をmessage解析なしに導出できる
- [x] unknown outcomeではapply再試行より`verify-artifact`を優先する
- [x] diagnostic件数の重複fieldを置かず、status/code/retryability/next actionの不整合をschemaで拒否する
