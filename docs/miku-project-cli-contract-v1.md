---
title: miku-project CLI contract v1
description: Gate G3で承認された、v1 CLIの最小command matrix、artifact flow、意味上の副作用分類。
topics:
  - miku-project
  - cli
  - agent-skills
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
    path: docs/miku-project-semantic-contract-v1.md
    label: Gate G1承認済みsemantic contract
    checked: 2026-08-10
  - type: local-file
    role: primary
    path: docs/miku-project-format-and-loss-contract-v1.md
    label: Gate G2承認済みformat and loss contract
    checked: 2026-08-10
  - type: local-file
    role: primary
    path: docs/miku-project-change-contract-v1.md
    label: Gate G2承認済みchange contract
    checked: 2026-08-10
  - type: local-file
    role: scenario
    path: docs/miku-project-zero-base-scenarios-v1.md
    label: G0承認済みR1/C1シナリオ
    checked: 2026-08-10
  - type: local-file
    role: current-state
    path: scripts/miku-project-cli.mjs
    label: 現行CLIのcompatibility evidence
    checked: 2026-08-10
---

# miku-project CLI contract v1

## 文書の位置づけ

これは2026-08-11に`Gate G3`で承認されたCLI contractである。`ZB-P3.1`〜`ZB-P3.12`のcommand、I/O、publication、result/diagnostics/exit code、Node参照実装、runtime capability、conformance corpus、runtime manifest、副作用分類、human gateとsafe next actionを確定した。

この文書は新しいv1 CLI契約の正本である。現行の`ai / state / import / export / report` command群は移行判断までcompatibility surfaceとして扱い、このcommand matrixの設計根拠や暗黙aliasにはしない。

## runtime実装方針

v1はNode CLIを最初の実行可能な参照実装とする。Nodeは`G3`で承認された契約を先に実装・検証し、`G4`で固定したcontract releaseをJava適合の入力とする。Java CLIはmoving Node sourceの逐次移植ではなく、固定済みの共通契約とconformance corpusを実装するruntimeである。

仕様の正本は、承認済みの製品・semantic・format・change・CLI契約、JSON Schema、および`ZB-P3.9`で固定する共通fixture / goldenである。Nodeの実行結果がこれらと矛盾しても、その挙動だけで仕様を変更しない。不一致はNode defectまたは契約上の不整合として明示的に解決し、必要なら契約をreviewしてversion管理する。Javaの適合判定もlive Node出力だけをoracleにせず、同じ正本を使う。

共通v1 capabilityについては、command semantics、構造化result、diagnostics、exit状態、determinism、artifact publicationの意味的な適合をNodeとJavaの両方に求める。具体的なcore profile、command要求、静的capabilityと動的preflight、extension、runtime選択とfallbackの境界は[runtime capability contract v1](miku-project-runtime-capability-contract-v1.md)を正本とする。v1のNode-only / Java-only extension setは空である。

## v1 command設計の原則

- 人、shell script、CI、AI Agentが同じcommandとartifact schemaを使う
- command名から、read-only、計画、承認済み変更、artifact検証のどれかを判別できる
- command間の受け渡しはversioned artifactだけを使い、会話履歴、session、作業directory内の暗黙fileへ依存しない
- 入力project artifactと既存outputを変更しない
- 意味変更は`plan-change`とhuman gateを経た`apply-change`だけに限定する
- 成功したproject outputは`COMMITTED`を持つ新規artifact set directoryだけとする
- commandを再利用しやすくするために責務を分けるが、v1 scenarioに存在しない汎用変換やreport commandは増やさない
- 人間向け文言の解析ではなく、後続で定義するresult、diagnostics、artifact metadata、exit statusから次の操作を選べるようにする

## control plane

v1 runtimeは`--help`と`--version`を必須とする。これらはproject workflow commandではなく、入力projectを読まず、fileを生成しないcontrol operationである。

独立した`capabilities` commandはv1に設けない。capability IDと`miku-project-cli-core/v1`の内容は[runtime capability contract v1](miku-project-runtime-capability-contract-v1.md)、machine-readableなproduct contract version、runtime version、fixture suite version、artifact/source digest、capability setは[runtime manifest contract v1](miku-project-runtime-manifest-contract-v1.md)とそのJSON Schemaを正本とし、必要なruntime bindingを各result metadataからも参照できるようにする。

## 最小command matrix

v1のproject workflow commandは次の五つだけとする。command wordはlowercase kebab-caseの固定値である。

| command | 主責務 | logical input | 成功時のprimary output | 副作用class | human gate |
| --- | --- | --- | --- | --- | --- |
| `inspect` | validなprojectからpurpose別Projectionを生成する | external XMLまたはcommitted artifact set、Projection purpose、必要なscope | `miku_project_projection/v1` | read-only | 不要 |
| `validate` | projectのformat、semantic invariant、unsupported dataを検証する | external XMLまたはcommitted artifact set | validation resultとdiagnostics | read-only | 不要 |
| `plan-change` | requestをdry-runし、意味差と公開計画を確定する | current project、`miku_project_change_request/v1`、未使用destination | `miku_project_semantic_diff/v1`と`miku_project_output_plan/v1` | exchange artifact生成 | 出力後に必要 |
| `apply-change` | 承認済みの局所変更を再検証し、新しいproject artifact setを生成する | current project、request、diff、output plan、approval | committed `miku_project_artifact_set/v1` | 意味変更 + project artifact生成 | 実行前に必須 |
| `verify-artifact` | artifact setのpublication状態と整合性を判定する | artifact set directory path | `absent / incomplete / committed / corrupt`と検証結果 | read-only | 不要 |

`logical input`はartifact roleを示す。具体的なoption名、stdin/stdout/fileの割当て、複数primary artifactのtransportは次節、result内のfield schemaとstatus/exit codeは [CLI result and diagnostics contract v1](miku-project-cli-result-contract-v1.md) で固定する。

## invocation grammar

v1は次のlong optionだけを持つ。短縮option、位置引数、`--key=value`、暗黙stdin、環境変数による既定値、設定fileによるoption注入は設けない。

```text
miku-project --help
miku-project --version
miku-project <command> --help

miku-project inspect
  --project <path|->
  --purpose <project_overview|task_change_context>
  [--task-uid <uid>]
  [--result <path|->]

miku-project validate
  --project <path|->
  [--result <path|->]

miku-project plan-change
  --project <path|->
  --request <path|->
  --destination <new-directory-path>
  [--result <path|->]

miku-project apply-change
  --project <path|->
  --request <path|->
  --plan-result <path|->
  --approval <path|->
  [--result <path|->]

miku-project verify-artifact
  --artifact-set <path>
  [--expect-plan-result <path|->]
  [--result <path|->]
```

共通規則は次のとおりである。

- 各optionは高々一回だけ指定できる。必須optionの欠落、未知option、余分な位置引数、同じoptionの重複はusage errorとする。
- optionの順序はcommand wordより後なら意味を持たない。
- `--result`の省略は`--result -`と同じである。`-`はそのoptionで明示的に許可したstdinまたはstdoutだけを表す。
- stdinを読むinput optionは一回のinvocationで高々一つである。`--project - --request -`のような組合せは、stdinを読む前に拒否する。
- `--project <path>`は通常fileならexternal XML、通常directoryならartifact setとして判別する。directoryはcommitted検証を通らなければproject inputにできない。`--project -`はUTF-8 XMLだけを受け取り、artifact setをstream化しない。
- `project_overview`では`--task-uid`を許可しない。`task_change_context`では`--task-uid`をちょうど一つ必須とする。
- `--plan-result`は、成功した同一contract versionの`plan-change` result envelope全体を受け取る。callerにdiffとoutput planを抽出・再構成させない。
- `--expect-plan-result`を指定した`verify-artifact`は、committed setのprovenance bindingがそのplanning resultと一致するかも検証する。省略時はpublicationと内部整合性だけを検証する。
- `apply-change`は`--destination`を受け取らない。承認されたdestinationは`--plan-result`内のoutput planだけから得るため、apply時に別pathへ差し替えられない。
- v1に`--force`、`--overwrite`、`--yes`、interactive promptは存在しない。

`--help`、`--version`、`<command> --help`はcontrol operationである。`--version`は`miku-project <runtime-version>`とLF一件、helpはUTF-8 textをstdoutへ返し、project inputやresult fileを必要としない。machine-readableなversion/capability判断にはこれらの文言ではなくruntime manifestを使う。

## result transport

五つのworkflow commandは、成功・validation rejection・usage error・runtime errorを問わず、原則としてversioned `miku_project_cli_result/v1` JSON documentを一件だけ返す。result schema、diagnostics field、statusとexit codeは [CLI result and diagnostics contract v1](miku-project-cli-result-contract-v1.md) に定め、transportは次で固定する。

- `--result -`ではstdoutにJSON document一件と末尾LFを出す。stdoutへprogress、prompt、human message、複数JSON、JSON Linesを混ぜない。
- `--result <path>`では、domain operationを始める前に存在しない通常fileをexclusive createしてresult channelを予約する。stdoutは空にし、同じJSON documentと末尾LFを予約fileへ書く。
- result pathが既存、symlink、directory、親directory不存在、またはexclusive create不能ならdomain operationを開始しない。このpreflight errorだけは指定result pathへ書けないため、stdoutへstructured error resultを返す。
- process終了やI/O failureで予約済みresult fileが空または途中書込みになり得る。完全なJSON parseとresult schema validationを通らないfileは入力artifactとして信頼せず、CLIは後続実行で暗黙削除・再利用しない。
- result fileへの書込み失敗時は自身が予約したfileだけをbest-effort cleanupし、非zero exitとstderrのemergency messageを許す。このfailureはvalid result envelope内のdiagnosticとして表現できないschema外failureであり、架空の`next_action`を生成しない。`apply-change`でproject publication後にresult deliveryが失敗した可能性がある場合、callerはdestinationを`verify-artifact`で検査する。
- `plan-change --result plan.json`のfile全体を、human gate後に`apply-change --plan-result plan.json`へ渡せる。result envelopeに含まれるsemantic diffとoutput planはそれぞれ元のartifact schemaとdigestを保つ。
- `inspect`のProjection、`plan-change`のdiff/output plan、`validate`と`verify-artifact`の検証情報はresult envelope内へ埋め込む。これらの個別file出力optionはv1に設けない。
- `apply-change`のproject artifactは`--result`とは別であり、承認済みdestination directoryだけへ生成する。result fileをartifact set memberにしてはならない。

workflow diagnosticsはすべてresult envelopeに含め、通常のstderrを機械契約に使わない。stderrはresult channel自体を確立・完了できないlauncher、runtime、I/O failureのemergency text専用であり、callerやAgentはその文字列をparseして分岐してはならない。現行CLIの`--diagnostics text|json`はv1に引き継がない。

## input transport

| option | 許可するsource | content |
| --- | --- | --- |
| `--project` | 通常file、通常directory、`-` | file/`-`はexternal XML、directoryはartifact set |
| `--request` | 通常file、`-` | `miku_project_change_request/v1` JSON |
| `--plan-result` | 通常file、`-` | 成功した`plan-change` CLI result JSON |
| `--approval` | 通常file、`-` | `miku_project_change_approval/v1` JSON |
| `--artifact-set` | filesystem path（不存在可、stdin不可） | 検証対象artifact set |
| `--expect-plan-result` | 通常file、`-` | 成功した`plan-change` CLI result JSON |

- stdinは明示的に値`-`を指定したoptionだけがEOFまで読む。input optionを省略してTTYやpipeを暗黙に待たない。空stdinはinvalid inputとする。
- JSON artifactはdocument一件だけを許可し、前後のJSON whitespace以外のbyte、duplicate object key、未知schema versionを拒否する。
- CLIはJSON artifactやresult envelopeからfile pathを暗黙探索しない。必要なpathはcommand optionまたはversioned artifact内の承認済みfieldだけから得る。
- `--project`とJSON inputのdirect entryはsymlinkではない通常fileまたは通常directoryを要求する。symlinkを使う場合はcallerが解決済みpathを明示する。`--artifact-set`は不存在を`absent`として観測でき、存在するsymlinkや非directoryは`corrupt`として扱う。artifact set memberは従来どおりsymlink禁止である。
- internal semantic stateをpublic CLI fileとしてmaterializeするoptionはv1に設けない。debug/conformance専用artifactは通常command surfaceと分離して`ZB-P3.9`で決める。

## text、encoding、binary

- v1 workflowのstdin、stdout、JSON、XML、provenanceはすべてUTF-8 textである。不正UTF-8 byte列はreplacement characterへ置換せずerrorにする。
- JSON input、JSON result、`provenance.json`はBOMを許可しない。出力はUTF-8 BOMなし、LF改行、末尾LF一件とする。JSONのcross-runtime比較はbyte比較ではなくschemaに基づくsemantic比較とし、同一runtimeでは決定的なbyte列を要求する。
- external XML inputはBOMなしUTF-8または先頭にUTF-8 BOM一件を許可する。XML declarationのencodingが存在する場合は大文字小文字を区別せず`UTF-8`だけを許可する。UTF-16/32、他encoding、二重BOM、先頭以外のBOMはunsupported-errorとする。
- 入力XMLのUTF-8 BOM除去はsemantic changeではないtransport normalizationとして`text.utf8-bom-removed`をresult、C1 output plan、成功時provenanceへ記録する。
- 生成する`project.xml`はUTF-8 BOMなし、XML declarationのencodingを`UTF-8`、改行をLF、末尾LF一件へ固定する。XML element内textのUnicode scalar sequenceはsemantic contractどおり保持し、CLI transportでUnicode normalizationしない。
- NUL byteはJSON/XML/text input、path optionのいずれでも拒否する。
- v1にbinary input/outputはなく、`--in-base64`、`--out-base64`、Base64 field、raw binary stdoutを設けない。将来binary formatを採用するときに別capabilityとして再評価する。

## path resolution

- relative pathはprocess開始時のcurrent working directoryを基準に解決する。CLIは`~`、環境変数、globを展開しない。
- input file/directoryは存在確認後にreal pathを得てresult metadataへ記録する。direct entryがsymlinkなら拒否するが、既存ancestorの解決結果はreal pathへcanonicalizeする。
- `--destination`と`--result <path>`は、既存する親directoryのreal pathと未使用basenameからabsolute canonical pathを作る。basenameが空、`.`、`..`、root、NULを含む値は拒否する。親directoryを再帰作成しない。
- destinationの既存entryは通常file、directory、symlink、dangling symlinkのいずれでも拒否する。permission等で不存在を確定できない場合もfail closedとする。
- destinationをinput artifact setと同じpathまたはその子孫にできない。result pathもinput artifact set、検証対象artifact set、destinationと同じpathまたはその子孫に置けない。
- `plan-change`はcanonical absolute destinationをoutput planへ記録する。`apply-change`はcurrent working directoryに依存せず、その値とapply時に再解決した親real pathが一致することを要求する。
- pathのcase sensitivity、separator、Unicode normalizationはhost filesystemに従い、CLIが推測変換しない。runtime間でpath byte表現が異なる場合、output planはruntime-boundであるため再計画・再承認する。

## project artifact publication

`apply-change`はGate G2の`exclusive-directory-commit-marker/v1`を次のCLI-visible protocolとして実行する。

1. `--result <path>`を使う場合はresult channelを先にexclusive reservationする。
2. current project、request、plan result、approvalを読取り、schema、digest、precondition、runtime bindingを再検証する。
3. destination parentのreal pathとfilesystem capabilityを再確認し、destinationが不存在であることを確認する。
4. destination directoryをnon-recursiveかつexclusiveに新規作成する。競合で既存になった場合は一切触れず失敗する。
5. 予約したdirectory内へ`project.xml`と`provenance.json`をそれぞれexclusive createし、書込みを完了してhandleをcloseする。
6. markerなしの状態でXML再decode equivalence、provenance schema、二memberだけであること、通常file、非symlink、digestを検証する。
7. 空の通常file`COMMITTED`をexclusive createしてcloseする。これが唯一のlogical commit pointである。
8. 三member、marker size、schema、XML profile、digestを再検証してからsuccess resultを生成する。

`plan-change`はdestinationを作らず、read-onlyなpath/capability preflightだけを行う。preflight時点の不存在は予約ではないため、human gate中のraceを`apply-change`が必ず再検査する。runtime/filesystem capabilityはcallerの自己申告ではなく、runtime manifestの静的capabilityとdestination parentの動的preflightを使う。判定が`unsupported`または`unknown`ならfail closedとする。判定の意味は[runtime capability contract v1](miku-project-runtime-capability-contract-v1.md)、manifest bindingは[runtime manifest contract v1](miku-project-runtime-manifest-contract-v1.md)、具体的なNode/Java APIは`ZB-P4` / `ZB-P5`で実装する。

v1が要求するdurabilityはfile handleのcloseまでであり、fileまたはdirectoryの`fsync`、電源断後の永続性、三fileのfilesystem上の同時可視性は約束しない。process crashや電源断の後は必ず`verify-artifact`で状態を再判定する。

### failure、cleanup、outcome

| commit pointとの関係 | CLIの処理 | observable state | safe next action |
| --- | --- | --- | --- |
| destination reservation前に失敗 | destinationへ触れない | `absent` | 入力を直して再計画する |
| reservation後、marker作成前に失敗しcleanup成功 | 自身が作ったmemberとdirectoryだけを削除 | `absent` | diagnosticsに従い再計画する |
| reservation後、marker作成前に失敗しcleanup失敗 | それ以上変更しない | `incomplete` | 利用・再利用せず、人がpathを確認する |
| marker作成後の再検証失敗 | cleanup・repairしない | `corrupt` | 利用せず、人が隔離・調査する |
| marker作成後、result delivery前後に中断 | 自動cleanupしない | `committed`または`corrupt` | `verify-artifact --expect-plan-result`で回復する |

cleanup statusは`not-needed / succeeded / failed / prohibited-after-commit`のいずれかとしてstructured diagnosticsへ載せる。CLIが削除してよいのは、そのinvocationがexclusive reservationに成功し、まだ`COMMITTED`を作っていないdestination内の、自身が作った二memberとdirectoryだけである。既存entry、別invocationのdirectory、markerありdirectory、input、result inputは削除しない。通常signalを処理できた場合は同じbest-effort cleanupを試みるが、強制終了や電源断での実行は保証しない。

`verify-artifact`の四状態は次のようにCLI境界を補足する。

- `absent`: path entryが存在しない
- `incomplete`: symlinkではない通常directoryが存在し、`COMMITTED` entryが存在しない
- `committed`: 許可された三memberだけが正しいtype/contentで存在し、全schema/XML/digest検証が成功する
- `corrupt`: path entryが存在するが通常directoryでない、または`COMMITTED`が存在するのにcommitted条件を満たさない

permission、transient I/O、観測中の変更などで四状態を確定できない場合、第五のpublication stateを推測せず、result自体をinspection failureとして`publication_state = null`にする。`verify-artifact`はどの状態でもcleanup、repair、marker作成、member削除を行わない。

## command semantics

### `inspect`

`inspect`はR1で人またはAgentへ渡すProjectionを作るcommandである。purposeはGate G2で承認された`project_overview`または`task_change_context`だけを許可する。

- `project_overview`はproject全体の構造理解に必要な限定情報を返す
- `task_change_context`は指定したleaf task一件の変更判断に必要なscopeだけを返す
- decodeとsemantic validationを内部で必ず行い、invalidまたはunsupported dataを持つ入力から成功Projectionを返さない
- Projectionからsemantic stateを復元できることを約束しない
- 入力がartifact setの場合、committed判定を先に通す。incomplete/corrupt setを部分的にinspectしない

### `validate`

`validate`はprojectを変更せず、継続可否を機械判定するcommandである。

- external XMLではformat profile、lexical、mapping、semantic invariant、unsupported dataを検証する
- artifact setではpublication状態、member、provenance、digestを検証した後、含まれるXMLをsemantic validateする
- invalid/unsupportedのときも、成功Projectionやplanned stateを返さない
- repair、normalizationの適用、cleanup、marker作成は行わない

`inspect`も内部validationを行うが、`validate`はProjectionを必要としないCI、preflight、問題診断のために独立commandとして残す。

### `plan-change`

`plan-change`はC1のhuman gate直前までを担当する。projectの意味やfinal destinationを変更しない。

- current projectをdecode/validateし、requestのschema、selector、precondition、operation allowlistを検証する
- dry-run applyとpost-apply semantic validationを行う
- semantic diffを作る
- planned stateをpreflight encode/redecodeし、normalization、loss、unsupported、destination、runtime bindingを含むoutput planを作る
- loss/unsupported、unsafe destination、実行runtimeまたはfilesystemのcapability不足では、承認可能なdiff/output planの組を成功出力しない
- destination directory、`project.xml`、`provenance.json`、`COMMITTED`を作らない

diffとoutput planは同じplanning operationから生成する。v1に任意の二状態を比較するgeneric `diff` commandは設けない。

### `apply-change`

`apply-change`だけが、変更後の意味を持つproject artifactを生成できる。

- request、diff、output plan、approvalをすべて明示入力として要求する
- current projectを再読取りし、state/request/diff/output planのdigest、precondition、runtime bindingを再計算してapprovalと照合する
- human gateをCLI内部で代行せず、CLIがapprovalを作成しない
- final destinationを排他的に新規作成し、Gate G2のcommit marker protocolに従う
- input project、既存path、incomplete/corrupt directoryを置換・修復・再利用しない
- success resultにはcommitted artifact setだけを載せる

`apply-change`を含む全workflow commandはTTY非依存の非対話commandである。`--yes`、対話prompt、会話履歴からの承認推測はv1に存在しない。human gateの提示内容、承認/拒否/修正要求、retry/abort、resultの`next_action`は[human gate and next action contract v1](miku-project-human-gate-and-next-action-contract-v1.md)を正本とする。

### `verify-artifact`

`verify-artifact`はartifact setを利用可能か判断し、途中終了後のoutcomeを回復するcommandである。

- directoryの不存在、marker欠落、member/type/schema/XML/digest不一致を区別する
- committed setではprovenanceに記録されたrequest/diff/output plan digestとruntime bindingを返し、呼び出し側が直前のoperationと照合できるようにする
- `COMMITTED`作成後、CLIのsuccess resultを受信する前にprocessやtransportが終了した場合、同じdestinationをこのcommandで検査する
- verification resultがcommittedなら前回operationの成果を回収でき、incomplete/corruptなら利用を中止できる
- cleanup、repair、marker作成、member削除は行わない

`verify-artifact`はpublication protocol専用であり、project semantic validationを主目的とする`validate`とは分ける。ただしcommitted判定に必要なXML profileとdigest検証は省略しない。

## 副作用分類

| class | command | 許される効果 | 禁止する効果 |
| --- | --- | --- | --- |
| read-only | `inspect`、`validate`、`verify-artifact` | result/diagnosticsを返す | 入力変更、project output、cleanup、repair、marker作成 |
| exchange artifact生成 | `plan-change` | diffとoutput planを返す | projectの意味変更、final destination作成、approval作成 |
| 意味変更 + project artifact生成 | `apply-change` | 承認された意味を持つ新規committed artifact setを一件生成する | 入力変更、既存path置換、未承認変更、partial success |

stdoutや明示output fileへresult/exchange artifactを書くことはtransport上のI/Oであり、project semantics上の副作用とは区別する。transportの安全規則は`ZB-P3.2`で定める。

## v1 workflow

### R1

```text
validate
  → inspect（project_overviewまたはtask_change_context）
  → 人またはAgentが理解・判断
```

`inspect`は内部validationを含むため単独実行もできる。明示的な`validate → inspect`はCI、Agent Skills、問題調査でvalidation resultを独立保存したい場合の推奨順序である。

### C1

```text
inspect（task_change_context）
  → 人またはAgentがchange requestを作る
  → plan-change
  → semantic diff + output planを人が確認
  → callerがapproval artifactを作る
  → apply-change
  → verify-artifact
```

`verify-artifact`は通常成功後の確認にも、success resultを受け取れなかったunknown outcomeの回復にも使う。

## v1に置かないcommand

| command候補 | v1に置かない理由 | 再評価先 |
| --- | --- | --- |
| `convert` | 外部read/write profileが一つで、独立した汎用変換jobを選んでいない | G4以後のformat追加時 |
| `export` | XLSX、report、派生表示artifactをv1 coreから外した | G4以後 / G7以後 |
| `report` | 見た目や帳票をsemantic CLI contractから分離する | derived output backlog |
| generic `diff` | C1で承認するdiffはrequest/current state/output planと一体でなければならない | 複数state比較scenarioを承認するとき |
| generic `apply` | 旧Patchや任意state置換と誤認させず、approval必須のC1を明示するため`apply-change`に限定する | operation追加時も同じ安全契約を使う |
| `approve` | CLIは人を認証せず、approvalはhuman gateを担当するcallerが作る | Agent Skills / caller workflow |
| `cleanup` / `repair` | incomplete/corrupt setの暗黙削除・再利用を禁止する | 明示的な保守workflowを別途承認するとき |

## 現行CLIとの関係

現行CLIの次の性質は、v1契約へそのまま引き継がない。

- `ai / state / import / export / report`のscope階層
- workbook JSONとPatchを中心としたcommand grammar
- JSON primary outputと別stderr diagnosticsの具体形
- 既存output fileの暗黙overwrite
- message文字列から一部diagnostic codeを推測する実装

現行commandをいつalias、adapter、別entrypoint、廃止候補のどれにするかは`ZB-P7`で決める。G3/P4では、新commandと旧commandを同じ名前の暗黙mode切替で共存させない。

## 後続G3 taskへの引渡し

| task | この文書から渡す固定事項 | 次に決めること |
| --- | --- | --- |
| `ZB-P3.2` | option、stdin/stdout/stderr、result file、UTF-8/BOM、binary非対応 | result envelopeのfield schemaは`P3.4`へ渡す |
| `ZB-P3.3` | exclusive create、commit marker、cleanup status、四状態、fsync非保証 | diagnostic code、runtime capability、manifest bindingへ反映済み |
| `ZB-P3.4`〜`P3.6` | result/diagnostics JSON Schema、closed code catalog、statusとexit `0/1/2/3` | runtime manifestとのversion/digest結合をresult schemaへ反映済み |
| `ZB-P3.7` | Nodeを参照実装、Javaを固定済み共通契約への適合runtimeとする。正本は契約・Schema・共通fixture / golden | 決定済み |
| `ZB-P3.8` | 九件の閉じたcore profile、Node/Java extension空集合、静的capabilityと動的preflight、fallback境界 | 決定済み |
| `ZB-P3.9` | 21 workflow / harness case、31 schema / binding adversarial case、seed fixture/golden、canonical digest、command I/O/effect、Projection/state、byte/semantic/artifact/runtime-integrity比較、unknown outcome回復 | 決定済み。runnerと全semantic case materializationは`P4/P5` |
| `ZB-P3.10` | 固定名manifest、外部manifest pin、asset/source SHA-256、capability/fixture binding、Node referenceとJava conformingの関係 | 決定済み。生成・検証処理は`P4/P5/P6` |
| `ZB-P3.12` | 全command非対話、human gate提示/三結果、retryability集約、safe next action、schema外unknown outcome | 決定済み。workflow実装とtestは`P4/P6` |

## P3.1〜P3.11 review checklist

- [x] v1 scenarioを五つのproject workflow commandだけで完結できる
- [x] 各commandのlogical input、primary output、責務が重複なく説明されている
- [x] read-only、exchange artifact生成、意味変更 + project artifact生成が分類されている
- [x] `apply-change`だけが変更後project artifactを生成できる
- [x] `plan-change`とhuman gateを経ずに`apply-change`できる経路がない
- [x] `verify-artifact`がunknown outcomeをhidden stateなしに回復できる
- [x] current CLIとv1 commandを暗黙に混在させていない
- [x] deferred commandと再評価条件が明記されている
- [x] 全commandの必須option、stdin source、result destinationが一意である
- [x] `plan-change` resultを分解・再構成せず`apply-change`へ渡せる
- [x] 通常のstdoutはJSON一件、stderrはresult channel failure専用であり、message parseを要求しない
- [x] UTF-8、BOM、LF、binary/Base64非対応が明記されている
- [x] path解決、既存path拒否、artifact set内部へのresult/destination生成禁止が一意である
- [x] commit point前後のcleanup権限と`absent / incomplete / committed / corrupt`が一意である
- [x] `fsync`と電源断durabilityをv1の保証へ含めず、再開時の`verify-artifact`を定義している
- [x] result/diagnostics/status/exit codeの正本とmachine-readable schemaがリンクされている
- [x] Projection、request、diff、plan、approval、provenanceの全体schemaと別artifact間binding ruleがリンクされている
- [x] Nodeを最初の実行可能な参照実装とし、固定したNode contract releaseをJava適合の入力にする順序が明記されている
- [x] Nodeの偶発的な挙動ではなく、承認済み契約・Schema・共通fixture / goldenを正本とする
- [x] Javaもlive Node出力だけをoracleにせず、共通v1 capabilityを同じconformance corpusで検証する
- [x] core profileの部分実装やruntime固有extensionをv1適合runtimeとして扱わない
- [x] filesystem対応を静的manifest宣言だけで判断せず、destination固有preflightでfail closedにする
- [x] Node/Javaが同じchecked-in fixture、golden、case ID、comparison modeを使う
- [x] unknown outcomeをapply再試行ではなく`verify-artifact`で回復するconformance caseがある
- [x] runtimeを固定manifestと外部manifest digestから選び、newest探索、glob、PATH/vendor/network fallbackを行わない
- [x] executableとsourceのrole、basename、size、SHA-256を分離して記録する
- [x] result、output plan、provenanceがartifact/manifest digest、capability profile、fixture suite versionへ束縛される
- [x] Node reference releaseとJava conforming releaseの対象関係を固定Node manifest digestで表せる
- [x] 全workflow commandがTTY非依存で、prompt、`--yes`、会話履歴からの承認推測を行わない
- [x] `plan-change` success、human gate、approval materialize、`apply-change`の間に自動短絡経路がない
- [x] status、command、diagnostic retryabilityからretry、replan、verify、abortを一意に判断できる
