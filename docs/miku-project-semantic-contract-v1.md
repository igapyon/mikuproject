---
title: miku-project semantic contract v1
description: G0で承認されたR1/C1を実証するための最小意味、不変条件、保持範囲のG1ドラフト。
topics:
  - miku-project
  - cli
  - agent-skills
  - specification
category: specification
status: draft
audience:
  - maintainer
  - developer
  - agent
created: 2026-08-10
updated: 2026-08-10
sources:
  - type: local-file
    role: primary
    path: docs/miku-project-zero-base-scenarios-v1.md
    label: G0承認済みv1利用シナリオ
    checked: 2026-08-10
  - type: local-file
    role: primary
    path: docs/miku-project-zero-base-spec-v20260809.md
    label: ゼロベース再設計仕様
    checked: 2026-08-10
---

# miku-project semantic contract v1

## 文書の位置づけ

これは `Gate G1` のためのドラフトである。G0で承認されたR1（読み取り・理解）とC1（安全な局所変更）に必要な意味だけを定義する。入力・出力形式、中間表現、JSON schema、CLIの語彙、診断schema、詳細な損失表現は決めない。それらは `G2` と `G3` の責務である。

現行の `ProjectModel`、MS Project XML、workbook JSONは、有用な実装・検証の証拠ではあっても、この契約の正本ではない。

## G0からの固定事項

- primary actorは、人の目的と承認を受けてCLIを操作するAI Agentである。人、shell script、CIも同一契約を使う。
- R1は、計画の構造、dependency、resource、assignmentを目的別Projectionで理解可能にする。
- C1は、taskを安定したidentityで選び、`percentComplete`だけを更新する局所変更である。
- dependency、resource、assignmentは観測・保持対象であり、v1の編集面には含めない。
- 外部artifactの成功は、byte列の一致ではなく、この文書に定める意味が保たれたかで判断する。

## v1の意味の範囲

| domain | R1で観測する意味 | C1で変更できる意味 | v1での扱い |
| --- | --- | --- | --- |
| project | 名称、計画全体の開始・終了、既定calendar参照 | なし | 検証・保持対象 |
| task | UID、名称、順序・階層、開始・終了、duration、milestone、summary、`percentComplete` | UIDで選んだtaskの`percentComplete`だけ | 必須 |
| dependency | predecessorとsuccessorのtask UID、link type、lag | なし | 観測・保持対象 |
| resource | resource UID、名称 | なし | 観測・保持対象 |
| assignment | assignment UID、task UID、割当済みの場合のresource UID | なし | 観測・保持対象 |
| calendar | calendar UIDと参照関係 | なし | 参照整合性の検証・保持対象。calendar演算はしない |
| actual、EV、baseline、timephased、extended attribute | なし | なし | v1の対象外。取扱いはG2で形式ごとに決める |

この表の「保持」は、G0の代表fixtureに存在する意味を、許可された進捗変更以外で意図せず変えないという要件である。未知データを常にopaqueに保持できるという約束ではない。

## identity、順序、階層

- task UIDは、同一プロジェクト内で空でなく一意なstable identityとする。表示用のID、行番号、名称、outline numberはselectorに使わない。
- 外部形式にあるproject rootやplaceholderを表す疑似task（たとえばMS Project XMLのUID `0`）は、semantic taskではない。adapterはproject metadataへ対応づけるか明示的に除外し、C1のselectorに使わせない。対応方法の詳細はG2で形式ごとに定める。
- resource UIDとassignment UIDも、それぞれの集合内で空でなく一意とする。
- taskの意味上の構造は、順序付きの木である。外部形式のoutline levelやoutline numberは、その木へのadapter表現であり、task identityではない。
- 先行する親taskを基準に、各taskは一つの親またはrootを持つ。rootはlevel 1、子は親よりちょうど1深いlevelとする。最初のtaskがrootでないこと、levelを飛ばすこと、親のない子を作ることは不正とする。
- 同一親の子の順序は意味を持つ。adapterはこの順序を保持する。

## 日時、duration、進捗

- 日時はtimezone変換を伴わないlocal civil datetimeとして意味を扱う。timezoneの表記・変換・欠落の形式別規則はG2で定める。
- taskの開始は終了以前、durationは0以上とする。calendarに基づくduration再計算はv1では行わない。
- milestoneは開始と終了が等しく、durationが0のtaskとする。
- summaryは階層の親taskであることを意味する。summaryの開始・終了・進捗を子から再計算するかはv1では決めず、入力上の値を観測・検証対象とする。
- `percentComplete`は0から100までの整数とする。C1では、指定taskの値だけを別の同範囲の整数へ更新できる。

## dependency、resource、assignment、calendar

- dependencyは先行taskから後続taskへの有向edgeである。参照先taskがないedge、自己参照、cycleは不正とする。
- v1で必ず読取り・保持を証明するdependencyは、finish-to-startかつlag 0のedgeである。別のlink typeやlagを扱う可否・損失方針はG2で明示する。いずれもv1では編集しない。
- assignmentは既存taskを参照しなければならない。resourceを割り当てるassignmentは既存resourceを参照しなければならない。未割当はsemantic上はresourceを持たないassignmentとして表し、外部形式のsentinel値をresource UIDとして露出しない。
- taskまたはprojectのcalendar参照がある場合、参照先calendarが存在しなければならない。calendarの勤務時間、休日、例外の演算・編集は対象外とする。

## C1の許可変更と保持条件

C1の意味上の変更は次だけである。

```text
task UIDで対象を一意に選ぶ
  → 現在のpercentCompleteをpreconditionとして照合する
  → percentCompleteを0..100の整数へ更新する
```

- UIDが存在しない、重複する、またはpreconditionと現在値が一致しない変更は適用しない。
- `percentComplete`以外のtask field、taskの追加・削除・移動、dependency、resource、assignment、calendarを変更する要求はv1では未対応として扱う。
- C1後、対象外task、dependency、resource、assignment、calendar参照は意味上等価でなければならない。
- apply前後でこの文書の不変条件を検証する。後の検証に失敗した場合、次状態artifactを成功として公開しない。

## valid、invalid、boundaryの最小fixture

`G1`のfixture setは、少なくとも次の意味を例示する。実ファイル形式とgolden resultはG2/G3で決める。

| 種類 | 例 | 期待 |
| --- | --- | --- |
| valid | `dependency.xml`の二task、FS・lag 0、resource、assignment | R1で構造を観測でき、C1でUID 2の`0 → 50`だけを変更できる |
| boundary | `percentComplete`が0または100 | valid |
| invalid | `percentComplete`が-1、101、整数でない | reject |
| invalid | 重複または空のtask UID、未知UIDへの変更 | reject |
| invalid | rootなし、levelの飛び、親のない子 | reject |
| invalid | 開始が終了より後、負のduration、milestone不整合 | reject |
| invalid | 存在しないpredecessor、自己dependency、dependency cycle | reject |
| boundary | resourceを持たない未割当assignment | valid。外部形式のsentinelをsemantic resource UIDにしない |
| invalid | 存在しないtask/resourceへのassignment、存在しないcalendar参照 | reject |

## 現行資産との照合

この節は採用可否の根拠を明らかにするための現状記録であり、現行実装を契約の正本にするものではない。

| semantic contractの条件 | 現行資産の根拠 | v1実装への含意 |
| --- | --- | --- |
| `dependency.xml`の読取り、UID 2の進捗変更、XML再出力 | MS Project XML codec、Patch適用、roundtrip testが存在する | 最初のNode vertical sliceの再利用候補になる |
| dependency、assignment、calendar参照の基本的な参照整合性 | 現行validatorは欠損参照を検出する | stable code、severity、locationを持つ新しいdiagnosticsへ置換する |
| UIDの一意性、開始・終了の順序、進捗範囲 | 現行validatorは一部を検出する | 現行ではwarning扱いが混在するため、v1のreject条件として再定義する |
| `percentComplete`は整数 | 現行Patchは0..100の有限数を許し、小数も拒否しない | 新契約では整数チェックを追加する |
| ordered tree、summary整合、dependencyの自己参照・cycle | 現行validatorは完全なtree/cycle検証を持たない | v1 validatorと共有fixtureで新規に検証する |
| C1での操作を進捗だけへ限定 | 現行Patchはtask、project、dependency、assignmentなど広い編集面を持つ | 現行operationを互換面として隔離し、v1のallowlistを別途実装する |

したがって、現行coreはR1/C1の実現可能性を示すが、そのままのvalidator、Patch surface、CLI契約を採用する理由にはならない。

## G1で確認する事項

このドラフトを承認する前に、次を確認する。

1. C1を`percentComplete`更新だけに限定すること
2. dependency、resource、assignment、calendarを編集せず、観測・保持・整合性検証の対象にすること
3. task UID、順序付き階層、日時、duration、進捗の不変条件
4. v1対象外domainを、未対応のまま暗黙に失わせないこと
5. 上表のvalid/invalid/boundary例を共有fixtureへ落とすこと

G1の承認後に、G2で形式・損失・中間表現・変更要求の詳細を決定する。
