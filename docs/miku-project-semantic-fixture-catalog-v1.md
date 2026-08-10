---
title: miku-project semantic fixture catalog v1
description: Gate G1で承認されたsemantic contract v1の不変条件、境界、C1 rejectを追跡するfixture catalog。
topics:
  - miku-project
  - cli
  - specification
  - testing
category: specification
status: approved
audience:
  - maintainer
  - developer
  - agent
created: 2026-08-10
updated: 2026-08-10
sources:
  - type: local-file
    role: primary
    path: docs/miku-project-semantic-contract-v1.md
    label: semantic contract v1
    checked: 2026-08-10
  - type: local-file
    role: scenario
    path: docs/miku-project-zero-base-scenarios-v1.md
    label: G0承認済みv1利用シナリオ
    checked: 2026-08-10
---

# miku-project semantic fixture catalog v1

## 文書の位置づけ

これは `Gate G1` で承認されたsemantic contractの全不変条件とC1のreject条件を追跡するcatalogである。Gate G3の[conformance corpus v1](miku-project-conformance-corpus-v1.md)で`S-V001`、`S-I012`、`S-I020`のseed file、semantic golden、diagnostic対応を固定した。残りのIDと各行の複数変種は、ここにある期待される意味を変えずにP4/P5で実行可能なfixtureへmaterializeする。

| 期待status | 意味 |
| --- | --- |
| `valid` | v1 semantic contractを満たす。C1はhuman gate以外の前提を満たせば進める。 |
| `invalid` | required fieldまたは不変条件に違反する。R1はdiagnosticsを返し、C1は進めない。 |
| `unsupported` | 入力または操作がv1の対応範囲外である。黙って変換・破棄・成功扱いせず、C1はfail closedにする。 |

各行の「baseまたは入力差分」は、原則として`S-V001`または階層用の`S-V002`を基準にした意味上の差分である。独立した境界例は行内ですべての前提を明記する。外部ファイルの具体的な作り方はG3で選ぶ。

## catalog

| fixture ID | 種類 | baseまたは入力差分 | 検証対象 | 期待status | 関連不変条件 |
| --- | --- | --- | --- | --- | --- |
| `S-V001` | valid / C1 | `testdata/dependency.xml`。ordered forestの二root、二つのleaf task、FS・lag 0 dependency、resource、assignment、calendarを含む。C1でUID 2を`0 → 50`にする。 | R1の観測、C1の局所変更、対象外意味の保持 | `valid` | field scope、forest、参照整合性、C1 equivalence |
| `S-V002` | valid / hierarchy | `S-V001`のUID 1をsummary root、UID 2をその第一leaf childにし、UID 3の第二leaf childとUID 4の独立leaf rootを追加する。順序は`1, 2, 3, 4`とし、既存dependencyとassignmentの参照は維持する。各taskはvalidなrequired fieldを持ち、summaryの日時・進捗は宣言値として保持する。C1ではUID 2を対象にする。 | parent/child、descendant連続性、root/sibling order、summary/children同値、nested leafへのC1 | `valid` | ordered forest、summary、C1 equivalence |
| `S-B001` | boundary | validなleaf taskの`percentComplete = 0`。 | 下限の受理とC1 precondition | `valid` | integer percent 0..100 |
| `S-B002` | boundary | validなleaf taskの`percentComplete = 100`。 | 上限の受理とC1 precondition | `valid` | integer percent 0..100 |
| `S-B003` | boundary | assignmentにresource UIDを持たないunassignedを追加する。 | unassignedがsentinelではなくresourceなしとして読まれる | `valid` | assignment参照規則 |
| `S-B004` | boundary | project `currentDate`、task calendar参照、resource name/type/calendar参照、assignment resource/start/finish/units/work、calendar name/isBaseCalendarなどoptional-preserved fieldを省略する変種。 | optional fieldの欠落を許容する | `valid` | field scope、日時規則 |
| `S-B005` | boundary | requiredなproject fieldだけを持ち、task、dependency、resource、assignment、calendarの各collectionを空にし、すべてのentity参照を省略する。 | 空forestと空collection | `valid` | collection境界、参照規則 |
| `S-B006` | equivalence | 二件以上のdependency、resource、assignment、calendarを持つvalidな意味状態の対を用意し、task順と全member/fieldは同じまま、非task collectionの並び順だけを入れ替える。 | collection順序を無視したsemantic equivalence | `valid` | collection順序、C1 equivalence |
| `S-I001` | identity | task UIDを空または重複させる二変種。 | task selectorの存在と一意性 | `invalid` | task UID一意性 |
| `S-I002` | C1 reject | 存在しないtask UIDをC1 targetにする。 | targetのstable identity解決 | `invalid` | C1 selector |
| `S-I003` | hierarchy | `S-V002`を基準に、childのparentを不在UIDにする、親より前にchildを置く、別rootをparentのdescendant区間へ挟む、sibling orderを欠落・重複させる、または外部形式の疑似taskをsemantic taskとして扱う変種。 | root/parent/descendant連続性/forest順序と疑似taskの除外 | `invalid` | ordered forest、疑似task規則 |
| `S-I004` | hierarchy | `S-V002`を基準に、childを持つtaskを`summary = false`にする、またはchildを持たないtaskを`summary = true`にする。 | summaryとchildrenの同値 | `invalid` | summary整合性 |
| `S-I005` | identity | resource UID、assignment UID、calendar UIDをそれぞれ重複または空にする六変種。 | collectionごとのUID一意性 | `invalid` | resource/assignment/calendar UID一意性 |
| `S-I006` | datetime | project startをfinishより後にする。 | project日時の順序 | `invalid` | project start ≤ finish |
| `S-I007` | datetime | task startをfinishより後にする。 | task日時の順序 | `invalid` | task start ≤ finish |
| `S-I008` | required field | project、task、dependency、resource、assignment、存在するcalendar entityのrequired fieldを一つずつ省略する変種。dependency endpointの欠落と不在UIDは別変種にする。 | 全domainのrequired field欠落 | `invalid` | field scope |
| `S-I009` | duration | durationを負値にする。 | declared durationの境界 | `invalid` | nonnegative duration |
| `S-I010` | milestone | milestoneでstart/finishまたはdurationを不整合にする。 | milestoneの三条件 | `invalid` | milestone整合性 |
| `S-I011` | progress | `percentComplete = -1`。 | 整数下限外 | `invalid` | integer percent 0..100 |
| `S-I012` | progress | `percentComplete = 101`。 | 整数上限外 | `invalid` | integer percent 0..100 |
| `S-I013` | progress | `percentComplete`を小数にする。 | 整数性 | `invalid` | integer percent 0..100 |
| `S-I014` | dependency | predecessorまたはsuccessorを不在task UIDにする。 | endpoint参照整合性 | `invalid` | dependency参照規則 |
| `S-I015` | dependency | predecessorとsuccessorを同一task UIDにする。 | 自己dependency拒否 | `invalid` | dependency自己参照禁止 |
| `S-I016` | dependency | 二task以上の有向cycleを作る。 | cycle拒否 | `invalid` | dependency acyclic |
| `S-I017` | assignment | assignmentのtask UIDまたはresource UIDを不在UIDにする、または両方あるassignment start/finishを逆順にする変種。 | assignment参照整合性と日時順序 | `invalid` | assignment参照・日時規則 |
| `S-I018` | calendar | project、task、resourceのcalendar参照を不在calendar UIDにする。 | calendar参照整合性 | `invalid` | calendar参照規則 |
| `S-I019` | unsupported | FS以外のlink type、またはnonzero lagを持つdependencyの二変種。 | dependency対応範囲とfail-closed | `unsupported` | FS・lag 0のみ対応 |
| `S-I020` | unsupported | unknown field、actual、EV、baseline、timephased、extended data、calendarの勤務時間/休日/例外/work weekを一つずつ含む変種。 | 未対応dataの検出とC1 fail-closed | `unsupported` | unsupported data規則 |
| `S-I021` | C1 reject | target leaf taskの現在`percentComplete`と異なるpreconditionを渡す。 | precondition照合 | `invalid` | C1 precondition |
| `S-I022` | C1 reject | `S-V002`のvalidなsummary taskへの`percentComplete`変更要求。 | C1編集面の限定 | `unsupported` | leaf taskのみ変更可 |
| `S-I023` | C1 reject | target leaf taskの現在値と同じ`percentComplete`を新値に指定する。 | no-opを意味変更として扱わない | `invalid` | C1は別の整数への更新のみ |
| `S-I024` | value/type | identity tokenまたはnon-empty textの空値、比較不能なdatetime、boolean以外、resource typeの範囲外、負または非有限のunits/workを一つずつ与える変種。 | semantic typeと値域 | `invalid` | field scopeのsemantic type |
| `S-I025` | dependency | `S-V001`のFS・lag 0 dependencyと同じsemantic tupleをもう一件追加する。 | dependency重複を暗黙に正規化しない | `invalid` | dependency tuple一意性 |

## traceability

| semantic contractの節 | 最低限対応するfixture ID |
| --- | --- |
| v1 field scope | `S-V001`、`S-B003`〜`S-B006`、`S-I001`、`S-I005`、`S-I008`、`S-I019`、`S-I020`、`S-I024` |
| identity、順序、階層 | `S-V001`、`S-V002`、`S-B005`、`S-I001`〜`S-I005` |
| 日時、duration、進捗 | `S-B001`、`S-B002`、`S-B004`、`S-I006`〜`S-I013` |
| dependency、resource、assignment、calendar | `S-V001`、`S-B003`、`S-B005`、`S-B006`、`S-I014`〜`S-I019`、`S-I025` |
| unsupported dataのfail-closed規則 | `S-I019`、`S-I020` |
| C1の許可変更とsemantic equivalence | `S-V001`、`S-V002`、`S-B006`、`S-I002`、`S-I021`〜`S-I023`、`S-I025` |

## G3/P4への引渡し

G3は共通directory、case ID、比較mode、canonical digest、seed入力、C1前後golden、期待diagnostics、公開してよいartifactを固定した。P4/P5は各fixture IDを同じcorpusへmaterializeし、外部形式の無条件なbyte一致ではなくsemantic contractで定めた等価性を判定する。catalogのID、期待status、検証対象を変える場合は、G1 contractの改訂と再承認を要する。
