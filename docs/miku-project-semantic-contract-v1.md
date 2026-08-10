---
title: miku-project semantic contract v1
description: G0で承認されたR1/C1を実証するために、Gate G1で承認された最小意味、不変条件、保持範囲。
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

これは `Gate G1` で承認されたv1のsemantic contractである。G0で承認されたR1（読み取り・理解）とC1（安全な局所変更）に必要な意味だけを定義する。入力・出力形式、中間表現、JSON schema、CLIの語彙、診断schema、詳細な損失表現は決めない。それらは `G2` と `G3` の責務である。

現行の `ProjectModel`、MS Project XML、workbook JSONは、有用な実装・検証の証拠ではあっても、この契約の正本ではない。

## G0からの固定事項

- primary actorは、人の目的と承認を受けてCLIを操作するAI Agentである。人、shell script、CIも同一契約を使う。
- R1は、計画の構造、dependency、resource、assignmentを目的別Projectionで理解可能にする。
- C1は、taskを安定したidentityで選び、`percentComplete`だけを更新する局所変更である。
- dependency、resource、assignmentは観測・保持対象であり、v1の編集面には含めない。
- 外部artifactの成功は、byte列の一致ではなく、この文書に定める意味が保たれたかで判断する。

## v1 field scope

この節の分類は、R1/C1で扱う**意味上のfield**だけに適用する。外部形式の文字列表現、JSON schema、diagnostic code、serialization normalizationは定義しない。

| 分類 | 意味 |
| --- | --- |
| `required` | v1 artifactに必ず存在し、型・参照・不変条件を満たす。欠落または違反は`invalid`である。C1後も対象の進捗以外は保持する。 |
| `optional-preserved` | 存在しなくてもよい。存在する場合は意味を解釈し、型・参照・不変条件を満たしたままC1後も保持する。 |
| `unsupported` | v1は意味を解釈・保持保証しない。R1は存在を報告し、C1は存在するartifactを成功としてapply/exportしない。 |

この契約で使うsemantic typeは次のとおりである。外部形式での文字列表現や数値表現はG2でadapterごとに定める。

| semantic type | validな値 |
| --- | --- |
| identity token | 空でないopaqueな値。同一collection内ではadapterが復号した値の完全一致で同一性を判定する。 |
| non-empty text | 一つ以上のUnicode scalar valueを持つ文字列。trim、大小文字変換、Unicode normalizationは行わない。 |
| local civil datetime | timezone変換を伴わない日付と時刻。比較可能でなければならない。 |
| boolean | `true`または`false`。外部形式の`0 / 1`などはadapter表現である。 |
| working duration | 0以上の有限な時間量。calendarを使った再計算は行わない。 |
| units | 0以上の有限な比率。上限は設けない。 |
| resource type | `work`、`material`、`cost`のいずれか。外部形式のcodeはadapter表現である。 |
| sibling order | root間または同一parentのchildren間で全taskを一意に並べる全順序。 |

| domain | `required` | `optional-preserved` | `unsupported` |
| --- | --- | --- | --- |
| project | name（non-empty text）、start/finish（local civil datetime） | currentDate（local civil datetime）、scheduleFromStart（boolean）、calendar参照（identity token） | actual、EV、baseline、timephased、extended data、明示していないfield |
| task | UID（identity token）、name（non-empty text）、parent relation（rootはnull、それ以外はtask UID）、sibling order、start/finish（local civil datetime）、duration（working duration）、milestone/summary（boolean）、`percentComplete`（0..100の整数） | calendar参照（identity token） | actual、EV、baseline、timephased、extended data、明示していないfield |
| dependency | predecessor/successor（task UID）、type = FS、lag = 0（working duration） | なし | FS・lag 0以外のlink typeまたはlag |
| resource | UID（identity token） | name（non-empty text）、type（resource type）、calendar参照（identity token） | actual、EV、baseline、timephased、extended data、明示していないfield |
| assignment | UID（identity token）、task UID | resource UID（identity token）、start/finish（local civil datetime）、units、work（working duration） | actual、EV、baseline、timephased、extended data、明示していないfield |
| calendar | 存在するcalendar entityのUID（identity token） | name（non-empty text）、isBaseCalendar（boolean） | 勤務時間、休日、例外、work week、明示していないfield |

外部ID、行番号、outline level、outline number、外部形式のsentinelはsemantic fieldではない。adapterがordered forestや参照へ対応づけるための表現であり、C1のselectorには使わない。

## identity、順序、階層

- task構造は、複数rootを許す**ordered forest**である。各taskは`parent = null`のroot、または一つの既存taskをparentとして持つ。parentを持つtaskはparentより後に現れ、parentの全descendantは連続した区間を占める。
- taskを一つも持たない空のordered forestはR1ではvalidである。C1ではtargetを解決できないため、変更要求は`invalid`になる。
- rootと同一parentを持つtaskの順序は意味を持つ。adapterは入力のsibling orderを保持する。
- `summary = true`であることと、taskが一つ以上のchildを持つことは同値とする。不一致は`invalid`である。summary taskの日時と進捗は子から再計算しないが、C1では変更しない。
- task UIDは空でなくforest内で一意なstable identityとする。stableである範囲は、一つの入力artifactを読み、次artifactを生成する処理単位である。名称、行番号、外部ID、outline numberはselectorに使わない。
- resource UID、assignment UID、calendar UIDは、それぞれのcollection内で空でなく一意とする。calendarは存在するentityだけがUIDを持つ。
- project rootやplaceholderを表す外部形式の疑似task（たとえばMS Project XMLのUID `0`）はsemantic taskではない。adapterはproject metadataへ対応づけるか明示的に除外し、C1のselectorに使わせない。形式ごとのmappingはG2で定める。

## 日時、duration、進捗

- datetimeはtimezone変換を伴わないlocal civil datetimeとして扱う。文字列表現、timezone表記、欠落の外部形式上の表現はG2で定める。
- projectとtaskのstart/finishは`required`であり、startはfinish以前とする。欠落または逆順は`invalid`である。
- durationは入力が宣言した非負のworking durationである。calendar演算による再計算、およびstart/finishとの差との一致はv1では要求しない。欠落または負値は`invalid`である。
- milestoneはstartとfinishが等しく、durationが0のtaskとする。不一致は`invalid`である。
- currentDate、assignment start/finishは`optional-preserved`である。assignmentの両方がある場合はstartがfinish以前とする。片方または両方の欠落は許容する。
- `percentComplete`は0から100までの整数とする。小数、-1、101、欠落は`invalid`である。unitsとworkは上表のsemantic typeを満たさなければならない。

## dependency、resource、assignment、calendar

- dependencyはpredecessor taskからsuccessor taskへの有向edgeである。v1で正式対応するedgeはfinish-to-start（FS）かつlag 0だけである。別のlink typeまたは非zero lagは`unsupported`であり、暗黙変換・無視・成功扱いをしない。
- dependencyのsemantic identityは`(predecessor UID, successor UID, type, lag)`のtupleであり、dependency collectionはこのtupleの集合とする。同じtupleの重複は`invalid`であり、暗黙に一件へ正規化しない。collection内の並び順は意味を持たない。
- predecessor/successorの欠損参照、自己参照、二task以上のcycleは`invalid`である。dependency編集はC1に含めない。
- assignmentは既存taskを参照しなければならない。resource UIDがあるassignmentは既存resourceを参照しなければならない。resource UIDがないunassigned assignmentはvalidである。外部形式のsentinel値（たとえば`-65535`）をsemantic resource UIDとして露出しない。
- project、task、resourceのcalendar参照がある場合、参照先calendarが存在しなければならない。calendarの勤務時間、休日、例外、work weekの演算・編集はv1の対象外である。
- dependency、resource、assignment、calendarの各collectionは空でもvalidである。ただし、存在しないentityへの参照はinvalidである。外部形式でcontainerを省略するか空で表すかはG2で定める。
- resource、assignment、calendarのcollection内の並び順は意味を持たない。UIDで識別するmemberと、存在していた`required`および`optional-preserved` fieldを保持する。

## unsupported dataのfail-closed規則

- unknown field、actual、EV、baseline、timephased、extended data、未対応dependency、未対応calendar内容はopaque preservationの対象にしない。
- R1の`inspect`はunsupported dataの存在と範囲を報告できる。必要なら部分観測を返してよいが、それをvalidなProjection成功成果物として公開しない。
- unsupported dataを含むartifactは、C1のvalidate、diff、apply、exportを成功させない。入力を変更せず、unsupportedであることを構造化diagnosticsとして返す。

## C1の許可変更とsemantic equivalence

C1の意味上の変更は次だけである。

```text
stable UIDでleaf taskを一意に選ぶ
  → 現在のpercentCompleteをpreconditionとして照合する
  → percentCompleteを0..100の別の整数へ更新する
```

- 対象UIDの不存在・重複、summary task、現在値不一致、新値と現在値の一致、整数範囲外、unsupported dataの存在は`apply`しない理由である。同値更新は意味変更ではないため`invalid`とし、human gateへ進めない。
- `percentComplete`以外のtask field、taskの追加・削除・移動、dependency、resource、assignment、calendar、projectを変更する要求はv1では未対応とする。
- apply前後で不変条件を検証する。対象taskの`percentComplete`以外の`required` fieldと、存在していた`optional-preserved` fieldはsemantic equivalentでなければならない。後の検証に失敗した場合、次状態artifactを成功として公開しない。
- semantic equivalenceではtaskのroot順とsibling orderを比較する。dependency、resource、assignment、calendarはcollection順を比較せず、semantic identityで対応づけたmemberの集合と各fieldを比較する。
- byte列、XML要素順、空白、外部ID、outline表記などadapterのserialization差はsemantic equivalenceに含めない。normalizationの詳細はG2で定める。

## semantic fixture catalog

valid、invalid、boundary、unsupported、C1 rejectの全例は [semantic fixture catalog](miku-project-semantic-fixture-catalog-v1.md) でIDを付けて管理する。この段階では製品用IRやJSON schemaを作らない。各IDを実行可能なconformance fixtureへ実装するのはG3の責務である。

## 現行資産との照合

この節は採用可否の根拠を明らかにするための現状記録であり、現行実装を契約の正本にするものではない。

| semantic contractの条件 | 現行資産の根拠 | v1実装への含意 |
| --- | --- | --- |
| `dependency.xml`の読取り、UID 2の進捗変更、XML再出力 | MS Project XML codec、Patch適用、roundtrip testが存在する | 最初のNode vertical sliceの再利用候補になる |
| dependency、assignment、calendar参照の基本的な参照整合性 | 現行validatorは欠損参照を検出する | stable code、severity、locationを持つ新しいdiagnosticsへ置換する |
| UIDの一意性、開始・終了の順序、進捗範囲 | 現行validatorは一部を検出する | 現行ではwarning扱いが混在するため、v1のreject条件として再定義する |
| `percentComplete`は整数 | 現行Patchは0..100の有限数を許し、小数も拒否しない | 新契約では整数チェックを追加する |
| ordered forest、summary整合、dependencyの自己参照・cycle | 現行validatorは完全なforest/cycle検証を持たない | v1 validatorと共有fixtureで新規に検証する |
| C1での操作を進捗だけへ限定 | 現行Patchはtask、project、dependency、assignmentなど広い編集面を持つ | 現行operationを互換面として隔離し、v1のallowlistを別途実装する |

したがって、現行coreはR1/C1の実現可能性を示すが、そのままのvalidator、Patch surface、CLI契約を採用する理由にはならない。

## G1で確認する事項

このドラフトを承認する前に、次を確認する。レビュー結果は `ZB-P1.8` の完了記録とする。

| 確認項目 | この文書での根拠 | fixture catalogでの根拠 | review結果 |
| --- | --- | --- | --- |
| ordered forest、summary、stable identity、疑似taskの規則 | 「identity、順序、階層」 | `S-V001`、`S-V002`、`S-B005`、`S-I001`〜`S-I005` | レビュー済み（2026-08-10） |
| fieldのsemantic typeと値域 | 「v1 field scope」 | `S-B004`、`S-I001`、`S-I005`、`S-I008`、`S-I024` | レビュー済み（2026-08-10） |
| project/taskの日時、duration、milestone、整数進捗 | 「日時、duration、進捗」 | `S-B001`、`S-B002`、`S-B004`、`S-I006`〜`S-I013`、`S-I024` | レビュー済み（2026-08-10） |
| FS・lag 0 dependency、tuple重複、参照整合性 | 「dependency、resource、assignment、calendar」 | `S-V001`、`S-I014`〜`S-I019`、`S-I025` | レビュー済み（2026-08-10） |
| resource、assignment、calendar、unassigned、空collection、collection順序の保持境界 | 「v1 field scope」「dependency、resource、assignment、calendar」 | `S-V001`、`S-B003`、`S-B005`、`S-B006`、`S-I005`、`S-I017`、`S-I018` | レビュー済み（2026-08-10） |
| unsupported dataを黙って捨てず、C1をfail closedにする | 「unsupported dataのfail-closed規則」 | `S-I019`、`S-I020` | レビュー済み（2026-08-10） |
| C1をleaf task一件の進捗更新に限定し、collection意味を含む保持条件を検証する | 「C1の許可変更とsemantic equivalence」 | `S-V001`、`S-V002`、`S-B006`、`S-I002`、`S-I021`〜`S-I023`、`S-I025` | レビュー済み（2026-08-10） |
| G1に外部形式、IR、schema、diagnostics、serializationの決定を混入させない | 「文書の位置づけ」および各節のG2への委譲 | fixtureの実ファイル・golden resultをG3へ委譲 | レビュー済み（2026-08-10） |

`ZB-P1.8`の文書横断reviewは2026-08-10に完了した。`Gate G1` は同日に承認され、このsemantic contractとcatalogはv1の意味契約として確定した。続くG2で、形式・損失・中間表現・変更要求の詳細を決定する。
