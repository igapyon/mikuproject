---
title: miku-project format and loss contract v1
description: Gate G2で承認された、v1 artifactの役割、形式、schema version、変換可能範囲、損失規則。
topics:
  - miku-project
  - cli
  - agent-skills
  - specification
  - data-format
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
    label: Gate G1承認済みsemantic contract
    checked: 2026-08-10
  - type: local-file
    role: scenario
    path: docs/miku-project-zero-base-scenarios-v1.md
    label: G0承認済みR1/C1シナリオ
    checked: 2026-08-10
  - type: local-file
    role: current-state
    path: docs/miku-project-current-capability-matrix-v20260810.md
    label: 現行形式・CLIの証拠
    checked: 2026-08-10
  - type: web
    role: primary-reference
    url: https://learn.microsoft.com/en-us/office-project/xml-data-interchange/predecessorlink-element
    label: Microsoft Project XML PredecessorLink element
    checked: 2026-08-10
  - type: web
    role: primary-reference
    url: https://learn.microsoft.com/en-us/office-project/xml-data-interchange/durationformat-element
    label: Microsoft Project XML duration format values
    checked: 2026-08-10
  - type: web
    role: primary-reference
    url: https://learn.microsoft.com/en-us/office-project/xml-data-interchange/type-element-multiple-parents
    label: Microsoft Project XML predecessor/resource Type values
    checked: 2026-08-10
---

# miku-project format and loss contract v1

## 文書の位置づけ

これは `Gate G2` で承認されたformat and loss contractである。[semantic contract v1](miku-project-semantic-contract-v1.md) の意味を、R1/C1で受け渡すartifact、形式、version、損失規則へ対応づける。G3で確定したcommand、I/O、encoding/publication境界は [CLI contract v1](miku-project-cli-contract-v1.md)、JSON result/diagnostics/statusは [CLI result and diagnostics contract v1](miku-project-cli-result-contract-v1.md) に従う。semantic stateとexchange artifactのmachine-readableな全体構造は[artifact JSON Schema v1](schemas/miku-project-artifacts-v1.schema.json)、digest用canonical JSONは [conformance corpus v1](miku-project-conformance-corpus-v1.md) に定める。

## G2の決定

- v1は `miku_project_semantic_state/v1` を採用する。これはNode/Java共通の**internal-only IR**であり、意味検証、semantic diff、conformance fixtureの基準である。
- このIRは、利用者の長期保存形式、AIへの標準hand-off、既存外部artifactの置換物にはしない。CLIが通常利用者へ全stateを渡す代わりに、purpose別Projectionを出力する。
- v1の外部プロジェクト形式は、MS Project XML vocabularyの明示的なsupported subsetである `miku-project-ms-project-xml-subset/v1` だけを `read / write` 対象とする。この文書で「MS Project XML」と略記するときはこのprofileを指す。これはv1の唯一の外部入力例・出力先であって、semantic contractの正本ではない。
- `mikuproject_workbook_json`、XLSX、帳票、SVG、Markdown、Mermaidはv1 coreのformat contractから外す。現行互換資産として保持するが、v1の成功成果物・往復保証・AI hand-offとして約束しない。
- 変更は全state置換ではなく、versioned change requestに含めた許可operationだけで表す。具体的な変更安全性は [change contract v1](miku-project-change-contract-v1.md) に定める。
- C1の成功出力は、既存pathを置換する単一XMLではなく、新規directoryとして公開するartifact setとする。set内の`project.xml`と`provenance.json`を完成・再検証した後、空の`COMMITTED`を排他的に新規作成する。`COMMITTED`を持ち、memberとdigestの検証を通るsetだけを公開済みと認める。

## artifactの役割と寿命

| artifact | format / schema version | 役割 | 生成者・消費者 | 寿命 | v1での位置 |
| --- | --- | --- | --- | --- | --- |
| external project artifact | `ms-project-xml` + `miku-project-ms-project-xml-subset/v1` + `ms-project-xml-adapter/v1` | 利用者が持つ入力、またはartifact set内の次状態XML | 人・Agent・CLI | caller-managed | `read / write` |
| semantic state | `miku_project_semantic_state/v1` | G1意味を表すIR、検証・比較・変換の内部基準 | CLI runtime、conformance tests | operation中のtransient。明示debug/test出力以外は公開しない | internal-only |
| Projection | `miku_project_projection/v1` | 人・AI AgentがR1で理解・変更意図を作るための限定view | CLI → 人・Agent | request/approvalの間だけ保持可。stateの代替保存ではない | exchange output |
| change request | `miku_project_change_request/v1` | C1で許可された操作とprecondition | 人・Agent → CLI | diffとapprovalに束縛されるまで保持可 | exchange input |
| semantic diff | `miku_project_semantic_diff/v1` | 予定される意味変更、人の承認判断 | CLI → 人・Agent | requestごとの短期artifact。監査用保存はcaller判断 | exchange output |
| output plan | `miku_project_output_plan/v1` | 出力形式、公開先、preflight結果をhuman gateへ提示する | CLI → 人・Agent | request/diffと同じ承認単位 | exchange output |
| change approval | `miku_project_change_approval/v1` | human gateが済んだstate/request/diff/output planへの明示的な束縛 | caller → CLI | 束縛したtupleと完全一致する間だけ有効。単回使用を主張しない | exchange input |
| provenance record | `miku_project_provenance/v1` | 出力の入力・変換・loss/normalizationを追跡 | CLI → caller | artifact setと同じ | `provenance.json` |
| published artifact set | `miku_project_artifact_set/v1` | C1の論理的な公開単位 | CLI → caller | caller-managed | 新規directory。`COMMITTED`がcommit marker |

`semantic state`はG3のpublic CLIでファイルとしてmaterializeしない。debug/conformance専用artifactを必要とする場合は通常command surfaceと分離してG3のfixture設計で決める。ただし、state schemaとそのsemantic digestはNode/Javaで共通でなければならない。成功した`miku_project_artifact_set/v1`のdirectoryには、固定名の`project.xml`、`provenance.json`、空fileの`COMMITTED`だけを置く。v1のC1は既存directoryの置換を許可しない。

artifact setの状態は次のとおりである。

| 状態 | 判定 | 利用可否 |
| --- | --- | --- |
| absent | destination directoryが存在しない | artifactなし |
| incomplete | directoryは存在するが`COMMITTED`がない | 利用禁止。失敗または中断した生成途中 |
| committed | 三memberだけがsymlinkではない通常fileとして存在し、`COMMITTED`が0 byteで、provenance schema、XML profile、全digestが一致する | 利用可能 |
| corrupt | path entryが存在するが通常directoryでない、または`COMMITTED`はあるがmember数・type、schema、XML、digestのいずれかが不正 | 利用禁止。自動修復・部分利用しない |

directoryの存在だけを公開判定に使ってはならない。Agent、CLI、他のconsumerは`COMMITTED`、provenance、digestをすべて検証する。incomplete/corrupt directoryをProjection、次のC1入力、成功artifactとして扱わない。

## semantic stateの論理schema

`miku_project_semantic_state/v1` は次のlogical structureを持つ。これはinternal IRのschemaであり、外部XML、Projection、change requestのschemaとは別である。

```json
{
  "kind": "miku_project_semantic_state",
  "schema_version": "1",
  "semantic_contract_version": "1",
  "project": {
    "name": "Dependency Project",
    "start": "<local-civil-datetime>",
    "finish": "<local-civil-datetime>",
    "current_date": "<local-civil-datetime>",
    "schedule_from_start": true,
    "calendar_uid": "1"
  },
  "tasks": [
    {
      "uid": "1",
      "name": "Prepare",
      "parent_uid": null,
      "start": "<local-civil-datetime>",
      "finish": "<local-civil-datetime>",
      "duration": "<working-duration>",
      "milestone": false,
      "summary": false,
      "percent_complete": 100
    },
    {
      "uid": "2",
      "name": "Execute",
      "parent_uid": null,
      "start": "<local-civil-datetime>",
      "finish": "<local-civil-datetime>",
      "duration": "<working-duration>",
      "milestone": false,
      "summary": false,
      "percent_complete": 0
    }
  ],
  "dependencies": [
    { "predecessor_uid": "1", "successor_uid": "2", "type": "FS", "lag": "<zero-working-duration>" }
  ],
  "resources": [{ "uid": "1", "name": "Miku", "type": "work" }],
  "assignments": [{ "uid": "1", "task_uid": "2", "resource_uid": "1" }],
  "calendars": [{ "uid": "1", "name": "Standard", "is_base_calendar": true }]
}
```

- `project`、各taskのG1 required field、dependency endpoint/type/lag、resource UID、assignment UID/task UID、calendar entity UIDは必須である。optional-preserved fieldは存在するときだけ書き、`null`を代用しない。例外としてroot taskの`parent_uid`は必須の`null`である。
- `tasks`はroot/sibling orderを表すpreorder arrayであり、この配列順は意味を持つ。`dependencies`はtuple集合、`resources`、`assignments`、`calendars`はUID集合であり、空arrayを許す。
- semantic equivalenceではnon-task collectionの入力順を比較しない。digest用canonical serializationでは、dependencyをsemantic tuple、他のcollectionをUIDで決定的に整列する。taskの順序は整列で変更してはならない。
- 各scalarのfixture/golden JSON表現とcanonical byte serializationは[conformance corpus v1](miku-project-conformance-corpus-v1.md)に固定する。ここで示したfield名、構造、null/省略規則、collectionの意味はG2の契約である。
- state digestにはこのsemantic stateだけを含める。source path、XML字句、timestamp、runtime、diagnostics、provenanceは含めない。

## format support matrix

| format / artifact | read | write | semantic roundtrip | loss / unsupportedの扱い | v1の約束 |
| --- | --- | --- | --- | --- | --- |
| MS Project XML supported subset `miku-project-ms-project-xml-subset/v1` | ○ | ○ | ○。byte一致は不要 | G1の`unsupported` dataを含む入力は成功変換しない。adapter表現はnormalization可 | R1/C1の外部入出力 |
| `miku_project_semantic_state/v1` | runtime内部のみ | runtime内部のみ | ○。G1のsemantic equivalenceで判定 | unsupported dataを持てない | Node/Java・fixtureの共通IR |
| `miku_project_projection/v1` | stateの復元には使わない | ○ | 該当しない。意図的なscope限定 | scope外dataはlossではなく非公開 | R1のAgent/human hand-off |
| `miku_project_change_request/v1` | ○ | 人・Agentが作成可 | 該当しない | allowlist外field/operationはerror | C1入力 |
| `miku_project_semantic_diff/v1` | approval bindingのために読む | ○ | 該当しない | raw format差を含めない | human gateの判断材料 |
| `miku_project_output_plan/v1` | approval bindingのために読む | ○ | 該当しない | loss/unsupportedが非空なら承認対象にしない | human gateの出力判断材料 |
| `miku_project_change_approval/v1` | ○ | callerが作成 | 該当しない | request/diff/output plan/state digest不一致はerror | applyの明示許可 |
| `miku_project_provenance/v1` | callerが監査時に読む | ○ | 該当しない | loss/unsupportedを隠さない | artifact set内の必須file |
| `miku_project_artifact_set/v1` | callerが利用 | CLIが新規directoryを排他的に予約し、commit markerを最後に作成 | set内XMLを再読取り可能 | markerなしはincomplete、検証不一致はcorrupt | C1の成功出力 |
| `mikuproject_workbook_json` | v1 coreでは約束しない | v1 coreでは約束しない | なし | 現行互換はG7まで別系統で維持 | legacy evidence / compatibility |
| XLSX、派生report | v1 coreでは約束しない | v1 coreでは約束しない | なし | v1の成功判定に使わない | deferred |

## MS Project XML adapterの契約

`ms-project-xml-adapter/v1` は `miku-project-ms-project-xml-subset/v1` とsemantic stateの間だけを担当する。

このprofileはMicrosoft Project XML Data Interchangeのelement vocabularyを使うが、Microsoft Projectの全exportを受理するprofileでも、`mspdi_pj12.xsd`など特定世代の完全なXSD conformanceを主張するprofileでもない。採用するのは下表に列挙したfield、unversioned namespace、lexical規則、canonical output順だけである。別namespace、必須metadataを含む特定世代XSDへの適合、Microsoft Project application各versionでのimport互換性は、G3以降に別fixtureとcapabilityとして実証されない限りv1の約束に含めない。Microsoftの特定世代schemaがProject childを`xsd:sequence`として定義していることは、[Microsoft Project element XML schema](https://learn.microsoft.com/en-us/office-project/xml-data-interchange/xml-schema-for-the-project-element)を設計根拠とする。

- rootはnamespace `http://schemas.microsoft.com/project` の`Project`一件とする。namespace宣言、comments、processing instructions、要素間whitespace以外の、下表にないelement/attributeは`unsupported-error`とする。
- 同じparent直下のsingleton elementが重複する場合は`invalid`とする。`Tasks`、`Resources`、`Assignments`、`Calendars` containerの省略は空collectionへdecodeする。containerが存在する場合は一件以上の対応memberを必須とし、空containerは`invalid`とする。
- decodeは下表のG1 required/optional-preserved field、task forest、dependency、resource、assignment、calendar参照をsemantic stateへ対応づける。
- XMLの`ID`、outline表現、疑似task、`0 / 1` boolean、日時・durationの字句表現などはadapter表現であり、semantic stateにはG1で定めた意味だけを渡す。
- decode時にunknown field、actual、EV、baseline、timephased、extended data、未対応dependency、未対応calendar内容を検出した場合、R1はその存在を報告できるが、validate、Projection成功、C1、publishを成功させない。
- decodeでは、同じparent内のsupported child elementの順序を意味として扱わない。encodeはsemantic stateに存在するrequired/optional-preserved meaningだけを、後述のcanonical output順でXMLへ書く。入力の要素順、空白、namespace prefix、日時・durationの字句、外部ID・outlineの表現はnormalizationしてよい。
- raw XMLにしかなかったunknown dataをopaqueに再出力する保証はしない。したがって、unknown dataを含むXMLをdecodeしてから成功出力する経路をv1は持たない。
- XMLへのwrite後は、再decodeしたsemantic stateがwrite前stateとsemantic equivalentであることを確認する。確認に失敗すれば外部artifactを成功として公開しない。

G3のtext transportはUTF-8だけを許可する。入力XMLはBOMなし、または先頭のUTF-8 BOM一件を許可し、BOM除去を`text.utf8-bom-removed` normalizationとして記録する。XML declarationにencodingがあれば`UTF-8`だけを許可する。出力XMLはUTF-8 BOMなし、XML declarationのencoding `UTF-8`、LF、末尾LF一件へ固定する。JSON input/result/provenanceはUTF-8 BOMを拒否し、出力はBOMなし・LF・末尾LF一件とする。詳細は [CLI contract v1](miku-project-cli-contract-v1.md#textencodingbinary) に従う。

### XML lexical profile

| semantic type | v1 input lexical | v1 output lexical | 不一致 |
| --- | --- | --- | --- |
| identity token | `0`または先頭zeroを持たない非負base-10整数。例外はassignmentの`ResourceUID=-65535`だけ | 同じ整数文字列 | `invalid`。未対応sentinelは`unsupported-error` |
| non-empty text | XML entity展開後のtextをtrim/Unicode normalizationせず使う | 同じUnicode scalar sequenceをXML escapeする | 空なら`invalid` |
| local civil datetime | `YYYY-MM-DDTHH:mm:ss`。timezone suffixと小数秒はv1非対応 | 同じ秒精度形式 | 値不正は`invalid`、対応外字句は`unsupported-error` |
| boolean | `0`または`1` | `0`または`1` | その他は`invalid` |
| working duration | `PT{H}H{M}M{S}S`。H/M/Sは非負整数、M/Sは0..59 | total durationを同形式へcanonicalize | 値不正は`invalid`、別ISO 8601形は`unsupported-error` |
| percent | `0`から`100`のbase-10整数 | 同じ整数 | 小数・範囲外は`invalid` |
| units | 0以上の有限base-10 decimal | 不要な末尾zeroを持たないdecimal | 値不正は`invalid` |
| dependency type | `1`だけをFSへdecode | FSを`1`へencode | 他codeは`unsupported-error` |
| dependency lag | 正式形は`LinkLag=0`と`LagFormat=3`。現行fixture互換として`LinkLag=PT0H0M0S`かつ`LagFormat`欠落もzeroへdecode可 | `LinkLag=0`と`LagFormat=3` | 非zero値・他のlegacy durationは`unsupported-error` |
| resource type | `0` = material、`1` = work、`2` = cost | semantic typeに対応するcode | 他codeは`unsupported-error` |

dependency/resourceのType codeは[MicrosoftのType定義](https://learn.microsoft.com/en-us/office-project/xml-data-interchange/type-element-multiple-parents)に従う。`LinkLag`は10分の1分単位の整数で、`LagFormat`を伴う。[MicrosoftのPredecessorLink定義](https://learn.microsoft.com/en-us/office-project/xml-data-interchange/predecessorlink-element)と[duration format値](https://learn.microsoft.com/en-us/office-project/xml-data-interchange/durationformat-element)に従い、v1出力はzero lagを`LinkLag=0`、minutesを表す`LagFormat=3`へ固定する。現行`testdata/dependency.xml`の`PT0H0M0S`は既存資産をR1/C1 fixtureとして継続利用するためのread-only compatibility lexicalであり、出力では必ず正式形へnormalizationし、provenanceへ`legacy-linklag-duration-to-tenths-minute`を記録する。

### XML field mapping

XML pathはnamespaceを省略して表す。`?`はoptional-preserved、`*`はcollection memberである。

| XML path | semantic path / role | read / write規則 | preservation |
| --- | --- | --- | --- |
| `Project/Name` | `project.name` | required singleton | `required` |
| `Project/StartDate` / `FinishDate` | `project.start` / `finish` | required singleton | `required` |
| `Project/CurrentDate?` | `project.current_date` | 存在時decode/write | `optional-preserved` |
| `Project/ScheduleFromStart?` | `project.schedule_from_start` | 存在時decode/write | `optional-preserved` |
| `Project/CalendarUID?` | `project.calendar_uid` | 存在時decode/writeし、calendar参照を検証 | `optional-preserved` |
| `Project/Tasks/Task*` | `tasks[]` | document orderをpreorderとして使う | `required container meaning` |
| `Task/UID` / `Name` | `task.uid` / `name` | required singleton | `required` |
| `Task/OutlineLevel` | `task.parent_uid`とsibling order | required。1はroot、N>1は直前のlevel N-1 taskをparentとする | `normalized adapter representation` |
| `Task/ID?` / `OutlineNumber?` | semantic fieldなし | 読取り時はidentity/selectorに使わず、出力時にpreorderから再生成 | `normalized adapter representation` |
| `Task/Start` / `Finish` / `Duration` | 同名のtask semantic field | required singleton | `required` |
| `Task/Milestone` / `Summary` / `PercentComplete` | 同名のtask semantic field | required singleton | `required` |
| `Task/CalendarUID?` | `task.calendar_uid` | 存在時decode/writeし、calendar参照を検証 | `optional-preserved` |
| `Task/PredecessorLink*` | `dependencies[]` | successorはcontainerのtask UID。下記三fieldを必須とする | `required when present` |
| `PredecessorLink/PredecessorUID` | `dependency.predecessor_uid` | required task参照 | `required` |
| `PredecessorLink/Type` | `dependency.type` | `1`だけをFSとして扱う | `required` |
| `PredecessorLink/LinkLag` | `dependency.lag` | 正式な整数`0`、または現行fixture互換の`PT0H0M0S`だけをzeroへdecode | `required / normalized compatibility` |
| `PredecessorLink/LagFormat` | dependency lagのadapter表現 | 正式形ではminutesを表す`3`を必須とする。legacy互換形では欠落を必須とする | `normalized adapter representation` |
| `Project/Resources/Resource*` | `resources[]` | document orderは意味を持たない | `collection` |
| `Resource/UID` | `resource.uid` | required singleton | `required` |
| `Resource/Name?` / `Type?` / `CalendarUID?` | 同名のresource optional field | 存在時decode/write。calendar参照を検証 | `optional-preserved` |
| `Resource/ID?` | semantic fieldなし | selectorに使わず、出力時にUID sortから再生成 | `normalized adapter representation` |
| `Project/Assignments/Assignment*` | `assignments[]` | document orderは意味を持たない | `collection` |
| `Assignment/UID` / `TaskUID` | 同名のassignment semantic field | required singleton。task参照を検証 | `required` |
| `Assignment/ResourceUID?` | `assignment.resource_uid` | 非負UIDはresource参照。`-65535`はfield欠落へdecode | `optional-preserved / normalized sentinel` |
| `Assignment/Start?` / `Finish?` / `Units?` / `Work?` | 同名のassignment optional field | 存在時decode/write | `optional-preserved` |
| `Project/Calendars/Calendar*` | `calendars[]` | document orderは意味を持たない | `collection` |
| `Calendar/UID` | `calendar.uid` | required singleton | `required` |
| `Calendar/Name?` / `IsBaseCalendar?` | 同名のcalendar optional field | 存在時decode/write | `optional-preserved` |

### hierarchy、pseudo task、unknown data

- semantic taskは`Tasks/Task`のdocument orderで処理する。`OutlineLevel`の先頭は1、増加は一度に1まで、減少は既存ancestor levelまでとする。違反は`invalid`である。levelと実際のpreorderからparentとsibling orderを一意に導出する。
- `OutlineNumber`が存在する場合は、導出したroot/sibling positionを1始まりのdot区切りで表した値と一致しなければ`invalid`である。`ID`が存在する場合は正の整数かつtask間で一意でなければ`invalid`である。
- `UID=0`のTaskは、一件だけ、Tasksの先頭、`OutlineLevel=0`、`Summary=1`の場合にproject summary pseudo taskとして除外する。このtaskをsemantic state、selector、dependency endpointへ露出しない。条件不一致または二件目のUID 0は`invalid`である。
- pseudo task以外のtask UID 0、dependency endpoint UID 0、resource/calendar/assignment UID 0は`invalid`である。
- `Calendar`の勤務時間、休日、例外、work week、actual、EV、baseline、timephased、extended data、および上表にないProject/Task/Resource/Assignment/Calendarのchild elementは`unsupported-error`である。XML comment、processing instruction、要素間whitespaceだけは意味を持たず無視する。
- encodeのcanonical Project child順は、存在するものだけを `Name`、`ScheduleFromStart`、`StartDate`、`FinishDate`、`CalendarUID`、`CurrentDate`、`Calendars`、`Tasks`、`Resources`、`Assignments` の順とする。これはMicrosoft Project schemaにおけるsupported element間の相対順を保つ。空collectionのcontainerは省略する。
- `Task` child順は `UID`、`ID`、`Name`、`OutlineLevel`、`OutlineNumber`、`Start`、`Finish`、`Duration`、`Milestone`、`Summary`、`PercentComplete`、`CalendarUID`、`PredecessorLink` とする。`ID`はpseudo taskを含めずsemantic preorderの1始まり連番、`OutlineNumber`はroot/sibling positionの1始まりdot区切りで必ず出力する。project summary pseudo taskは出力しない。`PredecessorLink` child順は `PredecessorUID`、`Type`、`LinkLag`、`LagFormat` とする。
- `Resource` child順は `UID`、`ID`、`Name`、`Type`、`CalendarUID`、`Assignment` child順は `UID`、`TaskUID`、`ResourceUID`、`Start`、`Finish`、`Units`、`Work`、`Calendar` child順は `UID`、`Name`、`IsBaseCalendar` とする。optional fieldは欠落時に省略する。non-task collectionはUIDのbase-10数値順、dependencyはsemantic tuple順、taskだけはsemantic preorderを保持する。

## field / loss matrix

| transformation | required | optional-preserved | adapter表現 | unsupported | 成功条件 |
| --- | --- | --- | --- | --- | --- |
| XML → semantic state | decodeして検証 | 存在すればdecodeして保持 | XML字句からsemantic typeへnormalization | `unsupported-error` | valid stateとprovenanceを得る |
| semantic state → XML | semantic equivalentに出力 | 存在していた値をsemantic equivalentに出力 | XML字句・要素順・outline等はnormalization可 | stateに存在不可 | 再decode後にsemantic equivalent |
| semantic state → Projection | purposeが要求する範囲だけ出力 | 同左 | JSONのfield順は非意味 | scope外は非公開。lossではない | source digest・scope・capabilityを添える |
| change request → planned state | requestのoperationだけ変更 | 変更しない | schema字句はG3で決定 | field/operationがallowlist外ならerror | pre/post validateとdigest一致 |
| before/after state → semantic diff | 意味上の差だけ記録 | 同左 | JSONのfield順は非意味 | diff対象stateがunsupportedならerror | raw XML差を混ぜない |

### preservation class

| class | v1の意味 | 例 |
| --- | --- | --- |
| `required` | 成功した変換・C1後にsemantic equivalentでなければならない | task UID、階層、FS/lag 0 dependency、target以外の`percentComplete` |
| `normalized` | adapterの字句・container・外部表現だけが変わってよい | XML日時表記、booleanの`0 / 1`、outline表現、要素順 |
| `lossy-with-warning` | v1では成功変換に使わない予約class。採用にはG2改訂が必要 | 将来のXLSX限定列など |
| `unsupported-error` | 読取りで検出し、成功したProjection/apply/publishを出さない | actual、baseline、unknown field、non-FS dependency |
| `opaque-preserved` | v1では使用しない。opaque復元の約束はしない | raw XMLの未解釈要素 |

`ignored change` はpreservation classではない。change requestの未知field、allowlist外operation、未知kindは、warningやno-opへ落とさず`unsupported-error`または`invalid`にする。

## Projection contract

Projectionはsemantic stateの全量置換物ではなく、目的・範囲・根拠を明示したread-only artifactである。すべてのProjectionに次を含める。

- `kind = "miku_project_projection"` と `schema_version = "1"`
- `semantic_contract_version = "1"`、`source_state_digest`、生成日時ではないdeterministicな`scope`
- `purpose`、対象task UID、含めたdomain、意図的に含めなかったdomain
- 編集判断用Projectionでは、v1で許可されるchange operationとchange requestの必要precondition
- unsupported dataの有無。存在する場合は成功した編集用Projectionにしない

v1のpurposeは次の二つだけとする。

| purpose | scope | 必ず含める意味 | 含めないもの |
| --- | --- | --- | --- |
| `project_overview` | project全体の構造理解 | project required field、全taskのUID/name/parent/order/summary/進捗、FS dependency、capability | taskごとの編集用詳細、全raw XML。diagnosticsはresult envelopeだけに置く |
| `task_change_context` | 指定したleaf task一件のC1判断 | targetの全G1 field、ancestor chain、前後dependency、target assignment/resource、current progress、source digest、許可operation | 他taskの不要な詳細、全state、自由編集field |

`task_change_context`を作る対象はvalidなleaf taskだけである。scopeを狭めることはdata lossではないが、Projectionからsemantic stateを復元しようとしてはならない。

Projectionはartifact schemaへ適合するだけでは十分でない。生成元semantic stateをcollection canonicalizationしたcanonical JSONのSHA-256が`source_state_digest`と一致し、purposeごとに次のcontent bindingを満たすことを共通validatorで検査する。

- `project_overview`は`target_task_uid = null`と固定scopeを持ち、projectを完全一致、semantic preorder上の全taskをoverview fieldへ射影し、全dependencyを過不足なく含める。taskの`order`はpreorderの0始まりindexである。
- `task_change_context`はscopeのtarget UIDと`target_task.uid`が一致し、そのtaskが生成元stateに存在するleafでなければならない。projectとtargetの全G1 field、rootからparentまでのancestor chain、targetをpredecessorまたはsuccessorに持つ全dependency、targetの全assignment、そのassignmentが参照する全resourceを過不足なく含める。
- `capability.unsupported_data`は入力検査で得た構造化結果と一致させる。unsupported dataを検出した入力から、成功した編集用Projectionを生成しない。
- scope/content/source digestのどれかが一致しないProjectionは、schema-validでも契約違反として拒否する。この検査は[result contractの`RB-012`](miku-project-cli-result-contract-v1.md#artifact-schema%E3%81%A8cross-artifact-binding)で固定する。

```json
{
  "kind": "miku_project_projection",
  "schema_version": "1",
  "semantic_contract_version": "1",
  "purpose": "task_change_context",
  "source_state_digest": { "algorithm": "sha-256", "value": "<64-lowercase-hex>" },
  "scope": {
    "target_task_uid": "2",
    "included_domains": ["project", "target_task", "ancestors", "dependencies", "assignments", "resources"],
    "omitted_domains": ["other_task_details", "raw_external_artifact", "unsupported_data"]
  },
  "project": {
    "name": "Dependency Project",
    "start": "<local-civil-datetime>",
    "finish": "<local-civil-datetime>"
  },
  "target_task": {
    "uid": "2",
    "name": "Execute",
    "parent_uid": null,
    "start": "<local-civil-datetime>",
    "finish": "<local-civil-datetime>",
    "duration": "<working-duration>",
    "milestone": false,
    "summary": false,
    "percent_complete": 0
  },
  "ancestors": [],
  "dependencies": [
    { "predecessor_uid": "1", "successor_uid": "2", "type": "FS", "lag": "<zero-working-duration>" }
  ],
  "resources": [{ "uid": "1", "name": "Miku", "type": "work" }],
  "assignments": [{ "uid": "1", "task_uid": "2", "resource_uid": "1" }],
  "capability": { "unsupported_data": [] },
  "supported_change_requests": [
    {
      "kind": "set_task_percent_complete",
      "required_preconditions": ["source_state_digest", "expected_percent_complete"]
    }
  ]
}
```

Projectionをsemantic state、change request、approval artifactとして入力してはならない。unknown purpose、source state digestの欠落、summary/nonexistent taskをtargetにした`task_change_context`、unsupported dataを隠した編集用Projectionはerrorとする。

## versionと互換性

- `semantic_contract_version`、artifact schema version、external format ID、adapter IDは別のversionとしてrecordする。
- v1 artifactのkindまたはrequired fieldの意味を破壊的に変える場合は、`/v2`または`schema_version = "2"`を新設する。旧artifactを推測変換しない。
- versionが未知、またはartifact kindとschema versionの組合せが未対応なら、CLIはerrorを返す。fallbackや会話履歴による補完をしない。
- Projection、change request、semantic diff、output plan、approval、provenance、conformance用semantic stateはartifact JSON Schemaへ全体適合させる。kind/schema versionだけ合う空objectやrequired fieldを欠く途中artifactをsuccess input/outputとして扱わない。
- result/diagnostics schemaは [CLI result and diagnostics contract v1](miku-project-cli-result-contract-v1.md) に固定した。runtime bindingの値と検証規則は[runtime manifest contract v1](miku-project-runtime-manifest-contract-v1.md)を正本とする。output planとprovenanceはruntime family/version、artifact/manifest digest、capability profile、fixture suite versionを持ち、artifact schema versionと同一視しない。

## G2 review checklist

2026-08-10のG2文書横断reviewと人による`Gate G2`承認で、次を確認した。

- [x] external XML、internal IR、Projection、change request、diff、output plan、approval、provenance、artifact setの役割と寿命に曖昧さがない
- [x] XMLが正本ではなく、semantic stateも長期保存・AI全量handoffではないことが明記されている
- [x] すべてのv1変換について`required / normalized / unsupported-error`が対応づけられている
- [x] `lossy-with-warning`と`opaque-preserved`をv1成功経路に混入させていない
- [x] Projectionのscopeと意図的非公開が、lossと混同されていない
- [x] MS Project XML supported subsetのprofile ID、非目標、namespace、lexical、field、canonical child順、hierarchy、pseudo task、sentinel、unknown dataのmappingが一意である
- [x] output plan、approval、provenance、artifact set、commit markerの役割と寿命が対応づいている
- [x] artifact setが新規directory一個を物理container、`COMMITTED`を論理的な公開境界とし、directory外のsidecar候補を残していない
- [x] approvalの寿命がhidden ledgerを必要としないtuple bindingとして定義されている
