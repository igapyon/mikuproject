# mikuproject

`mikuproject` は、MS Project XML 形式の入出力を扱うプロジェクト管理アプリとして設計する。

この文書は `mikuproject` の仕様メモであり、README の代わりではない。利用方法やビルド手順は `README.md` に置き、未完了タスクは `docs/TODO.md` に置く。

配置先:

- `docs/spec.md`

アプリ名:

- `mikuproject`

前提:

- このリポジトリ流儀の single-file web app とする
- ローカルで動作する HTML ツールとして構築する
- まずは UI よりも、MS Project XML の入出力と内部モデル化を優先する
- `MS Project XML` を意味の基軸として扱う
- 内部では `ProjectModel` を中立表現として扱う
- `.xlsx` は確認・可視化・限定編集のための周辺表現として扱う
- 仕様判断に迷った場合は、独自都合よりも `MS Project 仕様` に立ち返って判断する
- `MS Project` 実機は未保有である

## STEP 1 の目的

STEP 1 の目的は、`MS Project XML` を意味的に往復できる状態を作ること。

ここでいう「往復できる」とは、次を意味する。

- `MS Project XML` を読める
- 必要な情報を内部モデルへ落とせる
- 内部モデルから `MS Project XML` を再生成できる
- 再生成した XML を、少なくとも `mikuproject` 自身で再読込できる
- 主要フィールドが壊れず往復できる

注意:

- 目標は「元の XML と完全一致」ではない
- 目標は「意味的に往復できる」ことである

## `.xlsx` の位置づけ

`mikuproject` における `.xlsx` は、`MS Project XML` の代替正本ではない。

- `MS Project XML` は意味の基軸
- `ProjectModel` は内部の中立表現
- `.xlsx` は確認・可視化・限定編集のための周辺表現

したがって、`.xlsx` 対応は `MS Project XML` の仕様を置き換えるためではなく、`ProjectModel` を介した補助入出力として追加する。

現時点で想定する経路は次のとおり。

- `MS Project XML -> ProjectModel -> .xlsx`
- `.xlsx -> ProjectModel -> MS Project XML`

ただし、`.xlsx -> ProjectModel` は自由編集をそのまま受け入れるのではなく、編集可能な列を限定した部分更新として扱う。

現時点の `.xlsx` 周りは、実装済みの限定 import/export として次のように整理できる。

### 現状実装

- 構造忠実な汎用 workbook export/import
  - `Project / Tasks / Resources / Assignments / Calendars` を `ProjectModel` 構造に沿って扱う
- 表示専用の `WBS` workbook export
  - `Tasks` 中心の別 workbook として `.xlsx` 出力できる
  - 現時点では export 専用であり、import は扱わない
  - `WBS XLSX Export` では、`ProjectModel` から補完した既定祝日と、UI で指定した追加祝日を合成して扱う
  - 指定した祝日は WBS 日付帯で祝日色として表示する
  - sample 生成では、`Calendar.Exceptions` のうち非稼働日例外を祝日候補として WBS workbook へ反映する
  - レイアウトは表示専用として継続改善し、上部サマリ、凡例、Week/BaseDate ガイド、日付帯の視認性を段階的に整える

### 新規作成時の既定非稼働日

`mikuproject` は `MS Project XML` を意味の基軸として扱うため、新規 project 作成時の既定非稼働日も、できる限り `MS Project XML` の calendar 表現にそのまま載る形で扱う。

新規 project を作成する場合、明示的な calendar 指定がなければ、既定 calendar を 1 つ作り、その中に次を合成して持たせる前提とする。

- 土日を非稼働日とする週次ルールを `WeekDays` へ設定する
- 日本の祝日を非稼働日例外として `Exceptions` へ設定する

これらは独自の「非稼働日種別」や別 calendar 概念を正本側へ追加するのではなく、最初から `MS Project XML` として自然な 1 つの calendar にまとめて扱う。

この既定 calendar の表示名は、当面 `Standard` を既定値として扱う。新規 project に calendar が存在せず、project 側にも明示的な `CalendarUID` 指定がない場合に限り、この `Standard` calendar を自動補完する。

補完時は少なくとも次を行う。

- `Calendars` に既定 calendar を 1 件追加する
- `Project.CalendarUID` をその calendar の `UID` に設定する
- task / resource に個別 calendar 指定がない場合は、project 既定 calendar を継承する前提で扱う

この既定 calendar に含める祝日例外は、無制限に生成するのではなく、project の `StartDate` を基準に 5 年先までを対象範囲とする。

意図:

- 暗黙の「土日休み」を仕様化して、生成AI による新規計画作成でも前提を揃えやすくする
- 日本の業務予定として自然な初期状態を作る
- `MS Project XML` の `WeekDays / Exceptions` 表現をそのまま使い、独自概念への依存を増やさない
- 将来の実装では、明示的な calendar がある場合にこの既定値を上書きまたは置換できる余地を残す

### WBS workbook の非稼働日反映方針

`WBS XLSX Export` では、単に祝日色を塗るだけでなく、表示上の期間計算にも非稼働日を反映する方向で扱う。

- 期間帯の表示では、非稼働日を作業期間から除外する
- 進捗帯の表示でも、同じ非稼働日基準を使う
- 祝日色の表示と、営業日ベースの期間/進捗表示とで基準が食い違わないようにする

ここでいう非稼働日には、少なくとも次を含める。

- `Calendar.WeekDays` による週次の非稼働日
- `Calendar.Exceptions` による祝日その他の非稼働日例外

土日と祝日は、WBS 上では表示都合で別色にしてよいが、そのために `MS Project XML` 正本へ別概念を追加しない。色分けは `mikuproject` 側の表示ルールで扱う。

現時点で `XLSX Import` の反映対象としている列は次のとおり。

- `Project`: `Name / Title / Author / Company / StartDate / FinishDate / CurrentDate / StatusDate / CalendarUID / MinutesPerDay / MinutesPerWeek / DaysPerMonth / ScheduleFromStart`
- `Tasks`: `Name / Start / Finish / PercentComplete / PercentWorkComplete / Notes`
- `Resources`: `Name / Group / MaxUnits`
- `Assignments`: `Units / Work / PercentWorkComplete`
- `Calendars`: `Name / IsBaseCalendar / BaseCalendarUID`
- `NonWorkingDays`: `Name / Date / FromDate / ToDate / DayWorking`

一覧で見ると次のとおり。

| Sheet | Editable Columns | Notes |
| --- | --- | --- |
| `Project` | `Name / Title / Author / Company / StartDate / FinishDate / CurrentDate / StatusDate / CalendarUID / MinutesPerDay / MinutesPerWeek / DaysPerMonth / ScheduleFromStart` | project 単位の部分更新として扱う |
| `Tasks` | `Name / Start / Finish / PercentComplete / PercentWorkComplete / Notes` | `UID` をキーに部分更新する |
| `Resources` | `Name / Group / MaxUnits` | `UID` をキーに部分更新する |
| `Assignments` | `Units / Work / PercentWorkComplete` | `UID` をキーに部分更新する |
| `Calendars` | `Name / IsBaseCalendar / BaseCalendarUID` | `UID` をキーに部分更新する |
| `NonWorkingDays` | `Name / Date / FromDate / ToDate / DayWorking` | `CalendarUID + Index` をキーに部分更新する |

### Calendar 編集方針

`Calendars / Exceptions` は、業務上は重要だが壊しやすい領域でもあるため、当面は `mikuproject` の画面上で直接編集しない方針とする。

- 画面上では、calendar の存在、件数、参照状況、既定祝日の補完結果などの read-only 確認を主とする
- `Calendars / Exceptions / WeekDays / WorkWeeks` の実編集は、`MS Project XML` または `XLSX Import` 経由で行う
- 画面側に独自の calendar editor を持ち込まず、`MS Project XML` を意味の基軸とする設計を優先する

### 現時点で反映対象外のもの

これ以外の列や、未対応シートの編集は、現在の `XLSX Import` では反映対象としない。特に `Calendars` では、`WeekDays / WorkWeeks` はまだ反映対象外とする。`Exceptions` は `NonWorkingDays` シートとして限定的に扱う。

## WBS ステータスの扱い方針

`WBS` 用の業務ステータスは、`PercentComplete` の派生値としてではなく、`Task.ExtendedAttribute` に保持する前提で扱う。

- `Complete` と `Cancelled` を区別できるようにする
- `PercentComplete=100` とは別軸の状態として保持する
- `MS Project XML` の round-trip で保持しやすい形を優先する

`MS Project` 互換の観点では、`Active=false` は「スケジュール対象外」として使いうるが、`WBS` 上での業務ステータス表示とは役割が異なる。そのため、`mikuproject` では `Cancelled` などの業務値を `ExtendedAttribute` に置く方針を採る。

具体的な `FieldID / FieldName / 値候補` は今後の設計項目とする。

### UI 上の確認手段

`XLSX Import` 後の validation では、`Calendars.BaseCalendarUID` が既存 Calendar を指していない場合や、自身を指している場合の warning も、差分要約と並べて確認できるようにする。

## STEP 1 の完了条件

STEP 1 の完了条件は次のとおり。

- `MS Project XML` を入力として読み込める
- XML から必要な情報を抽出し、内部モデルを生成できる
- 内部モデルから `MS Project XML` を出力できる
- 出力した XML を再読込しても例外にならない
- `xml -> model -> xml -> model` の往復後に、主要フィールドが保持されている

## 現時点の実装メモ

現時点では、STEP 1 の確認をしやすくするために、次の補助表示を持つ。

- `Project / Tasks / Resources / Assignments / Calendars` の件数サマリ
- 内部モデルの JSON 表示
- `Project / Tasks / Resources / Assignments / Calendars` の preview 表示
- validation メッセージ表示

preview / validation の現状メモ:

- project は `OutlineCodes / WBSMasks / ExtendedAttributes` の代表値を preview で追えるようにする
- task / resource / assignment は参照先の名前つきで追えるようにする
- calendar は `Project / Task / Resource / BaseCalendar` からの参照関係を追えるようにする
- validation は `UID` だけでなく、可能な範囲で名前つきで追えるようにする

注意:

- これらは STEP 1 の主目的そのものではなく、意味的ラウンドトリップを確認しやすくするための補助機能である
- `.xlsx` 表示や `.xlsx import/export` も、同様に確認と限定編集のための補助機能として扱う
- `XLSX Import` の反映結果は、`Tasks / Resources / Assignments` ごとの件数と `UID` 単位の差分要約で確認できるようにする
- `XLSX Import` 後も validation を走らせ、反映結果と検証メッセージを同時に確認できるようにする
- validation では、`PercentComplete` の範囲外や `Start > Finish` のような編集結果も UI 上で追えるようにする
- validation が残っていても、`XML Export` はその時点の XML をそのまま保存できるようにする

### 現在の UI 上の整理

現行 UI は、概ね次の 3 画面構成で整理している。

- `Input`
  - `MS Project XML`、`XLSX`、`CSV + ParentID` の読込
  - サンプル XML の読込
  - 生成AIが返した `project_draft_view` の取込
- `Overview`
  - 内部モデルの要約確認
  - validation の確認
  - Mermaid gantt プレビュー
  - preview 表示
- `Output`
  - `MS Project XML`、`XLSX`、`WBS XLSX`、`CSV + ParentID` の保存
  - 生成AI向け `project_overview_view` / `phase_detail_view` の出力

ここでいう `Overview` は、内部実装上の `transform` 相当タブを、ユーザー向けに読み替えた呼称である。

## STEP 1 の入力データ前提

`MS Project` 実機を保有していないため、STEP 1 の入力データ前提は次のとおりとする。

- Microsoft 公開の `MS Project XML schema` を基準にする
- 当面の基準スキーマは `Microsoft Office Project 2007 XML Data Interchange Schema` とする
- 具体的には `https://schemas.microsoft.com/project/2007/` および `mspdi_pj12.xsd` を基準とする
- STEP 1 で扱うファイル形式は、`.mpp` ではなく `.xml` の `MS Project XML 形式` とする
- `.mpp` は MS Project のネイティブ本体形式、`.xml` は外部連携や交換のための XML 表現と捉える
- STEP 1 の検証用 XML は、自作の最小サンプル XML を用いる
- まずは `mikuproject` 自身で意味的に往復できることを優先する
- 実際の `MS Project` 本体が出力した XML との互換確認は、将来課題として扱う

検証用データの参照元メモ:

- 一時的な検証用データの参照元として `https://github.com/rpbouman/open-msp-viewer/` を利用する
- ただし、Git 管理下へそのまま格納するかどうかは別途判断する
- `open-msp-viewer` プロジェクトのサンプルには大いに助けられた。感謝する
- 実例 XML から見えた保持項目ギャップは `docs/gap-notes.md` に整理する
- 仕様判断で迷った場合は、MicrosoftDocs の Project XML Data Interchange リファレンスも補助資料として参照する
  - `https://github.com/MicrosoftDocs/office-developer-msproject-xml-docs/tree/main/project-xml-data-interchange`

## STEP 1 で扱う対象

STEP 1 では、MS Project XML のうち、次の情報を優先して扱う。

- `Project` 基本情報
- `Tasks`
- `Resources`
- `Assignments`
- 必要最小限の `Calendars`
- `PredecessorLink` などの依存関係

## STEP 1 で優先する主要フィールド

### Project

- `Name`
- `Title`
- `Author`
- `Company`
- `CreationDate`
- `LastSaved`
- `SaveVersion`
- `CurrentDate`
- `StartDate`
- `FinishDate`
- `ScheduleFromStart`
- `DefaultStartTime`
- `DefaultFinishTime`
- `MinutesPerDay`
- `MinutesPerWeek`
- `DaysPerMonth`
- `StatusDate`
- `WeekStartDay`
- `WorkFormat`
- `DurationFormat`
- `CurrencyCode`
- `CurrencyDigits`
- `CurrencySymbol`
- `CurrencySymbolPosition`
- `FYStartDate`
- `FiscalYearStart`
- `CriticalSlackLimit`
- `DefaultTaskType`
- `DefaultFixedCostAccrual`
- `DefaultStandardRate`
- `DefaultOvertimeRate`
- `DefaultTaskEVMethod`
- `NewTaskStartDate`
- `NewTasksAreManual`
- `NewTasksEffortDriven`
- `NewTasksEstimated`
- `ActualsInSync`
- `EditableActualCosts`
- `HonorConstraints`
- `InsertedProjectsLikeSummary`
- `MultipleCriticalPaths`
- `TaskUpdatesResource`
- `UpdateManuallyScheduledTasksWhenEditingLinks`
- `CalendarUID`
- `OutlineCodes`
- `WBSMasks`
- `ExtendedAttributes`

### Tasks

- `UID`
- `ID`
- `Name`
- `OutlineLevel`
- `OutlineNumber`
- `WBS`
- `Type`
- `CalendarUID`
- `Priority`
- `Start`
- `Finish`
- `Duration`
- `ActualStart`
- `ActualFinish`
- `Deadline`
- `StartVariance`
- `FinishVariance`
- `Work`
- `WorkVariance`
- `TotalSlack`
- `FreeSlack`
- `Cost`
- `ActualCost`
- `RemainingCost`
- `RemainingWork`
- `ActualWork`
- `Milestone`
- `Summary`
- `Critical`
- `PercentComplete`
- `PercentWorkComplete`
- `Notes`
- `ConstraintType`
- `ConstraintDate`
- `ExtendedAttribute`
- `Baseline`
- `TimephasedData`
- `TimephasedData`
- `PredecessorLink`

### Resources

- `UID`
- `ID`
- `Name`
- `Type`
- `Initials`
- `Group`
- `WorkGroup`
- `MaxUnits`
- `CalendarUID`
- `StandardRate`
- `StandardRateFormat`
- `OvertimeRate`
- `OvertimeRateFormat`
- `CostPerUse`
- `Work`
- `ActualWork`
- `RemainingWork`
- `Cost`
- `ActualCost`
- `RemainingCost`
- `PercentWorkComplete`
- `ExtendedAttribute`
- `Baseline`
- `TimephasedData`

### Assignments

- `UID`
- `TaskUID`
- `ResourceUID`
- `Start`
- `Finish`
- `StartVariance`
- `FinishVariance`
- `Delay`
- `Milestone`
- `WorkContour`
- `Units`
- `Work`
- `Cost`
- `ActualCost`
- `RemainingCost`
- `PercentWorkComplete`
- `OvertimeWork`
- `ActualOvertimeWork`
- `ActualWork`
- `RemainingWork`
- `ExtendedAttribute`
- `Baseline`

### Calendars

- `UID`
- `Name`
- `IsBaseCalendar`
- `BaseCalendarUID`
- `WeekDays`
- `Exceptions`
- `WorkWeeks`

## STEP 1 で後回しにするもの

STEP 1 では、次のようなものは後回し候補とする。

- `.xlsx import` における自由編集の全面対応
- `Calendars / Baseline / TimephasedData / ExtendedAttributes` の `.xlsx` 編集反映

- 表示設定
- UI レイアウト情報
- 独自拡張要素
- 完全互換のために必要だが、主要データの意味保持に直結しない補助ノード群

## 内部モデル方針

内部モデルは、MS Project XML をそのまま保持するのではなく、意味的に扱いやすい正規化済みのモデルとする。

最小モデル案:

```ts
type ProjectModel = {
  project: {
    name: string;
    currentDate?: string;
    startDate: string;
    finishDate: string;
    scheduleFromStart: boolean;
    defaultStartTime?: string;
    defaultFinishTime?: string;
    minutesPerDay?: number;
    minutesPerWeek?: number;
    daysPerMonth?: number;
    statusDate?: string;
    weekStartDay?: number;
    workFormat?: number;
    durationFormat?: number;
    currencyCode?: string;
    currencyDigits?: number;
    currencySymbol?: string;
    currencySymbolPosition?: number;
    fyStartDate?: string;
    fiscalYearStart?: boolean;
    criticalSlackLimit?: number;
    defaultTaskType?: number;
    defaultFixedCostAccrual?: number;
    defaultStandardRate?: string;
    defaultOvertimeRate?: string;
    defaultTaskEVMethod?: number;
    newTaskStartDate?: number;
    newTasksAreManual?: boolean;
    newTasksEffortDriven?: boolean;
    newTasksEstimated?: boolean;
    actualsInSync?: boolean;
    editableActualCosts?: boolean;
    honorConstraints?: boolean;
    insertedProjectsLikeSummary?: boolean;
    multipleCriticalPaths?: boolean;
    taskUpdatesResource?: boolean;
    updateManuallyScheduledTasksWhenEditingLinks?: boolean;
    calendarUID?: string;
    outlineCodes: OutlineCodeModel[];
    wbsMasks: WBSMaskModel[];
    extendedAttributes: ProjectExtendedAttributeModel[];
  };
  calendars: CalendarModel[];
  tasks: TaskModel[];
  resources: ResourceModel[];
  assignments: AssignmentModel[];
};

type TaskModel = {
  uid: string;
  id: string;
  name: string;
  outlineLevel: number;
  outlineNumber: string;
  wbs?: string;
  type?: number;
  calendarUID?: string;
  priority?: number;
  start: string;
  finish: string;
  duration: string;
  actualStart?: string;
  actualFinish?: string;
  deadline?: string;
  startVariance?: string;
  finishVariance?: string;
  work?: string;
  workVariance?: string;
  totalSlack?: string;
  freeSlack?: string;
  cost?: number;
  actualCost?: number;
  remainingCost?: number;
  remainingWork?: string;
  actualWork?: string;
  milestone: boolean;
  summary: boolean;
  critical?: boolean;
  percentComplete: number;
  percentWorkComplete?: number;
  notes?: string;
  constraintType?: number;
  constraintDate?: string;
  predecessors: PredecessorModel[];
};

type PredecessorModel = {
  predecessorUid: string;
  type?: number;
  linkLag?: string;
};

type ResourceModel = {
  uid: string;
  id: string;
  name: string;
  type?: number;
  initials?: string;
  group?: string;
  workGroup?: number;
  maxUnits?: number;
  calendarUID?: string;
  standardRate?: string;
  standardRateFormat?: number;
  overtimeRate?: string;
  overtimeRateFormat?: number;
  costPerUse?: number;
  work?: string;
  actualWork?: string;
  remainingWork?: string;
  cost?: number;
  actualCost?: number;
  remainingCost?: number;
  percentWorkComplete?: number;
};

type AssignmentModel = {
  uid: string;
  taskUid: string;
  resourceUid: string;
  start?: string;
  finish?: string;
  startVariance?: string;
  finishVariance?: string;
  delay?: string;
  milestone?: boolean;
  workContour?: number;
  units?: number;
  work?: string;
  cost?: number;
  actualCost?: number;
  remainingCost?: number;
  percentWorkComplete?: number;
  overtimeWork?: string;
  actualOvertimeWork?: string;
  actualWork?: string;
  remainingWork?: string;
};

type CalendarModel = {
  uid: string;
  name: string;
  isBaseCalendar: boolean;
  isBaselineCalendar?: boolean;
  baseCalendarUID?: string;
  weekDays: Array<{
    dayType: number;
    dayWorking: boolean;
    workingTimes: Array<{
      fromTime: string;
      toTime: string;
    }>;
  }>;
  exceptions: Array<{
    name?: string;
    fromDate?: string;
    toDate?: string;
    dayWorking?: boolean;
    workingTimes: Array<{
      fromTime: string;
      toTime: string;
    }>;
  }>;
  workWeeks: Array<{
    name?: string;
    fromDate?: string;
    toDate?: string;
    weekDays: Array<{
      dayType: number;
      dayWorking: boolean;
      workingTimes: Array<{
        fromTime: string;
        toTime: string;
      }>;
    }>;
  }>;
};
```

注意:

- これは STEP 1 の最小モデル案であり、今後拡張の余地がある
- 日付・期間表現は、まず XML と往復しやすい文字列保持を優先する

## 実装方針

STEP 1 の中核処理は、次のような責務に分ける。

- `parseXmlDocument(xmlText): XMLDocument`
- `importMsProjectXml(xmlText): ProjectModel`
- `validateProjectModel(model): ValidationIssue[]`
- `exportMsProjectXml(model): string`
- `normalizeProjectModel(model): ProjectModel`

テストの基本方針:

- `xml -> model -> xml -> model` のラウンドトリップを確認する
- 比較対象は文字列一致ではなく、正規化後の内部モデル一致とする

実装判断の原則:

- 仕様や表現方法に迷った場合は、`MS Project XML` の持ち方を優先する
- 独自に扱いやすいモデル化は許容するが、`MS Project XML` との意味対応を壊さないことを優先する
- 特にタスク階層や依存関係は、独自表現へ寄せすぎず、まず `MS Project` 側の表現を基準に考える

## テスト方針

STEP 1 では、少なくとも次を確認する。

- サンプル XML を読み込める
- 内部モデルへ変換できる
- 最小妥当性チェック結果を確認できる
- 再生成 XML を出力できる
- 再生成 XML を再読込できる
- 主要フィールドが保持される

比較観点:

- `Project` 基本情報
- `Tasks` の主要フィールド
- `Resources` の主要フィールド
- `Assignments` の主要フィールド
- 依存関係

## 非目標

STEP 1 では、次は非目標とする。

- MS Project XML の完全再現
- 元 XML のノード順や空白や書式の完全保持
- フル機能の編集 UI
- すべての MS Project XML 要素の対応

## STEP 1 実装済みメモ

現時点の STEP 1 実装では、次が入っている。

- `types.ts`, `msproject-xml.ts`, `main.ts` への責務分離
- サンプル XML の読込
- XML 文字列の import
- 内部モデルから整形済み XML を再生成
- XML ファイルの export
- `Project / Tasks / Resources / Assignments / Calendars` の簡易プレビュー表示
- `project / tasks / resources / assignments / calendars` 単位の検証メッセージ表示
- `mikuproject` 独自の最小妥当性チェック
- `Calendar` の `BaseCalendarUID / WeekDays / WorkingTimes` の round-trip
- `Calendar` の `IsBaselineCalendar / Exceptions / WorkWeeks / Exception WorkingTimes` の round-trip
- `Resource` の `CalendarUID / StandardRate / CostPerUse` の round-trip
- `Resource` の `Work / ActualWork / RemainingWork / Cost / ActualCost / RemainingCost / PercentWorkComplete` の round-trip
- `Assignment` の `StartVariance / FinishVariance` の round-trip
- `Resource` の `WorkGroup` の round-trip
- `Assignment` の `Delay / Milestone / WorkContour` の round-trip
- `Assignment` の `OvertimeWork / ActualOvertimeWork` の round-trip
- `Task` の `Deadline / StartVariance / FinishVariance` の round-trip
- `Task` の `WorkVariance / TotalSlack / FreeSlack / Critical` の round-trip
- `Resource` の `StandardRateFormat / OvertimeRate / OvertimeRateFormat` の round-trip
- `Assignment` の `PercentWorkComplete / ActualWork / RemainingWork` の round-trip
- `Project` の `StatusDate / WeekStartDay / WorkFormat / DurationFormat` の round-trip
- `Project` の `CurrencyCode / CurrencyDigits / CurrencySymbol / CurrencySymbolPosition` の round-trip
- `Project` の `FYStartDate / FiscalYearStart` の round-trip
- `Project` の `CriticalSlackLimit / DefaultTaskType` の round-trip
- `Project` の `DefaultFixedCostAccrual / DefaultStandardRate / DefaultOvertimeRate` の round-trip
- `Project` の `DefaultTaskEVMethod / NewTaskStartDate` の round-trip
- `Project` の `NewTasksAreManual / NewTasksEffortDriven` の round-trip
- `Project` の `NewTasksEstimated / ActualsInSync` の round-trip
- `Project` の `EditableActualCosts / HonorConstraints` の round-trip
- `Project` の `InsertedProjectsLikeSummary / MultipleCriticalPaths` の round-trip
- `Project` の `TaskUpdatesResource / UpdateManuallyScheduledTasksWhenEditingLinks` の round-trip
- `Project` の `OutlineCodes / WBSMasks` の最小 round-trip
- `Project` の `ExtendedAttributes` の最小 round-trip
- `Task` の `ExtendedAttribute` の最小 round-trip
- `Resource` の `ExtendedAttribute` の最小 round-trip
- `Assignment` の `ExtendedAttribute` の最小 round-trip
- `Task` の `Baseline` の最小 round-trip
- `Assignment` の `Baseline` の最小 round-trip
- `Resource` の `Baseline` の最小 round-trip
- `Task` の `TimephasedData` の最小 round-trip
- `Resource` の `TimephasedData` の最小 round-trip
- `Assignment` の `TimephasedData` の最小 round-trip
- `Task / Assignment` の `Cost / ActualCost / RemainingCost` の round-trip
- round-trip テスト

## Mermaid gantt 出力メモ

現時点では、確認・共有向けの補助出力として `ProjectModel -> Mermaid gantt` の片方向出力を持つ。

目的:

- `MS Project XML` の全情報保持ではなく、task の時系列と大まかな依存関係を軽量に共有する
- `mikuproject` 内部モデルの内容を、Mermaid 対応環境へ持ち出しやすくする

現時点の出力方針:

- summary task は `section` として扱う
- summary ではない task のうち、`Start` と `Finish` を持つものを gantt のタスク行として出力する
- `critical=true` は `crit` として出力する
- `milestone=true` は `milestone` として出力する
- `percentComplete >= 100` は `done` として出力する
- task 名や title は Mermaid で壊れやすい一部記号を簡易正規化して出力する
- predecessor は、`単一 predecessor` かつ `FS` かつ `lag なし` かつ `duration` を Mermaid 向けへ素直に変換できる task のみ `after ...` でネイティブ出力する
- 上記に当てはまらない predecessor は、task 名を含むコメント行で補助出力する
- comment 側の `lag` は、可能な範囲で `2h` のような短い人間向け表現に整形して出力する
- `lag` がある場合は、`after Prep + 2h` のような擬似読解用 comment も追加する

現時点で意図的に落とすもの:

- `Resources`
- `Assignments`
- `Calendars`
- `Baseline`
- `TimephasedData`
- コスト系の詳細
- `PredecessorLink` の完全表現

注意:

- これはあくまで片方向の補助出力であり、`Mermaid gantt -> ProjectModel` の往復は対象外とする
- 現時点の dependency 表現は部分的にネイティブ化しているが、複数 predecessor、`FS` 以外の link type、lag あり、複雑な duration はコメント保持のままとする
- どの情報を落としているかは、将来の `CSV + ParentID` 等の交換形式検討と切り分けて扱う

## CSV + ParentID 交換形式メモ

`mikuproject` の次段候補として、`CSV + ParentID` を「まず押さえるべき、よくある交換形式」の第1候補とする。

目的:

- 人が表計算ソフトやスプレッドシートで編集しやすい形を持つ
- 独自記法を先に増やしすぎず、一般的な交換形式を先に押さえる
- task 階層を `ParentID` で素直に表現する

最小列候補:

- `ID`
- `ParentID`
- `Name`

実用列候補:

- `WBS`
- `Start`
- `Finish`
- `PredecessorID`
- `Resource`
- `PercentComplete`

現時点の整理方針:

- まずは単一 CSV を前提に考える
- task 階層の正本は `ParentID` とし、`WBS` は補助列として扱う候補とする
- `PredecessorID` は単一値か複数値区切りかを今後決める
- `Resource` は名前で持つか `ResourceID` で持つかを今後決める

単一 CSV で落ちやすいもの:

- `Assignments` の完全表現
- `Calendars`
- `Baseline`
- `TimephasedData`
- コスト系の詳細

注意:

- 現時点では仕様草案段階であり、`CSV + ParentID <-> ProjectModel` の完全往復仕様は未確定
- 将来必要であれば、`tasks.csv / resources.csv / assignments.csv` の複数表構成も比較対象にする
- 現在の UI では、`CSV + ParentID` は textarea ではなくファイルベースの補助入出力として扱う
  - `Input` 側は CSV ファイル読込
  - `Output` 側は CSV ダウンロード

複数 CSV 構成の比較メモ:

- `single CSV` の利点は、人が 1 枚の表で task 階層を編集しやすいこと
- `single CSV` の弱点は、`Resource` や `Assignment` を task 行へ押し込むため、正規化されず表現が崩れやすいこと
- `tasks.csv / resources.csv / assignments.csv` の利点は、resource と assignment を独立表現でき、`ResourceID` ベースの安全な往復へ寄せやすいこと
- `tasks.csv / resources.csv / assignments.csv` の弱点は、人が直接編集するには 1 ファイル増えて分かりにくくなること
- 現時点では、まず `single CSV` で task 中心の軽量交換を育て、resource / assignment の保持要求が増えた時点で複数 CSV を比較する方針とする
- その場合の最初の分割候補は `tasks.csv` と `resources.csv` と `assignments.csv` であり、calendar はさらに次段とする

複数 CSV の最小草案:

- `tasks.csv`
  - 最小列候補: `ID / ParentID / Name`
  - 実用列候補: `WBS / Start / Finish / PredecessorID / PercentComplete / PercentWorkComplete / Milestone / Summary / Critical / Type / Priority / Work / CalendarUID / ConstraintType / ConstraintDate / Deadline / Notes`
- `resources.csv`
  - 最小列候補: `ResourceID / Name`
  - 実用列候補: `Initials / Group / CalendarUID / MaxUnits / StandardRate / OvertimeRate / CostPerUse`
- `assignments.csv`
  - 最小列候補: `AssignmentID / TaskID / ResourceID`
  - 実用列候補: `Start / Finish / Units / Work / PercentWorkComplete`

草案メモ:

- `tasks.csv` は現在の `single CSV` の task 列をほぼそのまま引き継げる
- `resources.csv` は name だけでなく `ResourceID` を正本にすることで、同名 resource の衝突を避けやすい
- `assignments.csv` を分けることで、1 task に複数 resource が割り当たるケースを自然に表現できる
- 第1段では `calendar` と `baseline/timephased` は複数 CSV にも入れず、別段とする
- もし複数 CSV に進む場合、最初の実装順は `tasks.csv -> resources.csv -> assignments.csv` が妥当と考える

`tasks.csv` の最小仕様草案:

- 目的は task 階層と task 単体属性を、resource / assignment から切り離して安全に往復すること
- 正本の階層表現は `ParentID` とし、`WBS` は補助列扱いとする
- `ID / ParentID / Name` を必須列とする
- `ID` は CSV 内で一意でなければならない
- `ParentID` は空文字を root task とみなし、値がある場合は既存 `ID` を指さなければならない
- `ParentID` の自己参照と循環参照は import error とする
- `Name` は空不可とする
- `PredecessorID` は任意列とし、複数値は `|` を正規表現としつつ、import では `,` `;` `、` も受ける
- `Milestone / Summary / Critical` は `0/1` を正とし、import では `true/false/yes/no` も受ける
- `PercentComplete / PercentWorkComplete` は `0..100` を想定し、範囲外は validation 対象とする
- `Start / Finish / ConstraintDate / Deadline` は `MS Project XML` と同じ日時文字列を前提にする
- `Type / Priority / ConstraintType` は整数列とする
- `Work` は `PT...` 形式の duration 文字列を前提にする

`tasks.csv` の第1段 scope:

- 含める: 階層、日付、依存、進捗、milestone/summary/critical、主要 task 属性
- 含めない: `Baseline`, `TimephasedData`, `ExtendedAttributes`, task ごとの cost 詳細
- `CalendarUID` は保持対象に含めるが、calendar 実体は別表へ分けず参照値扱いに留める

`resources.csv` の最小仕様草案:

- 目的は resource 単体属性を task 行から切り離し、同名 resource を安全に区別できるようにすること
- 正本の識別子は `ResourceID` とし、`Name` は表示用の主要属性として扱う
- `ResourceID / Name` を必須列とする
- `ResourceID` は CSV 内で一意でなければならない
- `Name` は空不可とする
- `Name` の重複は直ちに import error とはしないが、運用上は非推奨とする
- `CalendarUID` は任意列とし、calendar 実体は別表へ分けず参照値扱いに留める
- `MaxUnits / CostPerUse` は数値列とする
- `StandardRate / OvertimeRate` は `MS Project XML` と同じ文字列表現を前提にする
- `Initials / Group` は任意の表示属性とする

`resources.csv` の第1段 scope:

- 含める: 識別子、表示名、group/initials、calendar 参照、基本 rate/cost 属性
- 含めない: `Baseline`, `TimephasedData`, `ExtendedAttributes`, resource ごとの cost 実績詳細
- `assignments.csv` が別にある前提で、task との紐付けは `resources.csv` に持たせない

`assignments.csv` の最小仕様草案:

- 目的は task と resource の関係を独立表現し、1 task に複数 resource が付くケースを正規化して扱うこと
- 正本の識別子は `AssignmentID` とし、参照の正本は `TaskID / ResourceID` とする
- `AssignmentID / TaskID / ResourceID` を必須列とする
- `AssignmentID` は CSV 内で一意でなければならない
- `TaskID` は `tasks.csv` の既存 `ID` を指さなければならない
- `ResourceID` は `resources.csv` の既存 `ResourceID` を指さなければならない
- `TaskID / ResourceID` の組が重複する assignment を許すかは未確定だが、第1段では重複非推奨とする
- `Start / Finish` は任意列とし、assignment 固有の期間がある場合のみ保持する
- `Units / PercentWorkComplete` は数値列とする
- `Work` は `PT...` 形式の duration 文字列を前提にする

`assignments.csv` の第1段 scope:

- 含める: task-resource 参照、多重割当、assignment 単体の期間と work/units/進捗
- 含めない: `Baseline`, `TimephasedData`, `ExtendedAttributes`, assignment ごとの cost 詳細
- 第1段では `Milestone / Delay / WorkContour / OvertimeWork` などは未保持でもよい

現時点の判断メモ:

- 当面は `single CSV` を主系統として維持する
- 理由は、いまの利用目的が「軽量な交換・編集」であり、1 枚の表で task 階層を扱える利点がまだ大きいからである
- `tasks.csv / resources.csv / assignments.csv` は有力な次段候補だが、現時点では仕様草案までに留める
- `single CSV` から複数 CSV へ切り替える判断条件は、少なくとも次のいずれかを満たしたときとする
  - 同名 resource を安全に往復したい要求が具体化した
  - 1 task に複数 resource を持つ assignment を lossless に扱いたい要求が増えた
  - assignment 単体属性を `single CSV` の task 行へ押し込むのが不自然になった
  - `ResourceID` 正本での連携が必要になった
- 逆に、task 中心の軽量編集が主目的である間は `single CSV` の方が実用的とみなす

現時点の実装メモ:

- `ProjectModel -> CSV + ParentID` の出力を持つ
- 現在の出力列は `ID / ParentID / WBS / Name / Start / Finish / PredecessorID / Resource / PercentComplete / PercentWorkComplete / Milestone / Summary / Critical / Type / Priority / Work / CalendarUID / ConstraintType / ConstraintDate / Deadline / Notes`
- `PredecessorID` は複数値を `|` 区切りで補助出力する
- `Resource` は assignment から task 単位で集約した resource 名を補助出力する
- `CSV + ParentID -> ProjectModel` の最小逆変換を持つ
- 最小逆変換では `ID / ParentID / Name` を必須とし、`WBS / Start / Finish / PredecessorID / Resource / PercentComplete / PercentWorkComplete / Milestone / Summary / Critical / Type / Priority / Work / CalendarUID / ConstraintType / ConstraintDate / Deadline / Notes` を可能な範囲で復元する
- 最小逆変換では `PredecessorID / Resource` の複数値区切りとして `|` に加えて `,` `;` `、` を受け付け、trim と重複除去を行う
- 最小逆変換では `ID` 重複、空 `Name`、自己参照 `ParentID`、欠落 `ParentID`、循環 `ParentID` を import error として扱う
- UI には `CSV` のダウンロード導線と、CSV ファイル読込導線を追加済み
- 現時点では `Project` 詳細、`Calendars`、`Baseline`、`TimephasedData`、assignment 詳細は CSV から完全復元しない

## 次に決めること

STEP 1 の次の検討項目:

1. サンプル XML の置き場所
2. 内部モデル型の確定
3. XML パーサ / シリアライザの実装方針
4. STEP 1 で実際に保持する必須フィールドの最終確定
5. ラウンドトリップ比較用の正規化ルール
