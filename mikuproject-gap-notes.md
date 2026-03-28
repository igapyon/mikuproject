# mikuproject gap notes

`mikuproject` の次段を考えるための、実例 XML ベースの棚卸しメモ。

前提:

- 参照元は `local-data/` に一時配置した実例 XML
- Git 管理対象の testdata ではない
- ここでは `Project / Task / Resource / Assignment / Calendar` ごとに、実例で見えた主なタグを整理する
- 目的は「現在の内部モデルで保持しているもの」と「今後候補になるもの」の差分を把握すること

## 対象実例

- `3PointPlan-example.xml`
- `01145024.xml`
- `Project_Grouping_and_Conditional_Formatting_Example.xml`
- `link types.xml`

## 現在の STEP 1 で保持済み

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
- `CalendarUID`

### Task

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
- `PredecessorLink`

### Resource

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

### Assignment

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
- `ActualWork`
- `RemainingWork`
- `Cost`
- `ActualCost`
- `RemainingCost`
- `PercentWorkComplete`
- `OvertimeWork`
- `ActualOvertimeWork`

### Calendar

- `UID`
- `Name`
- `IsBaseCalendar`
- `IsBaselineCalendar`
- `BaseCalendarUID`
- `WeekDays`
- `Exceptions`
- `WorkWeeks`

## 実例で見えた主な未保持タグ

### Project 候補

優先度が高そう:

- 直近の高優先候補は消化済み

メモ:

- `Project CalendarUID` と calendar 実体の整合 validation は追加済み
- preview で `OutlineCodes / WBSMasks / ExtendedAttributes` の代表値を追えるようにした
- `ProjectModel -> Mermaid gantt` の片方向補助出力に着手済み

後回し候補:

- `ExtendedAttributes` の完全対応
- `ExtendedAttributes`

### Task 候補

優先度が高そう:

- `Task CalendarUID` に紐づくカレンダー差分の可視化は前進済み

メモ:

- `Task CalendarUID` の存在 validation は追加済み
- preview で task ごとの calendar 名を表示するようにした
- predecessor の validation は task 名つきで追えるようにした

実例で頻出だが重い:

- `Baseline`
- `TimephasedData`

特記事項:

- `UID=0`
- 空 `Name`
- `OutlineLevel=0`

は実例で普通に出るため、validation では placeholder 扱いを考慮済み

### Resource 候補

優先度が高そう:

- 直近の高優先候補は消化済み

メモ:

- preview で resource ごとの calendar 名を表示するようにした
- validation 文言は resource 名つきで追いやすくした

実例で頻出だが重い:

- `Baseline`
- `TimephasedData`

### Assignment 候補

優先度が高そう:

- 直近の高優先候補は消化済み

メモ:

- preview で assignment から task/resource 名を追えるようにした
- validation 文言は assignment UID と既知の task/resource 名を併記するようにした

実例で頻出だが重い:

- `Baseline`
- `TimephasedData`
- `ActualCost`
- `RemainingCost`
- `OvertimeWork`
- `ActualOvertimeWork`

特記事項:

- `ResourceUID=-65535`

は実例で未割当を示す特別値として扱う前提

### Calendar 候補

優先度が高そう:

- 直近の高優先候補は消化済み

メモ:

- calendarPreview を追加済み
- `Project / Task / Resource / BaseCalendar` からの参照数を preview で見えるようにした
- `BaseCalendarUID` の自己参照 warning を追加済み

後回し候補:

- 直近の軽量候補は消化済み

## 次に拾う候補

現時点での優先順:

1. 実例 XML ベースの `Calendar` 差分整理
   - task/resource ごとの calendar 差分が実例でどう使われるかを棚卸しする
2. `ExtendedAttributes` の次段整理
   - project 以外の `ExtendedAttributes` をどこまで保持・表示するかを決める
3. preview / validation の最終整形
   - 現状でかなり揃ったので、残る文言や導線を必要最小限で整える
4. 重い構造の着手判断
   - `Baseline` / `TimephasedData` を STEP 2 に含めるか別段に切るかを決める
5. 補助交換形式の次段整理
   - `Mermaid gantt` の dependency 表現をどこまで育てるかを決める
   - `CSV + ParentID` を別軸の交換形式として整理する

補足メモ:

- `CSV + ParentID` は、`ID / ParentID / Name` を最小列とする「まず押さえるべき、よくある交換形式」の第1候補として整理開始した
- `Mermaid gantt` は可視化・共有向けの片方向補助出力、`CSV + ParentID` は編集・交換向けの候補として切り分けて考える
- `CSV + ParentID` は最小出力と最小逆変換を実装済みで、現時点の出力列は `ID / ParentID / WBS / Name / Start / Finish / PredecessorID / Resource / PercentComplete / PercentWorkComplete / Milestone / Summary / Critical / Type / Priority / Work / CalendarUID / ConstraintType / ConstraintDate / Deadline / Notes`
- UI には `CSV を生成` に加えて `CSV を解析` 導線を追加済み
- 逆変換では `ParentID` から task 階層を再構築し、`PredecessorID` と `Resource` から最小の dependency / resource / assignment を復元する
- `PredecessorID / Resource` は `|` に加えて `,` `;` `、` の複数区切りを受け、trim と重複除去をかける
- 構造エラーとして `ID` 重複、空 `Name`、自己参照 / 欠落 / 循環 `ParentID` を import 時点で弾くようにした
- `single CSV` の次段比較候補として `tasks.csv / resources.csv / assignments.csv` をメモ化した
- 判断軸は「人が 1 枚で編集しやすいか」と「resource / assignment を正規化して安全に往復できるか」のトレードオフになる
- 複数 CSV に進む場合の最小草案として、`tasks.csv(ID / ParentID / Name)`, `resources.csv(ResourceID / Name)`, `assignments.csv(AssignmentID / TaskID / ResourceID)` を置く想定にした
- 分割時の実装順は `tasks.csv -> resources.csv -> assignments.csv` が自然で、calendar はさらに次段と考える
- `tasks.csv` については、`ParentID` 正本、`WBS` 補助、`ID / ParentID / Name` 必須、自己参照 / 欠落 / 循環 `ParentID` は import error、という最小仕様草案まで具体化した
- `tasks.csv` の第1段では `Baseline / TimephasedData / ExtendedAttributes / cost 詳細` は含めず、task 単体属性に集中する
- `resources.csv` については、`ResourceID` 正本、`Name` は必須だが重複非推奨、task との紐付けは持たず `assignments.csv` へ分離する、という最小仕様草案まで具体化した
- `resources.csv` の第1段では `Baseline / TimephasedData / ExtendedAttributes / cost 実績詳細` は含めない
- `assignments.csv` については、`AssignmentID / TaskID / ResourceID` 必須、`TaskID / ResourceID` は既存表を参照、assignment 固有の `Start / Finish / Units / Work / PercentWorkComplete` を持てる、という最小仕様草案まで具体化した
- `assignments.csv` の第1段では `Baseline / TimephasedData / ExtendedAttributes / cost 詳細 / contour 系` は含めない

## 後回しでよいもの

- `Baseline` 系
- `TimephasedData`
- コスト系の完全保持
- 表示設定
- Project Server 連携系
- `ExtendedAttributes` の完全対応

## 判断メモ

- STEP 2 は、まず「実例で頻出し、意味が分かりやすく、XML 往復しやすい項目」から拾うのがよい
- `Baseline` や `TimephasedData` は重要だが、構造が重いため別段階が自然
- 実例 XML を読む限り、parser 自体よりも「どこまでを内部モデルで保持するか」の整理が次の主題
- preview / validation の detail 表示強化は一段進んだので、次は実例 XML を見ながら保持方針を詰める段階
- 補助交換形式については、現時点では `single CSV` を主系統に維持し、複数 CSV は仕様草案止まりにするのが妥当
- 複数 CSV へ切り替える条件は、同名 resource 衝突、多重 assignment の lossless 保持、`ResourceID` 正本連携の要求が具体化したとき
- `Mermaid gantt` は、単一 predecessor かつ `FS` かつ lag なし かつ変換しやすい duration の task だけ `after ...` を使う部分ネイティブ化まで進めた
- 複雑な predecessor は、`type` や `lag` を付けたコメントへ逃がす方針にした
- `lag` は `PT...` をそのまま見せるのではなく、`2h` のような短い表現へ寄せるところまで進めた
