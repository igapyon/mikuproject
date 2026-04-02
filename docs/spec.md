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

立ち位置:

- `mikuproject` は `MS Project` 代替製品を目指すものではない
- 中核は、`MS Project XML` をハブにして `XLSX / Markdown / JSON / Mermaid / 生成AI / MS Project` をつなぐ変換・可視化・橋渡しツールである
- したがって、重要機能の優先順位も「橋渡しを強くするか」で判断する
- 具体的には、差分可視化、Patch 適用、phase 単位の scoped export/import、import/export の扱い可視化は優先しやすい
- 一方で、重い管理機能や UI 内完結の大規模操作は優先度を上げすぎない

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

同様に、構造忠実 workbook の JSON 版も、`MS Project XML` の代替正本ではなく、`.xlsx` の写し身として扱う。

一方で、生成AI 連携の編集用 JSON はこれと別系統のものとし、当面は workbook JSON と混同しないよう `.editjson` 拡張子を推奨する。

現時点で想定する経路は次のとおり。

- `MS Project XML -> ProjectModel -> .xlsx`
- `.xlsx -> ProjectModel -> MS Project XML`
- `MS Project XML -> ProjectModel -> workbook JSON`
- `workbook JSON -> ProjectModel -> MS Project XML`

ただし、`.xlsx -> ProjectModel` は自由編集をそのまま受け入れるのではなく、編集可能な列を限定した部分更新として扱う。

`workbook JSON -> ProjectModel` も同様に、自由編集をそのまま受け入れるのではなく、`.xlsx import` と同じ編集可能列の部分更新として扱う。

現時点の `.xlsx` / workbook JSON 周りは、実装済みの限定 import/export として次のように整理できる。

### 現状実装

- 構造忠実な汎用 workbook export/import
  - `Project / Tasks / Resources / Assignments / Calendars` を `ProjectModel` 構造に沿って扱う
- 構造忠実 workbook の JSON export/import
  - `mikuproject_workbook_json` として、`XLSX` workbook の論理構造を JSON へ写して扱う
  - `format = "mikuproject_workbook_json"`、`version = 1` を持つ
  - `Project / Tasks / Resources / Assignments / Calendars / NonWorkingDays` を `sheets` 配下に持つ
  - import 時の反映対象列・キー・部分更新ルールは `XLSX Import` と同じにする
- 表示専用の `WBS` workbook export
  - `Tasks` 中心の別 workbook として `.xlsx` 出力できる
  - 現時点では export 専用であり、import は扱わない
  - `WBS XLSX Export` では、`ProjectModel` から補完した既定祝日と、UI で指定した追加祝日を合成して扱う
  - 指定した祝日は WBS 日付帯で祝日色として表示する
  - sample 生成では、`Calendar.Exceptions` のうち非稼働日例外を祝日候補として WBS workbook へ反映する
  - 現行レイアウトでは、先頭に `プロジェクト情報`、続いて `凡例` と `サマリ` を置き、その下に日付帯と task 一覧を並べる
  - `プロジェクト情報` ブロックは先頭に置き、`プロジェクト名 / カレンダ / 開始日 / 終了日 / 現在日 / 祝日` を持つ
  - `凡例` ブロックでは、進捗済み / 予定帯 / 当日 / 週頭 / 週末 / 祝日 / フェーズ / 進捗済みタスク / 予定タスク / マイルストーン / サマリ / クリティカル / 未設定 を見分けられるようにする
  - `サマリ` ブロックには `表示日 / 表示週 / 営業日 / 前日数 / 後日数 / 表示 / 進捗 / 基準日 / タスク / リソース / 割当 / カレンダ` を持つ
  - `サマリ` の値側は、長い日付がはみ出しにくいように結合セルで表示する
  - 日付帯では、日付行と曜日行を分けて表示し、週ラベル専用行は持たない
  - 曜日帯では `Sat` と `Sun` を週末色で表示し、祝日は別色で表示する
  - `タスク詳細` は `Task.Notes` を表示し、空の場合は `-` とする
  - `J2` には `出力日時 YYYY-MM-DD HH:mm` を出す
- 表示専用の `WBS Markdown` export
  - `WBS XLSX` と同じく、人が読むための帳票として扱う
  - 当面は export 専用であり、import は扱わない
  - `WBS XLSX` と同じ `ProjectModel` を入力にし、表示内容もできるだけ揃える
  - 少なくとも `プロジェクト情報`、`WBS 本体`、`サマリ` を持つ
  - `プロジェクト情報` では、見出し上の主名として `Project.Name` を使う
  - `Project.Title` は、Markdown 出力の主表示には使わない
  - 1 つの Markdown に、`WBS ツリー` と `WBS テーブル` の両方を含める
  - 出力順は、`プロジェクト情報`、`WBS ツリー`、区切り線、`WBS テーブル`、区切り線、`サマリ` を基本とする
  - 前半の主表示は `WBS ツリー` とし、`WBS テーブル` は確認用の詳細表、その後ろに `サマリ` を置く
  - table 形式では、`プロジェクト情報` と `サマリ` を小さな Markdown table に分ける
  - `WBS テーブル` は、可読性を優先して 1 個の大 table とする
  - 結合セル、塗り色、罫線、列幅は Markdown へは持ち込まない
  - `WBS ツリー` では、task 階層をインデントと記号で表す
  - `WBS ツリー` では、子 task を `┗　` で表し、階層が深い場合は `全角空白 + ┗　` を段数ぶん積む形とする
  - `WBS ツリー` は Markdown の list 記法に無理に寄せず、等幅で読める `WBS の文字` として出す
  - `WBS ツリー` の task 行には、少なくとも `WBS / 名称 / 開始 / 終了 / 進捗` を含められるようにする
  - `タスク詳細` は、`WBS テーブル` では列として保持し、`WBS ツリー` では task 本文の次行または後置表現で保持する
  - 日付帯や進捗帯のような色依存の表現は、そのまま再現しない
  - 色依存の表現は、日付文字列・数値へ落として表す
  - `WBS Markdown` は、`WBS XLSX` と競合するものではなく、軽量なテキスト共有用の派生表現として扱う
  - `WBS ツリー` と `WBS テーブル` の task 集合は一致させ、片方だけに存在する task を作らない
  - 当面の仕様検討では、まずこの両方入り構成で sample 出力を作り、前半と後半の情報差を確認する
### 現在の派生表示

- 表示専用の `月別WBSカレンダー SVG` export
  - `WBS XLSX` や `WBS Markdown` と同列の、表示専用の派生出力として扱う
  - 入力は `ProjectModel` とし、正本や限定 import/export の経路とは切り分ける
  - 出力は単月ではなく、project 期間に含まれるすべての年月を対象にした月別 `SVG` 群とする
  - 配布単位は全月分をまとめた `ZIP` とし、各 `SVG` は 1 か月単位の月間カレンダーとする
  - 月グリッドは当面 `日曜日始まり` で固定する
  - 各日付セルには、その日に関係する `task` や `milestone` を表示する
  - 通常 `task` は `Start` から `Finish` 当日までを表示範囲とし、日付範囲の解釈は現行の `Daily SVG` 出力と同様に扱う
  - 通常 `task` は平日の各日に表示するが、`土日祝日` の中間日には表示しない
  - ただし、`土日祝日` であっても、その日が `Start` または `Finish` に一致する場合は表示する
  - `milestone` は単日要素として、その該当日のみに表示する
  - `summary task` は期間中の各日には表示せず、`開始日` と `終了日` の日のみに表示する
  - 日セル内の表示文字列には `WBS` 番号を含めない
  - 長い `task` 名は、セル幅に応じて途中で切って表示してよい
  - 1 日セル内の表示件数には上限を設け、当面は `3` または `4` 件程度を上限目安とし、超過分は省略表示する

### 派生表示の次段候補

- 表示専用の `WBS記述書 Markdown` export
  - `WBS` の階層表現そのものとは別物として扱う
  - 目的は、各 task の意味・成果物・完了条件などを文章で確認しやすくすること
  - 当面は export 専用であり、import は扱わない
  - 最小設計では、保持元を `Task.ExtendedAttribute` 主体とし、長文補足のみ `Task.Notes` を使う
  - 最小設計で扱う項目は、少なくとも次の 5 つとする
    - `TaskPurpose`
    - `TaskDeliverable`
    - `TaskOutOfScope`
    - `TaskDoneDefinition`
    - `TaskOwner`
  - 上記 5 項目は `Task.ExtendedAttribute` の `FieldName` で識別する前提とする
  - 補足本文や自由記述は `Task.Notes` を使う
  - Markdown 出力では、task ごとに 1 節を持つ構成を基本とする
  - 各節では、少なくとも `WBS / 名称 / 目的 / 成果物 / スコープ外 / 完了条件 / 担当 / 補足` を必要に応じて出す
  - 未設定の項目は空欄のまま出さず、省略する
  - `WBS記述書 Markdown` は、`WBS Markdown` の代替ではなく、task 説明を補う別出力として扱う
  - 当面は新規入力 UI を先に増やさず、既存の `ExtendedAttribute` / `Notes` からの export を優先する

### WBS Report SVG / ZIP の扱い

- `WBS Report` には、少なくとも `XLSX` / `Markdown` / `Daily SVG` / `Weekly SVG` / `Monthly Calendar SVG` / `Mermaid` を置く
- `Daily SVG` は日単位の WBS 俯瞰用 SVG とする
- `Weekly SVG` は週単位の WBS 俯瞰用 SVG とする
- `Monthly Calendar SVG` は月別カレンダー SVG 群をまとめた `ZIP` とする
- `Overview` の preview は、`Daily` / `Weekly` / `Monthly Calendar` を切り替えられるようにする

ファイル名は当面次を基本とする。

- `Daily SVG`: `mikuproject-wbs-daily-<YYYYMMDDHHmm>.svg`
- `Weekly SVG`: `mikuproject-wbs-weekly-<YYYYMMDDHHmm>.svg`
- `Monthly Calendar SVG`: `mikuproject-monthly-wbs-calendar-<YYYYMMDDHHmm>.zip`
- `Mermaid`: `mikuproject-wbs-mermaid-<YYYYMMDDHHmm>.md`

`Daily SVG` のラベル配置は次の方針とする。

- 基本は task / phase の近接ラベルとする
- 右側へ自然に置ける場合は右側へ表示する
- 左側へ自然に置ける場合は左側へ表示する
- 左側配置は canvas 外へ無制限に広げず、左端で抑える
- 左右どちらにも逃がしにくい場合や、表示範囲いっぱいに近い帯・線については、帯や線の上へラベルを重ねてよい
- 線上ラベルを使う場合は、可読性のために薄い背景を敷いてよい

`Weekly SVG` のラベル配置も同様に、基本は近接ラベルとしつつ、左側へ無制限に広がらないように扱う。

### Output の ALL ZIP

- `Output` には、主要成果物をまとめて取得する `ALL` ボタンを置く
- `ALL` は、その時点の内部モデルから再生成した主要出力をまとめた `ZIP` とする
- `ALL` の保存名は `mikuproject-all-<YYYYMMDDHHmm>.zip` とする
- `ALL` ZIP の並び順は、概ね `README.txt -> XML -> workbook -> CSV -> WBS Report -> monthly-calendar -> .editjson` の順に固定してよい

`ALL` ZIP には、少なくとも次を含める。

- `README.txt`
- `MS Project XML`
- workbook `XLSX`
- workbook `JSON`
- `CSV + ParentID`
- `WBS XLSX`
- `WBS Markdown`
- `Daily SVG`
- `Weekly SVG`
- `Mermaid`
- `project_overview_view`
- `full bundle`
- `phase_detail_view full`

`Monthly Calendar SVG` の扱いは次のとおり。

- 単独出力では、月別 `SVG` 群をまとめた `ZIP` として配布してよい
- 単独出力の `ZIP` でも、`monthly-calendar/` ディレクトリ配下へ月別 `SVG` を格納する
- 単独出力の `monthly-calendar/` 配下のファイル名も `YYYY-MM.svg` のような短い名前でよい
- ただし `ALL` ZIP の中では、さらに入れ子の `ZIP` にせず、`monthly-calendar/` ディレクトリ配下へ月別 `SVG` を展開して含める
- `monthly-calendar/` 配下のファイル名は `YYYY-MM.svg` のような短い名前でよい

### workbook JSON の位置づけ

`mikuproject` は、構造忠実 workbook のテキスト版として `mikuproject_workbook_json` を持てるようにする。

これは生成AI向け projection JSON とは別物であり、`XLSX` の写し身として扱う。

- 目的は、構造忠実 workbook をテキストで扱いやすくすること
- `XLSX` と意味を揃えた補助入出力として扱う
- 新しい編集モデルや新しい意味体系は持ち込まない
- styling、列幅、merge、塗り色などの表示情報は持たない

`mikuproject_workbook_json` の最小外形は次のとおり。

```json
{
  "format": "mikuproject_workbook_json",
  "version": 1,
  "sheets": {
    "Project": [],
    "Tasks": [],
    "Resources": [],
    "Assignments": [],
    "Calendars": [],
    "NonWorkingDays": []
  }
}
```

sheet 名は、構造忠実 workbook の `XLSX` と同じく次で固定する。

- `Project`
- `Tasks`
- `Resources`
- `Assignments`
- `Calendars`
- `NonWorkingDays`

各 sheet の列名は、対応する `XLSX` workbook の header と完全一致で固定する。

つまり、`mikuproject_workbook_json` は `XLSX` workbook の論理構造を JSON へ写したものであり、列追加や別名導入は行わない。

import 時の扱いも `XLSX Import` と完全に揃える。

- 反映対象列は `XLSX Import` と完全一致とする
- 反映単位は部分更新とする
- `Tasks / Resources / Assignments / Calendars` は `UID` をキーに扱う
- `NonWorkingDays` は `CalendarUID + Index` をキーに扱う
- 未対応列や未知の列は反映対象にしない
- `workbook JSON` だからといって自由編集を全量反映しない

拡張子運用:

- `mikuproject_workbook_json` は、構造忠実 workbook の JSON として `.json` を推奨する
- 生成AI 連携の編集用 JSON は、workbook JSON と区別するため `.editjson` を推奨する
- `project_draft_view` は当面この編集用 JSON 群に含め、`.editjson` で受け渡す前提とする

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

この既定 calendar に含める祝日例外は、無制限に生成するのではなく、原則として project の `StartDate` から `FinishDate` までの期間内に入るものへ限定する。

この制限は、calendar が存在しない project に対して `mikuproject` が既定 calendar を自動補完する場合にだけ適用する。

すでに calendar が存在する場合や、ユーザーが明示的に指定した calendar / `Exceptions` については、`mikuproject` 側で自動的に再構成・再トリミングしない前提とする。

意図:

- 暗黙の「土日休み」を仕様化して、生成AI による新規計画作成でも前提を揃えやすくする
- 日本の業務予定として自然な初期状態を作る
- project 期間外の祝日を `MS Project XML` へ過剰に書き出さず、正本の見通しと差分の素直さを保つ
- 既存 calendar やユーザー指定内容を、`mikuproject` の都合で自動変更しない
- `MS Project XML` の `WeekDays / Exceptions` 表現をそのまま使い、独自概念への依存を増やさない
- 将来の実装では、明示的な calendar がある場合にこの既定値を上書きまたは置換できる余地を残す

### WBS workbook の非稼働日反映方針

`WBS XLSX Export` では、単に祝日色を塗るだけでなく、表示上の期間計算にも非稼働日を反映する方向で扱う。

- 期間帯の表示では、非稼働日を作業期間から除外する
- 進捗帯の表示でも、同じ非稼働日基準を使う
- 祝日色の表示と、営業日ベースの期間/進捗表示とで基準が食い違わないようにする
- 通常日の空セルは原則として無色とし、土日祝日・当日・週頭など意味を持つ日だけ背景色を持たせてよい

ここでいう非稼働日には、少なくとも次を含める。

- `Calendar.WeekDays` による週次の非稼働日
- `Calendar.Exceptions` による祝日その他の非稼働日例外

土日と祝日は、WBS 上では表示都合で別色にしてよいが、そのために `MS Project XML` 正本へ別概念を追加しない。色分けは `mikuproject` 側の表示ルールで扱う。

現時点で `XLSX Import` の反映対象としている列は次のとおり。

- `Project`: `Name / Title / Author / Company / StartDate / FinishDate / CurrentDate / StatusDate / CalendarUID / MinutesPerDay / MinutesPerWeek / DaysPerMonth / ScheduleFromStart`
- `Tasks`: `Name / Start / Finish / Duration / PercentComplete / PercentWorkComplete / Milestone / Summary / Critical / CalendarUID / Predecessors / Notes`
- `Resources`: `Name / Group / MaxUnits / CalendarUID`
- `Assignments`: `Units / Work / PercentWorkComplete`
- `Calendars`: `Name / IsBaseCalendar / BaseCalendarUID`
- `NonWorkingDays`: `Name / Date / FromDate / ToDate / DayWorking`

一覧で見ると次のとおり。

| Sheet | Editable Columns | Notes |
| --- | --- | --- |
| `Project` | `Name / Title / Author / Company / StartDate / FinishDate / CurrentDate / StatusDate / CalendarUID / MinutesPerDay / MinutesPerWeek / DaysPerMonth / ScheduleFromStart` | project 単位の部分更新として扱う |
| `Tasks` | `Name / Start / Finish / Duration / PercentComplete / PercentWorkComplete / Milestone / Summary / Critical / CalendarUID / Predecessors / Notes` | `UID` をキーに部分更新する |
| `Resources` | `Name / Group / MaxUnits / CalendarUID` | `UID` をキーに部分更新する |
| `Assignments` | `Units / Work / PercentWorkComplete` | `UID` をキーに部分更新する |
| `Calendars` | `Name / IsBaseCalendar / BaseCalendarUID` | `UID` をキーに部分更新する |
| `NonWorkingDays` | `Name / Date / FromDate / ToDate / DayWorking` | `CalendarUID + Index` をキーに部分更新する |

`mikuproject_workbook_json` の import も、この表と同じ反映対象列・キー・部分更新ルールを使う。

`Tasks` シートの header ごとの扱いは次のとおり。

| Tasks Header | 扱い |
| --- | --- |
| `UID` | 表示のみ |
| `ID` | 表示のみ |
| `Name` | import 可 |
| `OutlineLevel` | 表示のみ |
| `OutlineNumber` | 表示のみ |
| `WBS` | 表示のみ |
| `Start` | import 可 |
| `Finish` | import 可 |
| `Duration` | import 可 |
| `PercentComplete` | import 可 |
| `PercentWorkComplete` | import 可 |
| `Milestone` | import 可 |
| `Summary` | import 可 |
| `Critical` | import 可 |
| `CalendarUID` | import 可 |
| `Predecessors` | import 可 |
| `Notes` | import 可 |

`Predecessors` の workbook import は、現時点では `predecessorUid` の `,` 区切り一覧を読む最小対応とする。`type / linkLag` などの詳細な依存表現は将来拡張とする。

`Resources` と `Assignments` が 0 件の workbook では、どの列が import 対象か分かるように、editable 列だけ着色されたダミー行を 1 行出してよい。これは表示補助であり、`UID` 等のキーが空なので import 時には無視される。

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
  - `Load from file` から `MS Project XML`、`XLSX`、workbook JSON (`.json`)、生成AI向け編集用 JSON (`.editjson`)、`CSV + ParentID` を読込
  - サンプル XML の読込
  - 生成AIが返した `project_draft_view` の貼り付け取込
- `Overview`
  - 内部モデルの要約確認
  - validation の確認
  - `Daily / Weekly / Monthly Calendar` preview の確認
  - preview 表示
- `Output`
  - `MS Project XML`、`XLSX`、`WBS XLSX`、workbook JSON、`CSV + ParentID` の保存
  - Mermaid fenced code block を含む `.md` と `Daily / Weekly / Monthly Calendar SVG` の保存
  - 生成AI向け `project_overview_view` / `phase_detail_view` / `full bundle` の `.editjson` 出力

生成AI連携の現状実装範囲:

- 既存 project 向けには `project_overview_view` / `phase_detail_view` / `task_edit_view` / `full bundle` の export を持つ
- `full bundle` には少なくとも `project_overview_view` / `phase_detail_views_full` / `task_edit_views_full` を含めてよい
- 新規生成向けには `project_draft_view` の import を持つ
- 既存 project 向けには Patch JSON の `update_project` / `update_task` / `update_resource` / `update_assignment` / `update_calendar` / `link_tasks` / `unlink_tasks` first cut import / 適用を持つ
- `task_edit_view` は既存 task を `uid` で特定して出力する個別編集用 projection とする
- `task_edit_view` は少なくとも `project` / `phase` / `target_task` / `parent_task` / `sibling_tasks` / `predecessors` / `successors` / `assignments` / `rules` を含む
- `task_edit_view` の `target_task` では `name` / `parent_uid` / `position` / `is_summary` / `is_milestone` / `planned_duration` / `planned_duration_hours` / `planned_start` / `planned_finish` / `percent_complete` / `notes` / `calendar_uid` / `critical` を見せる
- `task_edit_view.rules.allow_patch_ops` には、少なくとも `update_task` / `move_task` / `link_tasks` / `unlink_tasks` / `add_task` / `delete_task` / `update_assignment` / `add_assignment` / `delete_assignment` を含めてよい
- Patch JSON のうち `add_task` / `add_resource` / `add_assignment` / `add_calendar` / `move_task` / `delete_task` / `delete_resource` / `delete_assignment` / `delete_calendar` は first cut を実装済みとする

Patch JSON の次段 MVP 方針:

- 既存 project 向けの AI 返却形式として `Patch JSON` を将来追加する
- first cut は `operations` 配列を持つ部分適用 JSON とする
- first cut では、少なくとも `update_task` の import / 適用から着手し、依存関係の最小操作として `link_tasks` / `unlink_tasks` も段階的に含める
- first cut では、対象 task は `uid` で特定する
- first cut では、task の基本計画項目
  - `name`
  - `notes`
  - `calendar_uid`
  - `percent_complete`
  - `percent_work_complete`
  - `critical`
  - `planned_start`
  - `planned_finish`
  - `planned_duration`
  - `planned_duration_hours`
  - `is_milestone`
  を優先して扱う
- first cut では、project の基本項目
  - `name`
  - `title`
  - `author`
  - `company`
  - `start_date`
  - `finish_date`
  - `current_date`
  - `status_date`
  - `calendar_uid`
  - `minutes_per_day`
  - `minutes_per_week`
  - `days_per_month`
  - `schedule_from_start`
  を `update_project` で扱う
- `update_project` は `uid` を持たず、project 全体に対する単一 op として `fields` を受ける
- `update_project.name` は空でない文字列として扱う
- `update_project.title` / `author` / `company` は文字列として扱い、空文字を指定した場合はクリアしてよい
- `update_project.start_date` / `finish_date` / `current_date` / `status_date` は date-only または date-time を受け、date-only の場合は project 既定勤務時間帯を補完して扱う
- `update_project.start_date > finish_date` になる変更は warning として無視する
- `update_project.calendar_uid` が既存 calendar を指していない場合は warning として無視する
- `update_project.minutes_per_day` / `minutes_per_week` / `days_per_month` は `0` より大きい数値として扱う
- `update_project.schedule_from_start` は boolean として扱う
- first cut では、assignment の基本項目
  - `start`
  - `finish`
  - `units`
  - `work`
  - `percent_work_complete`
  を `update_assignment` で扱う
- first cut では、resource の基本項目
  - `name`
  - `initials`
  - `group`
  - `calendar_uid`
  - `max_units`
  - `standard_rate`
  - `overtime_rate`
  - `cost_per_use`
  - `percent_work_complete`
  を `update_resource` で扱う
- first cut では、calendar の基本項目
  - `name`
  - `is_base_calendar`
  - `base_calendar_uid`
  を `update_calendar` で扱う
- first cut では、assignment の新規追加として `add_assignment` を扱う
- first cut では、task の基本計画項目は `update_task` に寄せ、依存関係は `link_tasks` / `unlink_tasks` に分離する
- `update_task.is_milestone` は既存 task の milestone 化 / 解除を扱ってよい
- `update_task.notes` は task の `Notes` を更新してよく、空文字を指定した場合は `Notes` をクリアしてよい
- `update_task.calendar_uid` は task の `CalendarUID` を更新してよく、空文字を指定した場合は個別 calendar 指定をクリアしてよい
- `update_task.calendar_uid` が既存 calendar を指していない場合は warning として無視する
- `update_task.percent_complete` / `percent_work_complete` は `0..100` の数値として扱う
- `update_task.critical` は boolean として扱う
- `update_task` では summary task を milestone 化してはならない
- `update_task.is_milestone=true` の場合は `planned_finish = planned_start`、`planned_duration = 0` に正規化する
- `update_resource` は既存 resource を `uid` で特定して部分更新する
- `update_resource.name` は空でない文字列として扱う
- `update_resource.initials` / `group` / `calendar_uid` は文字列として扱い、空文字を指定した場合はクリアしてよい
- `update_resource.calendar_uid` が既存 calendar を指していない場合は warning として無視する
- `update_resource.max_units` は `0` 以上の数値として扱う
- `update_resource.standard_rate` / `overtime_rate` は文字列として扱い、空文字を指定した場合はクリアしてよい
- `update_resource.cost_per_use` は `0` 以上の数値として扱う
- `update_resource.percent_work_complete` は `0..100` の数値として扱う
- `update_calendar` は既存 calendar を `uid` で特定して部分更新する
- `update_calendar.name` は空でない文字列として扱う
- `update_calendar.is_base_calendar` は boolean として扱う
- `update_calendar.base_calendar_uid` は文字列として扱い、空文字を指定した場合はクリアしてよい
- `update_calendar.base_calendar_uid` は既存 calendar を指す必要があり、自身を指してはならない
- `add_calendar` の first cut は `uid` / `name` を必須とする
- `add_calendar` では `is_base_calendar` / `base_calendar_uid` を任意で受けてよい
- `add_calendar.base_calendar_uid` は既存 calendar を指す必要があり、自身を指してはならない
- `delete_calendar` の first cut は `uid` を受け、project / task / resource / 他 calendar の `base_calendar_uid` から参照されていない calendar の単純削除だけを扱う
- `delete_calendar` が成功した場合、差分表示では対象 calendar の `Name` を `(deleted)` として示す
- `add_resource` の first cut は `uid` / `name` を必須とする
- `add_resource` では `initials` / `group` / `calendar_uid` / `max_units` / `standard_rate` / `overtime_rate` / `cost_per_use` / `percent_work_complete` を任意で受けてよい
- `add_resource.calendar_uid` は既存 calendar を指す必要がある
- `add_resource.max_units` は `0` 以上の数値として扱う
- `add_resource.standard_rate` / `overtime_rate` は文字列として扱う
- `add_resource.cost_per_use` は `0` 以上の数値として扱う
- `add_resource.percent_work_complete` は `0..100` の数値として扱う
- `delete_resource` の first cut は `uid` を受け、assignment が付いていない resource の単純削除だけを扱う
- `delete_resource` first cut では assignment が残っている resource は削除できない
- `delete_resource` が成功した場合、差分表示では対象 resource の `Name` を `(deleted)` として示す
- `update_assignment` は既存 assignment を `uid` で特定して部分更新する
- `update_assignment.start` / `finish` は date-only または date-time を受け、date-only の場合は project 既定勤務時間帯を補完して扱う
- `update_assignment.start > finish` になる変更は warning として無視する
- `update_assignment.units` は `0` 以上の数値として扱う
- `update_assignment.work` は空でない文字列として扱う
- `update_assignment.percent_work_complete` は `0..100` の数値として扱う
- `add_assignment` の first cut は `uid` / `task_uid` / `resource_uid` を必須とする
- `add_assignment.task_uid` は既存 task、`resource_uid` は既存 resource を指す必要がある
- `add_assignment` では `start` / `finish` / `units` / `work` / `percent_work_complete` を任意で受けてよい
- `add_assignment.start` / `finish` は date-only または date-time を受け、date-only の場合は project 既定勤務時間帯を補完して扱う
- `add_assignment.start > finish`、`units < 0`、空 `work`、`percent_work_complete` の範囲外は warning として無視する
- `delete_assignment` の first cut は `uid` を受け、1 件の assignment を単純削除する
- `delete_assignment` が成功した場合、差分表示では対象 assignment の `TaskUID / ResourceUID` を `(deleted)` として示す
- `add_task` の first cut は、単一 task の追加に絞って `uid` / `name` / `new_parent_uid` / `new_index` を受ける
- `add_task.new_index` は `0-based` とし、sibling 範囲外は warning として無視する
- `add_task.new_parent_uid` は root 追加では `null` を許容し、非 root の場合は summary task を指す必要がある
- `add_task` では通常 task / milestone / summary task を明示的に追加できる
- summary task の追加は「空 summary を 1 件追加する」形で扱い、subtree は同時投入しない
- summary task の子は、後続の `add_task` / `move_task` で段階的に入れる前提とする
- したがって summary task 追加の JSON は、first cut では `uid` / `name` / `is_summary=true` / `new_parent_uid` / `new_index` を基本とし、子 task 一括定義は持たない
- `add_task` では `is_summary=true` と `is_milestone=true` を同時に指定してはならない
- `add_task` では `planned_start` / `planned_finish` / `planned_duration` / `planned_duration_hours` を任意で受けてよい
- `add_task.planned_start` / `planned_finish` の形式不正や、`planned_start > planned_finish` は warning として無視する
- `add_task.is_milestone=true` の場合は `planned_finish = planned_start`、`planned_duration = 0` に正規化する
- `add_task` に未対応 key が含まれる場合は warning として無視する
- `delete_task` の first cut は `uid` を受け、葉 task の削除だけを扱う
- `delete_task` first cut では summary task や子 task を持つ task の削除は扱わない
- `delete_task` first cut では assignment が付いている task の削除は扱わない
- `delete_task` first cut では後続依存から参照されている task の削除は扱わない
- `delete_task` が拒否される場合、warning には `children` / `assignments` / `successors` などの blocker 情報を含めて返す
- `delete_task` が成功した場合、差分表示では削除対象 task の `Name / ParentUID / Position` を `(deleted)` として示す
- `delete_task` は cascade delete を行わず、削除したい subtree がある場合も、まず葉 task から順に削除する前提とする
- 削除前に依存関係や親子関係が blocker になる場合は、先に `unlink_tasks` や `move_task` で整理してから `delete_task` を行う
- `move_task` の first cut は `uid` / `new_parent_uid` / `new_index` を受け、subtree 単位の移動として扱う
- `move_task.new_index` は `0-based` とし、sibling 範囲外は warning として無視する
- `move_task.new_parent_uid` は root への移動では `null` を許容し、非 root の場合は summary task を指す必要がある
- `move_task` では task を自身または配下へ移動してはならない
- `move_task` で結果が変わらない no-op 移動は warning として無視する
- `predecessors` は `update_task.fields` には含めず、依存関係の変更は `link_tasks` / `unlink_tasks` へ分離する方針とする
- その理由は、依存関係が単なる field 更新ではなく task 間リンク更新であり、将来の `type / lag` 拡張や validation と相性がよいためである
- `link_tasks` の first draft は `from_uid` / `to_uid` を必須とし、`type` は省略時 `FS` を既定としてよい
- `unlink_tasks` の first draft も `from_uid` / `to_uid` を必須とし、必要なら `type` や `lag` / `lag_hours` を併記して解除対象を特定できる形とする
- `unlink_tasks` で複数の依存関係が同じ条件に一致した場合は、その条件に一致した link をすべて解除する方針とする
- first draft では、依存関係の変更は predecessor 一覧の全置換ではなく、追加と解除を op 単位で返す方針とする
- 未知 `op`、存在しない `uid`、不正 field、不正日付は validation または import warning/error の対象とする
- `project_draft_view` は新規草案作成を優先するため、非稼働日を厳密に考慮しなくてもよい
- `project_draft_view` では、粗い草案でもよいので task に仮の `planned_start` / `planned_finish` を入れてよい
- この仮日付は通常 task だけでなく、summary task と milestone にも入れてよい
- summary task には配下 task を大まかに包む期間、milestone には節目の日付を仮で入れる方針を許容する
- 人間が「とりあえずえいやで日付を入れてよい」と指示している場合は、細かい整合よりも日付充足を優先してよい
- `calendar_uid = "1"` は既定の `Standard` calendar を指し、土曜日と日曜日を非稼働日として扱う前提とする
- task / resource に個別 `CalendarUID` がない場合は、project 既定 calendar を継承する前提で扱う
- 稼働日・祝日を考慮した日付補正は、後続の Patch JSON による再計画で扱う
- Patch JSON の `planned_start` / `planned_finish` 更新は、原則として非稼働日を避ける前提とする
- ただし、人間が明示的に非稼働日での実施を指示した場合は、その指示を優先できる
- 非稼働日ルールより人間指示を優先した場合は、AI 側の説明文でもその例外適用を明示する

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
- `0 < percentComplete < 100` は `active` として出力する
- `percentComplete >= 100` は `done` として出力する
- task 名や title は Mermaid で壊れやすい一部記号を簡易正規化して出力する
- predecessor は、`単一 predecessor` かつ `FS` かつ `lag なし` かつ `duration` を Mermaid 向けへ素直に変換できる task のみ `after ...` でネイティブ出力する
- 上記に当てはまらない predecessor は、task 名を含むコメント行で補助出力する
- comment 側の `lag` は、可能な範囲で `2h` のような短い人間向け表現に整形して出力する
- `lag` がある場合は、`after Prep + 2h` のような擬似読解用 comment も追加する
- preview と SVG export は `WBS SVG` 描画を使う
- phase 背景は交互の淡色で視認補助する

`project_draft_view` からの補完メモ:

- 通常 task で `planned_start` / `planned_finish` が date-only の場合、内部化時に勤務時間帯として `09:00:00` / `18:00:00` を補完して扱うことがある
- `planned_finish` だけが与えられた通常 task は、まず同日の `planned_start` を補完し、その後に上記の勤務時間帯補完を適用する
- `is_milestone=true` の task には、この勤務時間帯補完を適用しない
- 必要に応じて、最小の `resources` と `assignments` を併記してよい
- 現時点の `assignments` は `task_uid / resource_uid / start / finish / units / work / percent_work_complete` 程度の最小表現を想定する

現時点で意図的に落とすもの:

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
