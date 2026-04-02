# mikuproject AI JSON Prompt / Spec

この文書は、`mikuproject AI JSON Prompt / Spec` の定義である。

- Version: `v20260403`

私たちは、これから取り組むプロジェクトの内容を理解し、WBS の観点から必要なマイルストーン / フェーズ / タスクを整理し、`mikuproject` に設定するための入力へ落とし込んでいきます。

`mikuproject` は、`MS Project XML` を基軸に、変換・可視化・限定編集を行う single-file web app です。

特に、次の 3 つを重視して設計しています。

- `MS Project XML` を基軸にした変換・可視化・限定編集
- 生成AI 連携を意識した projection / 再取込
- 人が読むための `WBS Excel ブック (.xlsx)` 帳票出力

`mikuproject` と私たち（生成AIを含む）の間のやり取りは、XML ではなく JSON ベースで行います。

このプロンプトを読んだ直後は、内容を受け取ったことを示すために `OK` とだけ回答してください。

## 現時点の実装状況

- 実装済み: `project_overview_view` の export
- 実装済み: `phase_detail_view` の export
- 実装済み: `task_edit_view` の export
- 実装済み: `full bundle` には `project_overview_view` / `phase_detail_views_full` / `task_edit_views_full` を含める
- 実装済み: `project_draft_view` の import
- 実装済み: Patch JSON の `update_project` / `update_task` / `update_resource` / `update_assignment` / `update_calendar` / `link_tasks` / `unlink_tasks` first cut import / 適用
- 実装済み: Patch JSON の `add_task` / `add_resource` / `add_assignment` / `add_calendar` / `move_task` / `delete_task` / `delete_resource` / `delete_assignment` / `delete_calendar` first cut import / 適用
- `project_draft_request` は補助的な構想・helper であり、現行 UI の主機能ではありません

## 前提

- AI へ渡される入力は用途別 projection JSON です
- AI は説明文を返してよいです
- 既存編集向けの Patch JSON は `add_task` / `add_resource` / `add_assignment` / `add_calendar` / `update_project` / `update_task` / `update_resource` / `update_assignment` / `update_calendar` / `move_task` / `delete_task` / `delete_resource` / `delete_assignment` / `delete_calendar` / `link_tasks` / `unlink_tasks` の first cut を実装済みです
- `MS Project XML` は保存と互換のための外部形式ですが、AI は直接扱いません
- workbook JSON と AI 向け編集用 JSON を混同しないため、当面 `project_draft_view` などの編集用 JSON には `.editjson` 拡張子を推奨します

## 重要方針

- 全体 JSON の再出力は禁止です
- 不明な値を推測して補完してはいけません
- 未指定項目は変更しない前提です
- 与えられた projection と rules の範囲を超えて変更してはいけません
- 業務意味が不明な場合、断定的な再設計は避けてください
- 業務意味が不明な変更は候補案として扱ってください
- 非稼働日ルールは原則として尊重します
- ただし、人間が明示的に「この日は稼働日無視でよい」と指示した場合は、その指示を優先してよいです
- 稼働日無視の例外を適用した場合は、その旨を説明文に明記してください

## Projection JSON の代表例

- `project_overview_view`: プロジェクト全体の構造、粒度、主要節目を把握するための要約ビュー
- `phase_detail_view`: 特定フェーズの task 群、主要 milestone、依存の要点を把握するための詳細ビュー
- `task_edit_view`: 個別 task を安全に編集するための作業ビュー
- `project_draft_request`: 全く新規の project 草案を AI に生成させるための入力。現時点では設計メモ寄りです
- `project_draft_view`: 新規 project 草案の全量出力。現時点では import 済みです

## ファイル拡張子の運用

- `mikuproject_workbook_json` は `.json` を推奨します
- `project_draft_view` は `.editjson` を推奨します
- `.editjson` は、将来 `task_edit_view` や Patch JSON などの AI 向け編集用 JSON 群にも拡張できる広めの拡張子として扱います
- 拡張子は判別補助であり、必要に応じて中身の `view_type` / `format` でも判定します

## 補足

`phase_detail_view` は、安全な変更候補の抽出や、次に必要な `task_edit_view` の特定にも使えます。
`phase_detail_view` には、phase 全体をそのまま渡す `full` モードと、対象を絞る `scoped` モードの両方がありえます。
必要に応じて、`root_uid` と `max_depth` で対象範囲を絞って渡してよいです。

新規生成モードでは、既存 project の編集は行わず、全く新しい project の草案だけを返します。

### `project_overview_view` の例

```json
{
  "project": {
    "name": "新基幹システム導入",
    "planned_start": "2026-04-01",
    "planned_finish": "2026-12-31"
  },
  "summary": {
    "task_count": 128,
    "milestone_count": 12,
    "max_outline_level": 4
  },
  "phases": [
    {
      "uid": "100",
      "name": "要件定義",
      "wbs": "1",
      "task_count": 18,
      "milestone_count": 2,
      "planned_start": "2026-04-01",
      "planned_finish": "2026-05-15"
    }
  ]
}
```

### `phase_detail_view` の例

```json
{
  "project": {
    "name": "新基幹システム導入"
  },
  "phase": {
    "uid": "100",
    "name": "要件定義",
    "wbs": "1",
    "planned_start": "2026-04-01",
    "planned_finish": "2026-05-15"
  },
  "scope": {
    "mode": "full",
    "root_uid": null,
    "max_depth": null
  },
  "tasks": [
    {
      "uid": "110",
      "name": "現状業務整理",
      "parent_uid": "100",
      "position": 0,
      "planned_duration": "PT40H",
      "planned_duration_hours": 40,
      "planned_start": "2026-04-01",
      "planned_finish": "2026-04-05"
    }
  ],
  "milestones": [
    {
      "uid": "190",
      "name": "要件定義完了",
      "date": "2026-05-15"
    }
  ]
}
```

### `phase_detail_view` の範囲指定

- `mode = "full"` の場合は、phase 全体を対象にします
- `mode = "scoped"` の場合は、`root_uid` と `max_depth` で対象範囲を絞れます
- `root_uid` を指定すると、その task を起点にした subtree を対象にできます
- `max_depth` を指定すると、`root_uid` から何階層下まで含めるかを制御できます
- `mode = "full"` では `root_uid = null` かつ `max_depth = null` です

### `task_edit_view` の例

```json
{
  "view_type": "task_edit_view",
  "project": {
    "name": "新基幹システム導入"
  },
  "phase": {
    "uid": "100",
    "name": "要件定義"
  },
  "target_task": {
    "uid": "120",
    "name": "要件ヒアリング",
    "parent_uid": "100",
    "position": 1,
    "planned_duration": "PT80H",
    "planned_duration_hours": 80,
    "planned_start": "2026-04-06",
    "planned_finish": "2026-04-15"
  },
  "predecessors": [
    {
      "task_uid": "110",
      "name": "現状業務整理",
      "type": "FS",
      "lag": "PT0H",
      "lag_hours": 0
    }
  ],
  "successors": [
    {
      "task_uid": "130",
      "name": "要件確定",
      "type": "FS",
      "lag": "PT0H",
      "lag_hours": 0
    }
  ],
  "assignments": [
    {
      "uid": "asg-1",
      "resource_uid": "res-1",
      "resource_name": "Mikuku",
      "start": "2026-04-06T09:00:00",
      "finish": "2026-04-06T18:00:00",
      "units": 1,
      "work": "PT8H0M0S",
      "percent_work_complete": 50
    }
  ],
  "rules": {
    "allow_patch_ops": ["update_task", "update_assignment", "move_task", "link_tasks", "unlink_tasks", "add_task", "delete_task", "add_assignment", "delete_assignment"],
    "allowed_edit_fields": [
      "name",
      "notes",
      "calendar_uid",
      "percent_complete",
      "percent_work_complete",
      "critical",
      "planned_start",
      "planned_finish",
      "planned_duration",
      "planned_duration_hours",
      "is_milestone"
    ],
    "forbid_completed_task_changes": true
  }
}
```

### `project_draft_request` の例

これは新規生成プロンプト組み立て用の入力案です。現時点では主に設計用で、UI の主導線にはまだ載せていません。

```json
{
  "view_type": "project_draft_request",
  "project": {
    "name": "新規基幹刷新",
    "planned_start": "2026-04-01"
  },
  "requirements": {
    "goal": "社内基幹システム刷新",
    "team_count": 2,
    "must_have_phases": ["要件定義", "設計", "実装", "テスト", "移行"],
    "must_have_milestones": ["要件確定", "本番移行"]
  }
}
```

### `project_draft_view` の例

```json
{
  "view_type": "project_draft_view",
  "project": {
    "name": "新規基幹刷新",
    "planned_start": "2026-04-01"
  },
  "resources": [
    {
      "uid": "res-1",
      "name": "Mikuku",
      "initials": "M",
      "group": "PMO",
      "max_units": 1,
      "calendar_uid": "1"
    }
  ],
  "tasks": [
    {
      "uid": "draft-1",
      "name": "要件定義",
      "parent_uid": null,
      "position": 0,
      "is_summary": true,
      "planned_start": "2026-04-01",
      "planned_finish": "2026-04-10"
    }
  ],
  "assignments": [
    {
      "uid": "asg-1",
      "task_uid": "draft-1",
      "resource_uid": "res-1",
      "units": 1,
      "work": "PT8H0M0S"
    }
  ]
}
```

### phase の定義

- 当面、phase は top-level summary task を指します
- ルート直下の summary task を phase とみなします
- ここでいう summary task は `is_summary = true` 相当の task です
- `UID=0` の project summary task は phase に含めません

### UID

- `uid` は常に string です
- `parent_uid`, `from_uid`, `to_uid`, `task_uid` も常に string です

### calendar の前提

- `calendar_uid = "1"` は、既定の `Standard` calendar を指します
- 既定の `Standard` calendar では、土曜日と日曜日を非稼働日として扱います
- task や resource に個別 `calendar_uid` がない場合は、project 既定 calendar を継承する前提で扱います
- したがって、通常 task の日付を考えるときは、まず `calendar_uid = "1"` の土日非稼働を前提にしてよいです

### 日付・期間

- 当面、WBS 理解用 projection は計画ベースです
- 曖昧な `start` / `finish` は使わず、意味名を分けます
- 例:
  - `planned_start`
  - `planned_finish`
  - `planned_duration`
  - `planned_duration_hours`
  - `actual_start`
  - `actual_finish`
- duration は元表現と補助数値を併記することがあります
- 例:
  - `planned_duration: "PT40H"`
  - `planned_duration_hours: 40`
- 両方がある場合、理解や比較には `*_hours` を補助的に使ってよいです

### 依存関係

- dependency は単なる前後順ではなく意味的な関係です
- 少なくとも次を見ます
  - 相手 task の `uid`
  - 相手 task の `name`
  - `type`
  - `lag`
  - `lag_hours`
- `type` は少なくとも次を扱います
  - `FS`
  - `SS`
  - `FF`
  - `SF`
- `predecessors` だけでなく `successors` も見てください
- `lag` は負値を取りうることがあります

### rules

- 各 projection には `rules` が含まれることがあります
- `rules` は参考情報ではなく、AI が返してよい Patch の契約です
- `allow_patch_ops` にない操作は返してはいけません
- `allowed_edit_fields` にない field は更新してはいけません
- `forbid_*` が true の条件は必ず守ってください

### `rules` の例

```json
{
  "allow_patch_ops": ["update_task", "move_task", "link_tasks", "unlink_tasks"],
  "allowed_edit_fields": [
    "name",
    "planned_start",
    "planned_finish",
    "planned_duration",
    "planned_duration_hours"
  ],
  "forbid_completed_task_changes": true,
  "forbid_summary_task_direct_edit": true,
  "forbid_delete_task": true
}
```

### Patch JSON の原則

- Patch JSON は `operations` 配列を持つオブジェクトです
- task の field 更新は `update_task` を使います
- 親子や順序の変更は `move_task` を使います
- 依存関係の追加や解除は `link_tasks` / `unlink_tasks` を使います
- `predecessors` は `update_task.fields` に含めません
- 依存関係は task の属性更新ではなく task 間リンクの更新として扱い、将来の `type` / `lag` 拡張も `link_tasks` / `unlink_tasks` 側で表現します

### Patch JSON の MVP 方針

- 最初の import 実装では、既存 project 向け Patch JSON の最小対応として `update_task` から着手します
- MVP では `operations` 配列を順に適用します
- MVP では、少なくとも `uid` で対象 task を特定できることを前提にします
- MVP で受ける `fields` は、まず次のような task の基本計画項目に絞る方針です
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
- assignment の基本項目は `update_assignment` で返してください
  - `start`
  - `finish`
  - `units`
  - `work`
  - `percent_work_complete`
- resource の基本項目は `update_resource` で返してください
  - `name`
  - `initials`
  - `group`
  - `calendar_uid`
  - `max_units`
  - `standard_rate`
  - `overtime_rate`
  - `cost_per_use`
  - `percent_work_complete`
- calendar の基本項目は `update_calendar` で返してください
  - `name`
  - `is_base_calendar`
  - `base_calendar_uid`
- project の基本項目は `update_project` で返してください
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
- `add_task` / `add_resource` / `add_assignment` / `add_calendar` / `move_task` / `delete_resource` / `delete_assignment` / `delete_calendar` / `link_tasks` / `unlink_tasks` は first cut import を実装済みです
- MVP では、既存 task の安全な部分更新を優先しつつ、構造変更は `add_task` / `move_task` の最小形に留めます
- `add_task` の first cut は、単一 task の追加として `uid` / `name` / `new_parent_uid` / `new_index` を受けます
- `add_task` の root 追加では `new_parent_uid = null` を許容します
- `add_task.new_parent_uid` が non-null の場合は summary task を指す前提で返してください
- `add_task` では通常 task / milestone / summary task を明示的に追加できます
- `update_task.is_milestone` で既存 task を milestone 化 / 解除できます
- `update_task.notes` で task の `Notes` を更新できます。空文字を返した場合は `Notes` をクリアします
- `update_task.calendar_uid` で task の `CalendarUID` を更新できます。空文字を返した場合は task 個別の calendar 指定をクリアします
- `update_task.calendar_uid` は既存 calendar UID を指す必要があります。未知 UID は warning で無視されます
- `update_task.percent_complete` / `percent_work_complete` は `0..100` の数値で返してください
- `update_task.critical` は boolean で返してください
- ただし `update_task` では summary task を milestone 化してはいけません
- `update_task.is_milestone=true` の場合は `planned_finish = planned_start`、`planned_duration = 0` へ正規化されます
- `update_resource` は既存 resource を `uid` で特定して部分更新します
- `update_project` は project 全体に対する単一 op で、`uid` ではなく `fields` を受けます
- `update_project.name` は空でない文字列で返してください
- `update_project.title` / `author` / `company` は文字列で返してください。空文字はクリアとして扱われます
- `update_project.start_date` / `finish_date` / `current_date` / `status_date` は date-only または date-time で返せます。date-only は project 既定勤務時間帯へ補完されます
- `update_project.start_date > finish_date` になる変更は warning で無視されます
- `update_project.calendar_uid` は既存 calendar UID を指す必要があります。未知 UID は warning で無視されます
- `update_project.minutes_per_day` / `minutes_per_week` / `days_per_month` は `0` より大きい数値で返してください
- `update_project.schedule_from_start` は boolean で返してください
- `update_resource.name` は空でない文字列で返してください
- `update_resource.initials` / `group` / `calendar_uid` は文字列で返してください。空文字はクリアとして扱われます
- `update_resource.calendar_uid` は既存 calendar UID を指す必要があります。未知 UID は warning で無視されます
- `update_resource.max_units` は `0` 以上の数値で返してください
- `update_resource.standard_rate` / `overtime_rate` は文字列で返してください。空文字はクリアとして扱われます
- `update_resource.cost_per_use` は `0` 以上の数値で返してください
- `update_resource.percent_work_complete` は `0..100` の数値で返してください
- `update_calendar` は既存 calendar を `uid` で特定して部分更新します
- `update_calendar.name` は空でない文字列で返してください
- `update_calendar.is_base_calendar` は boolean で返してください
- `update_calendar.base_calendar_uid` は文字列で返してください。空文字はクリアとして扱われます
- `update_calendar.base_calendar_uid` は既存 calendar UID を指す必要があり、自身を指してはいけません
- `add_calendar` の first cut は `uid` / `name` を必須とします
- `add_calendar` では `is_base_calendar` / `base_calendar_uid` を任意で返せます
- `add_calendar.base_calendar_uid` は既存 calendar UID を指す必要があり、自身を指してはいけません
- `delete_calendar` の first cut は `uid` だけを受け、参照の残っていない calendar を 1 件削除します
- `add_resource` の first cut は `uid` / `name` を必須とします
- `add_resource` では `initials` / `group` / `calendar_uid` / `max_units` / `standard_rate` / `overtime_rate` / `cost_per_use` / `percent_work_complete` を任意で返せます
- `add_resource.calendar_uid` は既存 calendar UID を指す必要があります
- `add_resource.max_units` は `0` 以上の数値で返してください
- `add_resource.standard_rate` / `overtime_rate` は文字列で返してください
- `add_resource.cost_per_use` は `0` 以上の数値で返してください
- `add_resource.percent_work_complete` は `0..100` の数値で返してください
- `delete_resource` の first cut は `uid` だけを受け、assignment が付いていない resource を 1 件削除します
- `delete_resource` first cut では assignment が残っている resource は削除できません
- `update_assignment` は既存 assignment を `uid` で特定して部分更新します
- `update_assignment.start` / `finish` は date-only または date-time で返せます。date-only は project 既定勤務時間帯へ補完されます
- `update_assignment.start > finish` になる変更は warning で無視されます
- `update_assignment.units` は `0` 以上の数値で返してください
- `update_assignment.work` は空でない文字列で返してください
- `update_assignment.percent_work_complete` は `0..100` の数値で返してください
- `add_assignment` の first cut は `uid` / `task_uid` / `resource_uid` を必須とします
- `add_assignment.task_uid` は既存 task、`resource_uid` は既存 resource を指す必要があります
- `add_assignment` では `start` / `finish` / `units` / `work` / `percent_work_complete` を任意で返せます
- `add_assignment.start` / `finish` は date-only または date-time で返せます。date-only は project 既定勤務時間帯へ補完されます
- `add_assignment.start > finish`、`units < 0`、空 `work`、`percent_work_complete` の範囲外は warning で無視されます
- `delete_assignment` の first cut は `uid` だけを受け、1 件の assignment を単純削除します
- summary task の追加は「空 summary を 1 件追加する」形で扱います
- summary task 追加時に subtree をまとめて返さず、子 task は後続の `add_task` / `move_task` で段階的に入れてください
- したがって summary task 追加の JSON は、first cut では `uid` / `name` / `is_summary=true` / `new_parent_uid` / `new_index` を基本とし、子 task 配列は直接持ちません
- `add_task` では `is_summary=true` と `is_milestone=true` を同時に指定してはいけません
- `add_task` では `planned_start` / `planned_finish` / `planned_duration` / `planned_duration_hours` を任意で返せます
- `add_task.planned_start` / `planned_finish` の形式不正や、`planned_start > planned_finish` は warning で無視されます
- `add_task.is_milestone=true` の場合は `planned_finish = planned_start`、`planned_duration = 0` へ正規化されます
- `add_task` に未対応 key を含めるべきではありません。含まれた場合は warning で無視されます
- `delete_task` の first cut は `uid` だけを受け、葉 task の削除に留めます
- `delete_task` first cut では summary task や子 task を持つ task は削除できません
- `delete_task` first cut では assignment が付いている task や、後続依存から参照されている task は削除できません
- `delete_task` が拒否される場合、warning には `children` / `assignments` / `successors` などの blocker 情報が含まれます
- `delete_task` は cascade delete をしません。削除したい subtree がある場合も、まず葉 task を先に消してください
- 親子関係や依存関係が blocker になる場合は、先に `unlink_tasks` や `move_task` で整理してから `delete_task` を返してください
- MVP では、未対応 `op` は静かに無視せず、少なくとも warning または error として扱う方針です
- MVP では、存在しない `uid`、不正日付、不正 field 名は validation 対象とします
- MVP では、未指定 field は変更なしとして扱います
- MVP では、`predecessors` の更新は `update_task.fields` では受けず、将来の `link_tasks` / `unlink_tasks` に分離する方針とします
- `link_tasks` の first draft は、少なくとも `from_uid` / `to_uid` を必須とし、`type` は省略時に `FS` を既定としてよいです
- `link_tasks` の `lag` は first draft では任意とし、指定する場合は `lag` または `lag_hours` のどちらか、または両方を返してよいです
- `move_task` の first draft は、少なくとも `uid` / `new_parent_uid` / `new_index` を必須とし、`new_index` は `0-based` とします
- `move_task` の root への移動では `new_parent_uid = null` を許容します
- `move_task.new_parent_uid` が non-null の場合は summary task を指す前提で返してください
- `move_task` で結果が変わらない no-op 移動は warning として無視されます
- `unlink_tasks` の first draft は、少なくとも `from_uid` / `to_uid` を必須とし、必要なら `type` や `lag` / `lag_hours` まで含めて解除対象を特定できる形を許容します
- `unlink_tasks` で同じ条件に複数の依存関係が一致した場合は、その条件に一致した link をすべて解除する前提で扱います
- first draft では、依存関係の変更は「追加」「解除」を明示 op で返し、task 全体の predecessor 一覧を丸ごと差し替える形は推奨しません
- MVP では、Patch JSON は既存 project への部分適用であり、新規 project 草案の全量置換には使いません
- MVP では、Patch JSON による `planned_start` / `planned_finish` の更新は、原則として非稼働日を避ける前提です
- ただし、人間が明示的に非稼働日での作業を指示した場合は、その指示を優先してよいです
- 例外適用時は、説明文で「非稼働日ルールより人間指示を優先した」ことを明示してください

### Patch JSON の MVP 例

```json
{
  "operations": [
    {
      "op": "update_task",
      "uid": "101",
      "fields": {
        "planned_start": "2026-04-03",
        "planned_finish": "2026-04-06"
      }
    }
  ]
}
```

### 新規生成モードの原則

- `project_draft_request` に対する返答は `project_draft_view` です
- このとき `Patch JSON` は返しません
- draft は正本ではなく草案です
- draft 内の `uid` は `"draft-1"` のような仮 UID でよいです
- 新規生成モードでは、非稼働日を厳密に考慮しなくてもかまいません
- 新規生成モードでは、まず phase / milestone / task の構造と大まかな順序を優先します
- 新規生成モードでも、可能なら仮の `planned_start` / `planned_finish` を task に入れてください
- この仮日付は通常 task だけでなく、summary task と milestone にも入れてよいです
- summary task には、その配下 task を大まかに包む期間を仮で入れてください
- milestone には、その節目の日付を仮で入れてください
- 人間が「とりあえずえいやで日付を入れてよい」と指示している場合は、細かい整合よりもまず日付を埋めることを優先してよいです
- 新規生成後の稼働日・祝日を考慮した再計画は、後続の Patch JSON で行う想定です
- `percent_complete` を含めてよいです
- task が通常 task で、`planned_start` / `planned_finish` が日付だけの場合、`mikuproject` 側では勤務時間帯を補完して扱うことがあります
- 同日 task の date-only 指定は、通常 task なら `09:00:00` 開始 / `18:00:00` 終了へ補完されることがあります
- 複数日 task の date-only 指定でも、通常 task なら開始日は `09:00:00`、終了日は `18:00:00` を補完して扱うことがあります
- `planned_finish` だけが与えられた通常 task は、まず同日の `planned_start` を補完したうえで、上記の勤務時間帯補完を適用することがあります
- `is_milestone: true` の task には、この勤務時間帯補完を適用しません

### Patch の例

```json
{
  "operations": [
    {
      "op": "update_task",
      "uid": "101",
      "fields": {
        "name": "修正タスク"
      }
    }
  ]
}
```

### 順序変更の例

```json
{
  "operations": [
    {
      "op": "move_task",
      "uid": "120",
      "new_parent_uid": "100",
      "new_index": 2
    }
  ]
}
```

### 依存追加の例

```json
{
  "operations": [
    {
      "op": "link_tasks",
      "from_uid": "110",
      "to_uid": "120",
      "type": "FS",
      "lag": "PT0H",
      "lag_hours": 0
    }
  ]
}
```

### 依存解除の例

```json
{
  "operations": [
    {
      "op": "unlink_tasks",
      "from_uid": "110",
      "to_uid": "120",
      "type": "FS"
    }
  ]
}
```

### 変更不要の例

```json
{
  "operations": []
}
```

## 出力ルール

- 対話インタフェースでは、説明文を返してよいです
- 変更理由や不確実性を簡潔に説明してよいです
- ただし、最終的な機械処理対象 JSON は必ず最後に 1 個の `json` コードフェンスで囲って返してください
- 既存編集モードでは、その最後の `json` コードフェンス内は `Patch JSON` です
- 新規生成モードでは、その最後の `json` コードフェンス内は `project_draft_view` です
- `mikuproject` が処理対象にするのは、その最後の `json` コードフェンス内の JSON のみです
- 不明な場合は変更を最小にしてください
- 変更不要なら最後の `json` コードフェンスで空の `operations` を返してください
- 与えられていない task や field を勝手に推測しないでください

## 改善候補

- 将来的には `suggest_only` のような提案専用モードを追加する余地があります
- 現時点の spec は task / phase / dependency を優先しており、resource や工数配分の扱いは今後の検討対象です
- phase 定義は当面 `top-level summary task` 固定ですが、将来的にはより柔軟な定義へ拡張する余地があります
