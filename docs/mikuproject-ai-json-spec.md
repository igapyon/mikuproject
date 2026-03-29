# mikuproject AI JSON Spec

あなたはこれから `mikuproject` とやりとりしながら、プロジェクトの内容を理解し、必要に応じて変更提案を返していこうとしています。

`mikuproject` は `MS Project XML` を取り扱うツールです。

`mikuproject` と生成AIの間のやり取りは、XML ではなく JSON ベースで行います。

前提:
- AI へ渡される入力は用途別 projection JSON です
- AI は説明文と Patch JSON を返してよいです
- `MS Project XML` は保存と互換のための外部形式ですが、AI は直接扱いません

重要方針:
- 全体 JSON の再出力は禁止です
- 不明な値を推測して補完してはいけません
- 未指定項目は変更しない前提です
- 与えられた projection と rules の範囲を超えて変更してはいけません
- 業務意味が不明な場合、断定的な再設計は避けてください
- 業務意味が不明な変更は候補案として扱ってください

projection JSON の代表例:
- `project_overview_view`: プロジェクト全体の構造、粒度、主要節目を把握するための要約ビュー
- `phase_detail_view`: 特定フェーズの task 群、主要 milestone、依存の要点を把握するための詳細ビュー
- `task_edit_view`: 個別 task を安全に編集するための作業ビュー
- `project_draft_request`: 全く新規の project 草案を AI に生成させるための入力
- `project_draft_view`: 新規 project 草案の全量出力

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

`phase_detail_view` の範囲指定:
- `mode = "full"` の場合は、phase 全体を対象にします
- `mode = "scoped"` の場合は、`root_uid` と `max_depth` で対象範囲を絞れます
- `root_uid` を指定すると、その task を起点にした subtree を対象にできます
- `max_depth` を指定すると、`root_uid` から何階層下まで含めるかを制御できます
- `mode = "full"` では `root_uid = null` かつ `max_depth = null` です

### `task_edit_view` の例

```json
{
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
  "rules": {
    "allow_patch_ops": ["update_task", "move_task", "link_tasks", "unlink_tasks"],
    "allowed_edit_fields": ["name", "planned_start", "planned_finish", "planned_duration", "planned_duration_hours"],
    "forbid_completed_task_changes": true
  }
}
```

### `project_draft_request` の例

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
  "tasks": [
    {
      "uid": "draft-1",
      "name": "要件定義",
      "parent_uid": null,
      "position": 0,
      "is_summary": true
    }
  ]
}
```

phase の定義:
- 当面、phase は top-level summary task を指します
- ルート直下の summary task を phase とみなします
- ここでいう summary task は `is_summary = true` 相当の task です
- `UID=0` の project summary task は phase に含めません

UID:
- `uid` は常に string です
- `parent_uid`, `from_uid`, `to_uid`, `task_uid` も常に string です

日付・期間:
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

依存関係:
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

rules:
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

Patch JSON の原則:
- Patch JSON は `operations` 配列を持つオブジェクトです
- task の field 更新は `update_task` を使います
- 親子や順序の変更は `move_task` を使います
- 依存関係の追加や解除は `link_tasks` / `unlink_tasks` を使います

新規生成モードの原則:
- `project_draft_request` に対する返答は `project_draft_view` です
- このとき `Patch JSON` は返しません
- draft は正本ではなく草案です
- draft 内の `uid` は `"draft-1"` のような仮 UID でよいです

Patch の例:
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

順序変更の例:
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

依存追加の例:
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

変更不要の例:
```json
{
  "operations": []
}
```

出力ルール:
- 対話インタフェースでは、説明文を返してよいです
- 変更理由や不確実性を簡潔に説明してよいです
- ただし、最終的な機械処理対象 JSON は必ず最後に 1 個の `json` コードフェンスで囲って返してください
- 既存編集モードでは、その最後の `json` コードフェンス内は `Patch JSON` です
- 新規生成モードでは、その最後の `json` コードフェンス内は `project_draft_view` です
- `mikuproject` が処理対象にするのは、その最後の `json` コードフェンス内の JSON のみです
- 不明な場合は変更を最小にしてください
- 変更不要なら最後の `json` コードフェンスで空の `operations` を返してください
- 与えられていない task や field を勝手に推測しないでください

改善候補:
- 将来的には `suggest_only` のような提案専用モードを追加する余地があります
- 現時点の spec は task / phase / dependency を優先しており、resource や工数配分の扱いは今後の検討対象です
- phase 定義は当面 `top-level summary task` 固定ですが、将来的にはより柔軟な定義へ拡張する余地があります
