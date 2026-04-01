# MS Project XML x 生成AI 連携設計メモ

この文書は、`mikuproject` における `MS Project XML` と生成AIの連携方針を整理するための設計メモである。

現時点では実装仕様の確定版ではなく、アーキテクチャ判断と今後の仕様化ポイントを明確にすることを目的とする。

現時点の実装範囲は次のとおり。

- 実装済み: `project_overview_view` export
- 実装済み: `phase_detail_view` export
- 実装済み: `project_draft_view` import
- 未実装: `task_edit_view`
- 未実装: Patch JSON の受信・適用

この設計では、生成AIとの会話は `JSON` ベースで行う方針とする。

- AI へ渡す入力は用途別 projection JSON
- AI から受け取る出力は、現状実装では `project_draft_view`、将来案としては Patch JSON を想定する
- workbook JSON と AI 向け編集用 JSON を区別するため、当面 `project_draft_view` などの編集用 JSON には `.editjson` 拡張子を推奨する

`MS Project XML` は保存と互換のための外部形式として保持し、AI との直接入出力には使わない。

## 目的

`mikuproject` は `MS Project XML` を忠実に扱う編集ソフトとして設計する。

すでに `.xlsx` の import/export は実装済みであり、次の課題として生成AIとの安全な連携を検討する。

- `WBS` やプロジェクト構造を、AIが扱いやすい形で渡す
- AIの編集結果を、安全かつ検証可能な形で受け取る
- `MS Project XML` との互換性を維持する

なお、現時点の UI で扱える生成AI連携は、既存 project に対する `project_overview_view` / `phase_detail_view` の保存と、新規草案として返ってきた `project_draft_view` の取込までである。

## 最重要方針

結論として、生成AIとの会話は `JSON` ベースで行い、`MS Project XML` をそのまま AI に渡さない。

代わりに次のレイヤで分離する。

- 外部形式: `MS Project XML`
- 内部形式: 正規化された canonical model
- AI 入出力形式: 用途別 projection JSON と、将来案としての Patch JSON

要点は次のとおり。

- XML は保存と互換のための正本として扱う
- 内部では意味ベースの正規化モデルを扱う
- AI とは JSON ベースでやり取りする
- AI からの返却は、現状では `project_draft_view` の取込、将来は差分 Patch の受取を想定する

一言でまとめると、`XML は保存、JSON は思考、Patch は将来の操作` である。

## レイヤ構造

想定するデータ構造のレイヤは次のとおり。

### 1. 外部形式

- `MS Project XML`
- 正本
- 互換維持のための保存形式

### 2. 内部モデル

- 正規化された `Project / Task / Resource / Assignment / Calendar` モデル
- `mikuproject` 内部で意味を保って扱う canonical representation

### 3. AI 入出力形式

- AI入力: 用途別に投影した JSON
- AI出力: Patch JSON

この分離により、AI側の都合と `MS Project XML` 側の都合を疎結合にできる。

## なぜ XML をそのまま AI に渡さないか

`MS Project XML` は互換形式としては有用だが、AIとの直接連携には不向きである。

理由は次のとおり。

- XML の階層と要素数が多くノイズが多い
- AI にとって意味単位が読み取りづらい
- XML 全量を書き換えさせると破損リスクが高い
- 差分検証や部分適用が難しい

したがって、AI 連携では XML の完全再構成を求めるのではなく、意味ベースの JSON と差分操作に変換して扱う。

## AI 向け JSON 設計方針

AI 向け JSON は、XML をそのまま JSON 化したものではなく、意味ベースで正規化した構造をもとに、目的別に投影した projection とする。

### NG

- XML を機械的に JSON 化する
- 互換性都合の細かいフィールドをそのまま露出する
- AI に全フィールドの再出力を求める
- AI 用 JSON を 1 種類だけで済ませようとする

### OK

- タスクや依存関係など、意味のある単位に正規化する
- 今回の判断に必要な情報だけを出す
- 型と命名を一貫させる
- AI に理解しやすい構造を優先する
- 用途ごとに projection を切り替える

例:

```json
{
  "tasks": [
    {
      "uid": 100,
      "name": "要件定義",
      "parent_uid": null,
      "duration": "PT40H",
      "duration_hours": 40,
      "predecessors": []
    }
  ]
}
```

AI は `PT40H` のような `ISO 8601 duration` をある程度理解できるが、比較や推論をより安定させるため、AI projection では元の表現に加えて補助の数値表現も併記する方針とする。

例:

```json
{
  "duration": "PT40H",
  "duration_hours": 40
}
```

この場合、元の意味を保つ正規表現として `duration` を残しつつ、AI が扱いやすい補助値として `duration_hours` を与える。

## JSON は独自仕様でよいか

AI 向け JSON は、業界標準に完全準拠している必要はない。

この用途では、次の条件を満たす独自 JSON で十分に成立する。

- フィールド名が自然で意味を推測しやすい
- 同じ意味に対して同じ名前を使う
- 型が安定している
- 暗黙ルールが少ない

そのため、`OpenProject` など既存製品の JSON を参考にしつつ、`mikuproject` に必要な範囲で自作する方針が現実的である。

## 標準との距離感

完全にそのまま使える標準 JSON は現時点では見当たらない。

参考になるものはあるが、いずれもそのまま採用するには不足がある。

| 種類 | 位置づけ |
| --- | --- |
| `MS Project XML` | 標準だが XML であり、AI 入出力には不向き |
| `OpenProject JSON` | 実務的だが製品固有 |
| `schema.org Project` | 抽象的で WBS 編集には不足 |

したがって方針としては、`OpenProject` 風のタスク中心 JSON を参考にしながら、`MS Project XML` へマッピングしやすい自前 schema を定義する。

## AI 入力は「最小」ではなく「十分」

AI 連携で重要なのは、単に情報量を減らすことではない。

`全部の真実` を渡すのではなく、`今回の判断に必要な真実` を渡すことが重要である。

AI に対して情報を過剰に渡すと、ノイズが増え、判断の焦点がぼやける。一方で、重要な制約や背景を削りすぎると、もっともらしいが不正確な提案を返しやすくなる。

したがって、設計原則は `minimal` ではなく `sufficient` である。

- 少なければよいわけではない
- 多ければ正確になるわけでもない
- 目的に対して十分な文脈を渡すことが重要

## AI 入力 JSON の考え方

AI に渡す JSON は、canonical model の単純縮小版ではなく、目的ごとに編集された `context-rich projection` として設計する。

この projection には、少なくとも次の 3 層を含める。

### 1. 編集対象

AI が直接読み取り、提案や変更の対象にするデータ。

例:

- task の `uid`
- `name`
- `parent_uid`
- 並び順
- predecessor
- duration
- start / finish

### 2. 制約コンテキスト

AI が守るべきルールや変更可能範囲。

例:

- 変更禁止項目
- 完了済み task は変更禁止
- summary task の直接編集禁止
- predecessor の循環禁止
- 削除操作禁止

### 3. 判断補助コンテキスト

AI が正しい解釈に到達するための背景情報。

例:

- project 名
- フェーズ概要
- 親 task 名
- 近傍 task
- 主要 milestone
- 対象プロジェクトの規模や種別

## 全体像が必要なケース

AI の判断は、局所編集だけで完結しない場合がある。

たとえば次のような依頼では、全体像を渡す必要がある。

- WBS の抜け漏れを見たい
- 工程構成が妥当か見たい
- フェーズ配分の偏りを見たい
- 類似プロジェクトの WBS と比較したい
- 同種案件のたたき台を作りたい

このようなケースでは、個別 task の列挙だけでは不十分であり、フェーズ構成、主要 milestone、依存のまとまり、件数集計、代表 task 群など、プロジェクト全体を理解するための情報を渡す必要がある。

## AI Projection は 1 種類ではない

AI 入力は 1 種類の簡約 JSON で統一するより、用途別 projection を複数持つ方がよい。

候補は次のとおり。

- `project_overview_view`
- `phase_detail_view`
- `task_edit_view`
- `dependency_edit_view`
- `schedule_review_view`
- `similar_project_view`
- `comparison_view`
- `project_draft_request`
- `project_draft_view`

必要に応じて、複数の projection を同時に AI へ渡してよい。

たとえば `task_edit_view` を主データとしつつ、`project_overview_view` を補助コンテキストとして併用する、という使い方が考えられる。

また、`project_overview_view` で全体像を把握した後、特定フェーズを深掘りするための中間ビューとして `phase_detail_view` を置く構成が考えられる。

このとき、`phase_detail_view` は `full` と `scoped` の 2 モードを持てるようにしておく方が実務的である。`full` は phase 全量を渡し、`scoped` は必要に応じて `root_uid` と `max_depth` で subtree 単位に絞る。

現行 UI では、既存 project 向けの生成AI連携は `Output` タブから扱う。

- `full bundle`
- `project_overview_view`
- `phase_detail_view full`
- `phase_detail_view scoped`

を `.editjson` として保存できる構成とし、`scoped` の場合は `phase UID`、`root UID`、`max depth` 相当の入力を伴う。

一方で、既存 project の安全編集とは別に、全く新規の project 草案を AI に生成させるための `project_draft_request` / `project_draft_view` 系を分離して持つことも有効である。

## 新規生成モード

既存計画の編集と、新規計画の生成は、同じ AI 連携でも別モードとして扱う方が安全である。

- 既存編集モード: 既存 project を前提にし、出力は `Patch JSON`
- 新規生成モード: 粗い要件から project 草案を生成し、出力は `project_draft_view`

この新規生成モードでは、AI は既存 project を変更しない。許されるのは、新しい project の全量草案を返すことだけである。

現行 UI では、この新規生成モードは `Input` タブ内の `生成AI連携` から扱う。生成AIとの会話自体は外部で行い、`mikuproject` 側では `project_draft_view` を

- `Load from file` から `.editjson` として開く
- JSON テキストを貼り付けて取り込む

のいずれかで受け付ける。

### 新規生成モードの原則

- 既存 project の編集には使わない
- 出力は `Patch` ではなく全量の draft JSON とする
- draft は正本ではなく草案として扱う
- `mikuproject` 側で validation を通してから内部モデルへ取り込む
- 既存 UID と混同しないよう、仮 UID を使ってよい

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
      "is_summary": true
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

## `project_overview_view` の役割

`project_overview_view` は、個別 task を厳密に編集するためのビューではなく、プロジェクト全体の構造と粒度を AI に把握させるための要約 projection である。

主な用途は次のとおり。

- WBS 全体レビュー
- 抜け漏れ検出
- フェーズ構成の妥当性確認
- フェーズ配分の偏り確認
- 類似プロジェクト比較の入力
- 個別編集前の全体コンテキスト付与

逆に、次の用途には向かない。

- 個別 task の厳密編集
- 全依存関係の詳細編集
- resource / assignment の詳細編集
- XML 互換維持のための保持項目確認

## `project_overview_view` に含めるべき情報

`project_overview_view` では、全 task の全属性を渡すのではなく、全体理解に効く情報を圧縮して渡す。

中核になるのは次の情報である。

- project の基本情報
- task 全体件数などの集計
- フェーズ一覧
- 主要 milestone
- top level の依存関係
- 判断ルールや制約

特に重要なのは `phases` であり、AI がまず全体像を把握するための中心データとなる。

## `project_overview_view` に含めすぎない方がよい情報

次のような情報は、原則として `project_overview_view` には入れない方がよい。

- 全 task の全属性
- assignment の詳細
- calendar の細粒度定義
- XML 互換維持用の細かな保持項目
- `TimephasedData` のような重い詳細

これらは別 projection で扱うべきであり、overview に混ぜると AI の注意が散りやすい。

## `project_overview_view` の最小構造案

```json
{
  "project": {
    "name": "新基幹システム導入",
    "domain": "enterprise-it",
    "planning_mode": "forward",
    "planned_start": "2026-04-01",
    "planned_finish": "2026-12-31",
    "status_date": "2026-06-01"
  },
  "summary": {
    "task_count": 128,
    "summary_task_count": 18,
    "milestone_count": 12,
    "max_outline_level": 4
  },
  "phases": [
    {
      "uid": 100,
      "name": "要件定義",
      "wbs": "1",
      "task_count": 18,
      "milestone_count": 2,
      "planned_start": "2026-04-01",
      "planned_finish": "2026-05-15"
    }
  ],
  "milestones": [
    {
      "uid": 190,
      "name": "要件定義完了",
      "parent_uid": 100,
      "date": "2026-05-15"
    }
  ],
  "top_level_dependencies": [
    {
      "from_uid": 100,
      "to_uid": 200,
      "type": "FS"
    }
  ],
  "rules": {
    "completed_tasks_exist": true,
    "allow_patch_ops": ["add_task", "update_task", "move_task"]
  }
}
```

## `project_overview_view` の必須候補

最低限、次の項目は候補になる。

- `project.name`
- `project.planned_start`
- `project.planned_finish`
- `summary.task_count`
- `summary.max_outline_level`
- `phases`
- `milestones`
- `top_level_dependencies`
- `rules`

## `phases` の重要性

`project_overview_view` の中心は `phases` である。

AI は全 task 一覧よりも、まずフェーズの塊を見た方が、全体構造、工程順序、粒度感を把握しやすい。

`phases` に持たせる候補は次のとおり。

- `uid`
- `name`
- `wbs`
- `task_count`
- `summary_task_count`
- `milestone_count`
- `planned_start`
- `planned_finish`
- `duration`
- `duration_hours`
- `percent_complete`

## `phase` の定義

`project_overview_view` や `phase_detail_view` で使う `phase` は、曖昧な概念のままにせず、抽出ルールを固定した方がよい。

当面の基本方針としては、`phase` を `top-level summary task` ベースで定義するのがもっとも実務的である。

つまり、原則として次を `phase` とみなす。

- `Summary = true`
- ルート直下にある task
- `UID=0` の placeholder task は除外

この定義にすると、MS Project XML の task 階層と自然に対応しやすく、`project_overview_view` と `phase_detail_view` を一貫した基準で構成しやすい。

## `phase` 定義の利点

この定義には次の利点がある。

- XML 由来の task 階層と対応づけやすい
- project 全体を自然な大区分に分けやすい
- `phase_detail_view` の対象を `uid` で直接指せる
- Patch 適用後の再集計も行いやすい

要するに、独自のフェーズ抽出ロジックを複雑に持つより、まずは `top-level summary task` を `phase` とみなす方が壊れにくい。

## `phase` とみなさないもの

次のものは、原則として `phase` には含めない。

- `UID=0` の project summary task
- leaf task
- ルート直下であっても summary ではない task

この場合、ルート直下の leaf task は、どこかの phase に属する task ではなく、phase 外の top-level task として別扱いする。

## 例外と fallback

実際の WBS では、必ずしも top-level summary task がきれいにフェーズを表していない場合がある。

そのため、将来的には次の fallback を持つ余地がある。

1. `top-level summary task` を優先する
2. それが不十分な場合は、明示設定された phase marker を使う
3. さらに必要なら、特定 `OutlineLevel` を phase とみなす補助ルールを使う

ただし、初期段階では fallback を増やしすぎると `phase` の意味がぶれるため、まずは `top-level summary task` に固定する方がよい。

## `phase_detail_view` との関係

`phase_detail_view` は、この定義で抽出された 1 つの `phase` を対象にする。

つまり `phase_detail_view.phase.uid` は、`project_overview_view.phases[].uid` のいずれかと一致する前提で扱う。

これにより、overview で見た phase を、そのまま detail へドリルダウンする導線が明確になる。

## 代表 task を含めるか

`project_overview_view` は単なる集計だけでなく、必要に応じて代表 task を少数含めた方がよい。

要約だけでは、AI がフェーズの具体的な作業内容や粒度感を誤解することがあるためである。

そのため、phase ごとに `sample_tasks` を数件だけ含める設計は有効である。

例:

```json
{
  "uid": 100,
  "name": "要件定義",
  "task_count": 18,
  "sample_tasks": [
    { "uid": 110, "name": "現状業務整理" },
    { "uid": 120, "name": "要件ヒアリング" },
    { "uid": 130, "name": "要件確定" }
  ]
}
```

## `project_overview_view` の位置づけ

要するに、`project_overview_view` は、プロジェクト全体の構造、粒度、主要節目を AI に理解させるための要約 projection であり、個別 task 編集のための完全データではない。

## `phase_detail_view` の位置づけ

`phase_detail_view` は、`project_overview_view` で把握した全体像のうち、特定フェーズだけを一段詳しく見るための projection である。

これは単なる task 一覧ではなく、フェーズの task 群、主要 milestone、依存の要点、必要に応じた要約情報をまとめて持てる詳細ビューとして位置づける。

そのため、`phase_task_list_view` のような限定的な一覧ビューよりも、将来の拡張に耐えやすい。

役割のイメージは次のとおり。

- `project_overview_view`: プロジェクト全体の把握
- `phase_detail_view`: 特定フェーズの深掘り
- `task_edit_view`: 個別 task の厳密編集

また、`phase_detail_view` は phase 全体を返すだけでなく、特定 summary task 配下へ対象範囲を絞る中間ビューとしても使えるようにしておくとよい。

### 範囲指定の考え方

`phase_detail_view` には、少なくとも次の 2 モードがあるとよい。

- `full`: phase 全体をそのまま返す
- `scoped`: 対象 subtree に絞って返す

`scoped` の場合は、次のような scope 指定を持たせる。

- `root_uid`: 対象 subtree の根 task UID
- `max_depth`: `root_uid` から何階層下まで含めるか

この形にしておくと、overview と task_edit の間にある「少し詳しいがまだ広い」ビューを、用途に応じて段階的に絞り込める。

## `phase_detail_view` の具体例

```json
{
  "project": {
    "name": "新基幹システム導入",
    "planned_start": "2026-04-01",
    "planned_finish": "2026-12-31"
  },
  "phase": {
    "uid": 100,
    "name": "要件定義",
    "wbs": "1",
    "planned_start": "2026-04-01",
    "planned_finish": "2026-05-15",
    "task_count": 18,
    "milestone_count": 2,
    "percent_complete": 20
  },
  "tasks": [
    {
      "uid": 110,
      "name": "現状業務整理",
      "parent_uid": 100,
      "position": 0,
      "is_summary": false,
      "is_milestone": false,
      "duration": "PT40H",
      "duration_hours": 40,
      "planned_start": "2026-04-01",
      "planned_finish": "2026-04-05",
      "percent_complete": 100,
      "predecessor_uids": []
    },
    {
      "uid": 120,
      "name": "要件ヒアリング",
      "parent_uid": 100,
      "position": 1,
      "is_summary": false,
      "is_milestone": false,
      "duration": "PT80H",
      "duration_hours": 80,
      "planned_start": "2026-04-06",
      "planned_finish": "2026-04-15",
      "percent_complete": 50,
      "predecessor_uids": [110]
    }
  ],
  "milestones": [
    {
      "uid": 190,
      "name": "要件定義完了",
      "date": "2026-05-15"
    }
  ],
  "dependency_summary": [
    {
      "from_uid": 110,
      "to_uid": 120,
      "type": "FS"
    }
  ],
  "rules": {
    "allow_patch_ops": ["add_task", "update_task", "move_task"],
    "forbid_completed_task_changes": true
  }
}
```

この例では、`tasks` が中心データでありつつ、`phase` の要約、主要 `milestones`、依存の要点、`rules` を同時に持たせている。

これにより、AI は特定フェーズの中身を task 単位で追いながら、そのフェーズ全体の意味も見失いにくくなる。

## `task_edit_view` の位置づけ

`task_edit_view` は、生成AIに個別 task の編集をさせるための projection である。

`project_overview_view` や `phase_detail_view` が理解と検討のためのビューであるのに対し、`task_edit_view` は実際の変更提案を安全に返させるための作業ビューと位置づける。

主な用途は次のとおり。

- task 名の修正
- task の親子変更
- task の順序変更
- task の期間変更
- 依存関係の追加や解除
- 個別 task 周辺の整合見直し

## `task_edit_view` に必要な情報

`task_edit_view` では、対象 task だけを孤立して渡すのではなく、編集判断に必要な局所文脈をあわせて渡す。

少なくとも次の情報が候補になる。

- 対象 task 本体
- 親 task の情報
- 近傍の sibling task
- 同フェーズ内の関連 task
- predecessor / successor
- 編集ルール

特に、task 単体の属性だけでは sibling order や依存文脈を誤解しやすいため、近傍文脈は重要である。

## `task_edit_view` の具体例

```json
{
  "project": {
    "name": "新基幹システム導入"
  },
  "phase": {
    "uid": 100,
    "name": "要件定義"
  },
  "target_task": {
    "uid": 120,
    "name": "要件ヒアリング",
    "parent_uid": 100,
    "position": 1,
    "is_summary": false,
    "is_milestone": false,
    "planned_duration": "PT80H",
    "planned_duration_hours": 80,
    "planned_start": "2026-04-06",
    "planned_finish": "2026-04-15",
    "percent_complete": 50
  },
  "parent_task": {
    "uid": 100,
    "name": "要件定義"
  },
  "sibling_tasks": [
    {
      "uid": 110,
      "name": "現状業務整理",
      "position": 0
    },
    {
      "uid": 130,
      "name": "要件確定",
      "position": 2
    }
  ],
  "predecessors": [
    {
      "task_uid": 110,
      "name": "現状業務整理",
      "type": "FS",
      "lag": "PT0H",
      "lag_hours": 0
    }
  ],
  "successors": [
    {
      "task_uid": 130,
      "name": "要件確定",
      "type": "FS",
      "lag": "PT0H",
      "lag_hours": 0
    }
  ],
  "rules": {
    "allow_patch_ops": ["update_task", "move_task", "link_tasks", "unlink_tasks"],
    "forbid_completed_task_changes": true,
    "forbid_summary_task_direct_edit": true
  }
}
```

この例では、`target_task` が中心だが、親、近傍 task、依存関係、ルールも同時に持たせている。

これにより、AI は対象 task を局所的に読みつつ、周辺文脈を踏まえた Patch を返しやすくなる。

## `task_edit_view` と Patch の接続

`task_edit_view` の目的は、AI に task を理解させることそのものではなく、妥当な Patch を返させることである。

そのため、`task_edit_view` では次を意識する。

- Patch で参照する識別子をそのまま見せる
- 移動や依存変更に必要な相手 task を見せる
- 編集禁止条件を `rules` として明示する
- 未指定項目は変更しない前提で、対象外項目を無理に並べない

つまり `task_edit_view` は、`Patch JSON` を安全に生成させるための最小作業面として設計する。

## 日付・期間・進捗の意味論

AI が WBS を安定して理解するためには、日付や期間の意味を曖昧にしない方がよい。

特に次の点は、名前で区別して表現する。

- それが計画値か実績値か
- duration が計画上の期間か実績上の期間か
- 進捗値が計画補助なのか実績反映なのか

当面の基本方針は、WBS 理解用 projection を `計画ベース` とすることである。

そのため、overview 系や phase 系の projection では、原則として計画値を主に渡す。

推奨:

- `planned_start`
- `planned_finish`
- `planned_duration`
- `planned_duration_hours`
- `actual_start`
- `actual_finish`

`start` や `finish` のような曖昧な名前は、計画と実績が混在し始めると意味衝突を起こしやすい。

生成AIは文脈から推測できることも多いが、推測に依存した schema は弱い。

そのため、値そのものを増やすことよりも、意味名を分けることを優先する。

実績情報は、初期段階では overview に全面展開せず、必要な用途でのみ別 projection または別レイヤとして重ねる方が扱いやすい。

例:

```json
{
  "uid": 120,
  "name": "要件ヒアリング",
  "planned_start": "2026-04-06",
  "planned_finish": "2026-04-15",
  "planned_duration": "PT80H",
  "planned_duration_hours": 80,
  "actual_start": "2026-04-07",
  "actual_finish": null,
  "percent_complete": 50
}
```

この方針により、AI にはまず計画構造を理解させ、その後に必要な場面だけ実績を重ねて見せられる。

## 依存関係の意味論

AI が WBS を理解するうえでは、依存関係を単なる前後順としてではなく、`type` と `lag` を持つ意味的な関係として見せた方がよい。

少なくとも次の情報を持たせる。

- 相手 task の `uid`
- 相手 task の `name`
- 依存種別 `type`
- ラグ `lag`
- ラグの補助値 `lag_hours`

依存種別 `type` は、少なくとも次の 4 種を扱えるようにする。

- `FS`: Finish-to-Start
- `SS`: Start-to-Start
- `FF`: Finish-to-Finish
- `SF`: Start-to-Finish

AI は `FS` を暗黙の既定として理解しがちだが、実際には `SS` や `FF` が工程設計上重要になる場合がある。

そのため、依存関係を扱う projection では、`type` を省略せず明示する方がよい。

## `lag` の扱い

`lag` も duration と同様に、元表現と補助数値を併記する方が実務的である。

推奨:

```json
{
  "type": "FS",
  "lag": "PT8H",
  "lag_hours": 8
}
```

これにより、MS Project に寄せた元表現を保ちながら、AI が比較や計算をしやすくなる。

負のラグを許す場合も、同じ考え方で表せる。

```json
{
  "type": "SS",
  "lag": "-PT8H",
  "lag_hours": -8
}
```

## predecessor と successor

AI に対象 task の編集をさせる場合は、`predecessors` だけでなく `successors` も見せた方が安全である。

理由は次のとおり。

- predecessor だけでは、変更が後続 task に与える影響を見落としやすい
- successor だけでも、先行条件を見落としやすい
- 双方向を見せることで、局所編集の破壊範囲を判断しやすくなる

そのため、`task_edit_view` では `predecessors` と `successors` を併記する方針が妥当である。

## 依存関係の Patch との接続

依存関係を変更する場合は、汎用の field 更新よりも、専用の Patch 操作として表した方が安全である。

候補:

- `link_tasks`
- `unlink_tasks`

たとえば追加は次のように表せる。

```json
{
  "operations": [
    {
      "op": "link_tasks",
      "from_uid": 110,
      "to_uid": 120,
      "type": "FS",
      "lag": "PT0H",
      "lag_hours": 0
    }
  ]
}
```

これにより、依存の追加・解除・変更を task 本体の field 更新と混同しにくくなる。

## 依存関係で追加検討が必要な点

今後さらに詰める必要がある点は次のとおり。

- `lag` の単位を常に時間換算で持つか
- `type` の既定値を許すか、常に明示するか
- 同一 pair に複数 link を許すか
- summary task への link を許すか
- 循環依存をどう validation するか

## `rules` の具体化

AI に渡す `rules` は、単なる参考情報ではなく、AI が返してよい Patch の範囲を制御するための明示的な制約セットとして扱う。

`rules` の目的は次のとおり。

- AI に許可された操作範囲を明示する
- 編集禁止条件を明示する
- Patch 生成時の推測を減らす
- validation 失敗を減らす

つまり `rules` は、projection に付随する説明文ではなく、AI との編集契約として扱う。

## `rules` に入れる情報の種類

少なくとも次の 3 系統に分けて持たせるとよい。

### 1. 許可操作

AI が返してよい Patch 操作の一覧。

例:

- `allow_patch_ops`

候補値:

- `add_task`
- `delete_task`
- `move_task`
- `update_task`
- `link_tasks`
- `unlink_tasks`

### 2. 禁止条件

特定条件下での編集禁止ルール。

例:

- `forbid_completed_task_changes`
- `forbid_summary_task_direct_edit`
- `forbid_delete_task`
- `forbid_dependency_changes`

### 3. 制約パラメータ

単純な可否だけでなく、編集範囲や閾値を表すルール。

例:

- `max_new_tasks_per_request`
- `allowed_edit_fields`
- `allowed_target_scope`
- `allowed_dependency_types`

## `rules` の具体例

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
  "allowed_target_scope": "current_phase",
  "allowed_dependency_types": ["FS", "SS", "FF", "SF"],
  "forbid_completed_task_changes": true,
  "forbid_summary_task_direct_edit": true,
  "forbid_delete_task": true,
  "max_new_tasks_per_request": 0
}
```

この例では、

- 許可 op を限定し
- 更新可能 field を限定し
- 対象範囲を現在フェーズに限定し
- 完了済み task や削除操作を禁止している

## `rules` の解釈方針

AI に対しては、`rules` に書かれていない操作を暗黙に許可しない方がよい。

したがって、基本は次の解釈を採る。

- `allow_patch_ops` にない op は返してはいけない
- `allowed_edit_fields` にない field は更新してはいけない
- `forbid_*` が `true` の条件は必ず守る
- 迷った場合は変更せず、Patch を小さく保つ

この方針により、AI が「たぶん許されるだろう」と推測して広く変更するのを防ぎやすくなる。

## `rules` をどのビューに持たせるか

`rules` は用途に応じて、各 projection に持たせてよい。

たとえば:

- `project_overview_view` では大域的な編集方針
- `phase_detail_view` では当該フェーズの編集範囲
- `task_edit_view` では対象 task 周辺の具体的な編集制約

同じ key 名でも、ビューごとにスコープが違ってよい。ただし意味自体は変えない方がよい。

## `rules` と validation の関係

`rules` は AI への指示であると同時に、Patch 検証側でも使える方がよい。

つまり、返却された Patch は次の順で確認できる。

1. schema に合っているか
2. `rules` に違反していないか
3. 参照整合性に問題がないか
4. 業務ルールに違反していないか

この順に見ることで、単なる型不正と、許可範囲逸脱と、業務不整合を分けて扱いやすくなる。

## `rules` で今後詰める点

今後さらに明確化したい点は次のとおり。

- `allow_patch_ops` を必須にするか
- `allowed_edit_fields` を常に持たせるか
- `forbid_*` を列挙型に寄せるか、boolean key で持つか
- スコープを `project / phase / task` のどこまで細かく切るか
- AI が `rules` 違反を検出したときに空 Patch を返すか、理由つき拒否を返すか

## AI とのやり取り

### AI へ渡すもの

- WBS またはプロジェクトの用途別 projection JSON
- その projection JSON の schema
- 正しい例
- 制約ルール

### AI から受け取るもの

- Patch JSON のみ

AI に対しては、全体 JSON の再出力ではなく、差分だけを返すよう明示する。

## Projection 設計の実務方針

projection 設計では、次の考え方を採る。

- canonical model 全量をそのまま AI に見せない
- AI の用途ごとに projection を作る
- projection には編集対象だけでなく制約と背景も入れる
- XML 互換維持用の細かい保持項目は原則として見せない
- 今回の判断に不要な `Resource / Assignment / Calendar` 全量は渡さない
- ただし、判断に必要なら全体像や類似案件情報も渡す

要するに、AI に必要なのは `縮小データ` ではなく、`目的に応じて編集された文脈付きデータ` である。

## Patch 形式

AI の返却は差分ベースの Patch 形式にする。

例:

```json
{
  "operations": [
    {
      "op": "update_task",
      "uid": 101,
      "fields": {
        "name": "修正タスク"
      }
    }
  ]
}
```

この方式の利点は次のとおり。

- 変更範囲が明確
- 検証しやすい
- 適用前レビューがしやすい
- 不要な書き換えを抑制できる
- XML 全体破壊のリスクを減らせる

## 識別子設計

識別子は次の考え方で整理する。

| 項目 | 扱い |
| --- | --- |
| `UID` | 正本。内部・Patch・参照で使う不変キー |
| `ID` | 表示用 |
| `WBS` | 派生値 |

AI が参照・更新に使う主キーは `UID` に寄せる。

## JSON Schema の位置づけ

`JSON Schema` は正式な標準仕様として採用できる。

主に `draft-2020-12` を前提候補とする。

### Schema でできること

- 型制約
- 必須項目の定義
- `enum` 制約
- 配列やオブジェクトの基本構造制約

### Schema だけではできないこと

- 参照整合性の保証
- 循環依存の検出
- 業務ルール検証
- 意味的妥当性の保証

そのため、`JSON Schema` は入口の構文検証として使い、その後に独自の validation を重ねる必要がある。

## Schema と例と制約

AI 制御では schema だけでは足りない。

精度はおおむね次の順で高くなる。

| 条件 | 期待精度 |
| --- | --- |
| Schema なし | 低い |
| Schema のみ | 中 |
| Schema + 例 | 高い |
| Schema + 例 + 制約 | 非常に高い |

そのため、AI 入出力仕様として最低限必要なのは次の 3 点である。

1. `JSON Schema`
2. 正しい例
3. 制約ルール

## AI への指示方針

AI に対しては次を明示する。

- JSON のみ返す
- 出力形式を固定する
- 不明値は推測しない
- 全体返却は禁止し、Patch のみ返す

抽象的な説明よりも、具体例を添えた方が精度は高い。

## 想定アーキテクチャ

```text
MS Project XML
      ↓
正規化モデル（内部）
      ↓
AI 用 JSON
      ↓
AI（編集）
      ↓
Patch JSON
      ↓
検証
      ↓
内部モデル更新
      ↓
XML 再生成
```

## 今後詰めるべき仕様論点

現時点の方向性は妥当だが、実装仕様としては次を詰める必要がある。

### 1. Patch 操作の粒度

`update_task` だけでは不足する。最低でも次の操作は候補になる。

- `add_task`
- `delete_task`
- `move_task`
- `reorder_task`
- `update_task`
- `link_tasks`
- `unlink_tasks`

### 2. 階層と順序の扱い

`parent_uid` だけでは sibling order を表現できない。

人間向けには次の表現が候補になる。

- `parent_uid + position`
- `before_uid`
- `after_uid`

ただし、生成AIとの Patch では `before_uid / after_uid` は整合性を崩しやすく、扱いがやや不安定である。

そのため、AI とのやり取りでは、順序変更を汎用 field 更新として表すよりも、専用操作として表した方がよい。

推奨:

- `move_task`
- `new_parent_uid`
- `new_index`

例:

```json
{
  "operations": [
    {
      "op": "move_task",
      "uid": 120,
      "new_parent_uid": 100,
      "new_index": 2
    }
  ]
}
```

`new_index` は `0-based` と明記し、適用時に sibling 範囲チェックを行う。

さらに AI に許す操作を絞る場合は、`move_up / move_down / indent_task / outdent_task` のような意味的な操作へ寄せる余地もある。

### 3. 部分更新の意味論

Patch の `fields` に含まれない項目をどう扱うかを固定する必要がある。

- 未指定は変更なし
- `null` は原則として通常の更新値としては使わない
- 値の削除や解除が必要な場合は、専用操作または専用フィールド意味論で表す

生成AIとのやり取りでは、`未指定は変更なし` を基本ルールにした方が安全である。

AI は全項目を安定して再出力するとは限らないため、未指定を削除扱いにすると意図しない消失が起きやすい。

この意味論が曖昧だと安全な適用ができない。

### 4. 業務ルール検証

Schema の後段で、少なくとも次を検証する必要がある。

- 参照先 `UID` が存在するか
- predecessor に循環がないか
- 日付や duration の整合性が取れているか
- summary task に対する制約を満たすか
- 削除や移動で不整合が起きないか

### 5. 用途別 JSON 投影

AI に見せる JSON は単一万能形式ではなく、用途別に投影を分ける余地がある。

例:

- 編集用ビュー
- 要約用ビュー
- 提案生成用ビュー

### 6. 競合検出

古いスナップショットを前提に AI が Patch を返す可能性がある。

そのため Patch には次のような競合検出情報を持たせる余地がある。

- `base_revision`
- `snapshot_hash`

### 7. 権限制御

AI に許す操作範囲をセッション単位で制限できるようにする。

例:

- タスク名変更のみ許可
- 依存関係変更は禁止
- 削除操作は禁止

### 8. エラー分類

Patch 適用失敗時は、単なる失敗ではなく理由を分類して返せるようにする。

例:

- schema error
- reference error
- business rule violation
- conflict

## 現時点の暫定結論

`mikuproject` における生成AI連携の基本方針は次のとおりとする。

- 外部互換と保存は `MS Project XML`
- 内部処理と AI 連携は正規化 JSON
- AI の返却は Patch JSON
- 妥当性確認は schema と独自 validation を組み合わせる

この方針により、XML 互換性を保ったまま、AI 編集に必要な安全性と実装容易性を両立しやすくなる。
