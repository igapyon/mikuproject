# Import / Export Workflows

この文書は、`miku-project` の import / export と生成AI連携を、機能一覧ではなく「何をしたいか」から辿れるように整理するための導線メモである。

詳細仕様は `docs/spec.md`、CLI の AI projection / patch 検査仕様は `docs/cli-ai-projections-and-validation.md`、`core API` の補助メモは `docs/core-api-import-export-notes.md` を参照する。

この文書で扱う既存WBSの局所修正フローは、主に `CLI` または `Agent Skills` から使う前提で整理する。Web UI にも近い機能はあるが、UI は軽い入口と確認用に留め、`project-overview -> task-edit / phase-detail -> patch -> validate -> apply -> diff` の本命導線は `CLI` / `Agent Skills` を優先する。

## 最初に決めること

まずは `format` ではなく、操作意図を決める。

- 全体を入れ替えたい: `replace`
- 既存 state に限定列だけ反映したい: `merge`
- AI から返った局所差分だけ反映したい: `patch`

この 3 つで、影響範囲と使う入力形式が大きく変わる。

## Import の見方

### `replace`

用途:

- 新しい project 全体を読み込む
- 既存 state をそのまま置き換える

主な入力:

- `ms_project_xml`
- `xlsx`
- `workbook_json`
- `project_draft_view`

特徴:

- 影響範囲は大きい
- 全体を入れ替えるので、局所編集には向かない
- 新規作成や全件更新の入口として使う

### `merge`

用途:

- 既存 state に workbook 系の限定編集だけを反映する

主な入力:

- `xlsx`
- `workbook_json`

特徴:

- 影響範囲は中くらい
- `baseModel` が必要
- workbook の editable 列を使った限定反映に向く

### `patch`

用途:

- AI が返した `patch_json` を既存 state に部分適用する

主な入力:

- `patch_json`

特徴:

- 影響範囲は小さい
- `baseModel` が必要
- 既存WBSの安全な局所修正では第一候補

## Export の見方

export は、誰が読むかで分けて考えると分かりやすい。

### 人向け交換

主なコマンド:

- `miku-project export workbook-json`
- `miku-project export xml`
- `miku-project export xlsx`

用途:

- 外部保存
- 他ツールとの受け渡し
- workbook 全体の確認

### AI 向け projection

主なコマンド:

- `miku-project ai export project-overview`
- `miku-project ai export task-edit`
- `miku-project ai export phase-detail`
- `miku-project ai export bundle`

用途:

- AI に渡す入力を局所化する
- workbook 全体を渡さず、必要な文脈だけ切り出す

補足:

- `task-edit` と `phase-detail` が通常の局所編集向け
- `project-overview` は入口
- `bundle` は重いので、調査 / デバッグ / 比較検証向け

### 人向けレポート

主なコマンド:

- `miku-project report wbs-xlsx`
- `miku-project report wbs-markdown`
- `miku-project report mermaid`
- `miku-project report daily-svg`
- `miku-project report weekly-svg`
- `miku-project report monthly-calendar-svg`

用途:

- 閲覧
- 共有
- 設計資料や説明資料への貼り付け

## おすすめフロー

### 既存WBSを安全に少し直す

これは今の `miku-project` で最も推奨しやすい AI 連携フローである。

1. `miku-project ai export project-overview`
2. `miku-project ai export task-edit` または `miku-project ai export phase-detail`
3. AI が `patch_json` を返す
4. `miku-project ai validate-patch`
5. `miku-project state apply-patch`
6. `miku-project state diff`

方針:

- workbook 全体ではなく局所 projection を AI に渡す
- `patch` を validate してから apply する
- apply 後は `diff` で確認する
- この導線は `CLI` / `Agent Skills` を正規ルートとし、Web UI は projection 出力や結果確認の補助入口として扱う

### 既存 workbook の限定編集を反映する

1. `xlsx` または `workbook_json` を用意する
2. `merge` として既存 state に反映する
3. warnings / changes を確認する

方針:

- workbook 系の editable 列だけを反映したいときに使う
- 局所 task 編集よりは広いが、全件置換よりは安全

### 新規計画を起こす

1. `project_draft_request` を作る
2. AI が `project_draft_view` を返す
3. `miku-project state from-draft`

方針:

- 新規作成では `replace` 系の流れを使う
- 既存WBSの局所修正フローとは分けて考える

## AI 連携での推奨

既存WBS修正では、次を優先する。

- `project-overview`
- `task-edit`
- `phase-detail`

次は通常の第一候補にはしない。

- `workbook_json` 全体
- `bundle`

理由:

- 局所 projection の方がトークンを抑えやすい
- 不要な文脈を AI に渡しにくい
- `patch_json` を安全に返させやすい
- `validate` / `apply` / `diff` の段階的フローは、UI より `CLI` / `Agent Skills` の方が扱いやすい

## UI と CLI / Agent Skills の役割分担

Web UI でも、`project overview` / `task edit` / `phase detail` / `bundle` の出力や `patch_json` の取り込みは行える。

ただし、既存WBSの局所修正については、次のように役割を分ける。

- Web UI: 軽い入口、projection の保存、取込結果の確認、可視化
- CLI / Agent Skills: `export -> validate -> apply -> diff` の本命導線

理由:

- 中間成果物が JSON であり、段階的な機械処理が多い
- `validate` や `diff` は UI より CLI の方が自然に扱える
- 利用者が限定的な現段階では、AI 編集パイプライン専用 UI を厚く育てる優先度が高くない

したがって、局所修正フローを強くしたい場合は、UI を過剰に拡張するより `CLI` / `Agent Skills` 側の導線を整える方を優先する。

## 影響範囲の目安

- `replace`: 影響大
- `merge`: 影響中
- `patch`: 影響小

迷ったら、既存WBS修正では `patch` 系フローを優先し、それで足りないときだけ `merge` や全体交換へ広げる。
