# CLI AI Projections And Validation Notes

この文書は、`miku-project` の CLI から生成AI向け projection と patch 検査を扱うための仕様メモである。

現時点では実装タスクではなく、CLI の責務、コマンド体系、引数設計、first cut の範囲を整理することを目的とする。

## 目的

既存 `WBS` を安全に少し修正するユースケースでは、次の 2 つが重要になる。

- AI に渡す入力を workbook 全体ではなく局所 projection に絞ること
- AI が返した Patch JSON を実適用前後に検査しやすくすること

この文書では、特に次を CLI で扱いやすくする方針をまとめる。

- `task_edit_view`
- `phase_detail_view`
- `project_overview_view`
- Patch JSON の validate
- patch 適用前後の state の要約 / 差分確認

## 前提

- `mikuproject_workbook_json` は CLI の主要 state 入力とする
- AI に渡す入力は workbook 全体ではなく、可能な限り projection JSON を優先する
- Patch JSON の validation は、実際の patch 適用ロジックと乖離しない dry-run apply を基本とする
- CLI は agent 利用を想定し、stdout / stderr だけでなく構造化 diagnostics も将来前提に置く
- 既存WBSの局所修正フローは、Web UI を厚く作り込むより `CLI` / `Agent Skills` を主導線にする

## ユースケース

想定する基本フローは次のとおり。

1. `project_overview_view` で全体像を把握する
2. `phase_detail_view scoped` または `task_edit_view` で局所文脈を取得する
3. AI が Patch JSON を返す
4. `validate-patch` で warning / error / change summary を確認する
5. `state apply-patch` で反映する
6. `state diff` で何が変わったかを確認する

この流れから見て、CLI に最初に必要なのは「局所 projection」と「patch 検査系」である。

Web UI にも近い projection export は残してよいが、`project-overview -> task-edit / phase-detail -> patch -> validate -> apply -> diff` の段階的フローは、UI の主導線というより `CLI` / `Agent Skills` の主導線として扱う方が自然である。

repo 内の最小サンプルとして、`scripts/cli-ai-workflow-example.mjs` と `scripts/cli-ai-stdio-example.mjs` を置く。
前者は file-based に `project-overview` / `task-edit` / `validate-patch` / `apply-patch` / `state diff` を順に呼ぶ。
後者は `--in -` / `--out -` を使い、patch を標準入出力で流す agent 向けの最小フローを示す。

## 優先順位

first cut:

- `ai export task-edit`
- `ai export phase-detail`
- `ai validate-patch`
- `ai export project-overview`

second cut:

- `ai detect-kind`
- `state diff`
- `state summarize`
- `ai export phase-detail --mode full`
- `ai export bundle`

deferred:

- full bundle の拡張オプション
- すべてのコマンドへの一括 `--diagnostics json` 適用
- diff の詳細粒度拡張

## コマンド体系

CLI は責務ごとに整理する。

- `ai`: projection export、document kind 判定、patch validation など AI 連携向け操作
- `state`: workbook state の生成、更新、要約、差分比較
- `export`: workbook / xml / xlsx のような主要交換形式
- `report`: 人が読む派生出力

`ai export bundle` は `ai` 配下に残す。
理由は、これは一般 export ではなく、AI 連携向けの重い projection bundle だからである。
一方で通常フローの第一選択ではなく、主用途は調査 / デバッグ / 比較検証とする。

AI 向け CLI では、階層型 subcommand を正規形とする。

- `miku-project ai export project-overview`
- `miku-project ai export task-edit`
- `miku-project ai export phase-detail`
- `miku-project ai detect-kind`
- `miku-project ai validate-patch`

state 系の候補は次とする。

- `miku-project state summarize`
- `miku-project state diff`

## CLI Naming 方針

利用者がまだ限定的な段階なので、後方互換より CLI 体系の整理を優先する。

方針:

- 正規形は `ai export task-edit` / `ai export phase-detail` / `ai export project-overview` に統一する
- `ai export-task-edit` / `ai export-phase-detail` のような旧形式は維持しない
- `--help` と docs は新しい階層型コマンドに統一する

この方針により、早い段階で CLI を一貫した形に揃える。

## 引数設計

共通:

- `--in`: 入力 JSON ファイル。`-` なら標準入力
- `--in-base64`: binary 入力。`-` なら標準入力の Base64 text を読む
- `--out`: 出力ファイル。text output では `-` なら標準出力。binary artifact では file path を指定する
- `--out-base64`: binary 出力。`-` なら Base64 text を標準出力へ書く
- `--diagnostics text|json`: diagnostics の出力形式
- 同一コマンドで標準入力を読める入力オプションは 1 つだけにする

入出力の優先順位:

- `--in path` が指定された場合は、そのファイルを読む
- `--in -` が指定された場合は、標準入力を読む
- `--in` 省略時に限り、対応コマンドでは標準入力を暗黙入力として読む
- binary input では `--in-base64 -` により標準入力の Base64 text を読む
- text output で `--out path` が指定された場合は、そのファイルへ書く
- text output で `--out -` または `--out` 省略時は、標準出力へ書く
- binary artifact では `--out path` に file output するか、`--out-base64 -` で Base64 text を標準出力へ書く

exit code:

- `0`: success / warning / noop
- `1`: validation error や処理失敗
- `2`: CLI の使い方エラー

選択系:

- `--select auto|first-task|first-phase|uid`

対象指定:

- `--task-uid`
- `--phase-uid`
- `--root-task-uid`
- `--max-depth`
- `--mode scoped|full`

Patch validation 系:

- `--state`

命名方針:

- `--root-uid` より `--root-task-uid` を優先する
- `--select default` より `--select auto` を優先する
- CLI では多少長くても、曖昧さより明示性を優先する

## 既定選択ルール

UID 省略時の既定選択を CLI から明示できるようにする。

候補:

- `auto`: UI 既定選択と同じ挙動
- `first-task`: 最初の通常 task を選ぶ
- `first-phase`: 最初の phase を選ぶ
- `uid`: 明示 UID 指定を要求するモード

first cut では、既存実装との整合を優先して次を基本とする。

- `task-edit` は `auto` を既定にする
- `phase-detail` は `phase-uid` 明示を推奨する
- `select` を導入する場合も、CLI 内で一貫した UID 省略時挙動を維持する
- `--select` は現時点では `task-edit` / `phase-detail` 専用とし、`project-overview` には広げない
- `bundle` / `detect-kind` / `validate-patch` / `state summarize` / `state diff` も、現時点では選択ポリシーを持たせない

## Projection Export

### `ai export project-overview`

目的:

- phase / milestone / top-level dependency を軽く把握する入口
- どの phase / task を次に掘るべきかを決める

入力:

- `mikuproject_workbook_json`

出力:

- `project_overview_view`

first cut では、mode や細かい filtering は持たない。

### `ai export task-edit`

目的:

- 1 task の安全な部分編集に必要な局所文脈を渡す

入力:

- `mikuproject_workbook_json`

出力:

- `task_edit_view`

引数:

- `--task-uid`
- `--select`

first cut では、UID 未指定時に UI と同様の既定選択を維持する。

### `ai export phase-detail`

目的:

- phase 配下の必要部分木だけを AI に渡す
- `task_edit_view` に入る前段の局所理解に使う

入力:

- `mikuproject_workbook_json`

出力:

- `phase_detail_view`

引数:

- `--phase-uid`
- `--mode scoped|full`
- `--root-task-uid`
- `--max-depth`
- `--select`

first cut では次を優先する。

- `mode=scoped`
- `phase-uid` を主入力とする
- `root-task-uid` と `max-depth` で subtree を絞る

`mode=full` は second cut で扱う。

## Patch Validation

### `ai validate-patch`

目的:

- Patch JSON を実適用前に検査し、warning / error / changes を確認する

入力:

- `--state workbook.json`
- `--in patch.json`

基本方針:

- validation は dry-run apply とする
- 実際の patch 適用ロジックで評価する
- state は更新しない
- 既定では patch 後 workbook を返さない

返却したい情報:

- `ok`
- `diagnostics_version`
- `command`
- `context`
- `status`
- `exit_code`
- `warning_count`
- `error_count`
- `io`
- `warnings`
- `errors`
- `changes_summary`

この方針により、`state apply-patch` と `validate-patch` の意味差分を最小化できる。

`status` は少なくとも `success` / `warning` / `noop` / `error` を使い、warning-only や no-op は exit code ではなく diagnostics で機械判定できるようにする。
`--diagnostics json` を指定した場合、usage error や処理失敗でも stderr に同系統の JSON diagnostics を返す。
異常系では `error_type=usage_error|processing_error` を返せるようにする。
異常系では `error_code` に安定識別子を返せるようにする。
異常系では top-level に `error_details` を返せるようにする。
異常系では `errors[]` の各要素にも `code` を返せるようにする。
異常系では `errors[]` の各要素に `details` を返せるようにする。
異常系でも `io` を返し、どの入力元と出力先で失敗したかを追えるようにする。
特に usage error は、可能な範囲で文言推定ではなく CLI 実装側で `error_code` を直接持たせる。
processing error でも、kind mismatch や kind 判定不能のような代表ケースは `error_code` を直接持たせる。
JSON parse failure も `invalid_json_input` として `details.context` 付きで返す。

warning と failure の境界:

- warning-only は `status=warning` かつ `exit_code=0` とする
- no-op は `status=noop` かつ `exit_code=0` とする
- usage error は `status=error` / `error_type=usage_error` / `exit_code=2` とする
- processing error は `status=error` / `error_type=processing_error` / `exit_code=1` とする

## Detect Kind

### `ai detect-kind`

目的:

- JSON を `project_draft_view` / `patch_json` / `workbook_json` に判定する

価値:

- agent 側が入力の取り回しを単純化できる
- 実装コストの割に便利

優先度:

- second cut

## State Summary

### `state summarize`

目的:

- workbook の要約を返す

返却候補:

- project name
- phase count
- task count
- summary task count
- milestone count
- planned start / finish
- 主要 milestone

優先度:

- second cut

## State Diff

### `state diff`

目的:

- patch 適用前後の差分を人と agent の両方が把握しやすくする

入力:

- `--before`
- `--after`

first cut の粒度は要約に留める。

候補:

- project fields changed
- tasks added / updated / moved / deleted
- dependencies added / removed
- assignments added / updated / deleted

最初から完全 diff を目指さない。

## Diagnostics

agent 利用を考えると、stderr の文字列だけでは足りない。

したがって、将来互換を見据えて次を設ける。

- `--diagnostics text|json`

方針:

- 既定は `text`
- first cut では `validate-patch` と `state apply-patch` を優先対象にする
- `json` では warnings / errors / changes を構造化して返す

## 標準入出力

CLI は file-based だけでなく、標準入出力も扱いやすくする。

方針:

- `--in` 未指定時は、対応コマンドでは stdin を読む
- text output は `--out` 未指定時に stdout へ出す
- text output は `--out -` で明示的に stdout へ出す
- XLSX / ZIP などの binary artifact は `--out <path>` へ file output する
- binary stdout が必要な場合は、`--out-base64 -` で Base64 text を出す
- XLSX などの binary input を stdin から読む場合は、`--in-base64 -` で Base64 text を渡す
- agent 利用では一時ファイルを減らせることを重視する

## 典型フロー例

全体把握:

```sh
miku-project ai export project-overview --in workbook.json
```

局所 phase の切り出し:

```sh
miku-project ai export phase-detail \
  --in workbook.json \
  --phase-uid 100 \
  --mode scoped \
  --root-task-uid 123 \
  --max-depth 2
```

個別 task の取得:

```sh
miku-project ai export task-edit --in workbook.json --task-uid 123
```

patch の事前確認:

```sh
miku-project ai validate-patch --state workbook.json --in patch.json
```

patch の適用:

```sh
miku-project state apply-patch --state workbook.json --in patch.json --out workbook.next.json
```

state 差分の確認:

```sh
miku-project state diff --before workbook.json --after workbook.next.json
```

## まとめ

この仕様メモでの結論は次のとおり。

- 最優先は局所 projection の CLI 化である
- 次点は patch の dry-run validation である
- `project_overview_view` は入口として重要だが、局所編集の即効性では `task_edit_view` と `phase_detail_view scoped` が先である
- `state diff` と `detect-kind` は重要だが second cut でよい
- TODO 化はこのメモを前提に、実装単位へ圧縮してから行う
