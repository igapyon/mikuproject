# Refactoring Playbook

## 目的

この文書は、`mikuproject` でリファクタリングを進めるときの実務手順をまとめたものである。

このリポジトリでは、単に file を分けること自体を目的にせず、次を同時に満たすことを重視する。

- 挙動を変えない
- 後続の分割を安全に続けられる
- single-file web app / CLI bundle / test harness の読み込み順を壊さない

## 基本方針

- 先に壊れやすい境界へ小さいテストを追加する
- そのテストで守られた範囲から責務を分割する
- 大きい塊を一度に崩さず、閉じた責務から順に切り出す
- 各区切りで `npm run build:full` を実行して閉じる
- test の成否だけでなく、妙に遅くなった test も確認する

## 着手前の確認

リファクタリング対象に入る前に、まず次を確認する。

- その file の責務が何種類混在しているか
- 既存 test がどの境界を守っているか
- 分割後に読み込み順の変更が必要か
- `src/js` 生成物、HTML、CLI loader、test harness に追随が必要か

特に次のような file は、いきなり分割しない。

- UI 全体を抱えた controller
- import / export / validation / parse が同居した巨大 module
- `globalThis.__mikuproject...` の公開面を持つ module

## 先にやるべきテスト追加

大きい分割の前には、外形仕様を固定する小さい test を先に足す。

優先するのは次の境界である。

- 公開 entrypoint
- import/export の代表ケース
- format 判別や dispatch
- normalize / validate の代表ケース
- archive / bundle / filename 生成
- UI 外へ切り出せる pure function 相当の処理

逆に、最初から重い統合 test を増やしすぎない。

- 小粒な確認は `test:fast` または `test:ui` へ寄せる
- 重い回帰確認だけを `test:extended` や `test:full` に残す

## 分割順の考え方

分割順は、依存が閉じている塊から先に行う。

推奨順は次の通り。

1. util / normalize / validation helper
2. zip / codec / package XML のような低レイヤ helper
3. import / export の片側だけで閉じる処理
4. UI event / preview / support のような補助 module
5. 最後に controller / facade / entrypoint

つまり、中心の orchestrator を先に崩すのではなく、周辺から薄くしていく。

## このリポジトリでの具体的な進め方

### 1. 小さい test を足す

例:

- `main.ts` の format 判別、archive 生成、round-trip 判定
- `project-xlsx` / `project-workbook-json` の限定列 import
- `msproject-xml` の AI views / CSV / calendar helper の代表ケース

### 2. 閉じた責務を 1 つだけ外す

例:

- `project-patch-json` から `util` を外す
- `excel-io` から `zip` を外す
- `excel-io` から `normalize` を外す

この段階では、1 回で複数の大責務を抜こうとしない。

### 3. build 順と boot 順を追随させる

分割後は次の追随が必要になることが多い。

- `scripts/build-project.mjs`
- `scripts/lib/core-api-loader.mjs`
- `mikuproject-src.html`
- `scripts/build-project-xlsx-sample.mjs`
- `tests/` 配下の harness
- `src/js` 生成物
- `mikuproject.html`
- `index.html`

`globalThis.__mikuproject...` を使う module は、読み込み順が 1 つずれるだけで壊れることがある。

### 4. 区切りで `build:full` を実行する

各区切りごとに、少なくとも次を実行する。

```bash
npm run build:js
npm run build:full
```

確認するのは次の 2 点である。

- build / test が通ること
- 実行時間が不自然に悪化していないこと

特に `tests/mikuproject-cli.test.js` は最長になりやすいので、継続監視対象とする。

## 止めどきの判断

分割には止めどきがある。

次の状態になったら、追加分割より stop point の見直しを優先する。

- file が controller / facade として自然になった
- 純粋計算や helper がほぼ外へ出た
- 残っているのが orchestration だけになった
- これ以上分けると、呼び出し先を追うコストのほうが増える

`main.ts` のような file は、長くても controller として整理されていれば、無理にさらに割らない判断がありうる。

## 再開判断

いったん止めどきに達した領域を再度触るときは、次を先に確認する。

- 本当に新しい責務混在が再発しているか
- 単に wiring が長いだけではないか
- 既存 facade / public module をこれ以上割っても、追跡コストのほうが増えないか

特に `controller`、`facade`、`public entrypoint` まで整理された file は、長さだけを理由に再分割しない。

## 注意点

- `src/ts` を正本とし、`src/js` は必ず追随させる
- generated JS や single HTML を更新し忘れない
- `globalThis` 公開面を変える場合は、loader / harness も同時に見る
- 既存の挙動を変える変更は、責務分離と同時に混ぜない
- 失敗した分割は、原因を小さく直してから再度 `build:full` で閉じる

## 推奨チェックリスト

- 小さい回帰 test は足したか
- 分割対象は 1 責務に絞れているか
- build 順と boot 順を更新したか
- `npm run build:js` は通ったか
- `npm run build:full` は通ったか
- test 実行時間に異常な悪化はないか
- その file はもう stop point か、まだ次の塊があるか

## 現在の優先順位の見方

個別の優先順位や残作業は `docs/TODO.md` を参照する。

この文書は方法論を置く場所であり、次にどの file を触るかは `TODO` 側で管理する。
