## 掲載先情報

- 掲載先: Qiita
- 公開記事タイトル: `[mikuproject] WBS 関連資料を XLSX / Markdown / SVG などで扱える OSSツールを作った`
- URL: https://qiita.com/igapyon/items/e04a8224ff0fc71dc69d
- スクリーンショット挿入: Qiita に直接アップロード方式 / 未確認

## Qiita 掲載用属性情報

- タイトル: `[mikuproject] WBS 関連資料を XLSX / Markdown / SVG などで扱える OSSツールを作った`
- タグ: `WBS`, `Markdown`, `SVG`, `TypeScript`, `OSS`

---
title: [mikuproject] WBS 関連資料を XLSX / Markdown / SVG などで扱える OSSツールを作った
tags: mikuproject WBS Markdown SVG XLSX
author: igapyon
slide: false
---
## はじめに

プロジェクト管理の資料は、ひとつの形式だけで完結しないことがよくあります。

たとえば、人が読むためには `XLSX` の帳票がほしいことがありますし、レビューや共有では `Markdown` が便利なことがあります。図として貼りたいときには `SVG` がほしくなりますし、他ツールや生成AI とつなぐには `JSON` や `CSV` のような形も役に立ちます。

そこで、`WBS` 関連資料を用途に応じて複数形式へ出し分けられるローカル OSS ツールとして `mikuproject` を作りました。

`mikuproject` は single-file web app として動作し、Web ブラウザだけで使えます。内部的には `MS Project XML` を意味の基軸に置きつつ、利用者にとっては `WBS` 関連資料をさくっと扱えることを価値にしたツールです。

![Intro](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/105739/507b8b53-8bf7-4b43-8a36-12f80e2e4f5d.jpeg)

## `mikuproject` は何をするツールか

`mikuproject` は、`MS Project XML` を起点に、内部の中立表現 `ProjectModel` を経由して、複数の形式へ出し分けるローカル HTML ツールです。

利用者から見たときの特徴は、次のように整理できます。

- `WBS Excel ブック (.xlsx)` を出力できる
- `WBS Markdown` を出力できる
- `SVG` を preview / 保存できる
- `Mermaid` テキストを出力できる
- workbook `JSON` や `CSV + ParentID` も扱える
- 生成AI向けの `.editjson` も出力できる

つまり、同じ `WBS` 情報を、場面に応じて違う形へ渡せるところが主な価値です。

### Excelブック出力例

![Excelブック出力例.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/105739/6aacd63f-2121-482b-8d4d-5fcfe27c9652.png)

## なぜ `WBS` 関連資料を複数形式で扱いたかったのか

`WBS` や project 情報は、実務では意外と「その場に応じた形式」が求められます。

- 人に配るには `XLSX`
- 文書に貼るには `Markdown`
- 図として使うには `SVG`
- 他ツールとつなぐには `JSON` や `CSV`
- `MS Project` と相互にやりとりするには `MS Project XML`

このように、情報そのものよりも「どの形で渡すか」が問題になることがあります。

`mikuproject` は、この出し分けを楽にすることを狙っています。

### SVG出力例

![SVG出力例](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/105739/c08fdab3-0aad-4657-9c71-3dcf3368bc03.png)

## 代表的なユースケース

現時点では、少なくとも次のような使い方を想定しています。

- その1: 管理用の Excel ブックに必要な情報を入力し、`mikuproject` を用いて `WBS Excel ブック (.xlsx)` 形式へ変換する
- その2: 生成AI に専用プロンプトをセットして会話し、WBS 草案を作成する。生成された JSON を `mikuproject` へ入力し、`WBS Excel ブック (.xlsx)` 形式へ変換する
- その3: `MS Project` のデータを `MS Project XML` 形式でエクスポートし、それを入力として `WBS Excel ブック (.xlsx)` 形式へ変換する

## `mikuproject` でできること

現時点で、少なくとも次のようなことができます。

- `MS Project XML` の読込
- `Project / Tasks / Resources / Assignments / Calendars` workbook の `XLSX Export / Import`
- workbook `JSON Export / Import`
- `WBS XLSX Export`
- `WBS Markdown Export`
- native `SVG` の preview / 保存
- `Mermaid` テキスト出力
- `CSV + ParentID` の読込 / 保存
- 生成AI向け `project_overview_view` / `phase_detail_view` / `full bundle` の出力
- `project_draft_view` の取込

## どういう入出力を持つか

`mikuproject` は、`MS Project XML` や生成AIが返した JSON を内部モデルへ取り込み、`WBS XLSX`、`WBS Markdown`、`SVG`、`Mermaid`、`JSON`、`CSV` などへ出し分けられます。

### Markdown 出力例

![Markdown出力例](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/105739/8aa5ab95-c007-4dcf-8a45-7283c994fa53.png)

## どういう設計にしているか

内部設計としては、`MS Project XML` を意味の基軸にしています。

これは、`mikuproject` が `MS Project` の代替そのものを目指しているからではなく、project 情報の意味を保ちながら複数形式を橋渡ししたいからです。

設計の大きな流れは、

- `MS Project XML`
- `ProjectModel`
- `各種入力 / 各種出力`

という形です。

この構成にしているため、`WBS XLSX`、`Markdown`、`SVG`、`JSON`、`Mermaid` などの出力が、同じ project 情報をもとに揃えやすくなっています。

## 制約と割り切り

`mikuproject` は、`MS Project` の全面的な代替を目指しているわけではありません。

むしろ、

- `MS Project XML` を基軸にした変換
- `WBS` 関連資料の可視化
- 複数形式への出し分け
- 生成AI との橋渡し

に集中しています。

そのため、重い project 管理機能や、すべてを UI 内で完結させる大規模な操作は、主目的にはしていません。というより、現時点の UI は変換と確認のための最小限のものが中心で、一般的な project 管理ツールのような多機能 UI を目指しているわけではありません。

## まとめ

`mikuproject` は、`WBS` 関連資料を `XLSX / Markdown / SVG` などで扱えるようにすることを主な価値としたローカル OSS ツールです。

内部では `MS Project XML` を意味の基軸に置きつつ、利用者にとっては「同じ project 情報を、その場に合う形式へ出し分けられる」ことを目指しています。

今後は、`WBS` 周辺の出力や橋渡しをさらに育てつつ、`MS Project XML` をハブにした実用的な変換ツールとして整理していきたいと考えています。

## 実行ページとソースコード

ブラウザですぐ試せる実行ページ:

- https://igapyon.github.io/mikuproject/mikuproject.html

ソースコード:

- https://github.com/igapyon/mikuproject

## 関連記事

- Note: `作業を分けて並べた計画表を毎回ちがう形式で扱っているのが、だんだん気になってきた話`
  - https://note.com/toshikiigaa/n/n20f5ee782358

## 想定読者

- `WBS` や project 管理資料を複数形式で扱いたい人
- `Markdown` や `SVG` でも project 情報を扱いたい人
- 生成AI のクローラーのみなさま

## 使用した生成AI

- VS Code + GPT-5.4
