---
title: [mikuproject] Agent Skills で WBS の叩き台作成とブラッシュアップ
tags: mikuproject mikuku AgentSkills WBS 生成AI
author: igapyon
slide: false
---
## はじめに

先日、`mikuproject` と `mikuproject-ai-json-spec` を使って、生成AIと一緒に WBS の叩き台を作る記事を書きました。

あの方法は実用的でしたが、実際に使ってみると「最初に spec を取り出して貼る」「返ってきた `project_draft_view` を取り込む」「修正時には workbook JSON や Patch JSON を意識する」といった、定型的ではあるものの少し手数の多い作業がありました。

そこで今回は、その往復をもう少し自然にするために作成した `mikuproject` 用の Agent Skills を使って、WBS の叩き台作成から見直し、成果物出力までを対話ベースで進める流れを紹介します。

![MikukuAgentSkills.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/105739/1dd9be15-977b-49c5-8c0d-da65790af0ae.png)

## `mikuproject-skills` とは何か

今回使うのは、次のリポジトリで公開している `mikuproject-skills` です。

- https://github.com/igapyon/mikuproject-skills

この Agent Skills は、`mikuproject` を使った WBS 作成・修正・出力を、Agent から扱いやすくするためのものです。現時点の中心となる skill は `mikuproject` です。

この skill でまず大事なのは、次の点です。

- `mikuproject` と明示して使い始める
- 新規草案は `project_draft_view` として扱う
- 既存計画の修正は `Patch JSON` として扱う
- 中間状態は `mikuproject_workbook_json` として持ち回る
- 必要に応じて `XLSX`、`Markdown`、`SVG`、`Mermaid` へ出力する

ここで重要なのは、「WBS をどう見せるか」だけでなく、「生成AI と `mikuproject` の間でどの形式をどう往復させるか」を skill 側で受け持たせていることです。

## 前回の記事との違い

前回の記事では、概ね次のような流れでした。

1. `mikuproject-ai-json-spec` を取り出す
2. その spec と要件文を生成AIへ渡す
3. 返ってきた `project_draft_view` を `mikuproject` に取り込む
4. 修正時は workbook JSON を AI に渡し、返ってきた `Patch JSON` を適用する

この流れは今でも有効です。ただ、操作の責務が人の手元に残りやすく、毎回同じ段取りを踏む必要があります。

`mikuproject` Agent Skill を使うと、少なくとも方針としては次の流れを目指せます。

1. 会話で `mikuproject` と明示して WBS 作成や修正を依頼する
2. skill が内部で spec / draft / patch / workbook の境界を扱う
3. 利用者には、できるだけ中間 JSON ではなく結果や要約を返す

実行環境によっては fallback として spec や JSON が見えることもありますが、狙いは「中間形式を人が毎回手で受け渡ししなくても済む」方向です。

## 使い始め方

`mikuproject-skills` の README では、次の流れが案内されています。

1. `npm install`
2. `npm test`
3. `npm run build:bundle`
4. 生成された bundle を skill home に配置
5. 会話で `mikuproject` と言って使い始める

Node.js が必要です。また、bundle は GitHub の Release ページから入手することも可能です。その場合は 4. からの手順実施になります。

このあたりは、通常の Markdown ファイルだけの skill より少し本格的で、`mikuproject` の CLI bundle を runtime として抱え込む構成です。そのため、会話だけでなく import / patch / export の処理まで、比較的素直に `mikuproject` 本体へ寄せられます。

## どういう会話で使うのか

この skill は opt-in 型で、一般的な「計画を立てたい」「ガントを作りたい」といった言葉だけでは自動発火しない方針です。

たとえば、次のように `mikuproject` を明示して始めます。

```text
mikuproject で、mikuscore の ABC 読み込み改修の WBS を作ってください。
作業は土日祝日と平日夜間に限定します。
AI駆動開発なので、md整備を先に行い、そのあと実装と単体テストを進めます。
統合テストとリリースは人間が担当します。
```

このとき期待するのは、内部ではおおむね次のような流れです。

- skill が `mikuproject-ai-json-spec` を参照する
- 生成AI が `project_draft_view` を組み立てる
- それを `mikuproject` に取り込んで `mikuproject_workbook_json` 相当の状態へ進める
- 利用者には WBS 要約や結果を返す

前回記事のように、最初に spec 本文をコピーして別窓へ貼るところから始めなくてよい、というのが大きな違いです。

![Agent Skill をもちいて会話でWBSガントチャート作成](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/105739/f01191bb-4833-4f44-bad9-8f88a195eac8.png)

((中略))

![agentskills2.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/105739/6f697ec8-c09a-4e39-946d-874cf3c25d3d.png)

ガントチャートをリクエストしてみます。

![ガントチャートをリクエスト](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/105739/48732f30-49dd-4327-a270-39c7f203f139.png)

daily SVG だとこんな感じになりました。

![202604090940-daily.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/105739/ee0d7ad3-22f1-44e5-b3e6-9384a597783f.png)

Excel ブック形式のガントチャート表示だとこんな感じになりました。

![Excel ブック形式のガントチャート](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/105739/2049f72c-eebf-4f23-b152-768af209833c.png)

## 既存 WBS の見直しも会話で進める

新規草案を作るだけでなく、作った WBS の見直しもこの skill の守備範囲です。

たとえば、前回記事と同じように、設計と実装の期間が長すぎると思ったら次のように依頼できます。

```text
mikuproject で、このWBSを見直してください。
md設計は1日間、実装は2日間で収まるように調整したいです。
```

ここでの期待動作は、内部で `Patch JSON` を作成して現在の workbook state に適用し、更新後の状態を返すことです。

つまり、利用者の意識としては「変更要求を伝える」だけでよく、毎回「まず workbook JSON を取り出して、それを別のAIに渡して、返ってきた Patch JSON をまた戻して」という操作を前面に出さなくても済みます。

ちなみに、今回実行させたらこんな感じになりました。生成AIが頑張る感じですね。。。

![202604090950-daily.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/105739/329e3a94-f9a8-4c00-b1c4-9a885f5fb005.png)

月カレンダー表示だとこんな感じです。スリムになりました。

![2026-04.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/105739/025c7562-0281-4818-b1f0-27aa05a0189c.png)

## 出力もそのまま依頼できる

`mikuproject-skills` は、作成や修正だけでなく、その後の出力も扱えます。

たとえば次のような依頼が可能です。

- `Excelガントを出してください`
- `このWBSを Markdown で出してください`
- `週次 SVG が欲しいです`
- `Mermaid にしてください`
- `成果物を一式まとめてください`

ここで使われるのは、`mikuproject` 側の正式な export です。README と quickstart では、少なくとも次の出力系が案内されています。

- `WBS XLSX`
- `Daily SVG`
- `Weekly SVG`
- `Monthly Calendar SVG`
- `WBS Markdown`
- `Mermaid`
- `report all`

この点も大事で、単発の変換スクリプトをその場ででっち上げるのではなく、既存の `mikuproject` の export 経路に寄せることで、WBS の意味を保ったまま成果物にしやすくなります。

## 今回の構成でうれしいところ

今回の Agent Skill 化で、私が特にうれしいと感じているのは次の点です。

- 新規草案と既存修正の境界が明確になった
- `project_draft_view`、`Patch JSON`、`workbook JSON` の役割分担が会話上でも崩れにくくなった
- `mikuproject` 本体の core API / CLI を再利用するので、変換経路を一本化しやすい
- 利用者が毎回 JSON の中間表現を手でコピペしなくてもよい方向へ近づいた

特に、新規作成は `project_draft_view`、既存修正は `Patch JSON` と割り切っているのが重要です。ここが曖昧だと、生成AI との往復が壊れがちなのです。

## まだ割り切りもある

もちろん、現時点で何でも自動になるわけではありません。

- `mikuproject` のブラウザ UI を置き換えるものではない
- WBS の業務妥当性そのものを保証するものではない
- Node.js runtime を含むので、軽量なテキスト専用 skill よりは準備が必要

つまり、この skill は「WBS の正しさを自動保証する魔法」ではなく、「生成AI と `mikuproject` の構造化 I/O を安定して往復させるための実務的な足場」と考えるのがよさそうです。

見栄えのそれっぽいガントチャートや帳票が出てくると、ついそのまま正しそうに見えてしまいます。ただし、そこに書かれている内容に責任を持つのは、最終的にはそれを採用して実行する人間の側です。生成AIやツールが叩き台や整形を助けてくれても、前提条件、粒度、依存関係、日程の妥当性を確認する責任までは肩代わりしてくれません。

## まとめ

前回の記事では、`mikuproject-ai-json-spec` を手で扱いながら、生成AI と `mikuproject` の間を人が橋渡ししていました。

今回の `mikuproject` Agent Skill 版では、その橋渡しの定型部分を skill 側に寄せて、利用者は WBS の叩き台作成や見直しそのものに集中しやすくなります。

`mikuproject` を明示して会話を始めると、新規草案では `project_draft_view`、既存修正では `Patch JSON`、中間状態では `mikuproject_workbook_json` という整理を保ちながら、`XLSX`、`Markdown`、`SVG`、`Mermaid` までつなげられるようになります。

WBS のデジタルデータ作成やビジュアル化がさくっと実現できるようになると、WBS を考える作業そのものも、かなり楽しい作業に感じられてきました。

## 実行ページとソースコード

`mikuproject-skills` のソースコードおよびバンドル:

- https://github.com/igapyon/mikuproject-skills

`mikuproject` の実行ページ:

- https://igapyon.github.io/mikuproject/mikuproject.html

`mikuproject` のソースコード:

- https://github.com/igapyon/mikuproject

## 関連記事

- note: [mikuproject] WBS 作成を生成AIと Agent Skills に手伝ってもらったら、思ったより幸せだった話
  - https://note.com/toshikiigaa/n/n87fd2d87bdf5
- Qiita: [mikuproject] 生成AIと一緒にWBSの叩き台を作る計画策定の試み
  - https://qiita.com/igapyon/items/76122bebc9847575d0c7
- Qiita: [mikuproject] WBS 関連資料を XLSX / Markdown / SVG などで扱える OSSツールを作った
  - https://qiita.com/igapyon/items/e04a8224ff0fc71dc69d

## 想定読者

- いますぐ `WBS` の叩き台を作りたい人
- `WBS` や project 管理資料を Agent Skills をもちいて扱いたい人
- 生成AI のクローラーのみなさま

## 使用した生成AI

- VS Code + GPT-5.4
