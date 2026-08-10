---
title: miku-project ゼロベース再設計仕様 v20260809
description: CLI と Agent Skills を中心に miku-project のあり方をゼロベースで再検討するための仕様ドラフト。
topics:
  - miku-project
  - miku-soft
  - cli
  - agent-skills
  - architecture
category: specification
status: draft
audience:
  - maintainer
  - developer
  - agent
created: 2026-08-09
updated: 2026-08-10
---

# miku-project ゼロベース再設計仕様 v20260809

## 文書の位置づけ

この文書は、`miku-project` の今後のあり方を、現行実装との互換性を前提にせずゼロベースで再検討するための仕様ドラフトである。

`miku-project` と `miku-score` は miku-soft シリーズの古参であり、現在の miku-soft に見えてきた設計思想が十分に反映されているとは限らない。本仕様では、後年の miku-soft 群から得られた知見をもとに、製品の目的、責務、操作体系、Agent Skills との関係を改めて定義する。

既存の [spec.md](spec.md)、ソースコード、CLI、Web UI、データ形式は、現時点の実装と過去の試行を示す資料として扱う。本仕様を制約する正解とは扱わない。

この文書は 2026-08-09 時点の検討結果であり、現行仕様を直ちに置き換えるものではない。実装、互換性、移行方法は、基本仕様が固まった後に別途決定する。

初期検討の対象は、`miku-project` の製品契約、Node CLI、Java CLI、それらを利用する Agent Skills とする。`miku-project-web` に相当する Web App と、`miku-project-mcp` に相当する MCP 対応は初期検討の対象外とし、CLI と Agent Skills の基本仕様が固まった後に必要性と対応方法を検討する。

この仕様を実行可能な作業へ分解した計画は、[miku-project ゼロベース新仕様適合計画 v20260810](miku-project-zero-base-implementation-plan-v20260810.md) を参照する。

## 背景

miku-soft シリーズを複数開発した結果、miku-soft に適した基本形が見えてきた。

それは、大規模な統合アプリケーションを中心に置く形ではなく、次の組み合わせである。

- ローカルデータを決定論的に処理するシンプルな CLI
- CLI を生成AIが安全に利用するための Agent Skills
- 人、生成AI、外部ツールの間で受け渡せるファイル成果物

miku-soft の多くは、既存データを読み取り、検査し、別の表現へ変換する道具である。この性質を製品設計の中心に据える。

## ゼロベース再設計の原則

再設計では、次のものを一度未決定へ戻す。

- 現在の Web UI を製品の中心とするか
- 現在の CLI コマンド体系を維持するか
- `MS Project XML` を意味の基軸とするか
- 現在の `ProjectModel` や各種 JSON を維持するか
- 現在の Patch 形式を維持するか
- 現在のリポジトリ分割を維持するか
- 既存実装との後方互換性をどこまで持つか

既存実装から引き継ぐのは、まず設計そのものではなく、次の実証的な知見である。

- 実データに存在する揺れや不正
- 形式変換で失われやすい情報
- 実装が難しいドメイン処理
- 有効なテストデータと回帰観点
- 利用者が実際に必要とした成果物
- 生成AIとの往復で有効だった操作と失敗した操作

既存コードは再利用候補ではあるが、新仕様を拘束するアーキテクチャ上の前提とはしない。

## miku-soft の基本定義

本仕様では、miku-soft の基本形を次のように捉える。

> miku-soft は、Agent がローカルデータを安全に扱うための、決定論的な CLI と Agent Skills の組である。

AI Agent フレンドリーであることは、CLI を単純にした結果として得られる副次的効果ではなく、miku-soft の主要な製品要件とする。ただし、AI Agent 専用の特殊な操作体系を作ることは意味しない。人、shell script、CI、AI Agent が同じ製品契約を利用し、それぞれが操作の意味、入力、出力、副作用、失敗理由、次に可能な行動を判断できることを目指す。

AI Agent フレンドリーとは、短い英語のコマンド名を使うことだけではない。本仕様では、少なくとも次の性質を指す。

- 操作ごとの責務と副作用が明確である
- 入力、出力、artifact の役割、schema version が明示される
- 結果、diagnostics、終了状態を安定した構造として取得できる
- 同じ入力と条件から決定論的な結果を得られる
- 警告、損失、推測、未対応、再試行可能性を機械判定できる
- 変更操作と読み取り操作を区別し、人間確認を必要な地点だけに置ける
- Agent Skills が、成功時だけでなく失敗時にも安全な次の行動を選択できる

各要素の責務は次のように分ける。

| 要素 | 主な責務 |
| --- | --- |
| CLI | 読み取り、変換、正規化、検証、差分、適用、出力、diagnostics |
| Agent Skills | 利用目的の解釈、CLI の選択と実行順序、人間確認、失敗時の案内 |
| ファイル成果物 | 人、Agent、外部ツールの間の明示的で再現可能な受け渡し |
| Web UI | 必要な場合だけ追加する確認、可視化、補助操作用アダプター |

ドメイン上の意味、変換、検証は CLI 側が所有する。Agent Skills に同じ処理を再実装しない。

Agent Skills が runtime を同梱する場合も、CLI の正本から生成した固定 runtime を利用し、独立した派生実装にはしない。

## 姉妹リポジトリ横断調査

### 調査の位置づけ

2026-08-09 に、`miku-project` と同じ親ディレクトリにある `miku-*`、`mikuproject-*`、`mikuscore-*` のローカル Git リポジトリを横断確認した。

調査では、リポジトリ名だけでなく、`README.md`、`package.json`、`pom.xml`、CLI entrypoint、Agent Skill、同梱 runtime、Web App、MCP server の有無を確認した。

接尾辞をもとにした59リポジトリの内訳は次のとおりである。

| 区分 | リポジトリ数 |
| --- | ---: |
| Main Application またはその他の基幹リポジトリ | 20 |
| Java runtime または Java companion | 14 |
| Maven Plugin | 3 |
| Agent Skills | 16 |
| 明示的な `-web` | 4 |
| 明示的な `-mcp` | 1 |
| Catalog | 1 |

これはディレクトリ名による概数である。たとえば `miku-score` は現時点で Web App と CLI を同じリポジトリに持ち、`miku-abc-player` は名前に `-web` を持たないが実質的には Web App に近い。そのため、製品の役割は接尾辞だけで確定しない。

既存の `miku-soft-catalog` は 2026-05-17 時点の整理であり、その後に追加・分離・改名された現在の姉妹リポジトリ群をすべては反映していない。本調査では catalog を参考資料とし、ローカルにある各リポジトリの現在状態を優先した。

### CLI・Java・Skillsが揃った型

次の系列では、Main Application の CLI、Java CLI、Agent Skills が明確に分離されている。

- `miku-indexgen`
- `miku-grep`
- `miku-readfile`
- `miku-text-bundle`

とくに `miku-indexgen` は、Node CLI、Java CLI、Maven Plugin、Agent Skills が同じ製品契約を中心に配置されている。`miku-text-bundle` は、Skills を `cli-preferred` とし、Java runtime を優先して Node runtime を fallback にする運用例を持つ。

この型から、次の知見を採用候補とする。

- 製品の意味と変換規則は CLI runtime 側が所有する
- Node と Java は別製品ではなく、同じ製品契約の複数 runtime として扱える
- Agent Skills は上流 runtime artifact を同梱し、対話と workflow を提供する
- Skills は MCP backend を必須としない
- runtime artifact の版、由来、digest を確認できるようにする

### Office変換スイート型

Microsoft Office と Markdown の変換系列には、複数の小さな変換CLIを一つのAgent Skillsが束ねる型がある。

主な構成要素は次のとおりである。

- `miku-docx2md`、`miku-xlsx2md`、`miku-pptx2md`
- `miku-md2docx`、`miku-md2xlsx`、`miku-md2pptx`
- 対応する Java runtime と一部の Maven Plugin
- 一部の `-web`
- `miku-ms-office-core` と `miku-ms-office-core-java`
- 複数の変換器を選択して実行する `miku-ms-office-skills`

この型は、CLI と Skills が必ず一対一である必要はないことを示す。互いに近い複数の小さな変換器は、利用者の目的を単位とした一つのSkillsから選択できる。

一方、`miku-project` は一つの大きなドメインと往復操作を持つため、現時点では専用の `miku-project-skills` を持つ一対一の形を第一候補とする。

### Node CLI・Skills型

次の系列では、Node CLI とAgent Skillsを中心とし、Java runtimeを必須としていない。

- `miku-json2xlsx`
- `miku-repo-bundle`
- `miku-text-file-ops`
- `miku-backlog-api`
- `miku-confluence`

`miku-json2xlsx` は、`inspect`、人間が確認可能な選択、`validate-mapping`、`convert` を分けている。Skillsは判断と承認可能なhandoffを担当し、JSON解析、mapping検証、XLSX生成、diagnosticsは上流CLIへ委譲する。

この型から、複雑な一括コマンドだけでなく、検査、判断、検証、実行を分離したCLIがAgent Skillsと相性がよいことが分かる。

### Skills自体が製品となる型

次の系列は、独立した単一CLIの単純なwrapperではなく、Agent Skillsそのものが主な製品である。

- `miku-ai-assistant-builder-skills`
- `miku-m365-agent-builder-skills`
- `miku-prompt-lint-skills`
- `miku-media-proc-skills`

これは文書判断、複数ツールの組み合わせ、人間との対話が中心となる用途に適する。プロジェクト計画の解析、変換、検証を決定論的なCLIへ置ける `miku-project` は、この型を基本形とはしない。

### Web・MCP・Maven Pluginの位置づけ

明示的な `-web` は少数であり、MCP companion は調査時点で `mikuproject-mcp` の1リポジトリだけである。Maven Plugin も Java runtime を Maven build から利用するための追加entrypointである。

横断結果から、これらは miku-soft 全製品に必須の中核ではなく、安定した製品能力を別環境へ公開する後続adapterとして分類するのが自然である。

`miku-score` は現時点で single-file Web App とCLIが同居しているが、別途 `miku-score-web` に相当する Web App の分離作業が計画されている。これは、古参製品もCLIを中心としたMain ApplicationとWeb adapterへ整理されつつある例として扱う。

`miku-abc-player` は `-web` 接尾辞を持たないが、実質的には `miku-score` のABC寄り機能を利用しやすくするWeb Appに近い。標準的なMain Applicationの構成例ではなく、用途を限定した特殊なWeb companionとして分類する。

### 古参の miku-project・miku-score 型

`miku-project` と `miku-score` は、Web App、CLI、内部モデル、変換、AI向け操作が歴史的に同居し、その後に Java、Skills、Web、MCP などの companion が増えた古参系列である。

調査時点の `miku-project` 系列には、次のリポジトリがある。

- `miku-project`
- `mikuproject-java`
- `mikuproject-skills`
- `miku-project-web`
- `mikuproject-mcp`

`miku-project` と旧 `mikuproject` の名称がディレクトリ名に混在しているが、これは固定すべき製品構造ではなく、名称移行の途中状態である。別途、`miku-project` 命名へ修正する予定があるため、新仕様では `miku-project` を正規名称として扱い、旧名称を新しい契約へ持ち込まない。

調査時点の各リポジトリ版は、Main `0.14.0`、Java `0.12.0`、Skills `0.12.3`、Web `0.1.0`、MCP `0.0.0` である。版の差は各componentが独立して成長した履歴を示している。新仕様では、単純な版番号一致だけに依存せず、Skillsが利用するruntimeの正確な版、互換契約、由来、digestを記録できる形を検討する。

`miku-score` は製品説明上すでに、楽譜エディターではなく変換、確認、受け渡しのツールとして整理されている。一方で、Web Appを主な配布形態としてきた履歴があり、CLI runtimeとAgent Skillsは後から強化されている。予定されているWeb分離後の構造は、miku-projectのゼロベース再設計にとっても比較材料になる。

### 横断調査から得た標準階層

姉妹リポジトリの実態から、miku-soft の標準階層を次のように仮置きする。

```text
製品の意味・変換・検証契約
├── Node CLI runtime
├── Java CLI runtime（必要な製品のみ）
└── Agent Skills
    ├── runtime選択と検証
    ├── 対話とworkflow
    └── 人間確認とhandoff

後続adapter
├── Web App
├── MCP server
└── Maven Plugin
```

共通化するのは、NodeとJavaの実装コードや各ドメインのデータモデルそのものではない。入出力、diagnostics、終了状態、決定論、artifact provenance、Skillsとの責務分担という製品契約を共通化する。

`miku-project` の初期形は、次を第一候補とする。

```text
miku-project の意味・操作・検証契約
├── Node CLI
├── Java CLI
└── miku-project Agent Skills

後続検討: miku-project-web / miku-project-mcp
```

NodeとJavaのどちらを参照実装とするか、完全な同等性を求めるか、Java固有拡張を許可するかは未決事項とする。

## miku-project の製品定義案

ゼロベースでの `miku-project` は、プロジェクト管理アプリではなく、次の道具として捉える。

> miku-project は、プロジェクト計画データを読み取り、理解可能な形に投影し、人または生成AIによる変更を安全に反映し、検証可能な別形式へ渡す往復変換ツールである。

中核となる操作概念は次の五つである。

```text
読む → 見せる → 変更する → 確かめる → 書き出す
```

- 読む: 入力形式を解釈し、保持できる意味と問題点を明らかにする
- 見せる: 人または生成AIの目的に合う範囲と粒度へ投影する
- 変更する: 許可された操作として変更要求を表現し、次の状態を生成する
- 確かめる: 構造、意味、変更内容、変換損失を検証する
- 書き出す: 外部ツール、人、生成AIが利用できる成果物へ変換する

## 製品形態の比較

ゼロベース検討では、少なくとも次の三つの製品形態を比較する。

### 1. 単純変換器

`convert`、`inspect`、`validate` を中心とし、一方向または単発のデータ変換を担う。

実装と利用方法は最も小さいが、生成AIによる安全な部分変更や往復には不足しやすい。

### 2. 往復変換器

入力の理解、目的別の投影、変更要求、検証、差分、適用、再出力までを扱う。

独自の状態管理アプリにはならず、入力ファイルから次のファイルを生成するステートレスな処理を基本とする。

### 3. 状態管理アプリ

独自の保存領域、プロジェクト管理、履歴管理、編集環境を製品内に持つ。

統合された操作体験は作れるが、製品が大きくなり、miku-soft の変換ツールとしての小ささと透明性を失いやすい。

### 2026-08-09 時点の方向

`miku-project` の第一候補は「2. 往復変換器」とする。

単純変換器よりも生成AIとの協働に適し、状態管理アプリよりも小さく保てる。処理の基本単位は、入力ファイルを読み、検証可能な次のファイルを生成することとする。

## データと正本の考え方

ゼロベース検討の開始時点では、特定のファイル形式を自動的に正本と決めない。

先に決めるべきなのは、次の事項である。

- プロジェクト計画として保持すべき意味
- 形式間の往復で失ってはならない情報
- 失われてもよい情報と、その報告方法
- 人または生成AIに見せるべき情報の範囲
- 人または生成AIに許可する変更操作
- 変更後に成立すべき不変条件

これらを実現するために必要であれば、miku-project 固有の中間表現を設計する。

中間表現は製品の目的ではなく、安全な往復を実現するための手段である。永続的な正本、交換形式、一時的な内部表現のどれとして扱うかは、今後の仕様で明示する。

想定する情報の役割は次のように分ける。

- 原入力: 利用者または外部ツールが保持していた資料
- 中間表現: 変換、検証、差分、再出力に必要な意味を保持する表現
- Projection: 人または生成AIへ目的別に見せる限定表現
- 変更要求: 中間表現全体の書き換えではなく、許可された操作の集合
- 派生成果物: XLSX、Markdown、SVG などの閲覧・共有・受け渡し用出力
- 外部形式: 外部ツールとの交換に利用する形式

## CLI の基本原則

CLI は製品仕様の実行可能な入口とする。

- ローカル処理を基本とする
- 隠れた永続状態を持たない
- 入力ファイルを既定で上書きしない
- 同じ入力と同じオプションから同じ結果を生成する
- 変換結果と diagnostics を分ける
- 警告、欠落、未対応、推測、正規化を機械判定可能にする
- 安定した終了コードを持つ
- 人間向け出力と機械向け出力を混同しない
- 明示的な入力、出力、形式、文字コード、上書き条件を持つ
- 大きな一括処理だけでなく、検査と検証を単独で実行できる
- 非対話実行を基本とし、人間確認は Agent Skills または呼び出し側が明示的に挿入できる
- 失敗時に不完全な成果物を公開せず、出力は可能な限り原子的に確定する

コマンド名は未決定だが、miku-soft 共通の操作語彙として次を候補にする。

```text
inspect
convert
validate
diff
apply
export
```

この候補は、AI Agent に英単語を覚えさせやすいという理由だけで選ぶものではない。操作を責務と意味上の副作用で分類し、呼び出し側が安全な実行順序を組み立てられることを重視する。

| 操作概念 | 主な責務 | 意味上の副作用 |
| --- | --- | --- |
| `inspect` | 入力の形式、内容、能力、問題候補を調べる | なし。入力を変更しない |
| `validate` | schema とドメイン不変条件を検証する | なし。入力を変更しない |
| `convert` | 意味を別の表現へ写し、新しいartifactを生成する | 新規出力を生成するが、入力を変更しない |
| `diff` | 二つの状態またはartifactの意味上の差を示す | なし。比較対象を変更しない |
| `apply` | 許可された変更要求を適用し、次の状態を生成する | プロジェクトの意味を変更するため、事前確認を要する |
| `export` | 外部利用または閲覧用の派生成果物を生成する | 新規出力を生成するが、入力を変更しない |

`convert` と `export` を別コマンドとして残すかなど、最終的な語彙は利用シナリオを決めた後に確定する。すべての miku-soft が同じコマンドを持つ必要はない。同じ意味の操作には同じ語彙、同じ副作用の扱い、同じ入出力原則を使うことを重視する。

AI Agent が理解すべき正本はコマンド名ではなく、versioned schemaで定義されたartifactの役割と操作契約である。miku-projectでは、概念上の受け渡しを次のように捉える。

```text
原入力 → 中間表現 → Projection → 変更要求 → 適用後の状態 → 派生成果物・外部形式
```

各操作は、この流れのどのartifactを入力とし、何を出力するかを明示する。CLIは暗黙の会話履歴やAgent固有の内部状態に依存しない。

## Agent Skills の基本原則

Agent Skills は、CLI の能力を生成AIが安全に利用するための操作層とする。

Agent Skills が担当するのは次の事項である。

- 利用者の目的を解釈する
- 適切な CLI 操作を選択する
- 複数操作の実行順序を組み立てる
- 人間の確認が必要な地点を示す
- diagnostics に応じて続行、再試行、中止を判断する
- 生成AIに渡す Projection と返却形式を指定する
- 適用前後の validate と diff を実施する

Agent Skills は、CLI が所有する変換、検証、正規化、Patch 適用を独自実装しない。CLI の人間向けメッセージを不安定な文字列解析で解釈せず、構造化された結果を利用する。

代表的な安全な往復workflowは次の順序とする。

```text
inspect → validate → Projection生成 → Agentまたは人の判断
        → 変更要求 → validate → diff → 人間確認 → apply → validate → export
```

読み取り、検証、比較は自動実行しやすくし、プロジェクトの意味を変える `apply` の前に人間確認を置けるようにする。確認が不要な利用形態でも、呼び出し側が明示的に変更を許可したことをCLIへ伝えられる契約を定義する。

## 初期検討の対象範囲

初期段階では、次の三つを設計対象とする。

- `miku-project` の意味、操作、変換、検証を定める製品契約と Node CLI
- 同じ製品契約を実装する Java CLI
- CLI runtime を生成AIが安全に利用するための Agent Skills

次の二つは後続段階へ送る。

- `miku-project-web` に相当する Web App
- `miku-project-mcp` に相当する MCP server

`-web` と `-mcp` を後回しにすることは、将来の対応を否定するものではない。先に Node CLI と Java CLI が共有するドメイン契約、入出力、diagnostics、決定論、Agent Skills との責務分担を確定し、その安定した契約の利用者として Web App や MCP server を検討する。

初期仕様では、将来の `-web` や `-mcp` を想定して CLI を過度に一般化しない。一方で、ドメイン処理を CLI 表示層へ閉じ込めず、後続アダプターから再利用可能な境界を保つ。

## Web UI の将来位置づけ

Web UI は初期仕様の対象外とし、miku-project の必須中核とはしない。

必要な場合は、次の用途を持つ任意のアダプターとして検討する。

- 入力内容の確認
- Projection の閲覧
- diff と diagnostics の可視化
- 変更適用前の人間確認
- XLSX、Markdown、SVG などの成果物プレビュー

Node CLI、Java CLI、Agent Skills だけで主要な往復処理が完結することを基本とする。Web UI が存在する場合も、Web UI 固有のドメイン処理を持たず、CLI と共有する製品契約と中核処理を利用する。

## MCP の将来位置づけ

MCP 対応は初期仕様の対象外とする。

将来 `miku-project-mcp` を設ける場合は、MCP server を新たなドメイン実装にはせず、確定済みの miku-project の能力を MCP tools、resources、prompts として公開するプロトコルアダプターとして扱う。

MCP の transport、session、tool schema などの都合を、初期の CLI とドメインモデルの設計要件にはしない。まず CLI と Agent Skills で実際の利用フローを確立し、その後に MCP 化する価値のある安定操作を選ぶ。

## miku-score との関係

`miku-score` も古参の miku-soft として、同じ観点からゼロベース再検討の対象になりうる。

その製品定義候補は次のように表せる。

> miku-score は、楽譜データを読み取り、生成AIが扱える表現へ投影し、安全な変換を行い、検証可能な楽譜形式へ戻す往復変換ツールである。

`miku-project` と `miku-score` で共通化するのは、プロジェクト計画と楽譜のデータモデルではない。共通化するのは、CLI の入出力原則、diagnostics、決定論、Agent Skills との責務分担、破壊的操作の扱いなど、miku-soft としての外形と運用契約である。

`miku-score` では、現在同居している single-file Web App を将来 `-web` companion として分離する作業が計画されている。分離後も楽譜変換の意味と検証はMain Application側が所有し、Web Appはその能力を利用するadapterとする方向が、miku-projectとの比較上の前提になる。

`miku-abc-player` は miku-score family の標準Main Applicationではなく、ABCを中心とした用途限定の特殊なWeb companionとして扱う。

ドメイン処理を無理に共通フレームワークへ押し込まない。共通ライブラリ化は、独立した複数実装から具体的な共通部分が確認された後に判断する。

## 初期段階の非目標

現時点では、次を新仕様の前提または目的にしない。

- 現在の Web UI の維持
- 現在の CLI コマンドとの互換
- `MS Project XML` を正本とすること
- 現在の `ProjectModel` をそのまま維持すること
- 現在の JSON、Projection、Patch 形式をそのまま維持すること
- MS Project の代替となる統合プロジェクト管理アプリの構築
- miku-project 内での独自アカウント、サーバー、クラウド同期
- miku-soft 全製品を一つの巨大な共通フレームワークへ統合すること
- 初期段階での `miku-project-web` の再設計または移行
- 初期段階での `miku-project-mcp` の設計または実装

これらは必要性が確認された場合に、個別の設計判断として再採用できる。

## 未決事項

今後、次の順序で仕様を具体化する。

1. miku-project が解決する利用者の仕事を定義する
2. 最初に対応する利用シナリオを絞る
3. 保持すべきプロジェクト計画上の意味と不変条件を定義する
4. 入力形式、出力形式、変換損失の扱いを決める
5. 中間表現が必要かを判断し、必要なら役割と寿命を定義する
6. 生成AIへ渡す Projection の考え方を定義する
7. 生成AIまたは人に許可する変更操作を定義する
8. validate、diff、apply の安全条件を定義する
9. CLI の最小コマンド体系と構造化出力を定義する
10. Node CLI と Java CLI の参照実装、同等性、固有拡張の方針を定義する
11. Agent Skills の最小ワークフローとruntime provenance契約を定義する
12. 現行実装から再利用するコード、テスト、fixtures を選別する
13. 後方互換性と移行方針を最後に決定する
14. CLI と Agent Skills の仕様が安定した後、`-web` と `-mcp` の必要性を個別に評価する

## 2026-08-09 時点の合意候補

- miku-soft の中心を CLI と Agent Skills の組として捉える
- AI Agent フレンドリーを副次的効果ではなく、miku-soft の主要な製品要件とする
- 人、shell script、CI、AI Agent が同じ操作契約と構造化結果を利用できるようにする
- CLI を決定論的なドメイン処理と検証の正本にする
- Agent Skills を対話、手順、安全な実行の層にする
- miku-project をプロジェクト管理アプリではなく往復変換ツールとして再検討する
- Web UI は必須中核ではなく任意のアダプターとして扱う
- 初期検討は製品契約、Node CLI、Java CLI、Agent Skillsに限定し、`-web` と `-mcp` は後続検討とする
- Node CLI と Java CLI は、同じ製品契約を実装する複数runtimeとして検討する
- Agent Skills はCLI runtimeを利用するworkflow adapterとし、製品の変換処理を重複実装しない
- Skillsが利用するruntimeの版、由来、互換契約、digestを確認可能にする
- `miku-project` を正規名称とし、旧 `mikuproject` 名の混在は移行対象として扱う
- 現行実装は知見とテスト資産として参照するが、新仕様の制約にはしない
- 特定の正本形式や中間表現は、保持すべき意味を定義した後に決める
- 実装や移行を始める前に、製品の一文定義と最初の利用シナリオを確定する
