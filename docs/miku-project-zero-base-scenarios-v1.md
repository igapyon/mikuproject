---
title: miku-project ゼロベース v1 利用シナリオ
description: G0で承認されたv1の利用者ジョブ、AI Agent受入条件、非目標を記録する。
topics:
  - miku-project
  - cli
  - agent-skills
  - specification
category: specification
status: approved
audience:
  - maintainer
  - developer
  - agent
created: 2026-08-10
updated: 2026-08-10
---

# miku-project ゼロベース v1 利用シナリオ

## 文書の位置づけ

この文書は、[ゼロベース再設計仕様](miku-project-zero-base-spec-v20260809.md) と [実施計画](miku-project-zero-base-implementation-plan-v20260810.md) の `Gate G0` に必要な利用シナリオである。ここで選ぶのはv1で実証する利用者ジョブであり、MS Project XML、現行 `ProjectModel`、workbook JSONのいずれかを正本と決めることではない。形式、保持意味、CLIの最終語彙は後続の `G1`〜`G3` で確定する。

`G0` は 2026-08-10 に承認された。承認範囲はR1/C1、AI Agentをprimary actorとする同一CLI契約、ならびに本稿の非目標である。次の `G1` では、この承認範囲を超えずに最小semantic contractを定義する。

本稿の提案は、次の二つをv1のend-to-end scenarioとして採用することである。

1. 読み取り・理解: プロジェクト計画を検査、検証し、AI Agentまたは人に必要な範囲だけを渡す
2. 安全な局所変更: 明示された変更要求を検証、比較、人間確認したうえで、次のプロジェクト成果物を生成する

新規計画の全面生成、XLSXの自由編集、帳票生成、Web UI、MCPは、この二つを実証した後に評価する。

primary actorは、人の目的を受けてCLIを操作するAI Agentとする。人は変更の目的とmeaningful changeの承認を担い、shell scriptとCIは同じCLI契約を非対話で利用するsecondary actorとする。この区別は機能差を意味せず、すべてのactorが同じartifactと構造化resultを利用することを意味する。

## G0の代表fixtureと具体例

R1/C1を抽象的な要望に留めないため、最初の共通fixture候補には既存の [testdata/dependency.xml](../testdata/dependency.xml) を使う。このfixtureには、二つのtask（`UID 1: Prepare`、`UID 2: Execute`）、task間のdependency、resource、assignmentが含まれる。`G3`で新しいconformance corpusを作るまでは、既存fixtureを現状証拠として参照する。

| scenario | 入力 | 具体的な要求 | 成功時に観測できる結果 |
| --- | --- | --- | --- |
| R1 | `testdata/dependency.xml` | 全taskを対象に、構造とdependencyを理解するためのProjectionを得る | `Prepare` と `Execute`、`Prepare → Execute` のdependency、resource/assignmentの取扱い、対応可否・diagnosticsが構造化して分かる |
| C1 | `testdata/dependency.xml` と変更要求 | `UID 2` の進捗を `0` から `50` へ変更する。対象identity、現値、許可operationをpreconditionにする | diffにはこの進捗変更だけが示され、dependency、resource、assignment、`UID 1`は保持される。元のXMLは変更せず、承認後に別の次状態artifactを生成する |

この例でのPatch表現の具体的なJSON schemaや外部出力のbyte列は、まだ決めない。G0で固定するのは「どの意味を観測し、どの意味だけを変更し、何を保持すべきか」である。`G1`〜`G3`で、identity、precondition、診断、出力形式、semantic equivalenceを契約化する。

最初のC1で許可する意味変更は、taskを安定したUIDで選んだ `percentComplete` の更新だけとする。dependency、resource、assignmentはR1で観測し、C1で保持を検証する対象であり、v1の編集面には含めない。XMLの成功判定はbyte列の一致ではなく、この文書で指定した意味が保持・変更されたこととする。

## v1共通の受入原則

v1では、人、shell script、CI、AI Agentが同じCLI契約を利用する。Agent Skillsはこの契約を安全な順序で呼び出すworkflow adapterであり、ドメイン変換やPatch適用を再実装しない。

- CLIは非対話実行を基本とし、入力、出力、artifactの役割、上書き許可を明示する
- result、diagnostics、exit statusは機械可読かつversioned schemaで取得できる
- Agentは人間向けメッセージの文字列解析ではなく、status、diagnostic code、severity、retryability、artifact metadataから次の行動を判断できる
- 読み取りと検証は入力artifactを変更しない
- artifact生成は入力artifactを変更しない。既存出力の置換には明示許可を要する
- プロジェクトの意味を変更する操作は、`diff`を示した後のhuman gateを通過しなければ実行しない
- 失敗、validation error、unsafe overwriteでは、出力を公開せず、入力を変更しない

## Scenario R1: 読み取り・検証・目的別Projection

### 利用者ジョブ

利用者は既存のプロジェクト計画を渡し、「この計画の構造、問題、指定した範囲を理解したい」と依頼する。AI Agentは必要な情報だけを受け取り、変更は行わない。

### actorと入力

| actor | 入力 | 目的 |
| --- | --- | --- |
| 人またはAI Agent | 外部プロジェクトartifact | 内容と対応可能範囲を知る |
| CLI | 外部artifact、明示した入力形式または検出規則 | `inspect` と `validate` を行う |
| CLI | 検証可能な意味表現とProjection要求 | 用途を限定したProjectionを生成する |

最初のfixtureは、現行実装に読み書き資産がある小規模なMS Project XMLを候補とする。ただし、これはv1を実証するための外部入力例であり、XMLを正本とする決定ではない。

### 操作の流れ

```text
外部artifact
  → inspect
  → validate
  → Projection生成
  → 人またはAI Agentが理解・判断
```

| 区分 | 操作 | 入力 | 出力 | 意味上の副作用 |
| --- | --- | --- | --- | --- |
| read-only | `inspect` | 外部artifact | 形式、内容概要、能力、diagnostics | なし |
| read-only | `validate` | 解釈済みの意味表現 | valid/invalid、diagnostics | なし |
| artifact生成 | Projection生成 | 意味表現、目的、範囲 | AIまたは人向けProjection | 入力を変更しない |

### 成功条件

- 入力の形式、内容概要、利用可能な能力、diagnosticsを構造化して得られる
- validな入力から、目的と範囲を明示したProjectionを生成できる
- Projectionには、変更可能な範囲と、変更を求める場合の返却形式を示せる
- 同じ入力、同じ要求、同じruntimeでは、同じ結果または契約で定めたsemantic equivalenceの結果を得られる
- Agentは構造化結果だけで、`続行`、`入力修正を依頼`、`対象範囲を狭める`、`変更要求を作る`のいずれかを選べる

### 失敗時動作

- 解釈不能、未対応形式、validation failureは、stable code、severity、scope/path、必要ならretryabilityを持つdiagnosticsで返す
- 失敗時はProjectionを成功成果物として公開しない。部分的な観測情報を返す場合は、その状態と制約を明示する
- 入力artifactは一切変更しない

### 非目標

- プロジェクトを変更すること
- 全プロジェクトを常にAIへ渡すこと
- Web UIによる編集・可視化
- XLSX、SVG、Markdownなどの派生成果物をv1の必須出力にすること

## Scenario C1: Patchによる安全な局所変更往復

### 利用者ジョブ

利用者はR1で得たProjectionをもとに、AI Agentまたは人が作った局所変更要求を確認して、既存計画から次の計画artifactを生成したい。変更の内容と安全性を、適用前に確認できなければならない。

### actorと入力

| actor | 入力 | 目的 |
| --- | --- | --- |
| AI Agentまたは人 | R1のProjection、変更意図 | 許可された変更要求を作る |
| CLI | 現在の意味表現、変更要求 | precondition、許可operation、不変条件を検証する |
| 人 | semantic diff | 意味変更を承認または中止する |
| CLI | 承認済み変更要求 | 次の状態と外部artifactを生成する |

変更要求は中間表現全体の置換ではなく、許可されたoperationの集合とする。最初のC1で許可するoperationは、taskの`percentComplete`を更新する一種類だけである。追加、削除、移動、dependency変更はhuman gateが必要になる操作の例ではあるが、v1の操作候補には含めない。具体的な表現、selector、preconditionは `G2` で決める。

### 操作の流れ

```text
現在のartifact + 変更要求
  → 変更要求のvalidate
  → semantic diff
  → human gate
  → apply
  → post-apply validate
  → 次の外部artifactをexport
```

| 区分 | 操作 | 入力 | 出力 | 意味上の副作用 |
| --- | --- | --- | --- | --- |
| read-only | 変更要求の`validate` | 現在の状態、変更要求 | valid/invalid、diagnostics | なし |
| read-only | `diff` | 現在の状態、適用予定の状態 | semantic diff、loss/provenance | なし |
| 意味変更 | `apply` | 現在の状態、承認済み変更要求 | 次の状態 | プロジェクト意味を変更する |
| artifact生成 | `export` | 次の状態、明示した出力形式 | 次の外部artifact | 入力を変更しない |

### human gate

`apply`の前に、少なくとも次を人が確認できるようにする。

- 変更対象のidentityと変更前後の値
- 追加、削除、移動、依存変更などのoperation種別
- validation warning、normalization、loss、unsupported data
- 変更が許可範囲を超えないこと
- 生成する出力artifactの場所と、既存ファイルを置換する場合の明示許可

CLIは標準では対話的な確認を要求しない。Agent Skillsまたは呼び出し側がhuman gateを実現し、CLIへは承認済みであることと明示的な出力許可を渡す契約にする。

### 成功条件

- Patchが対象、操作、precondition、許可範囲を満たすか構造化して検証できる
- `diff`が変更をsemanticな単位で示し、人が承認可否を判断できる
- 承認後だけ次の状態と外部artifactを生成する
- apply後に不変条件を再検証し、入力artifactを上書きせずに出力する
- Agentは構造化結果だけで、`修正依頼`、`人へ承認依頼`、`適用`、`中止`を選べる

### 失敗時動作

- 不正Patch、precondition failure、validation failure、未対応operationでは、`apply`とexportを行わない
- 出力エラーまたはunsafe overwriteでは、部分成果物を公開しない
- 人が中止した場合は、変更要求とdiffを監査用artifactとして残すかを呼び出し側の明示方針で決める。元のプロジェクトartifactは変更しない

### 非目標

- Agentがhuman gateなしに意味変更を確定すること
- 既存artifactの暗黙上書き
- 変更要求以外の自由な全量JSON書換え
- 完全なXLSX roundtrip、帳票、Web UI、MCP

## G0で承認した事項

2026-08-10に、次を承認した。

1. v1はR1とC1の二つのscenarioを実証する
2. 最初の外部入力fixtureには小規模なMS Project XMLを使うが、正本形式は未決定とする
3. AIとの受け渡しは、外部形式そのものではなく目的別Projectionと変更要求を使う
4. `apply`前のhuman gate、構造化diagnostics、明示出力、失敗時非破壊をv1の必須条件とする
5. XLSX、帳票、SVG、Web、MCP、広範なドメイン拡張はv1の必須範囲から外す

承認後、`G1`ではこの二つのscenarioに必要な意味と不変条件だけを最小scopeとして定義する。
