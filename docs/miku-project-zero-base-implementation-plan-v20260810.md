---
title: miku-project ゼロベース新仕様適合計画 v20260810
description: 2026-08-09版ゼロベース再設計仕様を、製品契約、Node CLI、Java CLI、Agent Skillsへ段階的に反映するための実施計画。
topics:
  - miku-project
  - miku-soft
  - cli
  - java
  - agent-skills
  - migration
category: plan
status: draft
audience:
  - maintainer
  - developer
  - agent
created: 2026-08-10
updated: 2026-08-10
sources:
  - type: local-file
    role: primary
    path: docs/miku-project-zero-base-spec-v20260809.md
    label: miku-project ゼロベース再設計仕様 v20260809
    checked: 2026-08-10
  - type: local-file
    role: current-state
    path: docs/architecture.md
    label: 現行実装アーキテクチャ
    checked: 2026-08-10
---

# miku-project ゼロベース新仕様適合計画 v20260810

## 文書の役割

この文書は、[miku-project ゼロベース再設計仕様 v20260809](miku-project-zero-base-spec-v20260809.md) を実装可能な作業へ分解するための計画書である。

新仕様が「何を目指すか」を定義し、この計画が「どの順序で決め、作り、検証するか」を定義する。[TODO.md](TODO.md) は、この計画から現在着手可能な項目だけを抜き出した実行キューとして扱う。

この計画自体は、未決の製品仕様、互換性、改名、削除、リリースを承認するものではない。各段階の判断ゲートを通過するまでは、後続段階の実装を開始しない。

## 文書の責務分担

| 文書 | 役割 |
| --- | --- |
| `docs/miku-project-zero-base-spec-v20260809.md` | 新しい製品像、設計原則、対象範囲、非目標 |
| この文書 | 段階、依存関係、成果物、判断ゲート、完了条件 |
| `docs/TODO.md` | 現在着手できる未完了作業と計画ID |
| `docs/architecture.md` / `docs/spec.md` | 現行実装の構造と仕様を示す証拠 |
| `docs/migration-worklog.md` | 完了した移行作業と検証結果の記録 |
| `docs/refactoring-playbook.md` | 実装変更時の安全な分割・検証手順 |

新仕様が確定するまで、現行の `architecture.md` と `spec.md` を暗黙に書き換えて新仕様へ見せかけない。新しい契約が実装された段階で、現行仕様、移行ガイド、履歴資料の役割を改めて整理する。

## 計画の対象

初期計画の対象は次のとおりである。

- `miku-project` の意味、操作、変換、検証を定める製品契約
- 製品契約の Node CLI 実装
- 同じ製品契約の Java CLI 実装
- CLI runtimeを利用する `miku-project` Agent Skills
- 現行実装から再利用するコード、fixtures、tests、artifact buildの選別
- 旧CLI、旧データ形式、旧名称からの互換性と移行

次は初期計画の実装対象外とする。

- `miku-project-web` の再設計
- `miku-project-mcp` の再設計
- WebまたはMCPの都合を先取りした製品契約の一般化
- 現行 `mikuproject` 名の機械的な一括置換
- 未選択の帳票、見た目、ドメイン機能の拡張

WebとMCPについては、CLIとAgent Skillsが安定した後の再評価ゲートだけを置く。既存利用者を壊さないための互換性保守は継続できるが、新仕様の初期設計をそれらへ従属させない。

## 現在の基準状態

### 利用可能な資産

現行実装には、次の再利用候補がある。

- MS Project XMLのparse、normalize、validate、export
- `ProjectModel` と階層、依存、calendar、resource、assignmentの処理
- XLSXとworkbook JSONのimport/export
- AI向けProjection、Patch検証・適用、state summary、state diff
- WBS XLSX、Markdown、SVG、Mermaidなどの派生出力
- structured diagnostics、終了コード `0 / 1 / 2`、stdin/stdout、Base64 I/O
- 単一 `.mjs` CLI runtime、source archive、決定論的なbundle生成
- XML roundtrip、Patch、XLSX、帳票、CLI bundleの回帰tests

これらは再利用候補であり、新仕様の正本ではない。新しい意味契約とfixturesに合格した部分だけを採用する。

### 現行実装との主な差

- 現行文書は `MS Project XML` を意味の基軸、`ProjectModel` を中立表現としているが、新仕様ではどちらも未決である
- 現行CLIは `ai / state / import / export / report` を中心とし、新仕様候補の `inspect / validate / diff / apply / convert / export` と一致していない
- 現行CLIは出力先を対話なしで上書きするが、新仕様では安全な既定値と明示的な上書き条件が必要である
- 現行diagnosticsには再利用価値があるが、安定したcode、severity、location、loss、normalization provenanceが不足する
- 現行の `state diff` はworkbookの限定import対象を利用しており、完全なsemantic diffとは限らない
- CLI entrypointがcommand dispatch、I/O、diagnostics、diff、formattingを広く所有しているため、動作固定後に責務分割が必要である
- 現行test runnerでは `fast / full / all` の実質差が小さく、一部test fileが明示一覧から外れているため、suite契約を見直す必要がある
- CLI runtimeには、Skillsが機械的に検証できるartifact manifestが不足している
- Javaはmoving upstreamのstraight conversion、SkillsはCLIとMCPの両backendを含むため、新契約の確定前に追随を続けると旧設計を固定する

## リポジトリ別の責務

| Workstream | 主なリポジトリ | 責務 |
| --- | --- | --- |
| Product Contract | `miku-project` | 利用シナリオ、意味、不変条件、形式、損失、変更、CLI契約、共通fixtures |
| Node CLI | `miku-project` | 参照実装候補、単一`.mjs`、manifest、Node contract tests |
| Java CLI | 現 `mikuproject-java`、将来 `miku-project-java` | 固定済み契約のJava実装、fat JAR、共通conformance |
| Agent Skills | 現 `mikuproject-skills`、将来 `miku-project-skills` | CLI workflow、runtime選択、provenance/digest検証、human gate |
| Web | `miku-project-web` | 後続評価。初期計画では既存互換性保守のみ |
| MCP | 現 `mikuproject-mcp`、将来 `miku-project-mcp` | 後続評価。初期計画では新規設計を行わない |

## 依存関係と判断ゲート

```text
ZB-P0 利用者ジョブと現状棚卸し
  ↓ G0: v1シナリオ承認
ZB-P1 意味と不変条件
  ↓ G1: semantic contract承認
ZB-P2 形式・損失・中間表現・変更契約
  ↓ G2: artifact/change contract承認
ZB-P3 CLI・diagnostics・共通conformance契約
  ↓ G3: 実装開始承認
ZB-P4 Node CLI vertical sliceとv1完成
  ↓ G4: Node contract release固定
ZB-P5 Java CLI適合
  ↓ G5: cross-runtime適合承認
ZB-P6 Agent Skills適合
  ↓ G6: isolated skill bundle承認
ZB-P7 互換性・改名・release移行
  ↓ G7: vNext完了
ZB-P8 Web / MCP再評価
```

`ZB-P0`から`ZB-P3`までは仕様と契約を決める段階である。製品コードの大幅な再編を始めない。Java実装はmoving Node sourceを追いかけず、`G4`で固定された契約releaseを入力とする。Skillsのworkflow草案は早期に検討できるが、runtime統合はNode/Java artifactが固定されるまで完了扱いにしない。

## ZB-P0 利用者ジョブと現状棚卸し

依存: なし  
主担当: Product Contract
状態: 完了（`G0` 通過: 2026-08-10）

### 作業

- [x] `ZB-P0.1` 新仕様、実施計画、TODO、現行仕様、worklogの文書authorityを確認する
- [x] `ZB-P0.2` 現行CLI command、core API、入力形式、出力形式、diagnostics、testsをcapability matrixにする
- [x] `ZB-P0.3` 現行module、fixture、test、artifactを `reuse / evidence / rewrite / defer / drop` に仮分類する
- [x] `ZB-P0.4` R1/C1を、v1で証明する利用者ジョブのG0提案として選ぶ
  - 既存XMLを読み、inspect、validate、AI向けProjectionを得る
  - Projectionから変更要求を作り、validate、diff、human gate、apply、再validate、exportする
  - 新規draftをvalidateし、project stateを作る候補はv1 scenarioに選ばず後続へ送る
- [x] `ZB-P0.5` 各候補について、actor、入力、出力、成功例、失敗例、人間確認点、非目標を書く
- [x] `ZB-P0.6` 現行TODOを選択シナリオとの関係で再分類する
- [x] `ZB-P0.7` AI Agentフレンドリーを主要要件として、人、shell script、CI、AI Agentが同じ操作契約を利用する受入観点を定義する
- [x] `ZB-P0.8` 各scenarioの操作を `read-only / artifact生成 / 意味変更` に仮分類し、構造化結果から選べる次の行動とhuman gateを対応づける

### 成果物

- [docs/miku-project-zero-base-scenarios-v1.md](miku-project-zero-base-scenarios-v1.md)
- [docs/miku-project-current-capability-matrix-v20260810.md](miku-project-current-capability-matrix-v20260810.md)
- [docs/miku-project-zero-base-reuse-inventory-v20260810.md](miku-project-zero-base-reuse-inventory-v20260810.md)
- triage済みの `docs/TODO.md`

### Gate G0

次が明示的に承認されれば `G0` 通過とする。

- 一文の製品定義
- primary actorと利用者ジョブ
- v1で証明する1〜2個のend-to-end scenario
- scenarioごとの入力、出力、human gate、失敗時動作
- 人、shell script、CI、AI Agentが同じ契約を使い、構造化結果から安全な次の行動を判断できること
- v1に含めない機能と形式
- Web/MCPを必要とせずscenarioが完結すること

結果: 2026-08-10に通過。R1（読み取り・理解）とC1（安全な局所変更）をv1 scenarioとし、primary actorは人と協働するAI Agentとする。最初のC1で許可する意味変更はtaskの`percentComplete`更新だけであり、dependency、resource、assignmentは観測・保持対象とする。XMLの成功判定はsemantic equivalenceで行い、byte列一致を要求しない。

## ZB-P1 意味と不変条件

依存: `G0`  
主担当: Product Contract
状態: 進行中

### 作業

- [ ] `ZB-P1.1` project、task、階層、identity、順序の最小semantic scopeを定義する
- [ ] `ZB-P1.2` 日付、日時、duration、milestone、summary、進捗の意味と不変条件を定義する
- [ ] `ZB-P1.3` dependencyの最小scopeとcycle、link type、lagの規則を定義する
- [ ] `ZB-P1.4` resource、assignment、calendar、actual、baseline、timephased dataをv1に含めるか個別に決める
- [ ] `ZB-P1.5` unknown field、拡張field、opaque preservationの扱いを定義する
- [ ] `ZB-P1.6` stable identity、selector、並び順、文字列normalization、timezoneの規則を定義する
- [ ] `ZB-P1.7` valid、invalid、boundaryのsemantic fixturesを作る

### 成果物

- `docs/miku-project-semantic-contract-v1.md`
- semantic fixture set
- v1 domain scope table

### Gate G1

すべてのv1 fieldとoperation candidateが、意味、不変条件、identity、順序、valid/invalid境界へ対応づけられていること。現行 `ProjectModel` の形を理由にscopeを決めていないこと。

## ZB-P2 形式・損失・中間表現・変更契約

依存: `G1`  
主担当: Product Contract

### 作業

- [ ] `ZB-P2.1` 原入力、中間表現、Projection、変更要求、派生成果物、外部形式の役割を確定する
- [ ] `ZB-P2.2` miku-project固有の中間表現が必要か判断する
- [ ] `ZB-P2.3` 中間表現を採用する場合、internal、exchange、persistentのどの役割かとschema versionを定義する
- [ ] `ZB-P2.4` v1形式ごとに `read / write / roundtrip / loss / unsupported` matrixを作る
- [ ] `ZB-P2.5` preservationを `required / normalized / lossy-with-warning / unsupported-error / opaque-preserved` に分類する
- [ ] `ZB-P2.6` AI向けProjectionのpurpose、範囲、情報量、規則を定義する
- [ ] `ZB-P2.7` whole-state replacementとoperation-based changeの境界を決める
- [ ] `ZB-P2.8` 許可operation、selector、atomicity、precondition、validation、diff、apply後検証を定義する
- [ ] `ZB-P2.9` loss、normalization、ignored change、unsupported dataのprovenance表現を定義する
- [ ] `ZB-P2.10` 各operationについて入力artifact、出力artifact、役割、寿命、schema versionを対応づけ、会話履歴やAgent固有のhidden stateへ依存しないことを確認する

### 成果物

- `docs/miku-project-format-and-loss-contract-v1.md`
- `docs/miku-project-change-contract-v1.md`
- schemaとvalid/invalid examples
- format/field/loss matrix

### Gate G2

v1 scenarioの全artifactに役割、寿命、schema versionがあり、operation間の受け渡しが明示され、各変換に損失規則があり、各変更に検証とatomicity規則があること。CLI操作が会話履歴やAgent固有のhidden stateへ依存しないこと。`MS Project XML`、現行 `ProjectModel`、workbook JSON、新しいIRのいずれも、検討なしに正本へ置かれていないこと。

## ZB-P3 CLI・diagnostics・共通conformance契約

依存: `G2`  
主担当: Product Contract

### 作業

- [ ] `ZB-P3.1` v1 CLIの最小command matrixを確定する
- [ ] `ZB-P3.2` args、stdin request、stdout、stderr、file output、binary、Base64、encoding、BOMの規則を定義する
- [ ] `ZB-P3.3` hidden stateなし、既定上書きなし、`--force`等の明示条件、atomic write、部分成果物非公開を定義する
- [ ] `ZB-P3.4` versioned result/diagnostics schemaを定義する
- [ ] `ZB-P3.5` code、severity、scope、path、status、I/O metadata、loss、normalization、retryabilityを定義する
- [ ] `ZB-P3.6` exit codeとsuccess、validation-failed、invalid-usage、internal-errorの境界を定義する
- [ ] `ZB-P3.7` Nodeを参照実装とするか、Node/Javaを対称実装とするか決める
- [ ] `ZB-P3.8` Java固有extensionを許可する場合のcapability表現を定義する
- [ ] `ZB-P3.9` Node/Java/Skillsが共有するchecked-in conformance fixturesとgolden resultを設計する
- [ ] `ZB-P3.10` product contract version、runtime version、fixture suite version、asset、source、SHA-256、capability setを持つruntime manifestを定義する
- [ ] `ZB-P3.11` command候補を責務と意味上の副作用で `read-only / artifact生成 / 意味変更` に分類する
- [ ] `ZB-P3.12` 非対話実行を基本とし、意味変更の明示許可、human gateの挿入点、再試行・中止・次操作の機械判定条件を定義する

### 成果物

- `docs/miku-project-cli-contract-v1.md`
- result/diagnostics/runtime-manifest schema
- `testdata/conformance/` の共通fixturesとgolden
- cross-runtime capability matrix

### Gate G3

CLI examplesがschemaで検証でき、exit状態が一意で、unsafe overwriteや失敗が部分成果物を残さず、Node/Javaの適合対象と許可差分が明示されていること。人、shell script、CI、AI Agentが人間向けmessageの文字列解析なしに、操作の副作用、成功・失敗、再試行可能性、安全な次の行動を判断できること。このgate通過後に製品実装を開始する。

## ZB-P4 Node CLI vertical sliceとv1完成

依存: `G3`  
主担当: Node CLI

### 作業

- [ ] `ZB-P4.1` 現行CLIの互換動作をcontract testsで固定する
- [ ] `ZB-P4.2` 現行test suiteの `fast / full / all` を実態に合わせ、一部testの実行漏れを解消する
- [ ] `ZB-P4.3` semantic変更より先に、CLIのparser、command service、I/O、diagnostics、formattingを分離する
- [ ] `ZB-P4.4` G0で選んだscenario一つだけを新契約でend-to-end実装する
- [ ] `ZB-P4.5` whole-project inspect/validate、semantic diff、pre/post apply validationを選択scopeに応じて実装する
- [ ] `ZB-P4.6` safe output、atomic write、明示overwrite、structured loss reportingを実装する
- [ ] `ZB-P4.7` 既存coreの再利用部分を新conformance fixturesで検証する
- [ ] `ZB-P4.8` 選択scopeの残りを一sliceずつ追加し、capability matrixを更新する
- [ ] `ZB-P4.9` versioned single `.mjs`、sources、runtime manifest、SHA-256を生成する
- [ ] `ZB-P4.10` repository外のclean temporary directoryでbundle smokeを実行する

### 実装上の境界

- 現行commandは互換性方針が決まるまでcompatibility surfaceとして残す
- browser runtimeは既存 `miku-project-web` の互換性保守に必要な間は維持するが、新CLI契約を支配させない
- reportの見た目改善や未選択formatを、vertical sliceへ混ぜない
- CLI責務分割とsemantic変更を同じ変更単位で行わない

### Gate G4

- v1 scenarioがclean directoryで完結する
- 同一入力の反復実行が、契約どおりbyte-identicalまたはsemantic-equivalentになる
- invalid input、digest mismatch、unsafe overwriteで部分成果物を残さない
- repository外でsingle `.mjs` が動作する
- relevant tests、contract suite、`npm run build:full` が成功する
- Node contract releaseとartifact manifestが固定される

## ZB-P5 Java CLI適合

依存: `G4`  
主担当: Java CLI

### 作業

- [ ] `ZB-P5.1` moving Node subtree追随を止め、固定されたcontract release単位の入力へ切り替える
- [ ] `ZB-P5.2` 共通fixture、golden、result/diagnostics schemaをJava testから利用する
- [ ] `ZB-P5.3` required command、option、I/O、exit状態、atomicityを実装する
- [ ] `ZB-P5.4` JSON/textは必要箇所でexact比較し、XML/XLSXは契約に応じてsemantic比較する
- [ ] `ZB-P5.5` Node-only、Java-only、未対応差分をcapability matrixに記録する
- [ ] `ZB-P5.6` fat JAR、sources、runtime manifest、SHA-256を生成する
- [ ] `ZB-P5.7` standalone `java -jar` smokeをclean directoryで実行する

### Gate G5

共通contractのconformance suiteがNode/Javaの両方で成功し、差分がゼロまたは明示的に承認されていること。artifactがproduct contract、runtime、fixture suite、capabilityを識別できること。

## ZB-P6 Agent Skills適合

依存: workflow設計は`G2`後、runtime統合完了は`G5`後  
主担当: Agent Skills

### 作業

- [ ] `ZB-P6.1` G0で選んだscenarioだけをAgent workflowへする
- [ ] `ZB-P6.2` `inspect → validate → Projection生成 → Agentまたは人の判断 → 変更要求 → validate → diff → human gate → apply → post-validate → export` の安全な順序を定義する
- [ ] `ZB-P6.3` 初期backendをCLIだけに限定し、現行MCP backendはlegacy/deferredとして隔離する
- [ ] `ZB-P6.4` Node/Javaの選択順、fallback条件、fallback理由を構造化する
- [ ] `ZB-P6.5` `runtime-manifest.json` の明示assetだけを選び、最新版ファイル名の自動選択を廃止する
- [ ] `ZB-P6.6` version、compatibility、source、asset role、SHA-256を各operation開始前に検証する
- [ ] `ZB-P6.7` CLIの変換、validation、Patch logicをSkillsへ重複実装しない
- [ ] `ZB-P6.8` digest改変拒否、runtime欠落、capability不一致、fallback、human gateをtestする
- [ ] `ZB-P6.9` install後の隔離されたSkill bundleだけでrepresentative scenarioをsmokeする
- [ ] `ZB-P6.10` structured resultとdiagnosticsだけを使って続行、再試行、中止、人間確認を分岐できることをtestする

### Gate G6

固定runtimeだけでoffline実行でき、digestまたはcompatibility不一致でfail closedし、Skills内にproduct semanticsの派生実装がなく、human gateを含む代表workflowがisolated bundleで成功すること。

## ZB-P7 互換性・改名・release移行

依存: `G6`  
主担当: Product Contract、Node CLI、Java CLI、Agent Skills

### 作業

- [ ] `ZB-P7.1` 旧command、schema、global、bin alias、artifact名、`mikuproject_*` identifierを棚卸しする
- [ ] `ZB-P7.2` 各項目を `retain / deprecate / convert / remove` に分類し、互換期間を決める
- [ ] `ZB-P7.3` `mikuproject-java` と `mikuproject-skills` のrepo、directory、package命名を `miku-project` へ移行する
- [ ] `ZB-P7.4` persisted format identifierの改名を製品名移行と分離して判断する
- [ ] `ZB-P7.5` current formatからvNext formatへのconverterまたはactionable diagnosticを用意する
- [ ] `ZB-P7.6` README、architecture、current spec、migration guide、worklogのauthorityを更新する
- [ ] `ZB-P7.7` Node → Java → Skills のrelease順とcompatibility matrixを確定する
- [ ] `ZB-P7.8` browser runtimeを互換artifactとして維持するか、Main Applicationの既定buildから外すかをdownstreamと調整する
- [ ] `ZB-P7.9` clean install、`.mjs`、`.jar`、Skill bundle、provenanceのrelease acceptanceを実行する

### Gate G7

旧利用者がdocumented migrationを実行できるか、actionable diagnosticを受け取れること。新runtimeとSkillsがsource checkoutなしで動作し、すべてのartifactと互換関係をmanifestから追跡できること。

## ZB-P8 Web・MCP再評価

依存: `G7`

この段階までは、Web/MCPの新規実装taskを作らない。`G7`後に、実際に安定したCLI操作のうち、WebまたはMCPで公開する価値があるものだけを評価する。

- [ ] `ZB-P8.1` Webで必要な閲覧、可視化、human gateを評価する
- [ ] `ZB-P8.2` MCP tools/resources/promptsに適した安定操作を評価する
- [ ] `ZB-P8.3` adapter固有のsecurity、workspace、transport、artifact deliveryを別仕様にする

Web/MCPはproduct semanticsを実装せず、確定済みcontractのadapterとする。

## 現行TODOの取扱い

現行 [TODO.md](TODO.md) の詳細項目は削除せず、次の基準で再評価する。

| 現行TODOの種類 | 取扱い | 再評価時期 |
| --- | --- | --- |
| XML/XLSX roundtrip、dependency、calendar、validation | semantic/fixture候補として保持 | `ZB-P0`〜`ZB-P2` |
| partial apply、diff、scoped Projection | change workflowの証拠として保持 | `ZB-P0`〜`ZB-P2` |
| actual、EV、baseline、timephased、拡張field | domain scope候補。実装を約束しない | `ZB-P1` |
| XLSX、Markdown、SVG、Mermaid、sample | derived output候補。選択scenarioに必要なものだけ採用 | `ZB-P0`、`ZB-P4` |
| 帳票の色、線幅、layoutなどの見た目 | 初期core対象外 | `G7`後または関連adapter |
| Overview、Output、画面内task操作 | `miku-project-web` 側の後続候補 | `ZB-P8` |
| Skillsへのruntime受け渡し | provenanceを含めて再定義 | `ZB-P6` |
| source分割、build速度、test時間 | 選択sliceの実装安全性に必要な範囲だけ実施 | `ZB-P4` |
| `docs/spec.md` のdrift | currentとtargetのauthorityを先に明示し、移行時に解消 | `ZB-P0`、`ZB-P7` |

`G0`通過までは、現行TODOの「最優先」表現を新仕様の優先順位として扱わない。critical bug、security、既存releaseの互換性維持を除き、新機能や見た目改善へ着手しない。

## Verification Matrix

| 段階 | 必須検証 |
| --- | --- |
| 文書・契約 | Markdown link、schema example、`git diff --check`、authority整合 |
| Semantic Contract | valid/invalid/boundary fixtures、不変条件、loss matrix |
| Node CLI | contract tests、determinism、atomicity、clean bundle smoke、`npm run build:full` |
| Java CLI | 共通conformance、semantic roundtrip、`mvn test`、fat JAR smoke |
| Agent Skills | structure、manifest/digest、fallback、human gate、isolated bundle smoke |
| AI Agentフレンドリー | 副作用分類、schema validation、message非依存の分岐、human gate、失敗時の安全な次行動 |
| Release | Node/Java/Skills compatibility、asset provenance、clean install、migration smoke |

## 直近の実行キュー

現在着手可能なのは `ZB-P1` だけである。

1. R1/C1に必要なproject、task、階層、identity、順序の最小semantic scopeを定義する
2. 日時、duration、milestone、summary、進捗の意味と不変条件を定義する
3. dependency、resource、assignment、calendarの観測・保持・編集範囲を定義する
4. valid、invalid、boundaryのsemantic fixtureを設計する
5. Gate G1でsemantic contractを承認する

`G2`より後の項目は、semantic contractの判断を待つplanned workであり、現在の着手候補ではない。

## 計画の更新方法

- 詳細な依存、成果物、完了条件はこの文書で更新する
- 現在着手可能な項目だけを `docs/TODO.md` に置く
- 判断結果は該当するcontract文書へ記録する
- 完了した移行と検証結果は `docs/migration-worklog.md` へ記録する
- scopeを変える場合は、新仕様、計画、TODOを同じ変更で整合させる
- Web/MCPを前倒しする場合は、新仕様の初期対象範囲を明示的に改訂する
