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
updated: 2026-08-11
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
- 現行CLIは `ai / state / import / export / report` を中心とし、G3承認済みの `inspect / validate / plan-change / apply-change / verify-artifact` と一致していない
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
| Node CLI | `miku-project` | v1参照実装、単一`.mjs`、manifest、Node contract tests |
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
  - Projectionから変更要求を作り、validate、diff、output preflight、human gate、apply、再validate、artifact setをpublishする
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
状態: 完了（`Gate G1` 通過: 2026-08-10）

### 作業

- [x] `ZB-P1.1` task構造を空も許すordered forestとして定義し、root、parent、sibling order、summary、stable identity、外部形式の疑似taskを不変条件へ落とす
- [x] `ZB-P1.2` R1/C1が扱うfieldを `required / optional-preserved / unsupported` に分類したv1 domain scope tableを作る
- [x] `ZB-P1.3` project/taskの日時、declared duration、milestone、summary、整数`percentComplete`の意味とvalid/invalid境界を確定する
- [x] `ZB-P1.4` dependencyをfinish-to-start・lag 0へ限定し、tuple identity、重複、collection順序、欠損参照、自己参照、cycle、未対応type/lagの扱いを確定する
- [x] `ZB-P1.5` resource、assignment、calendar、unassigned、空collection、collection順序、unknown field、actual、baseline、timephased、extended dataのscopeとfail-closed規則を確定する
- [x] `ZB-P1.6` C1をleaf taskの`percentComplete`更新一種類へ限定し、selectorの安定期間、precondition、変更後のsemantic equivalenceを定義する
- [x] `ZB-P1.7` `docs/miku-project-semantic-fixture-catalog-v1.md`を作り、valid、invalid、boundary、unsupported、C1 rejectのfixture ID、入力差分、期待結果、不変条件を対応づける
- [x] `ZB-P1.8` scenario、semantic contract、fixture catalog、実施計画、TODOの用語とscopeを横断確認し、G1 review checklistを完了する

### G1修正の既定案

次をG1承認候補の既定案として文書化する。異なる結論を採る場合は、該当task内で代案、理由、R1/C1への影響を明記する。

- task構造は単一treeではなく、複数rootを許すordered forestとする
- 空のtask forestと空のdependency/resource/assignment/calendar collectionはvalidとする
- task UIDは一回の入力artifactから次artifactを生成する処理単位でstableとし、名称、行番号、outline numberをselectorにしない
- C1で変更できるのはleaf taskの整数`percentComplete`だけとする
- dependencyはfinish-to-start・lag 0だけをv1正式対応とし、それ以外を暗黙変換しない
- dependencyはsemantic tupleの集合として重複をinvalidにし、task以外のcollection順序は意味を持たせない
- resource、assignment、calendarは代表fixtureに存在する意味を観測・保持するが、編集面には含めない
- unknown、actual、EV、baseline、timephased、extended dataのopaque preservationはv1で約束しない。R1では存在を報告し、C1で保持を保証できない場合は成功成果物を公開しない
- semantic contractでは意味と境界を決め、外部形式、IR、JSON schema、diagnostic code、serialization normalizationはG2/G3へ残す

### 成果物

- `docs/miku-project-semantic-contract-v1.md`
- `docs/miku-project-semantic-fixture-catalog-v1.md`
- v1 domain scope table
- G1 review checklist

### Gate G1

次をすべて満たし、明示的に承認されれば`G1`通過とする。

- すべてのv1 fieldが `required / optional-preserved / unsupported` のいずれかに分類されている
- すべての不変条件とC1 operationがfixture IDへ対応づけられている
- ordered forest、identityの安定期間、日時・duration・進捗、dependency、参照整合性のvalid/invalid境界が曖昧でない
- unsupported dataを暗黙に破棄、変換、成功扱いしない
- 代表fixtureのC1で、対象taskの進捗以外に保持すべき意味が列挙されている
- 外部形式、IR、schema、diagnosticsなどG2/G3の決定をG1へ混入させていない
- 現行 `ProjectModel` の形を理由にscopeを決めていない

結果: 2026-08-10に通過。semantic contractとsemantic fixture catalogをv1の意味契約として承認した。

## ZB-P2 形式・損失・中間表現・変更契約

依存: `G1`  
主担当: Product Contract
状態: 完了（`Gate G2`通過: 2026-08-10）

### 作業

- [x] `ZB-P2.1` 原入力、中間表現、Projection、変更要求、派生成果物、外部形式の役割を確定する
- [x] `ZB-P2.2` miku-project固有の中間表現が必要か判断する
- [x] `ZB-P2.3` 中間表現を採用する場合、internal、exchange、persistentのどの役割かとschema versionを定義する
- [x] `ZB-P2.4` v1形式ごとに `read / write / roundtrip / loss / unsupported` matrixを作る
- [x] `ZB-P2.5` preservationを `required / normalized / lossy-with-warning / unsupported-error / opaque-preserved` に分類する
- [x] `ZB-P2.6` AI向けProjectionのpurpose、範囲、情報量、規則を定義する
- [x] `ZB-P2.7` whole-state replacementとoperation-based changeの境界を決める
- [x] `ZB-P2.8` 許可operation、selector、logical publication、precondition、validation、diff、output preflight、apply後検証を定義する
- [x] `ZB-P2.9` loss、normalization、ignored change、unsupported dataのprovenance表現を定義する
- [x] `ZB-P2.10` 各operationについて入力artifact、出力artifact、役割、寿命、schema versionを対応づけ、会話履歴やAgent固有のhidden stateへ依存しないことを確認する

### 成果物

- [format and loss contract v1](miku-project-format-and-loss-contract-v1.md)（artifact role、internal IR、format / field / loss matrix、Projection例）
- [change contract v1](miku-project-change-contract-v1.md)（request/diff/output plan/approval例、human gate、新規directoryのexclusive reservation、commit markerによる論理publish、provenance）
- internal semantic state、Projection、request/diff/output plan/approval/provenanceのschemaとvalid/invalid examples
- format/field/loss matrix

### Gate G2

v1 scenarioの全artifactに役割、寿命、schema versionがあり、operation間の受け渡しが明示され、各変換に損失規則があり、各変更に検証とlogical publication規則があること。CLI操作が会話履歴やAgent固有のhidden stateへ依存しないこと。`miku-project-ms-project-xml-subset/v1`、現行 `ProjectModel`、workbook JSON、新しいIRのいずれも、検討なしに正本へ置かれていないこと。internal IR、Projection、request、diff、output plan、approval、provenanceのlogical schemaとvalid/invalid例、XML subsetのnamespace / field / lexical / canonical child順 / 非目標、新規directoryを単位とするartifact setの公開規則が、この判断を実装者とAgentに再現可能な形で示していること。directoryと空の`COMMITTED`を排他的に新規作成し、marker、member、schema、digestの検証を通るsetだけをcommittedとすること。incomplete/corrupt setを利用せず、既存pathを置換しないこと。

結果: 2026-08-10に通過。format and loss contractとchange contractをv1のartifact / format / change契約として承認した。

## ZB-P3 CLI・diagnostics・共通conformance契約

依存: `G2`  
主担当: Product Contract
状態: 完了（`Gate G3`通過: 2026-08-11）

### 作業

- [x] `ZB-P3.1` v1 CLIの最小command matrixを確定する
- [x] `ZB-P3.2` args、stdin request、stdout、stderr、file output、binary、Base64、encoding、BOMの規則を定義する
- [x] `ZB-P3.3` hidden stateなし、C1では新規artifact set directoryだけを許可し、exclusive directory create、commit marker、`incomplete / committed / corrupt`、cleanup diagnosticsを定義する
- [x] `ZB-P3.4` versioned result/diagnostics schemaを定義する
- [x] `ZB-P3.5` code、severity、scope、path、status、I/O metadata、loss、normalization、retryabilityを定義する
- [x] `ZB-P3.6` exit codeとsuccess、validation-failed、invalid-usage、internal-errorの境界を定義する
- [x] `ZB-P3.7` Nodeをv1の参照実装とし、Javaを固定済み共通契約へ適合するruntimeとする（2026-08-10決定）
- [x] `ZB-P3.8` 共通core profile、静的capabilityと動的preflight、runtime固有extensionの許可範囲と表現を定義する。v1のNode/Java extension setは空とする
- [x] `ZB-P3.9` Node/Java/Skillsが共有するchecked-in conformance fixtures、golden、suite case、比較modeを設計する
- [x] `ZB-P3.10` product contract version、runtime version、fixture suite version、asset、source、SHA-256、capability setを持つruntime manifestを定義する
- [x] `ZB-P3.11` command候補を責務と意味上の副作用で `read-only / artifact生成 / 意味変更` に分類する
- [x] `ZB-P3.12` 非対話実行を基本とし、意味変更の明示許可、human gateの挿入点、再試行・中止・次操作の機械判定条件を定義する

### Node参照実装の決定

v1はNode CLIを最初の実行可能な参照実装とする。`G3`で製品契約を承認し、Nodeが`G4`でその契約を実証してcontract releaseを固定した後、Javaはその固定releaseと同じ共通conformance corpusへ適合する。したがって、Java実装の完了は`G4`を止めず、Nodeの実装変更を追い続けるmoving portにも戻さない。

参照実装という語は、Nodeの観察された挙動を仕様の正本にする意味ではない。正本は承認済みの製品・semantic・format・change・CLI契約、JSON Schema、および`ZB-P3.9`で固定する共通fixture / goldenである。Nodeがこれらと矛盾する場合はNodeの不具合、または明示的な契約改訂が必要な不整合として扱い、Nodeの挙動だけで契約を暗黙変更しない。Javaもlive Node出力だけをoracleにせず、同じ契約、Schema、fixture / goldenで検証する。

共通v1 capabilityのcommand semantics、result、diagnostics、exit状態、publication結果はNode/Javaで意味的に適合させる。[runtime capability contract v1](miku-project-runtime-capability-contract-v1.md) は九件の閉じた`miku-project-cli-core/v1`を定義し、Node/Javaとも部分実装をv1適合runtimeとして公開しない。共有v1 CLI上のruntime固有extensionは空集合とし、静的なruntime capabilityとdestination固有preflightを分離する。

### 成果物

- `docs/miku-project-cli-contract-v1.md`
- [CLI result and diagnostics contract v1](miku-project-cli-result-contract-v1.md)
- `docs/schemas/miku-project-cli-result-v1.schema.json`
- `docs/schemas/miku-project-cli-diagnostic-v1.schema.json`
- `docs/schemas/miku-project-artifacts-v1.schema.json`と`docs/examples/artifacts-v1/`（semantic / exchange artifactのclosed shapeとvalid example）
- [runtime capability contract v1](miku-project-runtime-capability-contract-v1.md)（core profile、command要求、cross-runtime target matrix、extension境界）
- [conformance corpus v1](miku-project-conformance-corpus-v1.md)（authority、canonical digest、比較mode、unknown outcome recovery）
- [runtime manifest contract v1](miku-project-runtime-manifest-contract-v1.md)と`docs/schemas/miku-project-runtime-manifest-v1.schema.json`
- [human gate and next action contract v1](miku-project-human-gate-and-next-action-contract-v1.md)
- `testdata/conformance/v1/` の21 workflow / harness case、31 schema / binding adversarial case、seed fixtures、semantic golden。command別I/O/effect、artifact path、Projection scope/content/source-state bindingを含む

### Gate G3

CLI examplesとProjection/request/diff/plan/approval/provenanceの全体がschemaで検証でき、command別の入力role / option / source / stdin / destination / side effectと、別artifact間のcanonical digest、runtime、artifact path / publication stateが固定されていること。Projectionのpurpose別scope、source-state digest、contentが生成元semantic stateへ共通ruleで束縛され、schema-validな虚偽Projectionを拒否できること。exit状態が一意で、unsafe overwriteや失敗がincomplete artifactを成功として返さず、cleanupできない場合もpathと状態をdiagnosticsで識別でき、Node/Javaの適合対象と許可差分が明示されていること。runtime manifestが固定asset/source、外部pin、capability、fixture suiteを束縛し、manifest不正・asset改変・capability不足・source欠落、楽観的なretryability、human gate短絡、expected plan不一致のsuccess化、unknown outcomeの再applyをschemaまたはconformance testで拒否できること。人、shell script、CI、AI Agentが人間向けmessageの文字列解析なしに、操作の副作用、成功・失敗、再試行可能性、安全な次の行動を判断できること。このgate通過後に製品実装を開始する。

結果: 2026-08-11に通過。CLI、result/diagnostics、runtime capability/manifest、conformance corpus、human gate / next action契約をv1正本として承認し、Node参照実装の`ZB-P4`を着手可能にした。

## ZB-P4 Node CLI vertical sliceとv1完成

依存: `G3`  
主担当: Node CLI
状態: 進行中（`ZB-P4.1`完了、次は`ZB-P4.2`）

### 作業

- [x] `ZB-P4.1` 現行CLIの互換動作をcontract testsで固定する
  - `tests/mikuproject-cli-compatibility-contract.test.js`にlegacy command surface、help/version、AI spec、stdin/stdout/stderr、named file outputと上書き、JSON usage diagnostics、draft→workbook変換を固定する6 caseを置く
  - legacy diagnosticsの`--out` I/O metadata欠落は観測済み互換挙動として固定し、新v1 result envelopeへ継承しない
- [ ] `ZB-P4.2` 現行test suiteの `fast / full / all` を実態に合わせ、一部testの実行漏れを解消する
- [ ] `ZB-P4.3` semantic変更より先に、CLIのparser、command service、I/O、diagnostics、formattingを分離する
- [ ] `ZB-P4.4` G0で選んだscenario一つだけを新契約でend-to-end実装する
- [ ] `ZB-P4.5` whole-project inspect/validate、semantic diff、pre/post apply validationを選択scopeに応じて実装する
- [ ] `ZB-P4.6` exclusive output directory、commit marker、incomplete/corrupt判定、cleanup diagnostics、structured loss reportingを実装する
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
- invalid input、digest mismatch、unsafe overwriteでcommitted artifactを生成しない。cleanup不能時のincomplete directoryはmarkerなしで識別され、diagnosticsにpathと状態が出る
- repository外でsingle `.mjs` が動作する
- relevant tests、contract suite、`npm run build:full` が成功する
- Node contract releaseとartifact manifestが固定される

## ZB-P5 Java CLI適合

依存: `G4`  
主担当: Java CLI

### 作業

- [ ] `ZB-P5.1` moving Node subtree追随を止め、固定されたcontract release単位の入力へ切り替える
- [ ] `ZB-P5.2` 共通fixture、golden、result/diagnostics schemaをJava testから利用する
- [ ] `ZB-P5.3` required command、option、I/O、exit状態、exclusive createとcommit markerによるlogical publicationを実装する
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
- [ ] `ZB-P6.2` `validate → inspect → Agentまたは人の判断 → 変更要求 → plan-change → human gate / approval → apply-change → verify-artifact` の安全な順序を定義する
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
| Node CLI | contract tests、determinism、commit-marker publication、incomplete/corrupt判定、clean bundle smoke、`npm run build:full` |
| Java CLI | 共通conformance、semantic roundtrip、`mvn test`、fat JAR smoke |
| Agent Skills | structure、manifest/digest、fallback、human gate、isolated bundle smoke |
| AI Agentフレンドリー | 副作用分類、schema validation、message非依存の分岐、human gate、失敗時の安全な次行動 |
| Release | Node/Java/Skills compatibility、asset provenance、clean install、migration smoke |

## 直近の実行キュー

`ZB-P1` / `Gate G1`、`ZB-P2` / `Gate G2`、`ZB-P3` / `Gate G3`は完了した。現在は`ZB-P4`（Node CLI vertical sliceとv1完成）が最優先であり、承認済み契約をNode参照実装として実証する。

P4へ渡す承認済みの前提は次の五つである。

1. `miku_project_semantic_state/v1`をNode/Java共通のinternal-only IRにし、`miku-project-ms-project-xml-subset/v1`を唯一の外部read/write profileに限定する判断
2. XML subsetのnamespace、field/lexical/hierarchy mapping、canonical child順、特定世代XSD完全準拠とapplication互換性を未実証の非目標とする境界を固定し、`required / normalized / unsupported-error`をv1成功経路としてlossy/opaque復元を約束しない判断
3. AI Agentには全stateではなく、purpose別のread-only Projectionだけを渡す判断
4. C1をleaf task一件の`set_task_percent_complete`だけに限定し、whole-state replacement/mergeを拒否する判断
5. current state、request、dry-run diff、output plan、approvalをdigestで束縛し、directoryを排他的に予約して`project.xml + provenance.json`を検証後、空の`COMMITTED`を排他的に作成して論理publishする判断。incomplete/corrupt setを利用せず、既存pathを置換しない

`ZB-P3.1`〜`ZB-P3.12`は[CLI contract v1](miku-project-cli-contract-v1.md)、[CLI result and diagnostics contract v1](miku-project-cli-result-contract-v1.md)、[runtime capability contract v1](miku-project-runtime-capability-contract-v1.md)、[runtime manifest contract v1](miku-project-runtime-manifest-contract-v1.md)、[conformance corpus v1](miku-project-conformance-corpus-v1.md)、[human gate and next action contract v1](miku-project-human-gate-and-next-action-contract-v1.md)として2026-08-11に承認された。次の作業は`ZB-P4.1`で現行CLI互換動作をcontract testへ固定し、続いて`ZB-P4.2`でtest topologyの実行漏れを解消することである。semantic変更や新command実装は、この二つのbaseline作業の後に進める。

## 計画の更新方法

- 詳細な依存、成果物、完了条件はこの文書で更新する
- 現在着手可能な項目だけを `docs/TODO.md` に置く
- 判断結果は該当するcontract文書へ記録する
- 完了した移行と検証結果は `docs/migration-worklog.md` へ記録する
- scopeを変える場合は、新仕様、計画、TODOを同じ変更で整合させる
- Web/MCPを前倒しする場合は、新仕様の初期対象範囲を明示的に改訂する
