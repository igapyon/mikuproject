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
updated: 2026-08-16
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
- `testdata/conformance/v1/` のP3完了時の21 workflow / harness case、31 schema / binding adversarial case、seed fixtures、semantic golden。command別I/O/effect、artifact path、Projection scope/content/source-state bindingを含む。現行case数はconformance corpus v1を正本とする

### Gate G3

CLI examplesとProjection/request/diff/plan/approval/provenanceの全体がschemaで検証でき、command別の入力role / option / source / stdin / destination / side effectと、別artifact間のcanonical digest、runtime、artifact path / publication stateが固定されていること。Projectionのpurpose別scope、source-state digest、contentが生成元semantic stateへ共通ruleで束縛され、schema-validな虚偽Projectionを拒否できること。exit状態が一意で、unsafe overwriteや失敗がincomplete artifactを成功として返さず、cleanupできない場合もpathと状態をdiagnosticsで識別でき、Node/Javaの適合対象と許可差分が明示されていること。runtime manifestが固定asset/source、外部pin、capability、fixture suiteを束縛し、manifest不正・asset改変・capability不足・source欠落、楽観的なretryability、human gate短絡、expected plan不一致のsuccess化、unknown outcomeの再applyをschemaまたはconformance testで拒否できること。人、shell script、CI、AI Agentが人間向けmessageの文字列解析なしに、操作の副作用、成功・失敗、再試行可能性、安全な次の行動を判断できること。このgate通過後に製品実装を開始する。

結果: 2026-08-11に通過。CLI、result/diagnostics、runtime capability/manifest、conformance corpus、human gate / next action契約をv1正本として承認し、Node参照実装の`ZB-P4`を着手可能にした。

## ZB-P4 Node CLI vertical sliceとv1完成

依存: `G3`  
主担当: Node CLI
状態: `ZB-P4.8`、`ZB-P4.9`承認済み（2026-08-13）。`ZB-P4.10`とGate G4は2026-08-14承認済み。次は固定済みNode contract releaseを入力として`ZB-P5` Java CLI適合へ進む

### 作業

- [x] `ZB-P4.1` 現行CLIの互換動作をcontract testsで固定する
  - `tests/mikuproject-cli-compatibility-contract.test.js`にlegacy command surface、help/version、AI spec、stdin/stdout/stderr、named file outputと上書き、JSON usage diagnostics、draft→workbook変換を固定する6 caseを置く
  - legacy diagnosticsの`--out` I/O metadata欠落は観測済み互換挙動として固定し、新v1 result envelopeへ継承しない
- [x] `ZB-P4.2` 現行test suiteの `fast / full / all` を実態に合わせ、一部testの実行漏れを解消する
  - `scripts/lib/test-suite-topology.mjs`を唯一のsuite inventoryとし、core API / loader testを編入する
  - `fast`は日常回帰、`full`はCLI統合testとbrowser runtime contractを含む完全回帰、`all`は全checked-in test fileを実行する安定aliasとする
  - `tests/mikuproject-test-suite-topology.test.js`で、全`*.test.js`の分類と重複なしを検証する
- [x] `ZB-P4.3` semantic変更より先に、CLIのparser、command service、I/O、diagnostics、formattingを分離する
  - 実行計画は下記「P4.3の実行計画」を正本とする。`ZB-P4.3.1`から順に、一変更単位ずつ実施する
- [x] `ZB-P4.4` R1の外部XML `validate → inspect(project_overview)` を最初の新契約vertical sliceとして実装する
  - 実行計画は下記「P4.4の実行計画」を正本とする。部分実装を完成済みcore runtimeとして公開せず、R1処理系と契約testを先に完成させる
- [x] `ZB-P4.5` C1のhuman gate直前までを実装する
  - `task_change_context`、strict change request、dry-run、pre/post semantic validation、semantic diff、XML encode/redecode preflight、read-only destination preflightを同じplanning operationへ接続する
  - 実際のapply、destination directoryの作成、`project.xml` / `provenance.json` / `COMMITTED`の公開は`ZB-P4.6`だけが担当する
- [x] `ZB-P4.6` approvalで束縛したC1 apply、exclusive artifact publication、read-only verification、committed artifact set入力を実装する
  - 実行計画は下記「P4.6の実行計画」を正本とする。read-only verifierをpublisherより先に完成させ、marker前後の状態判定に同じ検証primitiveを使う
- [x] `ZB-P4.7` 既存coreの再利用部分を新conformance fixturesで検証する
  - 実行計画は下記「P4.7の実行計画」を正本とする。旧coreをv1へ取り込むこと自体を目的にせず、v1のsemantic / artifact / CLI契約を満たす責務だけを採用する
- [x] `ZB-P4.8` 階層C1 sliceを追加し、v1 capability matrixを更新する
  - 実行計画は下記「P4.8の実行計画」を正本とする。承認済み`S-V002`とその階層不変条件だけをmaterializeし、新format、new CLI command、許可change operationは増やさない
- [x] `ZB-P4.9` versioned single `.mjs`、sources、runtime manifest、SHA-256を生成する（2026-08-13承認）
- [x] `ZB-P4.10` repository外のclean temporary directoryで、外側manifest pinを持つ配布runtime smokeを実行する（2026-08-14承認。第3回reviewのCLI result runtime binding完全一致補正、補正後full regression、`git diff --check`を確認）

### 実装上の境界

- 現行commandは互換性方針が決まるまでcompatibility surfaceとして残す
- browser runtimeは既存 `miku-project-web` の互換性保守に必要な間は維持するが、新CLI契約を支配させない
- reportの見た目改善や未選択formatを、vertical sliceへ混ぜない
- CLI責務分割とsemantic変更を同じ変更単位で行わない

### P4.3の実行計画

目的は、legacy CLIの外形を変えずに `scripts/miku-project-cli.mjs` をentrypoint/wiringへ縮小し、承認済みv1 commandを載せるための境界を作ることである。この工程で新v1 command、semantic scope、exit code、safe publication、legacy output上書き、diagnostics schemaを変更しない。特に、P4.1で固定したlegacy diagnosticsの`--out` I/O metadata欠落は観測済み互換挙動として保つ。

共通の作業規則:

- 各小項目の開始前に `tests/mikuproject-cli-compatibility-contract.test.js` と既存CLI testの対象境界を確認する
- 各小項目では移動・依存注入・import更新だけを行い、command名、option、標準入出力、出力byte列、diagnostics JSON、終了状態、上書き動作を変えない
- 各小項目ごとに `npm run test:fast`、対象のdirect test、`git diff --check` を実行する。command serviceまたはbundleに触れた区切りでは `npm run test:full` と `npm run build:cli-bundle` も実行する
- public entrypoint以外は `process.argv` を読まず、command serviceにはparse済み`command`、`options`、core API、I/O/diagnostics依存を明示して渡す。module間で同じerror classを二重定義しない
- module追加時はsingle `.mjs` bundleがsource treeに依存しないことを同じ変更で維持する。相対importを残したままbundleの動作を偶然に任せない

#### `ZB-P4.3.1` 入口とerror / argv境界を固定する

- `CliUsageError`、`CliProcessingError`、error details/codeの共通処理を一箇所へ移す。既存のerror codeとdiagnostics JSONはそのまま返す
- argv parse、`--help` / `--version`判定、requested diagnostics format、error時のcommand summaryを独立moduleへ移す
- entrypointだけが`process.argv`とtop-level `main().catch(...)`を持つ構成にする。optionの重複時の最後の値採用など、現行parse挙動を変更しない
- 正常系・usage errorのCLI compatibility testを通し、抽出moduleのdirect testではparse結果、missing option、`--diagnostics json`検出、command summaryを固定する

結果: 2026-08-12に完了。`scripts/lib/cli-errors.mjs`と`cli-argv.mjs`へ移し、`tests/mikuproject-cli-argv.test.js`を追加した。single-MJS bundleは内部moduleを明示順で内包し、source CLIとbundle CLIの`--version`、stdin `ai detect-kind --diagnostics json`を確認した。

#### `ZB-P4.3.2` text / binary I/Oと出力先記述を抽出する

- UTF-8 stdin/file読取り、binary/Base64読取り、stdin source重複検査、binary input/outputの制約、primary output書込みをI/O moduleへ移す
- diagnostics用のinput/output記述も同じI/O境界に置く。ただしlegacy commandが渡す`output: null`はstdoutとして記録される現行挙動を変えない
- `--out`の既存ファイル上書き、`--out-base64 -`、binary stdout拒否、stdioの空入力エラーを既存のcode/message/exit状態のまま維持する
- compatibility testのstdin/stdout/stderr・named outputを通し、既存CLI testのbinary/Base64ケースを対象確認する

結果: 2026-08-12に完了。`scripts/lib/cli-io.mjs`へ移し、`tests/mikuproject-cli-io.test.js`でBase64、binary target、stdin conflict、diagnostics I/O記述を固定した。legacy CLI target test、bundle、fast suiteを確認した。

#### `ZB-P4.3.3` diagnostics / formatting境界を抽出する

- diagnostics version、status/exit判定、warning/change summary、structured error、text/JSON formatter、help textをpresentation/diagnostics moduleへ移す
- `ai validate-patch`の成功・validation failureを含め、text diagnosticsとJSON diagnosticsの既存整合を保持する。message文字列からcodeを推測するlegacy処理は、v1への採用を意味しないまま移動だけする
- `--version`のpackage/bundle metadata読取りはentrypoint側に残すか、明示したmetadata providerへ渡す。bundle外で`package.json`を読めない場合の既存fallbackを保持する
- existing CLI testのdiagnostics、help/version、validation failureを対象確認する

結果: 2026-08-12に完了。`scripts/lib/cli-diagnostics.mjs`と`cli-presentation.mjs`へ移し、`tests/mikuproject-cli-diagnostics.test.js`でoption/status、structured error、validation formatterを固定した。`ai validate-patch`のvalidation failureを含むlegacy CLI target testを確認した。

#### `ZB-P4.3.4` legacy command serviceをfamilyごとに移す

- routerはcommand tokensの照合とunsupported command errorだけを担い、実処理を `ai`、`state`、`import/export`、`report` のfamily serviceへ一つずつ移す
- serviceはcore APIを引数で受け、filesystem・process・global stateを直接所有しない。workbook loading、AI view選択、state summary/diff、report生成などの既存helperは最小の所有moduleへ移す
- familyを移すたびに、P4.1 compatibility testと該当する既存CLI testを通す。処理順、JSON pretty-print、warnings/changesの順序を変えない
- 完了時にentrypointからlegacy command固有の`scope/action`分岐とdomain helperを除き、parse → control operation → core API lifecycle → router → diagnostics/output → disposeだけにする

結果: 2026-08-12に完了。`cli-ai-commands.mjs`、`cli-state-commands.mjs`、`cli-exchange-commands.mjs`、`cli-report-commands.mjs`へfamilyを移し、`cli-legacy-router.mjs`が照合・unsupported commandだけを担う。entrypointは73行のparse → control operation → core API lifecycle → router → diagnostics/output → disposeに縮小した。legacy compatibility contract 7 caseと既存CLI 57 caseを確認した。

#### `ZB-P4.3.5` single-MJS bundleをmodule graph対応へ更新する

- `scripts/build-cli-bundle.mjs`にCLI内部moduleの明示的・決定的なdependency orderを持たせ、bundleへ必要なsourceをentrypointより先に内包する
- source moduleのrelative importをbundle内へ残さない。bundle用にstripするimportとNode built-in依存を明示し、module名・初期化順・`BUNDLED_PACKAGE_VERSION`の可視性を検証する
- source archiveには新規CLI moduleとtestを自動的に含め、bundle生成結果が不必要なsource tree/pathへ依存しないことを維持する
- `npm run build:cli-bundle`、既存のrepository外single-MJS test、`--help`、`--version`、代表的なstdin commandをsource CLIとbundle CLIの双方で通す

結果: 2026-08-12に完了。`CLI_INTERNAL_MODULE_RELATIVE_PATHS`でinternal moduleの依存順を明示し、module syntaxを除去してentrypointより先にsingle-MJSへ内包する。既存のrepository外bundle test、bundleの`--help` / `--version` / stdin `ai detect-kind --diagnostics json`、source archive内のnew module/testを確認した。

#### `ZB-P4.3` 完了判定

- entrypoint、argv/error、I/O、diagnostics/formatting、legacy command family、bundleの責務と依存方向が上記どおりに分離されている
- P4.1 compatibility contract、既存CLI test、`npm run test:full`、`npm run build:full`、repository外bundle smokeが成功する
- `scripts/miku-project-cli.mjs`の外形互換を保ち、v1 commandやsafe publicationを先取りしていない
- 次の`ZB-P4.4`が、legacy routerとは別のv1 command serviceを追加する変更として開始できる

結果: 2026-08-12に通過。`npm run test:full`（17 files / 217 tests）と`npm run build:full`、P4.1 legacy compatibility contract、repository外single-MJS bundle smokeが成功した。legacy command、option、standard I/O、legacy overwrite、diagnostics envelope、exit behaviorは維持し、v1 command / semantic / safe publicationは実装していない。

### P4.4の実行計画

最初のvertical sliceは、G0で承認したR1のうち、`miku-project-ms-project-xml-subset/v1`のexternal XMLを`validate`し、validな同じ入力から`inspect --purpose project_overview`を得る経路とする。代表fixtureは`testdata/conformance/v1/fixtures/project/dependency-canonical.xml`であり、現行`ProjectModel`やlegacy `project_overview_view`を新契約の正本にはしない。

このsliceに含めるのはexternal XMLのfile/stdin入力、stdoutまたは新規result fileへのresult envelope、format/semantic validation、`project_overview` Projectionである。committed artifact set入力、`task_change_context`、change request、diff、approval、publication、`verify-artifact`は後続へ送る。legacy `ai / state / import / export / report`のcommand、option、I/O、上書き、diagnosticsは変更しない。

#### runtime適合を先取りしない境界

v1の`succeeded` / domain `rejected` resultは、全九capabilityを持つ`miku-project-cli-core/v1`のmanifestと実行assetを検証した`runtime.binding_status = verified`へ束縛しなければならない。一方、R1だけを実装した時点ではcore profile全体を実装済みとはいえない。したがってP4.4では次を守る。

- R1 command serviceはruntime bindingを明示依存として受け取り、固定したtest bindingでschema、semantic、I/O、determinismを検証する。このtest bindingはresult組立てのtest inputであり、release適合の証拠には数えない
- R1の部分実装だけを指すrelease `runtime-manifest.json`を生成せず、`miku-project-cli-core/v1`適合、Gate G4通過、Node reference releaseを名乗らない
- public bundleから五workflow commandを適合runtimeとして有効化するのは、P4.5〜P4.8で九capabilityが揃い、P4.9で実asset/source/manifest bindingを検証した後とする
- P4.4のconformance記録は`R1 service pass / runtime-integrity pending`と区別し、`CR-*`、release smoke、外部manifest pinをpass扱いしない

これはv1契約を緩和する措置ではなく、不完全なruntimeが契約適合を自己申告しないための実装順序である。

#### `ZB-P4.4.1` schema validatorとcanonical data基盤を作る

- `docs/schemas/`のartifact、result、diagnostic、runtime manifest schemaを一つのregistryとして読めるbuild/test基盤を作る
- JSON Schema validatorはbuild時にstandalone codeへ生成し、製品runtimeへpackage lookupやsource checkoutを要求しない。generator、生成物、drift checkを同じ変更に含め、生成物にrelative/bare importが残らないことを検査する
- canonical JSON serializer、canonical SHA-256、raw byte SHA-256、semantic collection canonicalizationを`conformance corpus v1`どおりに実装する。Unicode normalization、timestamp、host pathをsemantic digestへ混ぜない
- `dependency.state.json`のdigest、schema正例、`contract-cases.json`のうちR1に必要なschema/binding caseでdirect testを作る
- schema生成にbuild-time dependencyを追加する場合は、`package.json` / lockfile / `THIRD-PARTY-NOTICES.md`を同じ変更で更新し、runtime dependencyを増やさない

想定配置: `scripts/lib/v1/cli-v1-canonical-json.mjs`、`scripts/generated/cli-v1-schema-validators.mjs`、`scripts/generate-cli-v1-schema-validators.mjs`、`tests/mikuproject-cli-v1-contract.test.js`。

##### P4.4.1の実行順

`P4.4.1`は次の六単位を順に実施する。この段階ではCLI command surface、XML decode、runtime manifest、release assetを変更しない。

1. `ZB-P4.4.1a` build-time validator toolchainを固定する
   - `ajv`と`ajv-formats`を`devDependencies`へ追加し、JSON Schema 2020-12、strict validation、all-errors、URI format検証を使う。製品の`dependencies`へ追加しない
   - lockfileでgeneratorの依存版を固定し、Ajv本体、format plugin、生成物へ取り込まれるtransitive runtimeのlicenseを`THIRD-PARTY-NOTICES.md`へ記録する
   - `generate:cli-v1-schema-validators`は生成物を更新し、`check:cli-v1-schema-validators`はtemporary outputとのbyte比較だけを行うscriptとして分ける。通常testがtracked fileを書き換えないようにする
2. `ZB-P4.4.1b` 四schemaのregistryとstandalone moduleを生成する
   - artifact、CLI result、CLI diagnostic、runtime manifest schemaを固定順で読み、期待する`$schema`と`$id`、重複ID、未解決external `$ref`を生成前に検査する
   - diagnosticとartifact schemaをregistryへ先に登録し、external URN参照を持つresult schemaを同じregistryでcompileする
   - `validateArtifact`、`validateCliResult`、`validateCliDiagnostic`、`validateRuntimeManifest`の四entryを持つESMを生成する。Ajv standalone outputにruntime helper importが残る場合は既存esbuildで一fileへbundleする
   - checked-in生成物にはbare/relative import、absolute path、timestamp、hostnameを残さず、repository外から単独importできることをdirect testで確認する。生成moduleのsingle-MJS CLIへの内包は`ZB-P4.4.6`で行う
3. `ZB-P4.4.1c` canonical JSON byte serializerを実装する
   - object keyをUnicode code point順で再帰sortし、array順を保持し、insignificant whitespaceと末尾LFを付けないUTF-8 byte列を生成する
   - quotation mark、reverse solidus、U+0000〜U+001Fを契約どおりescapeし、solidusとその他のUnicode scalarを不要にescape/normalizeしない
   - `null`、boolean、string、array、plain objectと、10進表現へ曖昧なく戻せるintegerだけを受理する。Nodeの`Number`ではsafe integerを要求し、safe範囲外のraw JSON integerを丸めてdigest化しない。`undefined`、sparse array、non-finite number、fraction、`-0`、BigInt、non-plain object、unpaired surrogateはdigestを作らずrejectする
   - serializerは入力objectを変更せず、同値なkey挿入順の違いから同じbyte列を返す
4. `ZB-P4.4.1d` semantic state canonicalizationとdigestを実装する
   - task arrayはsemantic preorderを保持し、dependencyは`predecessor_uid / successor_uid / type / lag` tuple、resource / assignment / calendarはUIDで決定的に並べたcopyを作る
   - scalar comparatorはlocale、host、Unicode normalizationへ依存せず、Unicode scalar/code pointで比較する。collection memberや元stateをin-place sortしない
   - canonical JSONへSHA-256を適用して`{ algorithm: "sha-256", value: <64 lowercase hex> }`を返す。raw input byte SHA-256は別functionにし、semantic digestと混同させない
   - checked-in goldenの期待値を固定する。`dependency.state.json = a98f0c8b560382234572a61f360c9e96911bc75fc1d57b79968d6e60b5d751d0`、`dependency-percent-50.state.json = 1c72d70cc114853b2a61f1c4798794093e46419f16b6cc49819a9d050cb67a08`
5. `ZB-P4.4.1e` schema-layer contract harnessを実装する
   - `contract-cases.json`のRFC 6901 pointer mutationをtest-only helperでcopyへ適用し、元exampleとcase indexを変更しない
   - 四schemaの公式正例と、`validation_layer = json-schema`の18件すべてを期待どおりaccept/rejectする。schema errorのAjv文言・順序は製品診断契約に使わない
   - `validation_layer = cross-artifact-binding`の13件はindexへ残し、この段階でpass扱いしない。P4.4.1はcanonical digestのprimitiveだけを提供し、`RB-012`のProjection content bindingは`P4.4.5`、C1 bindingは`P4.5`以後で実装する
6. `ZB-P4.4.1f` repository integrationを完了する
   - direct testを`FAST_SUITE`へ一度だけ登録し、test topologyで未分類・重複がないことを確認する
   - generator check、対象direct test、`npm run test:fast`、`npm run test:full`、`npm run build:full`、`git diff --check`を順に通す
   - source schemaを一byte変更したnegative drift testと、生成物を再生成した後にdriftが解消することをtemporary copyで確認する。tracked schemaや生成物をtest中に変更しない

##### P4.4.1の完了判定

- 四schemaのstandalone validatorがsource checkoutやruntime package lookupなしでimportでき、公式正例と18 schema-layer caseを期待どおり判定する
- canonical serializerが契約のescape、Unicode scalar、integer、object key規則を満たし、不正なJSON-domain valueをfail closedにする
- semantic canonicalizationがtask順を保持し、非task collection順の差を吸収し、二つのgolden digestと一致する
- generator drift、dependency/license、test suite分類がrepository内で検査可能になっている
- public CLI behavior、legacy compatibility、runtime適合宣言は変更されず、13 cross-artifact binding caseは未完として可視のまま残る
- 次の着手点は`ZB-P4.4.2`のv1 strict argv / result transportであり、schema validatorとdigest APIを再実装せず利用できる

結果: 2026-08-12に完了。Ajv 8.17.1 / ajv-formats 3.0.1をbuild-time dependencyとして固定し、四schemaをregistryからcompileしてimport-freeな`cli-v1-schema-validators.mjs`へ決定的にbundleした。canonical JSON、raw / semantic SHA-256、semantic collection canonicalizationを追加し、二つのsemantic golden digestを固定した。direct testは公式正例、18 schema-layer case、temporary rootでのdrift検出・再生成、repository外importを通す。13 cross-artifact binding case、v1 command surface、runtime適合宣言は後続のままとする。`npm run check:cli-v1-schema-validators`、`npm run test:fast`、`npm run test:full`、`npm run build:full`が成功した。

#### `ZB-P4.4.2` v1専用argvとresult transportを作る

- raw argvから五workflow commandだけを識別するv1 dispatcherと、`inspect` / `validate`のstrict parserをlegacy parserとは別moduleに置く
- long optionだけ、必須option、unknown/duplicate option、余分な位置引数、`project_overview`での`--task-uid`拒否、明示stdin一件を、project inputを読む前に検証する
- `--result -`はJSON一件 + LFをstdoutへ、`--result <path>`は既存親directory内の未使用fileをexclusive createして返す。既存pathを上書きせず、stdout/stderrの役割をlegacy I/Oと混ぜない
- result envelope、diagnostic、status `succeeded / rejected / usage-error / runtime-error`、exit `0 / 1 / 2 / 3`、`next_action`をstable codeから決定的に組み立てる
- `process.argv`、stdin/stdout、filesystemをcommand serviceへ直接持たせず、entrypoint / v1 I/O境界から注入する

想定配置: `scripts/lib/v1/cli-v1-argv.mjs`、`cli-v1-io.mjs`、`cli-v1-result.mjs`、`cli-v1-router.mjs`、`tests/mikuproject-cli-v1-argv-io.test.js`。

結果: 2026-08-12に完了。五workflow commandのstrict argv grammar、control operation識別、explicit stdin一件制約、purpose/task UID制約をpure moduleとして実装した。result transportはstdoutまたは既存親directory内の新規fileをexclusive reservationし、既存path・不正path・親directory不在をdomain operation前に拒否する。result / diagnostic builderはgenerated schemaで自己検証し、status/exit、deterministic ordering、retryabilityからの`next_action`を固定した。legacy entrypointとlegacy command surfaceは変更せず、v1 routerの公開接続とsingle-MJS内包は`ZB-P4.4.6`へ残す。`npm run check:cli-v1-schema-validators`、`npm run test:fast`、`npm run test:full`、`npm run build:full`が成功した。

#### `ZB-P4.4.3` XML subset decodeとsemantic validationを実装する

- inputをraw byteで読みSHA-256を記録し、fatal UTF-8 decode、XML BOM normalization、XML declaration、namespace/root、許可element/attribute、singleton/container規則を検査する
- XMLを`miku_project_semantic_state/v1`へdecodeし、pseudo task、ordered forest、required/optional field、FS/lag 0 dependency、resource、assignment、calendarを契約どおり対応づける
- 現行XML DOM/codecのparse処理は再利用候補にできるが、現行validatorのwarning判定やmessage文字列をv1のoracleにしない。v1 validatorはstable diagnostic code、semantic rule ID、locationを直接生成する
- まず`S-V001`、`S-I012`、`S-I020`の三seedを通し、valid stateは`dependency.state.json`とsemantic exact比較する。invalid/unsupportedでは成功state digestやProjectionを返さない
- XML profile scanとsemantic invariant validationを分離し、P4.7で残るcatalog fixtureを追加できる構造にする

想定配置: `scripts/lib/v1/cli-v1-xml-adapter.mjs`、`cli-v1-semantic-validator.mjs`、`tests/mikuproject-cli-v1-xml-adapter.test.js`。

結果: 2026-08-12に完了。raw byte SHA-256、fatal UTF-8 decode、先頭BOM normalization、XML declaration encoding、root namespace、許可element / attribute、singleton / containerを明示的に検査する`ms-project-xml-adapter/v1`を追加した。adapterは外部XMLを`miku_project_semantic_state/v1`へ限定decodeし、project summary pseudo taskを除外してordered forest、FS/lag 0 dependency、resource、assignment、calendarを対応づける。semantic validatorはprofile scanのunsupported findingとstate不変条件を別入力として扱い、stableな`semantic.invalid` / `semantic.unsupported`、G1 rule ID、semantic locationを直接生成する。`S-V001`は`dependency.state.json`とexact一致し、`S-I012`と`S-I020`はそれぞれ正しい分類でrejectされる。v1公開entrypoint、result I/O integration、bundle内包は`ZB-P4.4.4`〜`ZB-P4.4.6`で行う。

#### `ZB-P4.4.4` `validate` resultを完成させる

- `validate --project <file|-> [--result <path|->]`を、strict parse → result reservation → XML decode → semantic validate → structured resultの順で実装する
- `CV-VALID-001`は`validation.valid = true`、format profile、canonical state digest、diagnostics空、`complete`を返す
- `CV-INVALID-001`は`semantic.invalid / S-I012`、`CV-UNSUPPORTED-001`は`semantic.unsupported / S-I020`として`rejected / 1`を返す
- 三caseすべてでinput不変、destinationなし、project artifactなし、I/O source/path/digest、result target、observations、next actionをschema/binding ruleと照合する

想定配置: `scripts/lib/v1/cli-v1-r1-commands.mjs`、`tests/mikuproject-cli-v1-validate.test.js`。

結果: 2026-08-12に完了。`runV1Validate`はP4.4.2のstrict invocationと予約済みresult transportを受け、direct external XML regular fileまたは明示stdinをraw SHA-256付きI/O metadataとして読んでから、P4.4.3のdecode / semantic validationへ接続する。`CV-VALID-001`はcanonical state digestと`complete`を、`CV-INVALID-001`と`CV-UNSUPPORTED-001`はstable diagnostic / rule ID、state digestなしの`rejected` resultを返す。missing、direct symlink、directory、read failureもdomain decode前に分離する。P4.4段階ではfixed test runtime bindingを注入するserviceであり、directory artifact setのcommitted検証は後続artifact verifier workstream、legacy public entrypointとsingle-MJS bundle内包はP4.4.6へ残す。

#### `ZB-P4.4.5` `inspect project_overview`を完成させる

- `inspect`は独自の緩い読取り経路を持たず、`validate`と同じXML decode / semantic validationを内部で必ず通す
- valid semantic stateから、project required/optional field、semantic preorder上の全task、全dependency、unsupported summaryを`miku_project_projection/v1`へ決定的に射影する
- `source_state_digest`、固定scope、included/omitted domain、taskの0始まり`order`を生成元stateへ`RB-012`で束縛する
- legacy `api.aiViews.exportProjectOverviewView()`は用途とschemaが異なるため直接返さない。再利用する場合も、v1 Projection builderの明示mappingを経由する
- `testdata/conformance/v1/golden/projection/dependency.project-overview.json`を追加し、`CI-OVERVIEW-001`をexact JSONで比較する。invalid/unsupported入力ではProjectionを返さない

想定配置: `scripts/lib/v1/cli-v1-projection.mjs`、`tests/mikuproject-cli-v1-inspect.test.js`。

結果: 2026-08-12に完了。`runV1Inspect`と`runV1Validate`は同じexternal-input preparationを通り、external XMLのdecodeとsemantic validationに差を持たない。valid stateだけを`miku_project_projection/v1`へ射影し、projectの許可field、semantic preorderの全taskと0始まり`order`、canonical dependency集合、固定scope、canonical state digestを含める。Projection builderはlegacy AI viewを返さず、`RB-012`用のexact binding checkerと`CI-OVERVIEW-001` goldenでsource digest / scope / contentを固定した。invalidまたはunsupported入力はdiagnosticを返すがProjectionを公開せず、同一runtimeのresult byte determinismもdirect testで確認する。

#### `ZB-P4.4.6` R1 sliceを統合して回帰を固定する

- v1 routerをlegacy routerの前段で識別できるようにするが、release manifest未完成のsource/bundleを適合runtimeとして公開しない境界を保つ
- fixed test bindingを注入したR1 harnessでfile/stdin、stdout/new result file、同一条件二回のbyte determinism、result path既存拒否、usage error時project未読を検証する
- direct testは`fast`、subprocess/I/O integrationは`full`へ分類し、suite topology testへ登録する
- `scripts/build-cli-bundle.mjs`へ新module/生成validatorを決定的な順序で含め、source archiveにもschema、generator、golden、testを含める。ただしrepository外の適合runtime smokeはP4.9/P4.10まで未完と記録する
- P4.1 legacy compatibility contract、既存CLI integration、`npm run test:fast`、`npm run test:full`、`npm run build:full`、`git diff --check`を通す

結果: 2026-08-12に完了。public source CLIとdevelopment bundleはv1 command wordをlegacy routerより先に認識する。ただしmanifestと九capabilityが未完成なので、project inputをreadせず、`--result` pathも予約せず、`command = cli`のunverified `runtime.capability-missing` resultでfail-closedにする。R1の成功経路は公開runtimeではなく、fixed verified test bindingを明示注入するsubprocess harnessに隔離した。このharnessでexternal XML file/stdin、stdout/new result file、existing result file拒否、usage error時input未読、同一runtime byte determinismを確認した。single-MJS bundleにはgenerated validatorを含むv1 module graphをlegacy helper名と衝突しないclosureへ決定的に内包し、source archiveにschema、generator、golden、test/harnessが入ることも確認した。R1 service passとruntime-integrity pendingを区別し、manifest、実asset/source binding、repository外適合runtime smokeはP4.9/P4.10まで未完とする。

#### `ZB-P4.4` 完了判定

- `CV-VALID-001`、`CV-INVALID-001`、`CV-UNSUPPORTED-001`、`CI-OVERVIEW-001`がR1 service harnessでschema、semantic/binding、I/O、determinismの期待を満たす
- external XMLのfile/stdinから`validate`と`project_overview`までが同じdecode/validation pipelineで完結し、入力と既存pathを変更しない
- invalid/unsupported/usage errorで成功payloadを返さず、stable diagnosticと安全な`next_action`を返す
- legacy CLIの外形互換と既存testが維持される
- partial runtime manifest、適合宣言、G4通過、`task_change_context`、C1、artifact publicationを先取りしていない
- 完了後の次作業は`ZB-P4.5`の`task_change_context`とC1 semantic planningである

結果: 2026-08-12に完了。R1 service harnessは`CV-VALID-001`、`CV-INVALID-001`、`CV-UNSUPPORTED-001`、`CI-OVERVIEW-001`を、schema・semantic / RB-012 binding・I/O・determinismの範囲で通過した。legacy surfaceとbundle/source archiveの回帰も通している。一方で、P4.4はNode reference runtime releaseやG4の根拠を作っておらず、public source/bundleがpartial core profileを有効化しないことをテストした。次は`ZB-P4.5`で`task_change_context`とC1 semantic planningを実装する。

### P4.5の実行計画

P4.5はC1のhuman gate直前、すなわち`inspect --purpose task_change_context`から`plan-change`の`semantic_diff + output_plan`までを閉じる。計画成功は新しいproject artifactを意味しない。実際のapply、排他的directory作成、provenance、commit marker、artifact set verificationはP4.6へ残す。

1. `task_change_context`をvalid semantic stateから決定的に生成する。対象leaf task、rootからのancestor chain、対象に接続するdependency、対象assignmentと対応resourceだけを含め、scope/source digest/contentをexact builderで`RB-012`へ束縛する。不存在targetは`change.request-invalid`、summary targetは`change.operation-unsupported`で成功Projectionを返さない。
2. `--request`をregular fileまたは明示stdinからread-onlyで読み、UTF-8、BOMなし、JSON document一件、duplicate keyなし、artifact kind/version/schemaをfail-closedに検証する。request source、canonical path、raw digestを`plan-change`の二番目のI/O inputへ記録する。
3. C1 allowlistの`set_task_percent_complete`だけをdry-runする。base state digest、target leaf、expected current value、新値が異なることを検証し、前後semantic validationを必須にする。planned whole stateはinternal-onlyとし、targetの`percent_complete`以外に差がないことをcanonical semantic stateで検証する。
4. semantic diffをbase/proposed/request digest、一件のbefore/after、empty loss/normalization/unsupported、preservation assertionとして生成する。planned stateをcanonical XML subsetへencodeし、再decode + semantic equivalenceを通したbyte digestだけをoutput planのpreflightへ載せる。
5. destinationは作成せず、existing/symlink/non-directory parent/解決不能なpathをrejectし、existing parentのreal path + unused basenameをcanonical absolute pathとしてoutput planへ記録する。runtime binding、diff/request/state digest、I/O destinationの`RB-001`〜`RB-005`をplanning result生成時に検査する。
6. fixed verified test bindingのC1 harnessで`CI-CONTEXT-001`、`CP-CHANGE-001`、stale precondition、duplicate request key、existing destination、result file、stdin/file transport、same-runtime determinismを検証する。public source/bundleはmanifestと九capabilityがそろうP4.9までfail-closedのままとする。

#### `ZB-P4.5` 完了判定

- `task_change_context`がvalid leafだけから生成され、`CI-CONTEXT-001` goldenと`RB-012` exact bindingを通す
- `CP-CHANGE-001`がsemantic diff、output plan、runtime/input/destination bindingをそろえたschema-valid resultとして成功し、planned semantic stateとpreflight XML再decode stateが`dependency-percent-50.state.json`と一致する
- stale base/current value、no-op、allowlist外/不正request、BOM/duplicate JSON、unsafe/existing destinationがsuccess diff/output planを返さない
- `plan-change`がdestinationもproject artifactも作らず、入力を変更しない。actual apply/publicationをP4.5の成功と数えない
- direct test、fixed-binding subprocess integration、legacy compatibility、`npm run test:fast`、`npm run test:full`、`npm run build:full`、`git diff --check`を通す

結果: 2026-08-12に完了。`task_change_context`をscope/content exact builderとして追加し、C1のchange requestをstrict JSON readerで読んで、base/precondition/leaf/no-opを検証するdry-runへ接続した。planned stateは公開せず、pre/post semantic validation、semantic diff、canonical XML encode/redecodeとraw byte digest、read-only destination preflightから`plan-change`の`semantic_diff + output_plan`だけを返す。`RB-001`〜`RB-005`、`CI-CONTEXT-001`、`CP-CHANGE-001`、stale/duplicate/existing-destination reject、file/stdin/result transportをfixed test bindingで検証した。directory作成、apply、provenance、`COMMITTED`、artifact verification、runtime release manifestは未実装であり、次作業`ZB-P4.6`へ残す。

### P4.6の実行計画

P4.6は、P4.5で承認待ちにしたC1 planを、明示的なapprovalと再計算結果へ束縛して新しいproject artifact setとして論理publishする。成功境界は`project.xml`や`provenance.json`のwrite完了ではなく、空の`COMMITTED`を排他的に作成し、その後のread-only再検証まで通過した時点である。既存pathの上書き、incomplete/corrupt setの再利用・修復、CLIによるapproval生成、`--force`、`fsync`保証はこの工程へ入れない。

実装は次の依存順で進める。各小項目のproduct serviceはfixed verified test bindingを明示注入して検証し、P4.9のruntime manifest完成まではpublic source CLI / development bundleを引き続き`runtime.capability-missing`でfail closedに保つ。

#### `ZB-P4.6.1` apply入力とapproval bindingを先に閉じる

- `--plan-result`は成功した`plan-change` result envelope全体、`--approval`は`miku_project_change_approval/v1`として、既存strict JSON readerのUTF-8 / BOM / duplicate key / regular-file / explicit-stdin規則で読む。schema、contract version、runtime binding、command/status、diff/output planの存在を検証し、入力role / option / source / path / raw digestを`RB-010`の順序で記録する
- current projectとrequestを再読取りし、P4.5のdry-run、pre/post semantic validation、semantic diff、canonical XML encode/redecode、destination parent/capability preflightを再実行する。承認時のbase/request/diff/output planとcanonical digestが一致し、current runtimeがplanのruntimeへ一致する場合だけ先へ進める
- approvalの四digestを`RB-006`で照合し、plan result改変、stale current state、request差替え、runtime差、destination parent差、非空loss/unsupportedをすべてdestination予約前の`rejected`にする。approvalの欠落・schema不正と、schema-validだがbinding不一致をstable diagnosticで区別する
- この段階ではplanned semantic stateとencoded bytesをinternal-onlyで保持し、directory、member、markerを一切作らない。binding失敗時の`effects.project_artifact`は`null`、cleanupは`not-needed`とする

結果: 2026-08-12に完了。`cli-v1-apply.mjs`のinternal preparation serviceがproject、request、plan result、approvalを`RB-010`順で読み、成功したplanning result envelopeと明示approvalだけを受理する。current stateからC1 planを再計算し、`RB-001`〜`RB-006`、runtime、approved canonical destination、空のloss/unsupportedを再検証した。approval schema不正は`change.approval-invalid`、schema-validなapproval / plan / runtime不一致は`change.binding-mismatch`、apply前に既存化したdestinationは`publication.reservation-conflict`として区別する。別cwdからのapply preparation、approval stdin、stale project、request / plan / approval / runtime改変、destination race / parent symlink化をdirect testで拒否し、いずれもdestinationを作成・置換しないことを確認した。public CLIとfixed subprocess harnessの`apply-change`成功経路はpublisher完成まで未接続のままである。

#### `ZB-P4.6.2` provenanceとstructured observationを純粋生成する

- revalidated current state、request、diff、output plan、approval、runtime、実際にencodeしたoutput bytes / stateから`miku_project_provenance/v1`を決定的に組み立てる。`RB-007`のruntime、input/change/output digest、target、before/afterを一箇所で検査し、schema-validなcanonical JSON bytesとraw digestを返す
- `transformations`は契約で固定した13段階を順序どおり記録する。provenance内の`normalizations`は再計算したapproved output planと一致させ、`losses` / `unsupported`は空でなければpublish前にrejectする。incomplete directory内のprovenanceだけを実行証跡として信用しない
- resultの`observations`は実行中に実測したnormalization / loss / unsupportedから組み立て、`code + path`で重複排除し決定順にする。人間向けmessage解析や、計画時に存在しなかったlossを黙って成功へ変換する経路を作らない

結果: 2026-08-12に完了。`cli-v1-provenance.mjs`がP4.6.1の再検証済みbase state、request、diff、output plan、runtime、実際に渡されたXML bytes / semantic stateから`miku_project_provenance/v1`とBOMなし末尾LF一件のcanonical JSON bytes / raw SHA-256を純粋生成する。生成時にXMLを再decode・semantic validateし、input/output raw digest、state digest、request/diff/plan digest、runtime、target、before/after、approved output normalizationを`RB-007`で相互照合する。loss/unsupported、output bytes/state/normalizationの差替えは`change.binding-mismatch / RB-007`で止める。structured observationsはmessage文字列を使わず`code + path`で決定順にsortし、同一内容だけを重複排除し、競合する同一keyはrejectする。recordはpublisher未接続の候補であり、directory作成、`provenance.json`書込み、marker作成、publication成功の宣言はまだ行わない。

#### `ZB-P4.6.3` read-only artifact verifierをpublisherより先に完成させる

- root pathを`lstat`し、`absent`、markerなしの通常directoryである`incomplete`、markerありで全条件を満たす`committed`、それ以外の確定した不整合である`corrupt`へ分類する。symlink、非directory、余分なmember、member symlink / 非regular file、非0-byte marker、schema/XML/profile/digest/binding不一致を`corrupt`とする
- committed判定では許可された`project.xml / provenance.json / COMMITTED`だけを読み、XML再decodeとsemantic digest、provenance schema、project/provenance raw digest、`RB-007`を検査する。permission、transient I/O、観測中の変更で状態を確定できない場合は推測せず`publication_state = null`の`runtime-error`にする
- `verify-artifact --expect-plan-result`はexpected plan result自体をstrictに検証したうえで`RB-008`を適用する。committed setでも不一致なら`rejected / matches_expected_plan = false / publication.expected-plan-mismatch`とし、実測bindingsは残す
- verifierは全状態でread-onlyとし、cleanup、repair、member生成、marker生成を行わない。`RB-011`によりinput path、verification path、effect path/stateを一致させ、committed以外を`data.artifact_set`へ載せない

結果: 2026-08-12に完了。`cli-v1-artifact-verifier.mjs`がartifact-set rootと固定memberを`lstat`し、markerなしdirectoryを`incomplete`、root/member type・member数・marker size・canonical JSON/XML・provenance schema・output raw/state digestの既知不一致を`corrupt`、不存在を`absent`、permission/transient read failureを`publication_state = null`として分類する。`committed`には三memberのraw digestを含むdescriptorと実測change/diff/plan bindingsだけを返す。`--expect-plan-result`相当のpure binding checkerは、plan runtime/destination、state/request/diff/output-plan digest、target/before/after、artifact digest、normalizationを`RB-008`で照合し、不一致時も`committed / matches_expected_plan = false / publication.expected-plan-mismatch`として実測bindingsを残す。検証経路はdirectory/memberの作成、cleanup、repair、marker作成を一切行わない。public CLI/harness接続はpublisherとapply serviceを完成する後続項目まで保留する。

#### `ZB-P4.6.4` exclusive publisherとcleanup state machineを実装する

- apply直前にdestination parent real path、承認済みcanonical path、runtime/filesystem capability、destination不存在を再確認し、non-recursive directory createで排他的に予約する。競合時は既存entryへ触れず`publication.reservation-conflict`とする
- 予約に成功したinvocationだけがownershipを持つ。directory内の`project.xml`と`provenance.json`はexclusive createでwrite/closeし、marker前に二memberだけであること、type/symlink、XML semantic equivalence、schema、digestをread-only primitiveで再検証する
- 空の`COMMITTED`をexclusive createしてcloseする操作だけをlogical commit pointとし、作成後は三memberとmarker sizeを同じverifierで再検証する。marker作成後は失敗時もcleanup / repairせず、`prohibited-after-commit`としてunknown outcomeを`verify-artifact`で回復できるようにする
- marker前の失敗では、in-memoryで追跡した「このinvocationが作成したmember」とdirectoryだけをbest-effort cleanupする。blind recursive deleteは使わず、既存entry、想定外member、別invocationのdirectoryを削除しない。cleanup成功は`absent / succeeded`、失敗はmarkerなし`incomplete / failed`として元failureと`publication.cleanup-failed`を両方返す
- filesystem操作は注入可能なadapter境界へ置き、write failure、cleanup failure、marker直前/直後、post-marker verification failureをtestで決定的に発生させる。v1のdurability保証はhandle closeまでで、file/directory `fsync`を成功条件へ追加しない

結果: 2026-08-12に完了。`cli-v1-publisher.mjs`がverified core runtimeと承認済みcanonical destinationだけを受け、dynamic destination preflight後にnon-recursive `mkdir`で予約する。`project.xml`と`provenance.json`は`wx`で個別write/closeし、marker前はexact二member topologyとP4.6.3から切り出したpure XML/provenance/digest検査を通す。空の`COMMITTED`を`wx`でcreate/closeした時点からcleanupを禁止し、post-markerではread-only verifierの`committed`判定だけを成功にする。marker前はin-memoryで追跡したregular memberだけを逆順unlinkし、空であることを確認したowned directoryだけを`rmdir`する。destination race、member/marker write failure、cleanup failure、post-marker verification failureをfilesystem adapter injectionで固定し、既存path、想定外entry、marker boundary後setを削除しない。apply/public `verify-artifact` serviceへの接続は次項まで行わない。

#### `ZB-P4.6.5` apply-change serviceをpublicationへ接続する

- result fileを使う場合はresult channelを最初にexclusive reservationし、`P4.6.1 → actual apply/post-validate → P4.6.2 → P4.6.4`の順で一回だけ実行する。actual output bytes/stateがapproved output planのpreflight digestへ一致しなければmarker前に失敗させる
- success resultはpost-marker verifierが返したcommitted descriptorだけを`data.artifact_set`へ載せ、destination、plan、effect path、artifact pathを`RB-011`で一致させる。inputは不変、`created_by_invocation = true`、cleanupは`prohibited-after-commit`、次操作は`verify-artifact`とする
- reservation前のreject、marker前cleanup成功/失敗、marker後のcorrupt判定をresult schemaの`status / exit / effects / observations / diagnostics / next_action`へ対応させる。marker後にresult deliveryを失った可能性は架空のfailure envelopeへ変換せず、emergency stderrと独立verifyで回復する
- `CA-CHANGE-001`、`CA-DEST-EXISTS-001`、`CA-BINDING-001`をfixed-binding service / subprocess harnessへ接続する。approvalなしの自動短絡や、同じapprovalで既存destinationを上書きする再実行経路を作らない

結果: 2026-08-12に完了。`runV1ApplyChange`がP4.6.1の四入力/RB-001〜RB-006再検証後、revalidated base stateから許可済みC1 operationを改めてmaterializeし、post-apply semantic validation、approved output planとのstate/XML raw digest一致、P4.6.2 provenance生成、P4.6.4 exclusive publicationを順に接続した。successはpost-marker verifierのcommitted descriptorだけを`data.artifact_set`へ載せ、`RB-011`のplan/destination/effect/descriptor path、input不変、`created_by_invocation = true`、`prohibited-after-commit`、next `verify-artifact`を検査する。fixed-binding harnessはresult fileをdomain inputより先に予約し、予約不可なら四input metadataを未読のままstdoutへ返す。applyが成功したplan resultを読んでcanonical destinationを得る前のnon-successでは、pathを推測せず`io.destination = null`を許すようresult contract/schemaを明記した。marker前write failureは`absent / succeeded cleanup`、delivery failure after commitはvalid envelopeを捏造せずthrowし、read-only verifierでcommitted outcomeを回復するdirect testを固定した。public source CLI/development bundleはP4.9までfail-closedのままである。

#### `ZB-P4.6.6` committed artifact setをproject入力へ接続する

- `--project <directory>`はP4.6.3のverifierでcommittedになったsetだけから`project.xml`を取り出す共通readerへ通す。`inspect / validate / plan-change / apply-change`の既存external XML pipelineへ接続し、incomplete/corrupt setのmemberを部分的に読まない
- artifact set由来XMLも外部XMLと同じprofile decode / semantic validationを通し、directory sourceとcanonical pathをI/O metadataへ正しく記録する。publication verificationとproject semantic validationの責務は混同せず、前者を通過してから後者を実行する
- external XMLとcommitted setから得た同一semantic stateが、契約で同値とするProjection、diff、output XMLを生むことをfixtureで確認する。入力set、marker、provenanceを変更しない

結果: 2026-08-12に完了。`readV1CommittedArtifactSetProject`はP4.6.3 verifierが`committed`を返した後だけ`project.xml`を読む共通boundaryとなり、raw digestをdescriptorと再照合してverification後のmember差替えも`publication.artifact-corrupt`で拒否する。markerなし・余分member・schema/XML/provenance/digest不整合・I/O判定不能はbytesを返さず、`--project`に対応するstructured diagnosticへ写像する。`readV1ExternalProjectInput`はdirect regular XML、explicit stdin、committed directoryの三入力を同じdecode / semantic validationへ渡すため、`inspect / validate / plan-change / apply-change`に別のdirectory用semantic pathを作らない。directory入力はcanonical artifact-set path、`source = directory`、実際に読んだ`project.xml`のraw digestをI/O metadataに残す。real publicationを使うdirect testでexternal XMLとcommitted directoryのProjection、semantic diff、preflight output XML digest、実際にpublishした`project.xml` bytesが一致し、入力artifactの`project.xml`、`provenance.json`、markerが不変であることを固定した。

#### `ZB-P4.6.7` failure matrix、conformance、bundle回帰を完了する

- workflow corpusの`CVF-ABSENT-001`、`CVF-INCOMPLETE-001`、`CVF-CORRUPT-001`、`CVF-COMMITTED-001`、`CVF-EXPECTED-PLAN-MISMATCH-001`、`CVF-EXPECTED-PLAN-INVALID-001`、`CA-CLEANUP-AGGREGATE-001`をmaterializeする。`CU-UNKNOWN-OUTCOME-001`ではcommit後のresultを破棄し、再applyせずverifyでcommitted outcomeを回収する
- binding corpusの`BC-APPROVAL-DIVERGENCE-001`、`BC-PLAN-BINDINGS-VALID-001`、`BC-APPLY-PATH-DIVERGENCE-001`、`BC-VERIFY-PATH-DIVERGENCE-001`、`BC-VERIFY-STATE-DIVERGENCE-001`を共通validatorへ接続する。diagnostic集約順、effect/cleanup state、result schema、artifact topologyをexactに固定する
- same-runtime反復で`project.xml`と`provenance.json`のbytes/digestが決定的であること、結果file/stdout、四入力のうち一つだけのstdin、usage時未読、既存destination/input不変、repository外bundleへのmodule包含を検証する
- public source CLI / development bundleのfail-closedを維持し、direct test、fixed-binding integration、legacy compatibility、`npm run test:fast`、`npm run test:full`、`npm run build:full`、`git diff --check`を通す。runtime manifest、release asset、公開workflow smokeはP4.9/P4.10の完了として数えない

結果: 2026-08-12に完了。`cli-v1-verify-artifact.mjs`を追加し、fixed verified binding harnessで`verify-artifact`を正式に実行できるようにした。artifact setはまずread-only verifierで観測し、optional `--expect-plan-result`はstrict JSON / schema-valid successful `plan-change` resultとして読んだ後にRB-008を再検証する。`absent / incomplete / corrupt`、committed+expected-plan mismatch、committed観測後のI/O・JSON・envelope/schema不正expected plan、I/O判定不能、result channel preflightを、`verification`・effect・I/O metadata・next actionの契約どおり区別し、repairや再applyを行わない。expected plan自体が不正ならRB-008を評価せず`committed / matches_expected_plan = null / actual bindings`とし、schema-valid planとの不一致だけを`false / publication.expected-plan-mismatch`とする。result schemaはverifyのresult channel予約失敗のようにartifact観測前のrejected resultを`data = null / project_artifact = null`として許可し、観測済みverificationを持つrejected resultだけにはstate/effect bindingを要求するよう補強した。`CA-CHANGE-001`、destination race、binding divergence、cleanup aggregate、全`CVF-*`、expected plan mismatch、`CU-UNKNOWN-OUTCOME-001`をdirect/fixed-binding testでmaterializeし、result deliveryを失ったapplyを再試行せずverifyで回復することを固定した。same canonical destinationでrunnerだけがfresh output directoryを再構成する二回実行では`project.xml`と`provenance.json`のbytesが一致する。13件すべてのcross-artifact contract caseをRB-001〜RB-006、RB-011、RB-012のpure validatorへ接続し、new verifier/apply bindingもbundle/source archiveに含めた。public source CLI/development bundleはP4.9までmanifest未検証runtimeとしてfail-closedのままである。

#### `ZB-P4.6` 完了判定

- current project、request、plan result、approval、runtime、destinationを再計算で束縛し、不一致時はdestination予約前に止まる
- apply successはexact三memberのcommitted artifact setだけを返し、loss/unsupportedを含む出力、markerなしdirectory、post-marker verification不一致を成功扱いしない
- verifierが`absent / incomplete / committed / corrupt / inspection failure(null)`を契約どおり区別し、expected plan不一致とunknown outcomeをhidden stateなしで回復できる
- cleanup権限がこのinvocationのmarker前生成物だけに限定され、cleanup失敗を元failureとともにstructured diagnostics / effectsへ残す。marker後は自動cleanupしない
- external XMLとcommitted artifact setの両project入力が同じsemantic pipelineを使い、入力と既存pathを変更しない
- listed workflow / binding case、failure injection、determinism、legacy回帰、full buildが成功し、partial runtimeをNode reference releaseとして公開していない
- 完了後の次作業は`ZB-P4.7`で、現行coreの再利用候補をv1 conformance fixturesに照らして採否判定する

### P4.7の実行計画

目的は、`ZB-P0.3`の仮分類をR1/C1で実証済みのv1契約に照らして更新することである。旧coreを再利用することは目的ではない。新しいsemantic state、structured result、human gate、committed artifact setのいずれかを曖昧にする候補は、互換性の証拠として残してもv1の実行経路へは採用しない。

- [x] `ZB-P4.7.1` 現行候補とv1実行経路を対応付ける
  - XML codec / normalizer / validator、AI view / Patch、core API loader、bundle builder、旧Agent workflow例、旧fixture / regression test、workbook / report / browser runtimeを候補単位に分ける
  - 各候補について、現在のv1実装箇所、使用するconformance case、採用しない場合の保持先、次の再評価gateを明記する。P0の仮分類を黙って上書きせず、判断記録から参照可能にする
- [x] `ZB-P4.7.2` v1 fixtureが要求する差を検証する
  - `S-V001`のexact semantic golden / raw provenance、`S-I012`のstable invalid rule、`S-I020`のunsupported分離、`CI-OVERVIEW-001` / `CI-CONTEXT-001`のProjection、`CP-CHANGE-001`とC1 apply / verificationを、候補の責務と照合する
  - legacy `dependency.xml`とv1 canonical XMLは同名入力の単純コピーとして扱わない。raw bytes、field order、lag representationを含め、v1 fixture / goldenが契約正本であることを確認する
  - legacy XML round-trip / core API loader testも実行し、旧経路を削らずに保守できることと、v1 conformanceを代替できないことの両方を確認する
- [x] `ZB-P4.7.3` 採否と後続条件を判断記録へ固定する
  - outcomeは少なくとも `v1実行経路へ採用`、`互換性の証拠として保持`、`後続scopeまでdefer` を区別する。候補を保留するだけの場合も、理由と再評価gateを記録する
  - v1 semantic pathがlegacy `ProjectModel`、global core API loader、旧Patch / AI view schemaへ暗黙依存しないことを確認する。配布builderの再利用はsemantic reuseと数えず、P4.9のmanifest / asset binding条件を残す
  - この項目では旧形式・旧command・名称を削除または改名せず、public CLIのmanifest未検証fail-closed境界も変えない
- [x] `ZB-P4.7.4` 判断記録と回帰結果をレビューし、人が承認した（2026-08-13）

成果物: [v1 core再利用採否 v20260812](miku-project-zero-base-reuse-verdict-v20260812.md)。

完了条件:

- すべてのP0 reuse候補が、R1/C1に対する現在の役割、根拠fixture / test、採否、再評価gateを持つ
- v1 semantic / projection / change / artifactの各境界で、旧coreをactive dependencyにしない根拠がある。また採用する非semantic build資産にはP4.9で満たすべき条件が残る
- v1 conformanceとlegacy compatibilityの双方が成功し、fixtureの役割と正本性が文書化される
- 人の承認前には、この採否をJava移植、Skills runtime選択、旧API削除の根拠にしない

### P4.8の実行計画

P4.8の最初のsliceは、承認済みsemantic fixture catalogの`S-V002`を用いた**階層C1 slice**とする。summary task配下のleaf taskをR1で観測し、ancestorを含む`task_change_context`から、既存の`set_task_percent_complete`だけをplan / approval / apply / verifyまで往復させる。

このsliceを先に選ぶ理由は、v1の中心であるWBS階層、限定Projection、leaf-only C1、canonical XML encode/redecode、committed artifact入力を、追加の製品契約なしに同時に検証できるためである。workbook/XLSX、新しいProjection purpose、dependency / resource / calendar編集、actual / baseline / timephased / extended dataは、このsliceに含めない。これらはformatまたはchange contractの改訂と別承認が必要になる。

旧`testdata/hierarchy.xml`はP4.8のv1入力に使わない。空の`Resources` / `Assignments` containerと`Notes`を含むため、v1 XML profileではそのまま成功入力にできない。現行互換fixtureとしては残し、P4.8では`S-V002`のrequired field、canonical child order、resource / assignment / calendar参照を明示した新しいv1 canonical fixtureを作る。

- [x] `ZB-P4.8.1` hierarchy corpusとcapability scopeを固定する
  - `testdata/conformance/v1/fixtures/project/`へ`S-V002`のcanonical XML、`golden/semantic/`へbefore / after state、`golden/projection/`へoverview / nested-leaf contextを追加する。対象leafはUID `2`、summary ancestorはUID `1`、C1は`0 → 50`とする
  - `suite-index.json`へ成功経路の`CV-HIERARCHY-VALID-001`、`CI-HIERARCHY-OVERVIEW-001`、`CI-HIERARCHY-CONTEXT-001`、`CP-HIERARCHY-CHANGE-001`、`CA-HIERARCHY-CHANGE-001`と、拒否経路の`CV-HIERARCHY-INVALID-PREORDER-001`、`CV-HIERARCHY-INVALID-SUMMARY-001`、`CP-HIERARCHY-SUMMARY-REJECT-001`を追加する。各caseは既存のschema / semantic-state / exact-json / cross-artifact-binding / artifact-topology / byte-same-runtime規則を使い、新しい比較modeを導入しない
  - `docs/miku-project-v1-capability-matrix-v20260813.md`を新設し、階層read / nested-leaf C1を`Node reference: implemented`、resource / assignment / calendarの編集、workbook/XLSX、new formatを`deferred`として記録する
- [x] `ZB-P4.8.2` hierarchy decode・semantic validationをfixtureで閉じる
  - canonical XMLからpreorder `1, 2, 3, 4`、parent relation、summary / child整合性、root / sibling order、既存dependency / assignment / calendar参照をexact semantic goldenへdecodeする
  - `S-I003`の不正parent / preorder、`S-I004`のsummary不整合を個別fixtureとしてmaterializeし、suite caseでも`semantic.invalid`、rule ID、diagnostic `location.path`を固定する。legacy coreのnormalizerやvalidatorをv1 pathへ呼び込まない
  - v1 encoderのoutline level / outline number導出後に再decodeし、canonical stateとsemantic equivalentであること、task順が変わらないことを確認する
- [x] `ZB-P4.8.3` R1のhierarchy Projectionを固定する
  - `project_overview`が全taskをpreorderの`order`、`parent_uid`、`summary`とともに返すことをexact goldenで確認する
  - UID `2`の`task_change_context`がancestor UID `1`を一件だけ含み、target、dependency、target assignmentとresourceを過不足なく返すことを`RB-012`で検証する
  - summary UID `1`または不存在UIDへの`task_change_context`をrejectし、read-only commandがproject / artifact / result既存pathを変更しないことを維持する
- [x] `ZB-P4.8.4` nested-leaf C1のplan / apply / verifyを接続する
  - UID `2`だけを`0 → 50`に変更するstrict request、semantic diff、output plan、approval、provenance、committed artifact setを既存のC1 schema / bindingで生成・照合する
  - apply後のXMLを再decodeし、UID `2`以外のtask field、parent relation、task order、dependency、resource、assignment、calendarがbefore stateとsemantic equivalentであることを確認する
  - summary UID `1`へのchange requestを`CP-HIERARCHY-SUMMARY-REJECT-001`で`S-I022` / `change.operation-unsupported`としてhuman gate前にrejectし、destinationを作らない。apply後のartifact set入力と直接`project.xml`入力へ同じC1 request / destinationを渡し、hierarchy Projection、semantic diff、output plan（preflight XML digestを含む）、RB-001〜005 bindingが一致することを確認する
- [x] `ZB-P4.8.5` regression、capability記録、承認資料を完成する
  - hierarchy direct tests、fixed-binding subprocess integration、existing flat R1/C1、legacy XML / core API / CLI / workbook/XLSX regression、bundle/source archive、`npm run build:full`、`git diff --check`を通す
  - P4.8でmaterializeしたfixture ID、Node capability、defer項目、public runtimeがP4.9までfail-closedであることを`docs/miku-project-v1-capability-matrix-v20260813.md`とTODOへ記録する
  - P4.8のscope外を実装済みと記さず、P4.9 manifest / asset binding、P4.10 clean runtime smoke、Java / Skills / Web/MCPの開始条件を変更しない
- [x] `ZB-P4.8.6` hierarchy C1 sliceを再レビューし、人が承認した（2026-08-13）

結果: `S-V002`のcanonical XML、before / after semantic golden、overview / nested-leaf context golden、成功経路の`CV/CI/CP/CA-HIERARCHY-*`五caseと、拒否経路の`CV-HIERARCHY-INVALID-PREORDER-001`、`CV-HIERARCHY-INVALID-SUMMARY-001`、`CP-HIERARCHY-SUMMARY-REJECT-001`を追加した。Node reference pathはordered forestをcanonical stateへdecodeし、encoderのoutline導出後も同じstateへredecodeする。`S-I003` / `S-I004`はmachine-readableなvalidate caseでcode、rule ID、diagnostic `location.path`まで固定し、`S-I022`はsummary change requestの拒否診断へ同じlocation規則で記録する。UID `2`のcontextだけがancestor UID `1`を含み、summary UID `1`と不存在UIDはread-only contextを受け取れない。nested leafの`0 → 50`はplan / approval / apply / verifyを通り、committed artifact setを入力にした次の`50 → 75` planも成功する。artifact set入力とその`project.xml`直接入力は、同じhierarchy Projection、semantic diff、output plan（preflight XML digestを含む）、RB-001〜005 bindingを返す。source archiveにはhierarchy corpusを含め、public source CLI / development bundleはP4.9までmanifest未検証のままfail closedである。2026-08-13に`npm run build:full`（29 files / 288 tests）と`git diff --check`が成功し、同日にP4.8.6は人が承認した。次はP4.9である。

完了条件:

- `S-V002`、`S-I003`、`S-I004`、`S-I022`がv1 corpus / testへmaterializeされ、拒否diagnosticのcode / rule ID / location pathを含めてflat fixtureだけでは検証できない階層契約を証明する
- nested leafのR1 Projection、C1 plan / approval / apply / verifyが、artifact set入力を含めて既存v1 contractに適合する
- summary編集、壊れたforest、unsafe outputでartifactを公開しない
- 新format、new command、新purpose、new change operation、legacy core dependency、public runtime有効化を導入しない
- 既存のflat R1/C1とlegacy compatibilityが回帰せず、人の承認を得る

### P4.9の実行計画

P4.9は、P4.8で実証したNode reference pathを**実行可能なversioned runtime artifact**へ束縛する工程である。ここでpublic v1 workflowを有効化できるのは、single `.mjs`、対応source archive、固定名`runtime-manifest.json`が同じdirectoryにあり、runtime自身が毎workflow開始前にmanifest / executable / sourcesの整合性を検証できる場合だけとする。release checksumまたはSkills lockによるmanifest digestの外側pin、repository外のclean smoke、公開ReleaseはP4.10以後であり、P4.9だけで配布元authenticityを主張しない。

現在のworking treeがdirty、または`package.json`のversionに対応する`v<version>` tagがHEADを指さない場合、release runtime builderはartifactを生成しない。この拒否は開発を止めるためでなく、source revision/tagがartifact内容を偽装しないためのpreflightである。テスト用の一時runtimeは明示的なtest-only source identityで作れても、release assetやmanifest digest pinとして扱わない。

- [x] `ZB-P4.9.1` versioned Node runtime builderを固定する
  - clean working tree、`git rev-parse HEAD`、`v<package.version>`のexact tagをrelease source preflightとして検証する。いずれかが満たされなければ、runtime directoryやassetを作成せずに失敗する
  - fresh output directoryだけに`miku-project-node-<runtime.version>.mjs`、`miku-project-node-<runtime.version>-sources.tgz`、`runtime-manifest.json`を生成する。既存directoryや既存assetを置換しない
  - bundle/source archiveは既存の決定論的builderを再利用するが、runtime artifactにはbundle時点のproduct/runtime version、contract descriptor、capability profile、fixture corpus digestを埋め込む。development bundleはこのmetadataを持たない
- [x] `ZB-P4.9.2` manifestとcorpus bindingを生成・検証する
  - `testdata/conformance/v1/`の通常fileだけをPOSIX relative path順にdigestし、契約のfive-step corpus digestを計算する。symlink/non-regular entryはrelease buildを拒否する
  - executable/sourceのraw byte sizeとSHA-256、nine capabilityのcanonical順、product / contract / schema / fixture suite、source contract/runtime revision/tag、`reference_runtime = null`をcanonical JSON + LFのmanifestへ記録する
  - runtime manifest JSON Schemaに加え、family/versionから導くbasename、artifact directory escape/symlink拒否、profileと九capabilityの完全一致、embedded metadataとの一致をrelease validationで検証する
- [x] `ZB-P4.9.3` versioned bundleのself-validationとpublic routingを接続する
  - versioned Node executableだけがadjacent manifestを通常fileとして読み、UTF-8/BOM/duplicate key/schema、manifest raw digest、basename/size/SHA-256、corpus/capability/versionを検証する。source archiveも各workflow前に必ず検証する
  - 成功時はmanifest/executable digestを含む`binding_status = verified` runtimeをR1/C1 serviceへ渡す。manifest/asset/source/capability不正は`runtime-error / 3`でproject、request、approval、artifact set、destination、result pathを読まず/作らずに返す
  - repository source CLIとmetadataを持たないdevelopment bundleは、P4.8までと同じ`runtime.capability-missing` fail-closedを維持する。manifest探索、glob、PATH fallback、別version fallbackを導入しない
- [x] `ZB-P4.9.4` runtime-integrity conformanceをmaterializeする
  - test-only temporary runtimeでvalid manifestを作り、`validate`をverified runtimeとして通す。result / output plan / provenanceのruntime bindingがmanifest/executable digest、profile、fixture suiteと一致することを確認する
  - `CR-MANIFEST-INVALID-001`、`CR-ASSET-DIGEST-001`、`CR-CAPABILITY-MISSING-001`、`CR-SOURCE-MISSING-001`をmanifest/asset/source mutationとして実行し、expected code、project未読、destination/result未作成、legacy fallbackなしを固定する。direct filesystem guardでfailure時のproject / result pathに対する`lstat`・`realpath`・`readFile`も監視する
  - 同じsource identityとoutput pathから二回作るruntime asset / source / manifestのbyte determinismを確認する。P4.10のrepository外のclean runtime smokeとは別のbuilder/service testとして扱う
- [x] `ZB-P4.9.5` 記録と承認資料を完成し、人が承認した（2026-08-13）
  - capability matrix、conformance corpus、TODOへNode reference runtimeの実装範囲と、manifest digestの外側pin・clean smoke・Releaseがまだ未完である境界を記録する
  - targeted runtime test、existing R1/C1・legacy regression、`npm run build:full`、`git diff --check`を通す。P4.10/P5/P6/Web/MCPの開始条件を先取りせず、人のレビューと承認を得る

完了条件:

- clean/tagged sourceだけから、固定layoutのversioned Node executable、sources、manifestを安全に生成できる
- five workflow commandはversioned bundleのmanifest / asset / source検証後だけにverified bindingで実行され、失敗時にdomain input/outputへ副作用を出さない
- `CR-*` runtime-integrity caseがNodeでmaterializeされ、manifest、asset、source、capabilityの別々のfailureをstable diagnosticで区別する
- source/development bundleのfail-closed境界、外側manifest pin、P4.10 clean smoke、Java/Skillsの後続条件が維持される

実装記録（2026-08-13承認）: `scripts/build-cli-v1-runtime.mjs`はcleanかつexact tagのrelease sourceだけを受け、`runtime/node/`の既存directoryを置換せずにversioned Node runtimeを生成する。test-only builderはtemporary runtimeを用意し、`CR-*`を`suite-index.json`のmutation定義へ接続する。failure caseはsubprocessでのartifact未作成確認に加え、direct filesystem guardでproject / result pathへの`lstat`・`realpath`・`readFile`が一度も起きないことを固定する。さらにclean・exact-tagのtemporary source repositoryからrelease builderを実行し、生成runtimeがrepository外のworking directoryから`validate`をverified bindingで実行できることを確認する。このtestはproduction Release / external checksumを作らないため、P4.10の配布artifact smokeを代替しない。versioned executableはembedded product/runtime versionとconformance corpus digestを使ってadjacent manifestを検証し、five workflowへverified bindingを渡す。source CLI / development bundleはmetadataを持たないため従来どおりfail closedである。2026-08-13に`npm run build:full`（30 files / 297 tests）と`git diff --check`が成功し、人が承認した。これはP4.10のrepository外clean smoke、Release checksum、Skills lock、公開Releaseを完了したという意味ではない。

### P4.10の実行計画

P4.10は、P4.9で生成・自己検証できることを確認したruntimeを、**source checkoutから切り離し、外側のtrust anchorでmanifestをpinしたconsumer配布物**として実行する工程である。release sourceの一時cloneをcleanかつexact tagとしてbuilderへ渡し、生成された`runtime-manifest.json`、versioned executable、source archiveの三fileだけを別のconsumer directoryへコピーする。consumer directoryへproject / request / plan / approvalを置くことは許可するが、runtime実行中にsource checkout、`node_modules`、repositoryの相対moduleへ依存してはならない。

2026-08-13に、外側manifest pinを追加する`ZB-P4.10.2`〜`ZB-P4.10.5`の補強計画を人が承認した。この承認は実装開始の承認であり、P4.10の完了・release・公開承認ではない。

同日の第2回reviewで、manifest pin成功後にexecutable / sourcesのsize・SHA-256をconsumerが検証しないままNodeを起動している問題を検出した。runtime自身のself-validationはJavaScriptの実行開始後であるため、改ざんされたexecutableの起動前検証を代替できない。`ZB-P4.10.2`〜`ZB-P4.10.4`を再度openにして再補強し、full regression・再review・人の承認を残す。この再補強はP4.10完了・release・公開承認ではない。

2026-08-14の第3回reviewでは、consumerによるmanifest / executable / sourcesの起動前境界は妥当と判断した一方、`assertPinnedNodeRuntimeResultBinding`が既知fieldだけを比較し、testも部分一致を併用しているため、CLI result schemaが禁止する余分なruntime fieldやnested digest fieldを受理できる問題を検出した。runtime manifest contractが要求する「CLI resultのruntime binding全体が検証済みmanifestと完全一致する」を実証できていないため、`ZB-P4.10.4`と`.5`をopenに戻した。以下の最終補正は`validateCliResult`によるenvelope全体のschema検証後にcanonical JSONで`result.runtime`とverified preflight bindingを完全比較するものとし、schema不正を`consumer.result-invalid`、schema-validな値不一致を`consumer.result-binding-mismatch`として区別する。これはP4.10完了・release・公開承認を意味しない。

P4.10では、`buildReleaseNodeRuntime`がmanifest生成直後に返したraw SHA-256を、配布runtime directoryや配布manifestから再導出しないtest-owned trust anchorとして保持する。consumerはruntime起動前にその値と配布manifest raw bytesを照合し、pin成功後だけmanifestをparseしてlauncherを解決する。これにより、manifest・executable・sourcesをまとめて書き換え、内部digestを合わせた改変も外側pinで拒否する。

これはproduction Releaseのpublish自体ではない。P4.10はRelease checksumと同じ二層trust modelをsmoke内で実証するが、公開checksum fileの発行、署名、network download、Skills lock、Java / Agent Skills integration、公開Releaseは後続scopeに保つ。また、検証完了後からprocess spawnまでの同時書換えを防ぐTOCTOU対策、file descriptorからの直接実行、配布directoryの権限設計はP4.10のtest-owned smokeへ広げず、production Skills/runtime配布設計でimmutableまたは管理下directoryを要求する後続論点として残す。

- [x] `ZB-P4.10.1` external consumer smokeをmaterializeする
  - clean / exact-tag temporary sourceから`buildReleaseNodeRuntime`でruntimeを生成し、manifest、versioned `.mjs`、sourcesだけをrepository外のconsumer runtime directoryへcopyする。配布memberがこの三fileだけであることを確認し、source checkoutを削除してから実行する
  - consumerが所有するcanonical project、materialized change request、plan result、approval、artifact destinationを使い、`validate`、`inspect`、`plan-change`、`apply-change`、`verify-artifact`を全て実行する。各resultの`binding_status = verified`、runtime digest / manifest digest / capability profile / fixture suiteが配布manifestと一致することを確認する
- [x] `ZB-P4.10.2` test-owned consumer preflightをmanifestと両assetの起動前検証として固定する
  - `buildReleaseNodeRuntime`がcopy前に返した`manifestDigest`をruntime directory外の明示引数としてconsumer verifierへ渡す。期待値は`sha-256` / 64 lowercase hexだけを受け、配布manifestやそのfilenameから生成しない
  - fixed `runtime-manifest.json`を`lstat`し、通常fileかつ非symlinkであることを確認してraw bytesを読む。raw SHA-256がtrust anchorと一致するまでJSON parse、artifact path解決、launcher構築、runtime spawnを行わない
  - pin成功後にstrict UTF-8 / BOM / duplicate key / schema / canonical JSON、Node reference identity、core capability、corpus、fixed basenameを検証する。既存のruntime自己検証primitiveを再利用しても、外側pin照合を省略または後置しない
  - pin済みmanifestのcanonical pathの親を唯一のruntime directoryとし、固定basenameのexecutable / sourcesをそこから解決する。両entryを`lstat`して通常file / 非symlinkを要求し、`realpath`後も同じcanonical directoryの直接memberであることを確認する。directory scan、glob、PATH、別version、alternate filenameを使わない
  - executable / sourcesのraw bytesを読み、manifest descriptorの`size_bytes`と`sha-256`へ一致させる。両方の検証が完了するまでverified bindingとlauncher descriptorを返さず、runtime self-validationは起動後の二重検証として維持する
  - test-owned helper内に同じ検証を重複実装する場合でも、`cli-v1-runtime-manifest.mjs`のfile descriptor / digest規則と一つの共通primitiveへ寄せる。既存self-verifierの検証順を外側pinより前へ移動しない
- [x] `ZB-P4.10.3` coordinated tamperとsingle-asset tamperをpre-launchで拒否する
  - versioned executableへ実行可能なcommentをappendし、その新size / SHA-256をmanifestへ反映してcanonical JSONで再serializeする。これによりbundle内部のmanifest / executable整合は成立しても、original trust anchorとのmanifest raw digestは不一致になるcaseを作る
  - consumer verifierがpin mismatchとして失敗し、注入したlauncher callbackが0回、project / request / plan / approval / result / destinationのfilesystem accessが0回、result / artifact setが未作成であることをassertする。legacy、PATH、別versionへfallbackしない
  - missing / malformed expected digestと、manifest missing / symlinkもruntime起動前に拒否する。これはCLI result diagnostic catalogへ新codeを追加せず、P4.10 consumer preflightのtest outcomeとして区別する
  - original manifest bytesと外側pinを変更せず、executableだけへcontentを追加するcaseを作る。manifest pinは成功してもexecutable size / digest不一致で拒否し、launcher callbackが0回であることを確認する。改ざんcode自身へ検証を委ねない
  - original manifest bytesと外側pinを変更せず、sourcesだけへcontentを追加するcaseも同じ境界で拒否する。必要に応じてexecutable / sourcesのmissing、symlink、size mismatchもtable-driven caseにし、どのcaseでもdomain I/Oとfallbackがないことを共通assertする
- [x] `ZB-P4.10.4` successful five-workflow smokeを完全なpinned consumer preflightと、schema-validかつ完全一致するconsumer result bindingへ接続する
  - [x] `ZB-P4.10.4a` operationごとにmanifest pin、manifest identity、executable / sourcesの通常file・非symlink・canonical path・size・SHA-256を再検証する。全検証後の`runtime.launcher = node`、検証済みcanonical executable path、固定`process.execPath`だけからcommandを組み立て、hard-coded artifact path、glob、directory scan、PATH、mtime、newest / alternate version fallbackを使わない。`validate`、`inspect`、`plan-change`、`apply-change`、`verify-artifact`の成功経路と、plan / provenanceのexact runtime bindingは維持する
  - [x] `ZB-P4.10.4b` `tests/helpers/mikuproject-v1-pinned-runtime-consumer.mjs`が受け取るCLI result envelopeを、既存の生成済み`validateCliResult`で最初に検証する。schema不正はtest-owned `consumer.result-invalid`としてfail closedにし、Ajvの可変なmessageを公開contractへ持ち込まず、CLI diagnostic catalogへ新codeを追加しない
  - [x] `ZB-P4.10.4c` schema検証後の`result.runtime`とconsumer preflightが返す`runtime`をdeep exact比較する。top-level key集合、`artifact_digest` / `manifest_digest`のnested key集合と値をすべて対象にし、schema-validだが期待bindingと異なる場合は既存のtest-owned `consumer.result-binding-mismatch`で拒否する。known fieldだけの列挙比較、`toMatchObject`、部分一致は受理判定に使わない
  - [x] `ZB-P4.10.4d` 成功resultのcopyを用いたtable-driven direct testで、required runtime field欠落、余分なtop-level runtime field、余分なnested digest field、schema-validなversionまたはdigest値不一致を固定する。前3件は`consumer.result-invalid`、値不一致は`consumer.result-binding-mismatch`となり、成功五workflowも同じschema validatorとexact比較を必ず通ることを確認した（P4.10 consumer対象13 cases、runtime manifest test file全22 tests）
  - [x] `ZB-P4.10.4e` result以外のplan runtime binding、approvalが束縛するoutput plan digest、apply provenance、verify resultの既存exact比較を維持し、source checkout削除、consumer `node_modules`なし、runtime三memberだけの条件を再確認した
- [x] `ZB-P4.10.5` 記録と承認資料を完成する（2026-08-14承認）
  - TODO、capability matrix、README、developmentへexternal pin成功、両assetの起動前検証、coordinated / single-asset tamper拒否と、公開Release checksum、署名、Skills lock、公開Releaseが未完である境界を記録する
  - exact result-binding negative testとtargeted smokeで、成功五workflow、executable-only tamper、sources-only tamper、coordinated tamper、anchor / manifest / asset entry failureを確認した。existing R1/C1・legacy regressionを含む`npm run build:full`は2026-08-14に30 test files / 310 testsで成功し、同じ補正後差分に対する`git diff --check`も成功した。P5/P6/Web/MCPの開始条件を先取りせず、同日に人の再reviewと承認を得た
  - manifest-pin-only baselineとして2026-08-13に`npm run build:full`（30 files / 303 tests）と`git diff --check`が成功したが、起動前asset検証不足により承認証拠には数えない。再補強後にtest数を含む全証拠を再取得する

manifest-pin-only実装記録（第2回reviewで差戻し）: test-owned consumer preflightは、copy前にrelease builderが返したraw `manifestDigest`だけを外側trust anchorとして受け、fixed `runtime-manifest.json`のregular-file / non-symlink確認とraw SHA-256照合をparse、artifact path解決、launcher構築、runtime spawnより先に行う。pin成功後にstrict JSON、schema、canonical serialization、Node reference identity、canonical capability、corpus、fixed basenameを検証し、manifestの`runtime.launcher = node`と`artifacts.executable.path`から固定`process.execPath`によるcommandを構成した。five workflow、coordinated tamper、anchor / manifest failureは通り、2026-08-13に`npm run build:full`（30 files / 303 tests）と`git diff --check`が成功した。しかしconsumerはexecutable / sourcesのsize・SHA-256を起動前に検証しておらず、改ざん済みJavaScriptをNodeが実行してからruntime自身にself-validationさせる境界だったため承認しない。

起動前asset再補強の実装記録（第3回reviewでresult bindingを差戻し）: consumer preflightはpin済みmanifestのcanonical parentを唯一のruntime directoryとして、固定basenameのexecutable / sourcesを同じregular-file descriptor primitiveで読む。各assetについてnon-symlink regular entry、canonical directoryの直接member、manifest記載size、raw SHA-256を検証し、すべて成功するまでverified bindingとlauncher descriptorを返さない。runtime自己検証は起動後の二重検証として残す。外側pinを変更しないexecutable-only / sources-onlyのsize改変とsame-size digest改変、missing executable、symlinked sourcesはすべてlauncher 0回かつdomain I/O / result / artifactなしで拒否する。coordinated manifest/executable tamper、anchor / manifest failure、source checkout削除・consumer `node_modules`なしの五workflow成功も維持する。第3回reviewの補正として、consumer helperはCLI result envelope全体を生成済みschemaで検証してからruntime bindingをcanonical JSONで完全比較するよう変更した。runtime field欠落、runtime余分field、nested digest余分fieldは`consumer.result-invalid`、schema-validな値不一致は`consumer.result-binding-mismatch`で拒否し、成功五workflowも同じ経路を通す。2026-08-14にP4.10 consumer対象13 cases（runtime manifest test file全22 tests）、`npm run build:full`（30 files / 310 tests）、補正後差分の`git diff --check`が成功し、再reviewと人の承認を得た。補正前の2026-08-13 P4.10 target 12 tests、`npm run build:full`（30 files / 309 tests）、`git diff --check`はbaselineとして保持する。これは公開Release checksum、署名、Skills lock、公開Releaseを意味しない。

完了条件:

- clean / exact-tag sourceで生成したruntimeの三配布memberだけが、source checkoutを削除したconsumer directoryで、外側pin成功後にfive workflowをverified bindingとして実行できる
- consumerのproject / request / plan / approval / artifact setはすべてconsumer側にあり、source checkout、runtime build dependency、`node_modules`への実行時依存がない
- coordinated manifest / executable改変を、runtime起動とdomain I/Oより前に外側pinで拒否できる
- manifest pinが成功するexecutable-only / sources-only改変も、consumerによるsize・SHA-256検証でruntime起動とdomain I/Oより前に拒否できる
- verified bindingとlauncher descriptorは、manifest pin、manifest identity、executable / sourcesのentry・size・digest検証がすべて成功した後だけ得られる
- 成功CLI resultは生成済みv1 result schemaに適合し、`result.runtime`の全keyとnested digestがconsumer preflightのverified bindingへ完全一致する。schema不正とschema-validなbinding不一致をtest-owned errorで区別し、余分なfieldも拒否する
- P4.10で二層trust modelを実証するが、公開Release checksumの発行、署名、Skills lock、公開Release、Java / Agent Skills integrationの完了とは取り違えない

補強前実装記録（reviewで差戻し）: clean / exact-tag temporary sourceから生成したP4.9 runtimeをconsumer directoryへ三fileだけ配布し、source checkoutを削除した後にconsumer側のflat canonical projectとC1 artifactsだけで五workflowを実行するtestを追加した。consumer側には`node_modules`を置かず、五resultはすべて`binding_status = verified`だった。しかし期待manifest digestを配布manifest自身から導出しており、coordinated tamperを検出する外側trust anchorになっていなかった。2026-08-13のtargeted test（10 tests）と`npm run build:full`（30 files / 298 tests）、`git diff --check`は補強前baselineとして保持し、P4.10の承認証拠には数えない。

### Gate G4

- v1 scenarioがclean directoryで完結する
- 同一入力の反復実行が、契約どおりbyte-identicalまたはsemantic-equivalentになる
- invalid input、digest mismatch、unsafe overwriteでcommitted artifactを生成しない。cleanup不能時のincomplete directoryはmarkerなしで識別され、diagnosticsにpathと状態が出る
- repository外でsingle `.mjs` が動作する
- relevant tests、contract suite、`npm run build:full` が成功する
- Node contract releaseとartifact manifestが固定される

2026-08-14にsource freeze commit `693b4ecd7d4328d77f3b2eada9c4965a9c9b15f5`へ`v1.0.3` tagを付与し、exact-tag clean sourceからactual Node runtimeを生成した。actual三memberは`distribution_status = internal-reference-only`のcandidateとして`workplace/gate-g4/v1.0.3/runtime/`へ保持する。Git管理する[miku-project Node reference runtime lock v1.0.3](miku-project-node-reference-runtime-lock-v1.0.3.json)がsource/tag、build toolchain、package-lock / corpus、三memberのsize / SHA-256を固定し、lock SHA-256は`95cd11cc4460348fa066908994430adba5983384c06c75679855120e5c5ea3d5`である。`verify:cli-v1-release-candidate`はlockを外側pinとしてactual三memberを隔離consumerへcopyし、五workflowとresult / output plan / provenance bindingを再検証する。runtime manifest test file全25 testsと`npm run build:full`（30 files / 313 tests）は成功し、Gate G4は2026-08-14に人が承認した。現行GitHub Release workflowはこの三memberをbuild / uploadしないため、`v1.0.3`を新v1 runtimeの公開Releaseとして扱わない。詳細は[Gate G4 readiness v20260814](miku-project-gate-g4-readiness-v20260814.md)を正本とする。

## ZB-P5 Java CLI適合

依存: `G4`  
主担当: Java CLI
状態: Gate G4承認により開始済み。2026-08-14に[Java contract handoff v1.0.3](miku-project-java-contract-handoff-v1.0.3.md)をmain側の実装入力として固定し、P5-AのJava repositoryへのsnapshot取り込みは完了・承認済み。P5-B harnessは、2026-08-15の再reviewで見つけたsemantic collection canonicalizationとRB-012 leaf制約を補正し、focused test、全体test、snapshot post-verification、両repo diff checkを通した再技術reviewを完了し、同日に人が承認した。P5-C1の`validate` vertical sliceは初回・再reviewのJava側指摘を修正し、`CU-USAGE-001`の現行Node/Java source挙動もcorpusどおり一致させ、全体回帰まで完了した。frozen Node `v1.0.3` release identityは修正前なので、corrective releaseと新snapshot前にcross-runtime / Gate G5の適合は主張しない。現在は再技術reviewと人の承認待ちである。承認前に`inspect`以降を開始しない

### 作業

- [x] `ZB-P5.1` moving Node subtree追随を止め、固定されたcontract release単位の入力へ切り替えた
  - [x] `ZB-P5.1.1` Node `v1.0.3` / revision `693b4ecd7d4328d77f3b2eada9c4965a9c9b15f5`、contract version `1`、fixture suite `1`、corpus digest、snapshot allowlist、Java側配置、受入条件をhandoff文書へ固定する
  - [x] `ZB-P5.1.2` Java repositoryにimmutable `vendor/miku-project-contract/v1.0.3/` snapshot importerとcanonical `SOURCE.json` generatorを追加した。Git tagでidentityを確認した後はfull revisionだけを読み、Node source working treeをcopyしない。expected manifestと全memberが完全一致する既存snapshotだけを無変更no-opにし、その他はhard errorにする。exclusive directory creationとrepository外・symlink親の拒否、ownership確認済みfailure cleanupにより既存treeを上書きしない
  - [x] `ZB-P5.1.3` snapshotのmember set、exact directory topology、regular/non-symlink、size、SHA-256、source tag/revision、corpus digest、`SOURCE.json` raw SHA-256外側pinを検証するJava testと、synthetic Git repositoryでimporter boundaryを検証するNode testを追加した。extra / missing / tampered / symlink / fractional size / rewritten manifest、異なる既存tree / empty directory、tag / revision、working tree、確認後tag移動、出力path、marker / member / manifest failure、unknown file / empty directoryを検証し、Node 12 testsとfocused Java verifier 8 testsが成功した
  - [x] `ZB-P5.1.4` Java README / upstream docsで、新v1 authorityをcontract snapshot、既存`vendor/mikuproject/`をlegacy referenceとして区別した。repository-wide `sh scripts/test-all.sh`をRelease workflowへ接続し、Node 12 testsとJava 140 tests（failures 0、errors 0、skipped 4）が成功した
- [x] `ZB-P5.2` 共通fixture、golden、result/diagnostics schemaをJava testから利用する
  - [x] `ZB-P5.2.1` `suite-index.json` / `contract-cases.json` loaderをJava test-side harnessとして実装した。public entrypointはP5-A snapshot verifier後のimmutable treeだけを読み、30 workflow caseと31 contract/binding caseを型付きで読み出す。root / case / known parameter / input / mutationのunknown field、workflowとcontract indexを通じたcase ID重複、snapshot外escape、allowlist外・symlink参照、未知role、不正JSON Pointerをfail-closedにする。実snapshotとcopied mutationを用いたfocused 8 tests、repository-wide `sh scripts/test-all.sh`（Node 12 tests + Java 148 tests、failures 0、errors 0、skipped 4）が成功した
  - [x] `ZB-P5.2.2` result / diagnostic / artifact / runtime manifestの四Schemaを同一registryで検証するJava 8 test-side方式を実装した。外部依存、Node呼出し、networkなしで、pinned draft 2020-12の`$ref`（localと登録URNのみ）、type / const / enum / object / array / combinator / conditional / contains / URI formatを評価する。未対応語彙、外部schema refはregistry生成時にfail-closedにする。全checked-in positive example、四Schemaの代表negative example、`contract-cases.json`のJSON Schema layer 18 caseのexpected valid/invalidが成功した
  - [x] `ZB-P5.2.3` canonical JSON / SHA-256とcase materializerをtest-sideへ分離し、semantic collectionの正規化はdependency 4-field tuple、resource / assignment / calendar UID、Unicode scalar比較を用いる共通domain-aware canonicalizerへ集約する。`cross-artifact-binding` 13 caseで`RB-001`〜`RB-006`、`RB-011`、`RB-012`を実入力から検証する。`exact-json`、`semantic-state`、明示抽出済み値の`semantic-cross-runtime`、`byte-same-runtime`、`artifact-topology`、`runtime-integrity`を別assertion boundaryとして実装し、Node live outputを呼ばない。`RB-007`〜`RB-010`の実行時I/O・provenance・diagnostic集約はP5-C command runnerで初めて評価する
  - [x] `ZB-P5.2.4` mutationはmaterialize後のtest-side copyだけへ適用し、Java固有goldenやNode live outputをoracleにしない。repository-wide `sh scripts/test-all.sh`は全Java test後にsnapshot verifierを再実行し、immutable原本がtest後も無傷であることを確認する（Node 12 tests、Java 161 tests、failures 0、errors 0、skipped 4）
- [ ] `ZB-P5.3` required command、option、I/O、exit状態、exclusive createとcommit markerによるlogical publicationを実装する
- [ ] `ZB-P5.3.1` `validate`を最初のvertical sliceとして実装し、初回・再reviewのJava側指摘を修正した。再技術reviewと人の承認を待つ。`CU-USAGE-001`の現行Node/Java source挙動は一致したが、frozen Node `v1.0.3` release identityは修正前なのでcross-runtime / Gate G5適合は未達である
  - 2026-08-15にJava側へlegacy CLIと分離したcommand serviceを追加した。strict `validate --project <path|-> [--result <path|->]`、direct regular XML / stdin、UTF-8、MS Project XML subset、semantic validation、structured JSON result / diagnostics、canonical semantic digest、exclusive-create resultを実装した
  - runtime manifest / public launcherはP5-Eのため、verified Java runtime bindingが無ければproject inputを読む前に`runtime.manifest-invalid` / exit 3でfail closedにする。artifact-set directoryもP5-C5まで未対応と明示してrejectする
  - 初回reviewで見つかった不正lexical integer / boolean / resource typeのfalse success、dependency cycle未検出、欠落start milestoneのinternal error、assignment rule ID差分、`CU-USAGE-001`未実行を修正した。再reviewでは`CU-USAGE-001`をwhole CLI argvとして直接実行し、option scope / location、percent欠落とdependency endpointのrule境界、Node `Number.isSafeInteger`範囲、resource / assignment / calendarのUID付きdiagnostic pathを固定した。巨大outline levelはdecode済みmember数を上限にしてallocationをfail-safeにする
  - fixed corpusの`command = cli`、`arguments = ["--unknown-option"]`、`cli.unknown-option`を正本とし、現行Node parserも先頭unknown long optionをoption scope / option location付き`cli.unknown-option`に修正した。Java serviceと現行Node sourceの挙動は一致する。immutable `v1.0.3` snapshotは改変せず、後続Node corrective releaseと新snapshotでreference identityを更新する
  - fixed corpusのvalid / invalid / unsupported / hierarchy、`CU-USAGE-001`、runtime failure、malformed XML、result overwrite拒否を初回・再review回帰とともにJava focused 12 testsで検証した。Nodeのargv / XML adapter / R1 integration関連20 testsと全体`npm test`、`sh scripts/test-all.sh`のNode importer 12、Java 176（failure/error 0、skip 4）、snapshot post-verification 8 testsが成功した。これは人のP5-C1承認またはcorrective release identityを含むcross-runtime適合を代替しない
  - [ ] `ZB-P5.3.2` `inspect project_overview`と`task_change_context`を追加する
  - [ ] `ZB-P5.3.3` `plan-change set_task_percent_complete`とhuman gate bindingを追加する
  - [ ] `ZB-P5.3.4` `apply-change`のexclusive directory reservation、三member、empty `COMMITTED`、unknown outcome回復を追加する
  - [ ] `ZB-P5.3.5` read-only `verify-artifact`とexpected-plan bindingを追加する
- [ ] `ZB-P5.4` JSON/textは必要箇所でexact比較し、XML/XLSXは契約に応じてsemantic比較する
- [ ] `ZB-P5.5` Node-only、Java-only、未対応差分をcapability matrixに記録する
- [ ] `ZB-P5.6` fat JAR、sources、runtime manifest、SHA-256を生成する
- [ ] `ZB-P5.7` standalone `java -jar` smokeをclean directoryで実行する

### 実行順とreview gate

1. `P5-A`: `ZB-P5.1.2`〜`P5.1.4`だけを実装し、snapshot authorityをreviewする。domain commandを変更しない。完了・承認済み（2026-08-14）。
2. `P5-B`: `ZB-P5.2.1`〜`P5.2.4`でharnessを完成し、未実装caseをpass扱いにしない。2026-08-15の再reviewで見つけたwhole-object JSON textによるcollection整列とRB-012 leaf判定を補正した。Node参照実装と同じdependency tuple / UID / Unicode scalar順をdigestと三comparison modeで共有し、summary taskと非summary parentのschema-valid ProjectionをRB-012で拒否するfocused testを追加した。focused test、`sh scripts/test-all.sh`（Node importer 12 tests、Java 164 tests、failure/error 0、skip 4）、snapshot post-verification（8 tests）、両repo diff checkを終えた再技術reviewを完了し、同日に人が承認した。
3. `P5-C`: `ZB-P5.3.1`からcommandを一つずつ実装し、各sliceでshared caseとlegacy Java regressionを通す。P5-C1は初回・再reviewのJava側指摘を修正し、`CU-USAGE-001`の現行Node/Java source挙動も一致済みで再review/承認待ちである。承認されるまでは次commandを開始しない。frozen Node `v1.0.3`との差分はcorrective releaseの明示的な入力とする。
4. `P5-D`: command coverage後に`P5.4/P5.5`のcross-runtime比較と差分承認を行う。Node corrective release、新snapshot、Javaとの再比較が完了するまでrelease identityを含むcross-runtime適合を主張しない。
5. `P5-E`: 最後に`P5.6/P5.7`のartifact、manifest、standalone smokeを実施する。

各review単位で、Java repositoryの`mvn test`、対象focused test、`git diff --check`、snapshot原本不変を確認する。旧subtree更新、命名変更、Java-only extension整理を同じ差分へ混ぜない。

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
- [ ] `ZB-P7.10` `.github/workflows/release-runtime-bundles.yml`をversioned v1 Node runtimeのclean exact-tag build、Gate/Release lock照合、三member＋公開SHA256SUMSのuploadへ対応させる。対応前の`v1.0.3`内部reference tagを新v1 runtimeの公開Releaseとして扱わない

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

`ZB-P3.1`〜`ZB-P3.12`は[CLI contract v1](miku-project-cli-contract-v1.md)、[CLI result and diagnostics contract v1](miku-project-cli-result-contract-v1.md)、[runtime capability contract v1](miku-project-runtime-capability-contract-v1.md)、[runtime manifest contract v1](miku-project-runtime-manifest-contract-v1.md)、[conformance corpus v1](miku-project-conformance-corpus-v1.md)、[human gate and next action contract v1](miku-project-human-gate-and-next-action-contract-v1.md)として2026-08-11に承認された。`ZB-P4.1`〜`ZB-P4.6`は完了し、`ZB-P4.7`の再利用採否、`ZB-P4.8`の階層C1 slice、`ZB-P4.9`のversioned runtime manifest / asset bindingは2026-08-13に承認された。source CLI / development bundleは、承認後もversioned artifactではないためv1 workflowを適合runtimeとして有効化しない。`ZB-P4.10`は三回のreview差戻しを補強し、2026-08-14に承認された。Gate G4ではactual `v1.0.3`を内部reference candidateとして固定し、Git管理する外側lockと再実行可能なverifier、追加後full regressionを完成し、2026-08-14に人が承認した。次は`ZB-P5` Java CLI適合である。

## 計画の更新方法

- 詳細な依存、成果物、完了条件はこの文書で更新する
- 現在着手可能な項目だけを `docs/TODO.md` に置く
- 判断結果は該当するcontract文書へ記録する
- 完了した移行と検証結果は `docs/migration-worklog.md` へ記録する
- scopeを変える場合は、新仕様、計画、TODOを同じ変更で整合させる
- Web/MCPを前倒しする場合は、新仕様の初期対象範囲を明示的に改訂する
