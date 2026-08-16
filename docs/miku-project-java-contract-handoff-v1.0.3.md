---
title: miku-project Java contract handoff v1.0.3
description: Gate G4で固定したNode参照契約を、moving subtreeではなくimmutable snapshotとしてJava CLIへ渡すための実装入力。
topics:
  - miku-project
  - java
  - cli
  - conformance
  - migration
category: implementation-input
status: ready
audience:
  - maintainer
  - developer
  - agent
created: 2026-08-14
updated: 2026-08-16
---

# miku-project Java contract handoff v1.0.3

## 決定

Java v1は、Node repositoryのmoving subtreeや現在の`devel`を仕様入力にしない。Gate G4で承認した`v1.0.3`のcontract snapshotをimmutableな入力として取り込み、そのsnapshotに対して適合させる。

Nodeは参照実装だが、Nodeのlive出力だけをJavaのoracleにしない。承認済み契約、JSON Schema、fixture、golden、case indexが正本であり、Node/Java差分は契約の比較modeに従って判定する。

## 固定identity

| 項目 | 値 |
| --- | --- |
| contract version | `1` |
| fixture suite version | `1` |
| Node reference release | `1.0.3` |
| source repository | `https://github.com/igapyon/miku-project` |
| source tag | `v1.0.3` |
| source revision | `693b4ecd7d4328d77f3b2eada9c4965a9c9b15f5` |
| conformance corpus SHA-256 | `f0b4a821f80f155ba01afc1fdf7c5d2fa6ca5e744bbd0090819e0d591f1c47a3` |
| Gate runtime lock | `docs/miku-project-node-reference-runtime-lock-v1.0.3.json` |
| Gate runtime lock SHA-256 | `95cd11cc4460348fa066908994430adba5983384c06c75679855120e5c5ea3d5` |

Gate lockはNode runtime candidateの内部承認用pinであり、Java artifactのlockではない。ただしsource identity、contract release、fixture suite、capability profileをJava handoffと同じ値へ束縛する。

## snapshot allowlist

Java repositoryへ取り込むsnapshotは、`v1.0.3`の次のpathだけで構成する。

### 契約文書

- `docs/miku-project-semantic-contract-v1.md`
- `docs/miku-project-semantic-fixture-catalog-v1.md`
- `docs/miku-project-format-and-loss-contract-v1.md`
- `docs/miku-project-change-contract-v1.md`
- `docs/miku-project-cli-contract-v1.md`
- `docs/miku-project-cli-result-contract-v1.md`
- `docs/miku-project-runtime-capability-contract-v1.md`
- `docs/miku-project-runtime-manifest-contract-v1.md`
- `docs/miku-project-conformance-corpus-v1.md`
- `docs/miku-project-human-gate-and-next-action-contract-v1.md`

### Machine-readable入力

- `docs/schemas/*.schema.json`
- `docs/examples/artifacts-v1/*.json`
- `docs/examples/cli-v1/*.json`
- `docs/examples/runtime-manifest-v1/*.json`
- `testdata/conformance/v1/**`

Nodeのsource code、generated JavaScript validator、legacy `testdata/`、Web/browser artifact、report fixtureはsnapshotへ入れない。Java実装が必要とする意味をsnapshot外のNode sourceから読み取った場合も、それを新しい契約として扱ってはならない。

## Java repositoryでの配置

Java側では次のimmutable treeを新設する。

```text
vendor/miku-project-contract/v1.0.3/
├── SOURCE.json
├── docs/
│   ├── miku-project-*-contract-v1.md
│   ├── schemas/
│   └── examples/
└── testdata/conformance/v1/
```

既存`vendor/mikuproject/`は旧系列の互換・移植履歴として残すが、v1 command、result、diagnostic、runtime manifest、conformanceのauthorityから外す。P5では削除、subtree更新、directory renameを同時に行わない。

`SOURCE.json`はcanonical JSONで、少なくとも次を記録する。

- `kind = miku_project_contract_snapshot`
- `schema_version = 1`
- `contract_version = 1`
- source repository / full revision / exact tag
- fixture suite version / corpus digest
- Node reference release / Gate lock digest
- snapshot memberごとのrelative path、byte size、raw SHA-256

memberはpath昇順とし、absolute path、`..`、symlink、allowlist外memberを拒否する。snapshot directoryが既に存在する場合は置換せず、同一inventoryならno-op、異なる場合はhard errorにする。

## import手順

1. Java repositoryのbranchとworking treeを確認し、既存変更がないことを確認する。
2. Node sourceの`v1.0.3^{commit}`が上記full revisionと一致することを確認する。
3. 確認済みfull revisionからallowlistだけをfresh temporary directoryへ抽出する。tagはidentity確認だけに使い、確認後の`ls-tree` / `show`でtag名を再読しない。current checkoutのfileをcopyしない。
4. 全memberがregular fileかつnon-symlinkであることを確認し、size / SHA-256 inventoryを作る。
5. canonical `SOURCE.json`を作り、freshな`vendor/miku-project-contract/v1.0.3/`へexclusiveにinstallする。既にある場合は同一inventoryだけを無変更no-opとして受け入れる。
6. Java testからinventory、corpus digest、schema/example/fixtureの存在を検証する。
7. READMEとupstream docsで、v1のauthorityをこのsnapshotへ切り替える。旧subtreeはlegacy authorityと明記する。

この手順はJava repository内の固定scriptにし、手作業の`cp`一覧を運用手順にしない。networkを必要とするのはsnapshot sourceを取得するときだけとし、通常の`mvn test`はsnapshotだけで完結させる。

## P5.1受入条件

- Java v1の入力がfull revision、exact tag、member inventoryで一意に固定されている。
- Java docsがmoving subtree追随をv1の開発規則として要求しない。
- legacy `vendor/mikuproject/`を更新または削除していない。
- snapshot改変、member欠落、余分member、symlink、digest不一致をJava testが実装起動前に拒否する。
- snapshot更新は既存`v1.0.3`の上書きではなく、新しいsource release directoryとして行う。

## P5.2受入条件

- Java conformance harnessがsnapshotの`suite-index.json`と`contract-cases.json`を直接読む。
- fixture / goldenをJava test固有copyへ再複製しない。
- result / diagnostic / artifact / runtime manifestの四Schemaを同じregistryとして検証する。
- `exact-json`、`semantic-state`、`semantic-cross-runtime`、`byte-same-runtime`、`artifact-topology`、`runtime-integrity`を別assertionとして実装する。
- Nodeのlive出力をgolden更新へ使用せず、差分発生時はcontract、golden、Node defect、Java defectのどれかを明示する。

## 最初のJava実装単位

最初の変更はdomain commandの移植ではなく、次の順で行う。

1. immutable snapshot importer / `SOURCE.json` generator
2. snapshot inventory verifier test
3. suite index loaderとcase selection test
4. Schema registryの技術選定とpositive / negative example test
5. `validate` commandの一つ目のvertical slice

`inspect`、`plan-change`、`apply-change`、`verify-artifact`は、この基盤がreviewを通った後に一commandずつ追加する。fat JAR、Java runtime manifest、standalone smokeはcommand適合後の`ZB-P5.6/P5.7`で扱う。

## 開始時baseline

2026-08-14時点の姉妹repositoryはJava `0.12.0`であり、`docs/upstream-snapshot.md`は旧Node package `0.8.0` / revision `245deaa99d6d2ba970969a9359ce003386da3472`を互換sourceとしている。これはlegacy実装の出発点であり、v1 contract snapshotではない。

同日に変更前の`mvn test`を実行し、132 tests、failures 0、errors 0、skipped 4で成功した。skipされた4件はopt-inのNode parity testである。この結果はlegacy Java baselineがgreenであることだけを示し、Gate G5 conformanceの成功として数えない。

## P5-A実装記録

2026-08-14にJava repositoryへ`v1.0.3`の49-member snapshot、tagでidentityを確認後にfull revisionだけを読むimporter、canonical `SOURCE.json`、test-side snapshot verifierを追加した。importerはworking tree copyやlegacy subtreeの更新を行わず、tagが確認後に動いても別objectを読まない。existing snapshotがexpected manifest、全regular file、member pathから導出したdirectory topologyまで完全一致する場合だけ`"snapshot_action":"unchanged"`で無変更no-opにする。欠落・余分・改変・symlinkを含む既存destinationは`contract-snapshot.destination-conflict`でhard errorにし、exclusive directory creationで既存treeを上書きしない。repository外またはsymlink親への出力も拒否する。destination確保後のmarker、member、manifest、検証、marker削除の失敗はownership tokenと既知memberを照合し、既知regular fileだけを個別unlinkした後に既知directoryだけを深い順のnon-recursive `rmdir`でcleanupする。未知file / directoryを発見した場合は削除せずfail-closedにする。verifierもsource identity、corpus digest、member set、exact directory topology、`SOURCE.json` raw SHA-256外側pin、regular/non-symlink、size、SHA-256を確認し、extra / missing / tampered / symlink / fractional-size / rewritten-manifest mutationを拒否する。

synthetic Git repositoryを用いるNode importer test 12件は、exact directory topology、tag移動後のfull revision読込、marker / member / manifest書込失敗のcleanup、未知file / empty directory保全、inode / mtime / ctimeを含む無変更no-opを確認する。focused verifier 8 testsと、repository-wide `sh scripts/test-all.sh`（Node 12 tests + Java 140 tests、failures 0、errors 0、skipped 4）が成功した。Release workflowもこの全テスト入口をpackage前に実行する。これはP5-Bのloader / Schema registryやv1 commandの適合を示すものではなく、immutable input boundaryが成立した証拠である。

## P5-B1実装記録

2026-08-14にJava test-side harnessへ`ConformanceSuiteLoader`を追加した。public loaderはまずP5-Aのsnapshot verifierを通し、合格したimmutable snapshotからのみ`suite-index.json`と`contract-cases.json`を直接読む。root / case / known parameter / input / mutationの形をstrictに検査し、workflowとcontract indexを通じたcase ID重複、unknown field、snapshot外escape、allowlist外・symlink参照、未知input role、不正RFC 6901 pointerをfail-closedにする。P5-B1はindexの取得と構造検査だけであり、Schema評価、mutation実行、v1 domain commandの実装は含まない。focused loader 8 testsとrepository-wide `sh scripts/test-all.sh`（Node 12 tests + Java 148 tests、failures 0、errors 0、skipped 4）が成功した。

## P5-B2実装記録

2026-08-14に`ConformanceSchemaRegistry`を追加し、artifact、CLI diagnostic、CLI result、runtime manifestの四Schemaを同一registryに登録した。これはJava 8 test-sideのclosed evaluatorであり、外部dependency、Node呼出し、networkを用いない。pinned draft 2020-12に実際に現れる語彙とlocal / 登録URN参照だけを許可し、未対応語彙または登録外参照はregistry生成時に拒否する。四Schemaの全checked-in positive exampleと代表negative example、`contract-cases.json`のJSON Schema layer 18 mutation caseのexpected valid/invalidを検証した。続くP5-B3で`cross-artifact-binding` 13 caseとruntime / semantic / byte / topology comparisonを実装した。

## P5-B3実装記録

2026-08-14にtest-sideのcanonical JSON / SHA-256 serializerとcase materializerを追加し、pretty-print、object insertion order、historical JSON utilityの出力にdigestを依存させないようにした。serializerはUnicode code point順のkey、指定escape、integer-only number、unpaired surrogate拒否を実装し、snapshotのsemantic state、change request、semantic diff、output planの固定digestを直接照合する。binding validatorは各inputを四Schema registryで先に検証し、schema不適合ならbindingを評価しない。`contract-cases.json`の13件は`RB-001`〜`RB-006`、`RB-011`、`RB-012`を通じて正例と改変拒否を確認する。

comparison assertionは`exact-json`、`semantic-state`、callerが明示抽出したruntime-independent値の`semantic-cross-runtime`、`byte-same-runtime`、`artifact-topology`、`runtime-integrity`を別々に持つ。artifact topologyは三member、non-symlink regular file、empty commit marker、provenance schema、raw digest、descriptorを、runtime integrityは外側manifest digest pin、BOM、不正UTF-8 / duplicate key、schema、snapshot corpus digest、core capability、固定basename、size/digest、member setを検証する。process起動、project未読/destination未作成観測、actual provenance、diagnostic aggregationを要する`RB-007`〜`RB-010`は、未実装commandを仮想的にpassさせずP5-C runnerへ残す。

mutationはmaterializeしたtest-side copyへだけ適用し、snapshot原本やJava固有goldenを書き換えない。`sh scripts/test-all.sh`はNode importer 12 tests、全Java 161 tests（failures 0、errors 0、skipped 4）の後に`ContractSnapshotVerifierTest`を再実行し、test後のimmutable inventoryも確認する。P5-Bは実装完了であり、P5-C開始前にこのharness自体の最終reviewを行う。

P5-Bの技術reviewは2026-08-14に完了した。snapshot authority、loader/registryのpublic entrypoint、13 binding case、18 Schema case、semantic collection canonicalization、runtime manifestのstrict intake、comparison assertionの責務境界、test後snapshot verifierを再確認し、Node importer 12 tests、Java 161 tests（failures 0、errors 0、skipped 4）、両repositoryの`git diff --check`が成功した。これはP5-Bの人による承認を代替しない。`validate`を含むP5-C domain commandは承認後にのみ開始する。

## P5-B4再review補正（2026-08-15）

上記の初回技術review後、二つの不足を見つけたため、P5-B4を承認待ちではなく再技術reviewへ戻した。Java側はsemantic collectionをmember全体のcanonical JSON text順で整列していたが、v1 contractはdependencyを`(predecessor_uid, successor_uid, type, lag)`、resource / assignment / calendarを`uid`で整列し、比較はUnicode scalar順で行う。optional fieldが逆順になる複数memberでは両規則が異なる。`ConformanceCanonicalJson`のdomain-aware canonicalizerをsemantic digestと`exact-json` / `semantic-state` / `semantic-cross-runtime`で共有し、Node referenceで得たcanonical digest、task順保持、tuple / UID順、Unicode scalar順をfocused testで固定する。

またRB-012の`task_change_context` validatorはtargetが存在しcontentが一致することを確認していたが、`summary = false`かつdirect childなしのleaf制約を確認していなかった。schema-validなsummary task Projectionと、`summary = false`だがchildを持つProjectionを、同じRB-012 violationとして拒否するtestを追加した。contract snapshotは変更しない。focused test、`sh scripts/test-all.sh`（Node importer 12 tests、Java 164 tests、failures 0、errors 0、skipped 4）、その後のsnapshot post-verification（8 tests）、両repositoryの`git diff --check`を2026-08-15に通過し、再技術reviewを完了した。P5-C domain commandは人によるP5-B承認後にだけ開始する。

P5-Bは2026-08-15に人が承認した。承認範囲はimmutable snapshotを入力とするtest-side conformance harnessまでであり、次に開始できる製品実装は`ZB-P5.3.1`の`validate` vertical sliceだけとする。`inspect`以降のcommand、fat JAR、runtime manifest、release、命名変更、Web / MCPはこの承認に含めない。

## P5-C1実装記録（初回・再review指摘を修正、再技術review・人の承認待ち）

2026-08-15に、legacy `MikuprojectCli`とは別の`MikuProjectV1Validate` command serviceを追加した。この境界だけがstrictな`validate --project <path|-> [--result <path|->]`を受理し、legacy grammarやsilent overwriteを再利用しない。direct regular XMLまたはstdinをUTF-8として読み、先頭BOMだけをnormalizationとして記録する。MS Project XML subsetからsemantic stateを作り、v1 invariantを評価し、Unicode scalar順のcanonical JSONとdomain指定順（dependency tuple / UID collection）でstate SHA-256を作る。成功・rejected・usage-error・runtime-errorの構造化result / diagnostic / exit codeは固定schema registryで検証する。

このsliceはruntime manifestや公開JAR launcherを先取りしない。callerがartifact SHA-256とmanifest SHA-256を検証済みの`VerifiedRuntime`を供給できない場合、serviceはproject inputに触れず`runtime.manifest-invalid` / exit 3で停止する。これはP5-Eまでverified bindingを偽装しないための境界である。`--result`は既存pathを上書きせずexclusive createし、P5-C1のproject inputはdirect file / stdinだけとする。artifact-set directory、`verify-artifact`、`inspect`、`plan-change`、`apply-change`、fat JAR、manifest、releaseは後続scopeである。

初回技術reviewで、不正lexical integer / boolean / resource typeを欠落として受理するfalse success、dependency cycleの未検出、欠落startを持つmilestoneの`internal.unexpected-error`、assignment参照の安定rule ID差分、`CU-USAGE-001`を実行していない不足が見つかった。adapterは不正lexical値をsemantic validationへ保持する形に直し、cycle検出`S-I016`、totalなmilestone検証、assignmentの`S-I008` / `S-I017`を追加した。

再reviewでは、`CU-USAGE-001`をservice prefixなしのwhole CLI argvとして直接実行し、unknown optionの`location.scope = option`と`location.option`を固定した。percentの欠落を`S-I008`、欠損dependency endpointを`S-I008`、有効な字句だが既存しないendpointを`S-I014`へ分離した。Javaの数値変換はNodeの`Number.isSafeInteger`と同じ`0`〜`2^53-1`まで受け、unsafe valueをsemantic validationへ残す。resource / assignment / calendarのunknown fieldとcalendar参照はUID付きdiagnostic pathを用いる。malformed outlineの処理は、既にdecode済みのmember数を上限にして不正な巨大levelによるallocationを防ぐ。

固定corpusの`CU-USAGE-001`は`command = cli`と`arguments = ["--unknown-option"]`に`cli.unknown-option`を要求する。Java serviceをcorpusへ適合させた後、現行Node `parseV1Invocation`も先頭unknown long optionをoption scope / option location付き`cli.unknown-option`に修正し、source挙動を一致させた。frozen Node `v1.0.3` tagはこの修正前であるため、immutable snapshotを改変せず、後続のNode corrective releaseと新snapshotでreference identityを更新する。P5-D / Gate G5ではそのrelease identityを用いて再比較する。

focused testはfixed v1.0.3 corpusの`CV-VALID-001`、hierarchy valid、invalid、unsupported、hierarchy invalid、`CU-USAGE-001`に加え、usageでのproject未読、unverified runtimeでのproject未読、malformed XMLのschema-valid rejected result、`--result` overwrite拒否、初回・再review回帰を確認する（12 tests）。Node側のargv / XML adapter / R1 integration関連20 testsと全体`npm test`、Javaの`sh scripts/test-all.sh`内Node importer 12 tests、Java 176 tests（failures / errors 0、skipped 4）、snapshot post-verification 8 testsが成功した。ここでの完了は再技術reviewを可能にする実装完了を示すだけであり、P5-C1の技術review、人による承認、corrective release identityを含むcross-runtime適合の主張は残る。

## 対象外

- 公開Node Release、公開checksum、署名
- Java package / repository / legacy wire IDのrename
- Java-only extensionの削除
- Agent Skills runtime統合
- Web / MCP

これらはGate G5の成立に必要な作業へ混ぜない。
