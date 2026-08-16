---
title: miku-project v1 capability matrix v20260813
description: P4.8 hierarchy C1 slice、P4.9 runtime asset binding、P4.10 external pinned-consumer smokeとGate G4 source freeze時点の、承認済みv1契約に対するNode reference実装範囲とdefer範囲。
topics:
  - miku-project
  - cli
  - conformance
  - capability
category: reference
status: approved
audience:
  - maintainer
  - developer
  - agent
created: 2026-08-13
updated: 2026-08-14
---

# miku-project v1 capability matrix v20260813

## 目的とauthority

これはゼロベース仕様のv1契約に対して、**Node reference implementationで実証済みの範囲**と、意図的に未実装の範囲を分けて示すP4.8 / P4.9 / P4.10の記録である。将来機能の約束や現行Web/旧coreのcapability表ではない。

| 文書 | 役割 |
| --- | --- |
| [semantic contract v1](miku-project-semantic-contract-v1.md) | 意味・不変条件の正本 |
| [CLI result contract v1](miku-project-cli-result-contract-v1.md) | command/result/diagnostic/effectの正本 |
| [conformance corpus v1](miku-project-conformance-corpus-v1.md) | fixture、golden、比較規則の正本 |
| 本書 | P4.8 / P4.9 / P4.10時点の実装範囲とdeferの記録 |

source CLIとdevelopment bundleは、manifestを持つversioned runtime artifactではないため、引き続きv1 workflowをfail-closedにする。P4.9で生成するversioned Node runtimeだけが、adjacent manifest・executable・sources・embedded corpus digestを自己検証した後にv1 workflowを実行できる。P4.10はtest-owned consumerが外側manifest pinを照合した後、consumer自身もexecutable / sourcesのentry・size・digestをNode起動前に検証するよう再補強した。第3回reviewの補正で、consumerはCLI result envelope全体をschema検証し、`result.runtime`をnested digestまでverified preflight bindingと完全比較する。P4.10 consumer対象13 cases、補正後full regression、`git diff --check`を確認し、2026-08-14に再reviewと人の承認を得た。公開Release checksumと公開Releaseは後続である。

Gate G4のsource freezeでは、commit `693b4ecd7d4328d77f3b2eada9c4965a9c9b15f5`にexact `v1.0.3` tagを付け、actual runtimeを生成した。manifest SHA-256は`e577c11e8a9f5cc4c09ffe458cb7597c42f5450cf7d65fcde15cf500863db0b2`、executable SHA-256は`29afdbce020613e3ba65e2d967886e14aa385c46d3dde4e9bfaca43a4a2b57b8`、sources SHA-256は`8e804604c2923821aba19c54f4c1831b8b27ce6407147c9090a5f58673b29562`である。actual三memberは内部reference candidateとして`workplace/gate-g4/v1.0.3/runtime/`へ保持し、Git管理する[miku-project Node reference runtime lock v1.0.3](miku-project-node-reference-runtime-lock-v1.0.3.json)がsource/tag、build toolchain、package-lock / corpus、三memberを固定する。`verify:cli-v1-release-candidate`でouter lock、asset、五workflow、runtime/result/output plan/provenance bindingを再検証できる。追加後full regressionも成功し、Gate G4は2026-08-14に人が承認した。公開Release checksum、Skills lock、公開Releaseは含めない。

## Node reference capability

| 能力 | status | conformance evidence | 境界 |
| --- | --- | --- | --- |
| XML subsetのread / semantic validate | implemented | `S-V001`、`S-V002`、`CV-HIERARCHY-INVALID-PREORDER-001`、`CV-HIERARCHY-INVALID-SUMMARY-001`、`S-I012`、`S-I020` | 明示したMS Project XML subsetのみ。unsupported dataはfail closed |
| ordered forest / summary整合性 | implemented | `S-V002`、`S-I003`、`S-I004` | task preorderとparent relationはsemantic stateの意味。疑似taskはsemantic taskにしない |
| `inspect project_overview` | implemented | `CI-OVERVIEW-001`、`CI-HIERARCHY-OVERVIEW-001` | purpose追加なし。全taskのpreorder、parent、summaryをread-onlyで返す |
| `inspect task_change_context` | implemented | `CI-CONTEXT-001`、`CI-HIERARCHY-CONTEXT-001` | leaf taskだけ。nested leafにはancestorを必要十分に含める。summary / nonexistent targetはreject |
| leaf taskの進捗 C1 | implemented | `CP-CHANGE-001`、`CA-CHANGE-001`、`CP-HIERARCHY-CHANGE-001`、`CA-HIERARCHY-CHANGE-001`、`CP-HIERARCHY-SUMMARY-REJECT-001` | `set_task_percent_complete`のみ。base digest、precondition、explicit approval、pre/post validationを必須とする |
| canonical XML encode / re-decode | implemented | `S-V001`、`S-V002`、flat / hierarchy C1 after golden | task preorderと非変更collection意味をsemantic equivalenceで確認。外部XML bytesの一般的roundtripは約束しない |
| committed artifact set / verify | implemented | `CA-*`、`CVF-*`、`CU-UNKNOWN-OUTCOME-001` | create-new directoryと`COMMITTED` markerだけを成功として公開。verifyはrepairしない |
| versioned Node runtime asset binding | Gate G4 approved (2026-08-14) | `CR-MANIFEST-INVALID-001`、`CR-ASSET-DIGEST-001`、`CR-CAPABILITY-MISSING-001`、`CR-SOURCE-MISSING-001`、五workflowのverified binding test、Gate lock verifier | clean・exact tag sourceだけから三memberを作る。各workflow前の自己検証に加え、Gate lockから外側pinしてactual candidateを再検証できる |

## 意図的なdefer

| 能力 | status | 再評価gate | 理由 |
| --- | --- | --- | --- |
| resource / assignment / calendarの編集 | deferred | change contract改訂後 | P4.8ではnested leafの進捗一件だけを許可し、参照の保持を検証する |
| dependency編集、task追加/削除/移動、summary編集 | deferred | change contract改訂後 | 新operation、diff、影響範囲、human gateの承認が必要 |
| workbook / XLSX | deferred | format scope承認後 | v1 XML subset / semantic stateの正本性を先に守る |
| new external format、Projection purpose | deferred | G4以降の契約改訂 | command / artifact schema / goldenを先に固定しない |
| actual、baseline、EVM、timephased、extended data | deferred / unsupported | semantic scope拡張後 | 現v1では黙って保持・破棄しない |
| Web、MCP | deferred | ZB-P8 | CLI + Skillsの安定後にadapterとして再評価する |
| Java runtime、Agent Skills runtime | not started | ZB-P5 / ZB-P6 | Node reference contractとruntime manifestを先に固定する |
| source CLI / development bundleのv1 workflow | intentionally fail-closed | ZB-P4.9 implementation | manifest探索、glob、PATH、別version fallbackをしない。versioned artifact以外は`runtime.capability-missing` |
| public Release asset / external provenance | not released | ZB-P7.10 | Gate lockはGit履歴をtrust rootにする内部承認用である。現行Release workflowはv1三memberをbuild / uploadせず、公開Release checksum、Skills lock、公開Releaseは未完 |

## P4.8の不変条件

- hierarchy sliceは新format、新command、新Projection purpose、新change operationを追加しない。
- C1はsummary taskを編集せず、対象leaf以外のtask field、parent relation、task order、dependency、resource、assignment、calendarを変更しない。
- legacy `ProjectModel`、旧Patch / AI view、Web runtimeをv1 semantic pathのactive dependencyにしない。
- source CLIとdevelopment bundleは引き続きfail closedとし、versioned runtimeだけをmanifest verifiedで有効化する。
- artifact内部の自己整合性と、Release checksum / Skills lockによる配布元の外側trust anchorを混同しない。

## 実装・回帰の証拠

2026-08-13に、hierarchy direct tests、fixed verified-binding subprocess integration、既存flat R1/C1、legacy XML / core API / CLI / workbook/XLSX回帰、bundle/source archiveを含む`npm run build:full`を実行し、29 test files・288 testsが成功した。`git diff --check`も成功した。

P4.9では、clean / exact-tag release preflight、fresh runtime directory、single `.mjs` / sources / canonical manifestのbyte determinism、manifest schema・filename/version・artifact/source digest・corpus digestのbinding、五workflowのverified runtime resultを専用testで確認した。clean・exact-tagのtemporary source repositoryからrelease builderを実行し、生成runtimeをrepository外のworking directoryでverified `validate`として動かす成功経路も確認している。`CR-MANIFEST-INVALID-001`、`CR-ASSET-DIGEST-001`、`CR-CAPABILITY-MISSING-001`、`CR-SOURCE-MISSING-001`に加えsource digest不一致も、input / result path / destinationに触れず`runtime-error`で拒否する。failure pathはdirect filesystem guardでproject / result pathへの`lstat`・`realpath`・`readFile`がないことまで検証する。2026-08-13に`npm run build:full`を実行し、30 test files・297 testsが成功し、P4.9.5は人が承認した。**P4.9承認時点では**P4.10のexternal consumer smokeは未完だった。

P4.10の初回実装では、P4.9のclean / exact-tag temporary sourceから生成したruntimeの三memberだけをexternal consumer directoryへコピーし、source checkoutを削除してからconsumer側のproject / request / plan / approvalで五workflowを実行した。consumerに`node_modules`は置かず、すべてのresultが配布manifestに一致する`verified` bindingを返した。一方、期待manifest digestを配布manifest自身から導出していたため、coordinated tamperを拒否する外側trust anchorになっていなかった。2026-08-13のtargeted smoke（10 tests）と`npm run build:full`（30 test files・298 tests）、`git diff --check`は補強前baselineとし、P4.10の承認証拠には数えない。

外側pinの**補強途中時点では**、copy前に`buildReleaseNodeRuntime`が返すraw manifest SHA-256をtest-owned trust anchorとして保持した。consumer preflightはfixed `runtime-manifest.json`のregular-file / non-symlink確認とraw digest pinを先に行い、pin成功後だけstrict JSON、schema、canonical JSON、Node reference identity、capability、corpus、固定basenameを検証する。launcherにはmanifestの`runtime.launcher = node`と`artifacts.executable.path`だけを渡し、Node executableは固定の`process.execPath`を用いる。五workflowはoperationごとにpreflightを通し、CLI result、output plan、artifact provenanceのruntime bindingをpin済みbindingと照合する。executableのcomment追加とmanifest側size / digest更新を組み合わせたcoordinated tamper、missing / malformed anchor、manifest missing / symlinkは、launcher callback 0回かつdomain I/O / result / artifact setなしで拒否した。2026-08-13の`npm run build:full`（30 files / 303 tests）と`git diff --check`は、当時まだ人のreview・承認前のbaselineであり、公開Release checksum、署名、Skills lock、公開Releaseを代替しない。

この記録のP4.8部分はP4.8.6、P4.9部分はP4.9.5で2026-08-13に承認された。P4.10 external pinned-consumer smokeは第2回review差戻しの起動前asset再補強と、第3回review差戻しのCLI result runtime binding完全一致補正を実装し、補正後のP4.10 consumer対象13 cases、full regression、`git diff --check`を確認して2026-08-14に承認された。JavaはGate G4承認後、SkillsはGate G5後、Web / MCPはGate G7後に開始判定する。公開Releaseは`ZB-P7.10`まで未完である。

起動前asset再補強後のP4.10 target 12 tests、`npm run build:full`（30 files / 309 tests）、`git diff --check`は完全一致補正前baselineとして保持する。補正後はP4.10 consumer対象13 cases（runtime manifest test file全22 tests）、`npm run build:full`（30 files / 310 tests）、`git diff --check`が成功し、2026-08-14に承認された。

Gate G4補強ではactual `v1.0.3` candidateをローカル保持し、Git管理するGate lockと再実行可能なverifierを追加した。lockはbuild toolchainとpackage-lockも記録し、runtime directory自身から期待manifest digestを導出しない。verifierの成功五workflow、lockとmanifestのsource revision不一致、および不整合lockをlaunch 0回で拒否するcaseを追加後、runtime manifest test file全25 testsと`npm run build:full`（30 files / 313 tests）が成功した。全条件の再レビュー後、Gate G4は2026-08-14に人が承認した。
