---
title: miku-project Gate G4 readiness v20260814
description: Node参照実装のGate G4について、v1.0.3 source freeze・内部reference candidate・外側lock・再実行可能なconsumer検証の条件別証跡を記録する。
topics:
  - miku-project
  - cli
  - conformance
  - release
category: decision-record
status: approved
audience:
  - maintainer
  - developer
  - agent
created: 2026-08-14
updated: 2026-08-14
---

# miku-project Gate G4 readiness v20260814

## 結論

`ZB-P4.1`〜`ZB-P4.10`の実装・回帰・P4.10承認に加え、2026-08-14にsource freeze commit `693b4ecd7d4328d77f3b2eada9c4965a9c9b15f5`へ軽量tag `v1.0.3`を付与した。clean・exact-tag sourceから生成したactual runtimeは、公開Releaseではない**内部Node reference candidate**として`workplace/gate-g4/v1.0.3/runtime/`へ三memberのまま保持する。

Git管理する外側trust anchorは[miku-project Node reference runtime lock v1.0.3](miku-project-node-reference-runtime-lock-v1.0.3.json)である。lock SHA-256は`95cd11cc4460348fa066908994430adba5983384c06c75679855120e5c5ea3d5`であり、source revision/tag、Node `26.5.0` / npm `11.17.0` / esbuild `0.28.1`、package-lock / corpus digest、manifestを含む三memberのsize / SHA-256を固定する。`npm run verify:cli-v1-release-candidate`はこのlockから期待値を取り、actual candidateの三memberと五workflowを再検証できる。追加実装後のtargeted 25 tests、`npm run build:full`（30 files / 313 tests）、最終`git diff --check`は成功した。全条件の再レビュー後、Gate G4は2026-08-14に人が承認した。

## 条件別の判定

| Gate G4条件 | 判定 | 根拠 |
| --- | --- | --- |
| v1 scenarioがclean directoryで完結する | pass | `tests/mikuproject-cli-v1-runtime-manifest.test.js` のexternal consumer smokeが、source checkout削除・consumer `node_modules`なし・runtime三memberだけで五workflowを実行する。P4.10は2026-08-14承認済み。 |
| 同一入力が契約どおり決定的である | pass | `validate` stdout、C1 plan/provenance、committed `project.xml` / `provenance.json`、`verify-artifact` stdout、runtime asset/source/manifestについてsame-runtime byte determinismをdirect/integration testで固定する。XMLの一般的byte roundtripではなく、契約が指定するbyteまたはsemantic比較を使う。 |
| invalid input / digest mismatch / unsafe overwriteがcommitted artifactを作らず、incompleteを識別できる | pass | `CV-*`、`CP-*`、`CA-*`、`CVF-*`、`CU-UNKNOWN-OUTCOME-001`、`CR-*`とpublisher/verifier testsが、既存destination拒否、runtime failure時domain I/Oなし、markerなしincomplete、`publication.cleanup-failed`、repairなしverifyを固定する。 |
| repository外でsingle `.mjs`が動く | pass | P4.9のversioned runtime testとP4.10のthree-file consumer smokeが、source checkoutとruntime dependencyなしでNode executableを起動する。 |
| relevant tests、contract suite、`npm run build:full`が成功する | pass | lock/verifier追加後のruntime manifest test file全25 tests、`npm run build:full` 30 files / 313 tests、最終`git diff --check`が成功した。 |
| Node contract releaseとartifact manifestが固定される | pass | source freeze commit `693b4ecd…` / exact `v1.0.3`、actual三member、Git管理するGate lock `95cd11cc…a3d5`を固定した。lockを外側pinとしてactual candidateの三member、五result、output plan、provenanceを再検証するcommandも成功した。 |

## 実施記録

1. P4.7〜P4.10とfreeze準備をcommit `693b4ecd…`へまとめ、`package.json`の`1.0.3`と同じ軽量tag `v1.0.3`をそのcommitへ付けた。tagの付け替えはしていない。
2. `npm run build:cli-v1-runtime -- --out-dir <fresh-directory>`を一回実行し、既存runtime directoryを置換せずactual runtimeを生成した。manifestはcanonical JSON / schemaに適合し、executable `miku-project-node-1.0.3.mjs`（2,772,202 bytes、SHA-256 `29afdbce020613e3ba65e2d967886e14aa385c46d3dde4e9bfaca43a4a2b57b8`）とsources `miku-project-node-1.0.3-sources.tgz`（902,788 bytes、SHA-256 `8e804604c2923821aba19c54f4c1831b8b27ce6407147c9090a5f58673b29562`）を束縛した。
3. actual三memberを`workplace/gate-g4/v1.0.3/runtime/`へ保持し、Git管理するcanonical JSON lockへmanifest/executable/sources、source、toolchain、package-lock、corpusを記録した。lock自身のSHA-256も本書へ記録する。
4. `scripts/verify-cli-v1-release-candidate.mjs`はlockをruntime directory外のtrust anchorとして先に検証する。候補directoryが三memberだけであること、各entryが通常file / 非symlinkでsize / SHA-256一致であることを確認後、隔離consumerへcopyして`validate`、`inspect`、`plan-change`、`apply-change`、`verify-artifact`を実行する。五result、output plan、artifact provenanceのruntime bindingを完全一致で照合する。
5. verifierの五workflow成功と、不整合lockをruntime launch 0回で拒否する二caseを既存runtime manifest testへ直列追加した。別test fileで重いsource archive buildを並列化するとtimeoutしたため、同じruntime build系列へ統合してfull suiteの安定性を守る。
6. verifier追加後のruntime manifest test file全25 tests、`npm run build:full`（30 test files / 313 tests）、最終`git diff --check`は成功した。全条件の再レビュー後、Gate G4は2026-08-14に人が承認した。

## 境界

- Gate G4はNode参照実装と`distribution_status = internal-reference-only`のcandidate / lockを承認する。公開Releaseの発行ではない。
- Java CLI（G5）、Agent Skills（G6）、互換性/改名/release移行（G7）、Web/MCP（G8）は開始・完了扱いにしない。
- 現行`.github/workflows/release-runtime-bundles.yml`はこのv1三memberをbuild / uploadしない。`v1.0.3`を現行workflowから新v1 runtimeの公開Releaseとして扱わず、公開経路は`ZB-P7.10`で整備する。
- Gate lockはGit履歴をtrust rootにする内部承認用pinであり、実運用のRelease checksumまたはSkills lockではない。TOCTOU対策、配布directory権限設計も後続scopeに残す。
