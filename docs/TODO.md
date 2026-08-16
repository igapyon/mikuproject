# TODO

この文書には、現在着手可能な未完了作業と、再評価待ちの既存backlogだけを書く。新しい製品像は `docs/miku-project-zero-base-spec-v20260809.md`、詳細な順序と完了条件は `docs/miku-project-zero-base-implementation-plan-v20260810.md`、現行実装の仕様は `docs/spec.md` を参照する。

## ゼロベース再設計: G0 完了

- [x] `ZB-P0.1` 新仕様、実施計画、TODO、現行仕様、worklogの文書authorityを確認する
- [x] `ZB-P0.2` 現行CLI、core API、形式、diagnostics、testsのcapability matrixを作る
- [x] `ZB-P0.3` 現行module、fixture、test、artifactを `reuse / evidence / rewrite / defer / drop` に仮分類する
- [x] `ZB-P0.4` R1/C1を、v1で証明するend-to-end scenarioのG0提案として具体的な入力・出力例で書く
- [x] `ZB-P0.5` scenarioごとにactor、human gate、成功、失敗、非目標を定義する
- [x] `ZB-P0.6` 下記の現行backlogを選択scenarioとの関係でtriageする
- [x] `ZB-P0.7` AI Agentフレンドリーを主要要件として、同じ操作契約を人、shell、CI、Agentが利用できる受入観点を定義する
- [x] `ZB-P0.8` scenario内の操作を `read-only / artifact生成 / 意味変更` に仮分類し、構造化結果、次の行動、human gateを対応づける
- [x] `Gate G0` 一文の製品定義、primary job、v1 scenario 1〜2個、非目標、AI Agentフレンドリーの受入観点を承認する（2026-08-10）

G0の承認内容は、[v1利用シナリオ](miku-project-zero-base-scenarios-v1.md)、[現行capability matrix](miku-project-current-capability-matrix-v20260810.md)、[再利用資産棚卸し](miku-project-zero-base-reuse-inventory-v20260810.md)に記録する。詳細は [miku-project-zero-base-implementation-plan-v20260810.md](miku-project-zero-base-implementation-plan-v20260810.md) を参照する。

## ゼロベース再設計: G1 完了

`ZB-P1.1`〜`ZB-P1.8`と`Gate G1`は完了した。製品コードは変更していない。成果物は [semantic contract v1](miku-project-semantic-contract-v1.md)、[semantic fixture catalog v1](miku-project-semantic-fixture-catalog-v1.md)、[実施計画](miku-project-zero-base-implementation-plan-v20260810.md) である。

- [x] `ZB-P1.1` ordered forest、summary、identityを確定する
  - `docs/miku-project-semantic-contract-v1.md` の「identity、順序、階層」を編集する
  - 「順序付きの木」を「複数rootを許すordered forest」へ変更する
  - rootは`parent = null`、childは一つのparentを持ち、sibling orderは入力順として保持すると書く
  - taskを持たない空のordered forestはR1でvalidとする
  - `summary = true`と「子を一つ以上持つ」を同値にし、不一致をinvalidとする
  - task UIDは入力artifactから次artifactを生成する一回の処理単位でstableとし、名称、行番号、outline numberをselectorにしない
  - MS Project XMLのUID `0`など外部形式の疑似taskはsemantic taskに含めず、mapping詳細はG2へ送る
  - 完了条件: 複数rootを持つ`testdata/dependency.xml`がvalidになり、単一tree前提やraw `level 1`をsemantic invariantとして残していない

- [x] `ZB-P1.2` v1 domain scope tableをfield単位へ展開する
  - `required / optional-preserved / unsupported` の三分類を定義する
  - semantic typeとしてidentity token、non-empty text、local civil datetime、boolean、working duration、units、resource type、sibling orderを定義し、外部形式の字句表現はG2へ送る
  - `required`に分類する: projectのname/start/finish、taskのUID/name/parent/order/start/finish/duration/milestone/summary/percentComplete、dependencyの両端UID/type/lag、resource UID、assignment UID/task UID、存在するcalendar entityのUID
  - `optional-preserved`に分類する: project currentDate/scheduleFromStart/calendar参照、task calendar参照、resource name/type/calendar参照、assignment resource UID/start/finish/units/work、calendar name/isBaseCalendar
  - `unsupported`に分類する: actual、EV、baseline、timephased、extended data、および明示していないunknown field
  - task/resourceの外部ID、outline level/number、外部sentinelはadapter表現でありsemantic fieldではないと明記する
  - 完了条件: 代表fixtureに現れるproject/task/dependency/resource/assignment/calendarの各fieldが一つの分類に入り、未分類fieldがない

- [x] `ZB-P1.3` 日時、duration、milestone、summary、進捗の境界を確定する
  - projectとtaskのstartはfinish以前とする
  - datetimeはtimezone変換しないlocal civil datetimeとし、文字列表現はG2へ送る
  - `required`に分類したstart/finish/durationの欠落はinvalidとする。project currentDateとassignment start/finishは欠落可能だが、存在する値は保持し、start/finishの両方がある場合だけ順序を検証する
  - durationは入力が宣言した非負のworking durationとして保持し、calendarによる再計算やstart/finishとの差の一致をv1では要求しない
  - milestoneはstartとfinishが等しくduration 0、summaryは子を持つtaskとする
  - percentCompleteは0〜100の整数とし、小数、-1、101をinvalidとする
  - C1の対象を`summary = false`のleaf taskへ限定し、summary進捗は観測・保持のみとする
  - 完了条件: required/optionalな日時fieldの欠落可否と、各invalid条件が文章で判定可能になっている

- [x] `ZB-P1.4` dependencyのv1 scopeを閉じる
  - semantic edgeを`predecessor UID → successor UID`として定義する
  - v1正式対応をfinish-to-start・lag 0だけに限定する
  - 欠損参照、自己参照、二task以上のcycleをinvalidとする
  - 他のlink typeまたは非zero lagをunsupported errorとし、暗黙変換や無視を禁止する
  - semantic identityを`(predecessor UID, successor UID, type, lag)`とし、同じtupleの重複をinvalidにする
  - dependency collectionの並び順は意味を持たない
  - dependency編集はC1に含めない
  - 完了条件: link type/lagの扱いを「G2で決める」とする文がsemantic contractからなくなり、入力ごとのvalid/unsupported判定が一意になる

- [x] `ZB-P1.5` resource、assignment、calendar、unknown dataのfail-closed規則を確定する
  - resource、assignment、calendarのUIDは各collection内で空でなく一意とする
  - assignmentは既存taskを参照し、resource UIDがある場合だけ既存resource参照を必須とする
  - unassignedはresourceなしとしてvalidとし、`-65535`などのsentinelをsemantic UIDへ露出しない
  - project/task/resourceのcalendar参照がある場合は既存calendarを必須とする
  - dependency/resource/assignment/calendar collectionは空でもvalidとし、task以外のcollection順序は意味を持たない
  - unknown、actual、EV、baseline、timephased、extended dataのopaque preservationをv1では約束しない
  - R1ではunsupported dataの存在を報告し、C1では保持保証がない場合にapply/exportを成功させない
  - 完了条件: unsupported dataを無警告で破棄・正規化・成功扱いできる抜け道がない

- [x] `ZB-P1.6` C1の許可変更とsemantic equivalenceを確定する
  - 許可operationを「stable UIDで選んだleaf task一件のpercentComplete更新」だけにする
  - preconditionには対象UIDと現在のpercentCompleteが必要と定義する。JSON field名やschemaはG2へ送る
  - 対象UID不存在、重複、summary task、現在値不一致、同値更新、整数範囲外、unsupported data存在時はapplyしない
  - apply後は対象percentComplete以外の`required`と存在していた`optional-preserved`がsemantic equivalentであることを要求する
  - byte一致、XML要素順、空白、serialization表現はsemantic equivalenceに含めず、詳細なnormalizationはG2へ送る
  - 完了条件: C1の入力、許可変更、reject条件、保持条件が一つの節だけで追跡できる

- [x] `ZB-P1.7` semantic fixture catalogとtraceabilityを作る
  - `docs/miku-project-semantic-fixture-catalog-v1.md` を新規作成する
  - 各行に `fixture ID / 種類 / baseまたは入力差分 / 検証対象 / 期待status / 関連不変条件` を書く
  - valid/boundary: `S-V001` dependency.xml正例、`S-V002` validなsummary/child階層、`S-B001` 進捗0、`S-B002` 進捗100、`S-B003` unassigned、`S-B004` optional field欠落、`S-B005` 空forest/collection、`S-B006` 非task collection順序のsemantic equivalenceを登録する
  - identity/hierarchy: `S-I001` duplicate task UID、`S-I002` unknown target、`S-I003` hierarchy orphan/depth jump、`S-I004` summary/children不一致、`S-I005` duplicate resource/assignment/calendar UIDを登録する
  - datetime/progress: `S-I006` project date order、`S-I007` task date order、`S-I008` required field欠落、`S-I009` negative duration、`S-I010` milestone mismatch、`S-I011` percent -1、`S-I012` percent 101、`S-I013` percent小数を登録する
  - dependency/reference: `S-I014` missing predecessor、`S-I015` self dependency、`S-I016` cycle、`S-I017` orphan assignment、`S-I018` missing calendar、`S-I019` unsupported link/lagを登録する
  - C1/unsupported: `S-I020` unsupported extended/unknown data、`S-I021` precondition mismatch、`S-I022` summary taskへの進捗変更、`S-I023` 同値更新を登録する
  - type/value: `S-I024` identity token、non-empty text、datetime、boolean、resource type、units/workのinvalid値を登録する
  - dependency identity: `S-I025` duplicate semantic tupleを登録する
  - この段階では製品用IRやJSON schemaを作らず、実行fixture化の担当phaseをG3と記録する
  - 完了条件: semantic contractの全不変条件とreject条件が一つ以上のfixture IDへ対応する

- [x] `ZB-P1.8` G1文書横断reviewを完了する
  - `docs/miku-project-zero-base-scenarios-v1.md`、semantic contract、fixture catalog、実施計画、TODOでR1/C1、field scope、用語、非目標を照合する
  - semantic contractに「ordered tree」「FS・lag 0以外をG2で決める」「unknown dataの扱いをG2で決める」という未決表現が残っていないことを確認する
  - G2へ残すのが外部形式、IR、schema、diagnostics、serialization normalizationだけであることを確認する
  - `git diff --check`を実行し、変更文書の差分をreviewする。文書だけの変更なので製品testは必須にしない
  - 完了条件: review結果と未決事項がsemantic contractのG1 checklistに記録され、未決事項が0または明示的な承認依頼になっている

- [x] `Gate G1` semantic contract、domain scope table、fixture catalog、review checklistを人が承認する（2026-08-10）
  - 承認前に`ZB-P1.1`〜`ZB-P1.8`をすべて完了させる
  - 承認後、G2を「現在の最優先」へ切り替えた

## ゼロベース再設計: G2 完了

`ZB-P2`では製品コードを変更していない。成果物は [format and loss contract v1](miku-project-format-and-loss-contract-v1.md) と [change contract v1](miku-project-change-contract-v1.md) である。詳細な完了条件は [実施計画のZB-P2](miku-project-zero-base-implementation-plan-v20260810.md#zb-p2-形式損失中間表現変更契約) を参照する。

- [x] `ZB-P2.1` 原入力、中間表現、Projection、変更要求、派生成果物、外部形式の役割を確定する
- [x] `ZB-P2.2` miku-project固有の中間表現が必要か判断する
- [x] `ZB-P2.3` 中間表現を採用する場合、internal / exchange / persistentの役割とschema versionを定義する
- [x] `ZB-P2.4` v1形式ごとの `read / write / roundtrip / loss / unsupported` matrixを作る
- [x] `ZB-P2.5` preservationを `required / normalized / lossy-with-warning / unsupported-error / opaque-preserved` に分類する
- [x] `ZB-P2.6` AI向けProjectionのpurpose、範囲、情報量、規則を定義する
- [x] `ZB-P2.7` whole-state replacementとoperation-based changeの境界を決める
- [x] `ZB-P2.8` 許可operation、selector、logical publication、precondition、validation、diff、output preflight、apply後検証を定義する
- [x] `ZB-P2.9` loss、normalization、ignored change、unsupported dataのprovenance表現を定義する
- [x] `ZB-P2.10` artifactの役割、寿命、schema versionを対応づけ、hidden state非依存を確認する
- [x] `Gate G2` artifactの役割・損失規則・変更安全性を人が承認する（2026-08-10）
  - internal-only IR、`miku-project-ms-project-xml-subset/v1`のnamespace / field / lexical / canonical child順 / 非目標、Projectionの限定公開、C1 operation allowlist、diff/output planをdigestで束縛するhuman gateを承認対象とする
  - C1成功出力を`project.xml + provenance.json + COMMITTED`を持つ新規artifact set directoryとし、directoryとcommit markerを排他的に新規作成する。markerなしをincomplete、markerありで検証不一致をcorruptとして利用禁止にし、既存pathを置換しない判断を承認対象とする
  - 承認後、G3を「現在の最優先」へ切り替えた

## ゼロベース再設計: G3 完了

G3では、承認済みのsemantic / format / change contractをCLI・diagnostics・conformance契約へ落とし込む。まだ製品コードは変更しない。詳細な依存と完了条件は[実施計画のZB-P3](miku-project-zero-base-implementation-plan-v20260810.md#zb-p3-cli-diagnostics-共通conformance契約)を参照する。

- [x] `ZB-P3.1` v1 CLIの最小command matrixを確定する（[CLI contract v1](miku-project-cli-contract-v1.md)）
- [x] `ZB-P3.2` args、stdin/stdout/stderr、file output、encoding/BOMの規則を定義する
- [x] `ZB-P3.3` exclusive directory create、commit marker、`incomplete / committed / corrupt`、cleanup diagnosticsをCLI契約へ定義する
- [x] `ZB-P3.4` versioned result / diagnostics schemaを定義する（[result and diagnostics contract v1](miku-project-cli-result-contract-v1.md)）
- [x] `ZB-P3.5` diagnostic code、severity、scope/path、status、I/O metadata、loss、normalization、retryabilityを定義する
- [x] `ZB-P3.6` exit codeとsuccess、validation failure、invalid usage、internal errorの境界を定義する
- [x] `ZB-P3.7` Nodeをv1の参照実装とする。正本は承認済み契約・Schema・共通fixture / goldenであり、Nodeの偶発的な挙動は契約を上書きしない（2026-08-10決定）
- [x] `ZB-P3.8` [runtime capability contract v1](miku-project-runtime-capability-contract-v1.md)で共通core profile、静的capabilityと動的preflight、runtime固有extensionの境界を定義する。v1のNode/Java extension setは空とする
- [x] `ZB-P3.9` [conformance corpus v1](miku-project-conformance-corpus-v1.md)と`testdata/conformance/v1/`に共通fixture / golden、P3完了時の21 workflow / harness case、31 schema / binding adversarial case、比較modeを設計する。runtime manifest/asset/source failure、複数diagnostic集約、command別I/O/effect、Projection/state binding、expected plan不一致、`COMMITTED`後にresultを受け取れないunknown outcomeを含める。現行case数はcorpus v1を正本とする
- [x] `ZB-P3.10` [runtime manifest contract v1](miku-project-runtime-manifest-contract-v1.md)とJSON Schemaでproduct / runtime / fixture / asset / source / SHA-256 / capabilityの記録・検証規則を定義する
- [x] `ZB-P3.11` command候補を`read-only / artifact生成 / 意味変更`へ分類する
- [x] `ZB-P3.12` [human gate and next action contract v1](miku-project-human-gate-and-next-action-contract-v1.md)でnon-interactive実行、human gate、retry / abort / next actionの機械判定を定義する
- [x] `Gate G3` CLI契約、command別I/O/effect、artifact/result/diagnostic schema、Projection/source-stateを含むcross-artifact binding、conformance corpus、runtime manifest、human gate / safe next actionを文書横断reviewし、人が承認する（2026-08-11）

G3承認により、`ZB-P4`（Node CLI vertical slice）の製品実装へ進める。

## ゼロベース再設計: P4 現在の最優先

P4ではNode CLIを承認済み契約の参照実装にする。最初に現行互換動作とtest topologyを固定し、その後に責務分離と新契約のvertical sliceへ進む。詳細な境界とGate G4条件は[実施計画のZB-P4](miku-project-zero-base-implementation-plan-v20260810.md#zb-p4-node-cli-vertical-sliceとv1完成)を参照する。

- [x] `ZB-P4.1` 現行CLIの互換動作をcontract testsで固定する
  - `tests/mikuproject-cli-compatibility-contract.test.js`で、legacy command surface、help/version、AI spec、stdin/stdout/stderr、named file outputと上書き、JSON usage diagnostics、draft→workbook変換を固定した
  - legacy diagnosticsが`--out`をI/O metadataへ反映しない観測済み挙動は、新v1 result contractと混同せず明示的に固定した
- [x] `ZB-P4.2` 現行test suiteの`fast / full / all`を実態に合わせ、一部testの実行漏れを解消する
  - `fast`を日常回帰、`full`をCLI統合testとbrowser runtime contractを含む完全回帰、`all`を全checked-in test fileを実行する安定aliasとして固定した
  - `tests/mikuproject-core-api.test.js`と`tests/mikuproject-core-api-loader.test.js`をsuiteへ編入し、`tests/mikuproject-test-suite-topology.test.js`で新規test fileの未分類・重複を検出する
- [x] `ZB-P4.3` semantic変更より先に、CLIのparser、command service、I/O、diagnostics、formattingを分離する（詳細は[実施計画のP4.3](miku-project-zero-base-implementation-plan-v20260810.md#p43の実行計画)）
  - [x] `ZB-P4.3.1` entrypointからerror / argv parseを抽出し、`process.argv`とtop-level error handlingをentrypointだけに残す。`scripts/lib/cli-errors.mjs`と`cli-argv.mjs`へ移し、parse/errorの既存code・JSON diagnosticsをdirect testとcompatibility testで固定した
  - [x] `ZB-P4.3.2` text / binary / Base64 I/O、stdin source検査、primary output、diagnostics用I/O記述を`cli-io.mjs`へ抽出した。legacy `--out`上書きとdiagnostics I/O metadata欠落を変えない
  - [x] `ZB-P4.3.3` diagnostics/status/error/validation formatterを`cli-diagnostics.mjs`、help presentationを`cli-presentation.mjs`へ抽出した。legacyのmessage→code推測は移動だけとし、v1へ継承しない
  - [x] `ZB-P4.3.4` legacy commandを`ai`、`state`、`import/export`、`report`のfamily serviceへ一つずつ移し、`cli-legacy-router.mjs`と73行のentrypoint wiringへ縮小した
  - [x] `ZB-P4.3.5` CLI内部moduleを決定的な順序でsingle `.mjs` bundleへ内包し、既存のrepository外bundle smokeとsource archiveの内容確認を通した
  - [x] `ZB-P4.3` 完了確認: `npm run test:full`、`npm run build:full`、P4.1 compatibility contract、repository外bundle smokeが成功し、v1 command / semantic / safe publicationを先取りしていない
- [x] `ZB-P4.4` R1の外部XML `validate → inspect(project_overview)` を最初の新契約vertical sliceとして実装する（詳細は[実施計画のP4.4](miku-project-zero-base-implementation-plan-v20260810.md#p44の実行計画)）
  - [x] `ZB-P4.4.1` schema registry / standalone validator、canonical JSON / SHA-256、semantic collection canonicalizationを実装し、生成物driftとgolden digestをdirect testで固定する
    - [x] `ZB-P4.4.1a` `ajv` / `ajv-formats`をbuild-timeだけに追加し、generator、drift check、第三者告知を整備する
    - [x] `ZB-P4.4.1b` 四schemaを固定registryでcompileし、四validatorをexportするimport-free standalone ESMを決定的に生成する
    - [x] `ZB-P4.4.1c` Unicode code point key sort、契約escape、整数限定、unpaired surrogate拒否を持つcanonical JSON serializerを実装する
    - [x] `ZB-P4.4.1d` task順を保持し、dependency tupleと他collection UIDをsortするsemantic canonicalization、semantic/raw SHA-256を分離実装する
    - [x] `ZB-P4.4.1e` 公式正例と18 schema-layer contract case、二つのgolden digestをdirect testで固定する。13 cross-artifact binding caseは未完のまま残す
    - [x] `ZB-P4.4.1f` generated drift、repository外import、FAST_SUITE分類、fast/full/build:full、legacy回帰を確認する
  - [x] `ZB-P4.4.2` legacy parserと分離したv1 strict argv、exclusive result file、result envelope / diagnostic / status / exit / next-action builderを実装する
    - [x] 五workflow commandのlong option grammar、control operation、unknown / duplicate / missing / positional / stdin conflict / purpose scopeをproject未読でrejectする
    - [x] stdoutまたは新規result fileのexclusive reservation、canonical parent path、abort時の自身の未完file cleanupを実装する
    - [x] stable diagnostic、status / exit、deterministic `next_action`、schema-valid JSON + LF result builderを実装する
    - [x] legacy entrypointは未変更とし、v1 router公開接続とsingle-MJS内包を`ZB-P4.4.6`へ残す
  - [x] `ZB-P4.4.3` XML subset profile scan → semantic state decode → v1 invariant validationを実装し、`S-V001`、`S-I012`、`S-I020`を通す
    - [x] raw UTF-8 byte digest、先頭BOM normalization、XML declaration / namespace / attribute / singleton / container profile scanを追加した
    - [x] MS Project XML subsetを`miku_project_semantic_state/v1`へdecodeし、pseudo task除外、ordered forest、FS/lag 0 dependency、resource / assignment / calendar mappingを固定した
    - [x] semantic validatorがstable code・rule ID・semantic locationを直接生成し、`S-I012`を`semantic.invalid`、`S-I020`を`semantic.unsupported`として分離した
    - [x] canonical XML→semantic golden exact比較、invalid/unsupported seed、UTF-8/BOM/profile、forest/dependency direct testsをFAST_SUITEへ追加した。v1公開entrypoint接続は`ZB-P4.4.6`のままとする
  - [x] `ZB-P4.4.4` `validate`をfile/stdin・stdout/new result fileで実装し、`CV-VALID-001`、`CV-INVALID-001`、`CV-UNSUPPORTED-001`をcontract testへ接続する
    - [x] direct external XML regular file / explicit stdinをraw SHA-256付きI/O metadataへ安全に読み、missing / symlink / type / read failureをstable I/O diagnosticへ分類した
    - [x] fixed test runtime bindingを注入する`validate` serviceを追加し、strict invocation → reserved result transport → XML decode → semantic validate → resultの順を固定した
    - [x] validはstate digest付き`complete`、invalid / unsupportedはstate digestなしの`rejected`にし、normalization / unsupported observations、input不変、stdout/file result channelをdirect testで検証した
    - [x] directory artifact-set入力は後続のartifact verifier workstreamへ、public entrypoint/bundle接続は`ZB-P4.4.6`へ残す
  - [x] `ZB-P4.4.5` semantic stateからv1 `project_overview`を生成し、source digest / scope / content bindingと`CI-OVERVIEW-001` goldenを固定する
    - [x] `validate`と同じexternal XML decode / semantic validation preparationを共有し、valid stateだけからProjectionを生成する
    - [x] project overview専用builderと`RB-012` binding checkerで、source digest、固定scope、task preorder / 0始まりorder、全dependencyを一つの決定的mappingとして固定した
    - [x] `CI-OVERVIEW-001`のProjection goldenを追加し、exact JSON、invalid/unsupported時の`data = null`、同一runtime byte determinismをFAST direct testで検証した
  - [x] `ZB-P4.4.6` fixed test bindingによるR1 integration、byte determinism、usage時project未読、legacy回帰、bundle/source archive包含を検証する
    - [x] public CLI / development bundleはv1 command wordをlegacyより先に識別するが、manifest未完成のためproject/result pathを読まず・予約せず`runtime.capability-missing`でfail-closedにした
    - [x] fixed verified test bindingだけを受けるR1 subprocess harnessでfile/stdin、stdout/new result file、existing result file拒否、usage errorのinput未読、同一runtimeのbyte determinismを確認した
    - [x] v1 module graphとstandalone schema validatorをsingle-MJSへ決定順で内包し、v1 private helper名はlegacyと衝突しないclosureへ隔離した。source archiveにschema/generator/golden/test/harnessを含むことも確認した
    - [x] full-only integration test、legacy compatibility、既存CLI integration、fast/full/build:full、`git diff --check`を通した
  - [x] P4.4ではpartial runtimeを`miku-project-cli-core/v1`適合として公開しない。実asset/source/manifest bindingとrepository外のworkflow smokeはP4.9/P4.10で行う
- [x] `ZB-P4.5` C1のhuman gate直前までを実装する（詳細は[実施計画のP4.5](miku-project-zero-base-implementation-plan-v20260810.md#p45の実行計画)）
  - [x] `task_change_context`を対象leaf、ancestor、接続dependency、対象assignment / resourceに限定して生成し、source digest・scope・contentの`RB-012` exact bindingと`CI-CONTEXT-001` goldenを固定した
  - [x] `--request`をUTF-8 / BOMなし / duplicate keyなしのstrict JSON artifactとしてfileまたはexplicit stdinから読み、kind/version/schema、raw digest、source/pathをfail-closedにした
  - [x] `set_task_percent_complete`だけをbase digest・leaf・expected current value・no-opの検査後にdry-runし、pre/post semantic validation、対象外field preservation、semantic diffを固定した
  - [x] planned stateをinternal-onlyに保ち、canonical XML encode/redecode equivalenceとraw digest、read-only destination preflightを含む`output_plan`を生成した
  - [x] `RB-001`〜`RB-005`、`CP-CHANGE-001`、base/current stale・no-op・allowlist外・BOM/duplicate request・unsafe/existing destinationをdirect / fixed-binding integrationで検証した。`plan-change`はdestinationもproject artifactも作らない
  - [x] public source CLI / development bundleはP4.9のverified runtime manifestまでv1 workflowを`runtime.capability-missing`でfail-closedに保つ
- [x] `ZB-P4.6` approvalで束縛したC1 apply、exclusive artifact publication、read-only verification、committed artifact set入力を実装する（詳細は[実施計画のP4.6](miku-project-zero-base-implementation-plan-v20260810.md#p46の実行計画)）
  - [x] `ZB-P4.6.1` `plan-change` resultとapprovalをstrict JSONで読み、current projectからdiff/output planを再計算して`RB-001`〜`RB-006`、runtime、destination、loss/unsupportedをdirectory予約前に再検証した。approval schema不正、binding不一致、stale state、destination raceを区別し、internal preparationからpublicationへはまだ接続しない
  - [x] `ZB-P4.6.2` provenanceをrevalidated input/change/output/runtimeから純粋生成し、13 transformation、normalization、空のloss/unsupported、`RB-007`、schema、BOMなし末尾LF一件のcanonical bytes/raw digestを固定した。XML再decode、state/digest、target/before/after、normalizationの相互照合と、observationの決定的sort/重複排除を実装し、publisherへはまだ接続しない
  - [x] `ZB-P4.6.3` read-only verifierを先に実装し、root/memberのlstat、marker、exact三member、canonical XML/provenance/schema/output digest/state digestから`absent / incomplete / committed / corrupt / 判定不能(null)`を分類した。expected plan指定時はruntime、destination、change/output digest、target/before/after、normalizationを`RB-008`で照合し、mismatchでもcommitted実測bindingsを保持する。verifierは一切のrepair/writeを行わない
  - [x] `ZB-P4.6.4` destinationをnon-recursive exclusive createし、二memberのexclusive write/closeとmarker前検証、空の`COMMITTED`のexclusive create、marker後再検証を行うpublisherを実装した。marker前だけ追跡済みregular memberと空directoryをbest-effort cleanupし、race既存path・想定外entry・marker boundary後setは削除しない。write / cleanup / post-marker failureを区別する内部state machineを注入testで固定した
  - [x] `ZB-P4.6.5` fixed verified bindingの`apply-change` serviceへ再検証→actual apply/post-validate→provenance→publicationを接続し、post-marker verifierのcommitted descriptorだけをsuccess payloadへ載せた。result fileを先にexclusive予約し、reservation失敗時は四入力もdestinationも未読/未確定のままstdout resultへ退避する。marker前write failureは`absent / succeeded cleanup`、result delivery不明時は再applyせず`verify-artifact`で回復する
  - [x] `ZB-P4.6.6` `--project <directory>`をcommitted verifier経由で`inspect / validate / plan-change / apply-change`へ接続し、incomplete/corrupt setを部分利用せずexternal XMLと同じsemantic pipelineへ渡す。directory source/canonical path/project.xml raw digestをI/O metadataへ残し、external XMLとのProjection・diff・output XML bytes同値と入力artifact不変をdirect testで固定した
  - [x] `ZB-P4.6.7` `CA-CHANGE-001`、`CA-DEST-EXISTS-001`、`CA-BINDING-001`、`CA-CLEANUP-AGGREGATE-001`、`CVF-ABSENT-001`、`CVF-INCOMPLETE-001`、`CVF-CORRUPT-001`、`CVF-COMMITTED-001`、`CVF-EXPECTED-PLAN-MISMATCH-001`、`CVF-EXPECTED-PLAN-INVALID-001`、`CU-UNKNOWN-OUTCOME-001`と`BC-APPROVAL-DIVERGENCE-001`、`BC-PLAN-BINDINGS-VALID-001`、`BC-APPLY-PATH-DIVERGENCE-001`、`BC-VERIFY-PATH-DIVERGENCE-001`、`BC-VERIFY-STATE-DIVERGENCE-001`をfailure injection付きでmaterializeした。`verify-artifact`のfixed-binding service、runner管理下のsame-path byte determinism、result file/stdout、cross-artifact corpus、bundle/source archive、legacy回帰、full buildを固定した
  - [x] `ZB-P4.6` 完了確認: loss/unsupportedやmarkerなし/corrupt setを成功にせず、cleanup権限・effects・diagnosticsがcommit point前後で一意であり、public runtimeはP4.9までfail closedのままである（2026-08-12 承認）
- [x] `ZB-P4.7` 既存coreの再利用部分を新conformance fixturesで検証する（詳細は[実施計画のP4.7](miku-project-zero-base-implementation-plan-v20260810.md#p47の実行計画)）
  - [x] `ZB-P4.7.1` XML codec / validator、AI view / Patch、core API loader、bundle builder、旧Agent workflow例、旧fixture / regression、workbook / report / browser runtimeを、現在のv1経路・根拠fixture・再評価gateへ対応付けた
  - [x] `ZB-P4.7.2` `S-V001`、`S-I012`、`S-I020`、`CI-OVERVIEW-001`、`CI-CONTEXT-001`、`CP-CHANGE-001`、C1 apply / verificationとlegacy XML / core API / CLI / workbook / XLSX回帰、旧Agent workflow例を実行し、v1 contractを代替できる旧coreがないことと旧経路が保守可能なことを確認した（2026-08-13: v1/legacy core 9 files / 78 tests、legacy CLI 2 files / 64 tests、workbook/XLSX 2 files / 33 tests、Agent workflow例2本が成功）
  - [x] `ZB-P4.7.3` [v1 core再利用採否 v20260812](miku-project-zero-base-reuse-verdict-v20260812.md)へ`v1実行経路へ採用 / 互換性の証拠として保持 / 後続scopeまでdefer`、理由、P4.8/P4.9/P4.10以降の再評価条件を固定した。旧形式・名称・public fail-closed境界は変更しない
  - [x] `ZB-P4.7.4` 記録と回帰結果をレビューし、人が承認した（2026-08-13）
- [x] `ZB-P4.8` 階層C1 sliceを追加し、v1 capability matrixを更新する（詳細は[実施計画のP4.8](miku-project-zero-base-implementation-plan-v20260810.md#p48の実行計画)）
  - [x] `ZB-P4.8.1` `S-V002` canonical XML、before / after semantic golden、hierarchy overview / nested-leaf context golden、C1 requestと成功経路の`CV/CI/CP/CA-HIERARCHY-*`、拒否経路の`CV-HIERARCHY-INVALID-PREORDER-001`、`CV-HIERARCHY-INVALID-SUMMARY-001`、`CP-HIERARCHY-SUMMARY-REJECT-001`をsuite casesへ追加し、新設した[miku-project v1 capability matrix v20260813](miku-project-v1-capability-matrix-v20260813.md)へNode capabilityとdefer範囲を記録した
  - [x] `ZB-P4.8.2` `S-V002`、`S-I003`、`S-I004`をadapter / semantic validator / canonical encoderへ接続し、preorder・parent・summary・task order・XML encode/redecode equivalenceを固定した。不正二caseはsuite-indexの期待`semantic.invalid`、rule ID、diagnostic `location.path`までmachine-readableに照合する
  - [x] `ZB-P4.8.3` UID `2`のancestor付き`task_change_context`と全taskのhierarchy overviewを`RB-012` / exact goldenで固定し、summary / nonexistent targetをread-only rejectした
  - [x] `ZB-P4.8.4` nested leafの`0 → 50` C1をplan / approval / apply / verifyまで通し、`CP-HIERARCHY-SUMMARY-REJECT-001`で`S-I022` summary targetをartifact未作成でrejectし、target外の階層・collection意味を保持した。committed hierarchy artifactを入力にした`50 → 75` planでは、artifact set入力と直接`project.xml`入力のhierarchy Projection、semantic diff、output plan（preflight XML digestを含む）、RB-001〜005 bindingが一致することを確認した
  - [x] `ZB-P4.8.5` hierarchy / flat R1-C1 / legacy compatibility / bundle-source archive / full buildを回帰し、capability matrixとpublic fail-closed境界を更新した（2026-08-13: `npm run build:full`、29 files / 288 tests）
  - [x] `ZB-P4.8.6` hierarchy C1 sliceを再レビューし、人が承認した（2026-08-13）。次は`ZB-P4.9`のruntime manifest / asset bindingであり、public v1 runtimeは引き続きfail-closedとする
- [x] `ZB-P4.9` versioned single `.mjs`、sources、runtime manifest、SHA-256を生成する（2026-08-13承認。詳細は[実施計画のP4.9](miku-project-zero-base-implementation-plan-v20260810.md#p49の実行計画)）
  - [x] `ZB-P4.9.1` cleanかつ`v<package.version>`がHEADを指すsourceだけをrelease runtime buildの入力にし、freshな`runtime/node/`へ`miku-project-node-<runtime.version>.mjs`、対応sources、固定名`runtime-manifest.json`を生成する
  - [x] `ZB-P4.9.2` manifestへcanonical artifact/source SHA-256・size、corpus digest、9 capability、contract/version/source provenanceを記録し、schema・filename/version・canonical JSONを検証する
  - [x] `ZB-P4.9.3` versioned bundle自身がadjacent manifest、executable、sourcesをworkflow前に検証し、verified runtime bindingで五commandを実行する。source CLI/development bundleはmanifest未検証なら従来どおりfail-closedとする
  - [x] `ZB-P4.9.4` valid runtime、manifest不正、executable改変、source欠落・改変、capability不足を`CR-*` caseへ接続し、いずれもproject未読・destination未作成を確認する。direct filesystem guardで、runtime failure時にproject / result pathへの`lstat`・`realpath`・`readFile`が一度も起きないことも固定した
  - [x] `ZB-P4.9.5` build/test結果、manifest digestの外側pinがP4.10/Skillsで必要なこと、P4.10 clean smoke前にreleaseを宣言しないことを記録し、レビューと人の承認を得た（2026-08-13）
    - clean / exact-tag temporary sourceのrelease builder成功経路、`npm run build:full` 30 files / 297 tests、`git diff --check`を確認した。P4.10のclean smoke、Release checksum / Skills lockによる外側manifest pin、公開Releaseは未完のまま維持する
- [x] `ZB-P4.10` repository外のclean temporary directoryで、外側manifest pinを持つ配布runtime smokeを実行する（2026-08-14承認。第3回reviewのCLI result runtime binding完全一致補正、補正後full regression、`git diff --check`を確認。Gate G4の現在状態は下記G4 sectionを正本とする。詳細は[実施計画のP4.10](miku-project-zero-base-implementation-plan-v20260810.md#p410の実行計画)）
  - [x] `ZB-P4.10.1` clean / exact-tag temporary sourceからrelease runtimeを生成し、`runtime-manifest.json`、versioned `.mjs`、sourcesだけをconsumer directoryへ配布する。source checkoutを削除後、consumer自身のproject / request / plan / approvalだけで五workflowを成功させる
  - [x] `ZB-P4.10.2` test-owned consumer preflightを契約の検証順どおり完成する。copy前のmanifest raw SHA-256を外側trust anchorとして通常file / 非symlinkの配布manifestへ照合した後、canonical manifest directoryと固定basenameからexecutable / sourcesを解決し、両方の通常file / 非symlink、canonical path、size、raw SHA-256をmanifestと照合する。両assetの検証完了前にlauncher descriptorを返さない
    - manifest / assetとも同じregular-file descriptor primitiveを使う。runtime自己検証は、consumerの起動前検証に加わる起動後の二重検証として維持する
  - [x] `ZB-P4.10.3` pre-launch failure matrixを完成する。既存のmanifest + executable coordinated tamper、missing / malformed trust anchor、manifest missing / symlinkに加え、manifestはpin成功のままexecutableだけを変更したcaseとsourcesだけを変更したcaseを追加し、size / digest mismatchを区別して拒否する。各caseでlauncher 0回、project / request / plan / approval / result / destination access 0回、result / artifact未作成、fallbackなしを固定する
  - [x] `ZB-P4.10.4` 成功smokeを完全なconsumer preflightと、schema-validかつ完全一致するconsumer result bindingへ接続する
    - [x] `ZB-P4.10.4a` operationごとにmanifest pinと両assetを再検証し、検証済みcanonical executable path、`runtime.launcher = node`、固定`process.execPath`だけからcommandを組み立てる。五workflowをこの経路で成功させ、output plan / provenance runtime bindingは完全一致で照合する。glob、PATH、mtime、newest / alternate version fallbackを使わない
    - [x] `ZB-P4.10.4b` CLI result envelopeを既存の生成済み`validateCliResult`で検証してからbinding比較へ進める。schema不正はtest-owned `consumer.result-invalid`としてfail closedにし、CLI diagnostic catalogへ新codeを追加しない
    - [x] `ZB-P4.10.4c` `result.runtime`とconsumer preflightが返す`runtime`を、top-level key集合とnested digest objectを含めてdeep exact比較する。schema-validだが値が異なる場合は既存のtest-owned `consumer.result-binding-mismatch`で拒否し、known fieldだけの比較や部分一致を使わない
    - [x] `ZB-P4.10.4d` 正常resultを複製したtable-driven negative testで、runtime field欠落、余分なtop-level field、余分なdigest内field、schema-validな値不一致を固定する。前3件は`consumer.result-invalid`、値不一致は`consumer.result-binding-mismatch`となり、成功五workflowも同じvalidator / exact比較を通ることを確認した（P4.10 consumer対象13 cases、runtime manifest test file全22 tests）
    - [x] `ZB-P4.10.4e` result以外のoutput plan / provenance runtime bindingの既存完全一致を維持し、補正後の五workflowで再確認した
  - [x] `ZB-P4.10.5` README、development、capability matrixへ「P4.10は外側pinを検証するが、公開Release checksumの発行、Skills lock、公開Releaseは後続」の境界を記録する。target/full regression、`git diff --check`、再レビュー、人の承認を得た（2026-08-14）
    - 起動前asset再補強のbaseline: 2026-08-13にsuccess五workflow、coordinated tamper、asset単独のsize / digest / entry failureを含むP4.10 target 12 testsと`npm run build:full` 30 files / 309 tests、`git diff --check`が成功した。ただしCLI result bindingの完全一致補正前なので最終承認証拠には数えない
    - 補正後の実装証跡: 2026-08-14にP4.10 consumer対象13 cases（runtime manifest test file全22 tests）、`npm run build:full` 30 files / 310 tests、`git diff --check`が成功した。同日に再レビューと人の承認を得た
    - manifest-pin-only実装のbaseline: 2026-08-13にconsumer preflight / coordinated-tamper testを含む`npm run build:full`が30 files / 303 testsで成功し、`git diff --check`も成功した。ただし起動前asset検証がないため承認証拠には数えない
    - 補強前baselineは2026-08-13にexternal consumer smoke、`npm run build:full` 30 files / 298 tests、`git diff --check`が成功したが、P4.10の承認証拠には数えない
- [x] `Gate G4` Node参照実装、共通conformance、決定性、安全なpublication、bundle/manifestを検証し、2026-08-14に人が承認した（詳細は[Gate G4 readiness v20260814](miku-project-gate-g4-readiness-v20260814.md)）
  - [x] `ZB-G4.1` clean consumer scenario、決定性、安全なpublication、repository外single `.mjs`、relevant tests / `npm run build:full`をP4.1〜P4.10のapproved evidenceへ照合する。2026-08-14時点で五条件をpassと判定した
  - [x] `ZB-G4.2` P4.7〜P4.10をsource freeze commit `693b4ecd7d4328d77f3b2eada9c4965a9c9b15f5`へまとめ、`package.json`と一致する未使用のlightweight tag `v1.0.3`を同commitへ付けた。既存tagの付け替えはしていない
  - [x] `ZB-G4.3` exact-tag clean sourceから生成したactual三memberを内部reference candidateとして`workplace/gate-g4/v1.0.3/runtime/`へ保持し、Git管理する[miku-project Node reference runtime lock v1.0.3](miku-project-node-reference-runtime-lock-v1.0.3.json)へsource/tag、build toolchain、package-lock / corpus、三memberのsize / SHA-256を固定した。lock SHA-256は`95cd11cc4460348fa066908994430adba5983384c06c75679855120e5c5ea3d5`である
  - [x] `ZB-G4.4` `verify:cli-v1-release-candidate`を追加し、Gate lockを外側pinとしてactual candidateの三memberだけをexternal consumerへcopyした。`validate` / `inspect` / `plan-change` / `apply-change` / `verify-artifact`とruntime/result/output plan/provenanceの完全binding照合が成功した。不整合lockはruntime launch 0回で拒否する
  - [x] `ZB-G4.5` lock/verifier追加後のruntime manifest test file全25 tests、`npm run build:full`（30 files / 313 tests）、最終`git diff --check`が成功し、実測test数と結果をreadiness/capability matrixへ記録した
  - [x] `ZB-G4.6` Gate G4 evidenceを最終レビューし、2026-08-14に人が承認した。公開Release checksum、署名、Skills lock、公開Releaseは承認対象外として後続へ残す

## ゼロベース再設計: P5 Java CLI適合（現在工程）

詳細な依存関係と全taskは[実施計画のZB-P5](miku-project-zero-base-implementation-plan-v20260810.md#zb-p5-java-cli適合)、固定入力は[Java contract handoff v1.0.3](miku-project-java-contract-handoff-v1.0.3.md)を正本とする。Java側の現在の`vendor/mikuproject/`は旧系列のlegacy referenceであり、新v1の仕様正本にしない。

開始時baselineとして2026-08-14にJava `0.12.0`の`mvn test`を実行し、132 tests、failures 0、errors 0、skipped 4で成功した。skip 4件はopt-in Node parityであり、Gate G5 conformanceの成功には数えない。

- [x] `ZB-P5.1.1` Node `v1.0.3`のsource revision、contract / fixture suite、corpus digest、snapshot allowlist、Java側配置、受入条件を固定した
- [x] `ZB-P5-A1` Java repositoryへimmutable contract snapshot importerとcanonical `SOURCE.json` generatorを追加した。exact tag `v1.0.3` / revision `693b4ecd7d4328d77f3b2eada9c4965a9c9b15f5`のidentity確認後はfull revisionだけから49 memberを`vendor/miku-project-contract/v1.0.3/`へ取り込む。同一inventoryの既存snapshotだけは`"snapshot_action":"unchanged"`で無変更no-opとし、それ以外は`contract-snapshot.destination-conflict`でfail-closedにする。exclusive directory creation、repository外・symlink親の拒否、ownership確認済みfailure cleanupにより、旧subtreeや既存snapshotを変更しない
- [x] `ZB-P5-A2` Java snapshot inventory verifier testとNode importer testを追加した。source identity、corpus digest、member set、exact directory topology、canonical `SOURCE.json`、そのraw SHA-256外側pin、regular/non-symlink、size / SHA-256を実装起動前に検証し、extra / missing / tampered / symlink / fractional size / rewritten manifestを拒否する。importer側は正常install、metadata不変の同一rerun no-op、異なる既存tree / empty directory、tag / revision、working tree非参照、確認後tag移動、出力path、symlink親、決定論的manifest、Git symlink、marker / member / manifest書込失敗cleanup、未知file / empty directory保全を検証する（Node 12 tests、focused verifier 8 tests成功）
- [x] `ZB-P5-A3` Java README / upstream docsでv1 snapshotとlegacy subtreeのauthorityを分離した。repository-wide `sh scripts/test-all.sh`を正本としてRelease workflowへ接続し、Node 12 testsとJava 140 tests（failures 0、errors 0、skipped 4）が成功した。P5-Aは最終レビューを通過し、2026-08-14に承認した
- [x] `ZB-P5-B1` `suite-index.json` / `contract-cases.json` loaderをJava test-side harnessに追加し、verified snapshotの30 workflow caseと31 contract/binding caseを直接読むようにした。unknown field、case ID重複、path escape、allowlist外参照はfail-closedである（2026-08-14、focused 8 tests、Node 12 tests + Java 148 tests成功）
- [x] `ZB-P5-B2` result / diagnostic / artifact / runtime manifestの四Schema registryをJava 8 test-sideで実装した。四Schemaのchecked-in positive example、代表negative example、JSON Schema layerの18 mutation caseを評価し、未対応語彙や登録外refはfail-closedにする
- [x] `ZB-P5-B3` canonical JSON / SHA-256とcase materializerをtest-sideに置き、`cross-artifact-binding` 13 caseで`RB-001`〜`RB-006`、`RB-011`、`RB-012`を実入力から評価した。`exact-json` / `semantic-state` / `semantic-cross-runtime` / `byte-same-runtime` / `artifact-topology` / `runtime-integrity`も独立boundaryとしてtestし、Node live outputやJava固有goldenをoracleにしない（2026-08-14、Node 12 tests + Java 161 tests、failure/error 0、skip 4）
- [x] `ZB-P5-B4` P5-B harnessを再技術reviewした。前回review後に見つけたwhole-object JSON textによるsemantic collection整列を、Node参照契約どおりdependency tuple / UID / Unicode scalar順の共通domain-aware canonicalizerへ置換し、digestと全three comparison modeへ適用した。optional field逆順・複数member・Unicode scalar・assignment UID/task逆順、summary taskと`summary = false`だがchildを持つschema-valid ProjectionのRB-012 negative testを追加した。focused test、`sh scripts/test-all.sh`（Node importer 12 tests、Java 164 tests、failure/error 0、skip 4）、snapshot post-verification（8 tests）、両repositoryの`git diff --check`を2026-08-15に成功させた。P5-Bは同日に人が承認した
- [ ] `ZB-P5-C1` Java `validate` first vertical sliceは実装済みで、初回・再技術reviewのJava側指摘を修正した。再技術reviewと人の承認待ち。`CU-USAGE-001`のNode/Java source挙動はcorpusどおり一致したが、frozen `v1.0.3` release/snapshotは歴史的差分を含むため、cross-runtime / Gate G5の適合をここで主張しない
  - Java側はlegacy CLIと分離し、strict argv、direct XML / stdin、XML subset / semantic validation、structured result、canonical semantic digest、exclusive-create resultを実装した。runtime manifestがまだ無いため、verified runtime bindingを供給できなければproject未読の`runtime.manifest-invalid`としてfail closedにする
  - 初回reviewで見つけた不正lexical値のfalse success、dependency cycle未検出、欠落start milestoneのinternal error、assignment rule ID、共通`CU-USAGE-001`未実行を修正した。再reviewでは`CU-USAGE-001`のwhole CLI argv、usage scope / option location、percent欠落とdependency endpointのrule境界、Node `Number.isSafeInteger`範囲、resource / assignment / calendarのUID付きdiagnostic pathをJava側へ固定した。巨大outline levelはdecode済みmember数を上限にしてallocationをfail-safeにする
  - fixed corpusの`command = cli`、`arguments = ["--unknown-option"]`、`cli.unknown-option`を正本とし、現行Node parserも先頭unknown long optionをoption scope / option location付き`cli.unknown-option`に修正した。Java serviceと現行Node sourceの挙動は一致する。immutable `v1.0.3` snapshotは改変せず、後続Node corrective releaseと新snapshotで参照identityを更新する
  - fixed v1.0.3 corpusのvalid / invalid / unsupported / hierarchyと、`CU-USAGE-001`、runtime未検証、malformed XML、result overwrite拒否、初回・再review回帰をJava focused 12 testsでtestした。Nodeのargv / XML adapter / R1 integration関連20 tests、Node全体`npm test`、`sh scripts/test-all.sh`のNode importer 12、Java 176（failure/error 0、skip 4）、snapshot post-verification 8 testsが成功した。artifact-set directory、public JAR entrypoint、manifest、`inspect`以降は含めない。これはP5-C1の人による承認またはrelease identityを含むcross-runtime適合を意味しない

P5-Bは承認済みであり、次に開始できるのはP5-C1の`validate` vertical sliceだけである。fat JAR更新、version increment、旧subtree同期、命名変更はP5-C1の範囲に含めない。

## 現行仕様由来・再評価待ち

以下は過去の実装と利用経験から得た候補を失わないために保持する。新仕様での実装を約束するものではない。Web UI項目は後続の `miku-project-web` 候補、帳票や見た目はderived output候補、意味と往復に関する項目は`ZB-P0`〜`ZB-P2`の入力としてtriageする。

次の分類は、以下の個別backlogに最初に適用する再評価先である。同じ項目が複数の分類に関係する場合も、R1/C1を成立させるまで実装を開始しない。

| 現行backlogの種類 | disposition | 再評価先 |
| --- | --- | --- |
| sample、`local-data`、BOM、source分割、build/test時間、既存roundtrip回帰 | evidence / reuse候補。現行互換とfixtureの根拠として保持 | `ZB-P4` |
| Skills runtime受け渡し、bundle smoke | Agent Skillsのprovenance候補 | `ZB-P6` |
| XML/XLSX、dependency、calendar、validation、partial apply、diff、scoped Projection | semantic / format / change候補。現行JSONやCLIを正本にしない | `ZB-P1`〜`ZB-P3` |
| actual、Earned Value、baseline、timephased data、ExtendedAttribute | domain scope候補。v1実装を約束しない | `ZB-P1` |
| Overview、Output、画面内task操作 | `miku-project-web` の後続候補 | `ZB-P8` |
| WBS XLSX、SVG、Markdown、Mermaid、sample workbookの見た目、タイムチャート | derived output候補。v1 coreの後に評価 | `ZB-P4`後 / `G7`後 |
| `docs/spec.md` と現行実装のdrift | current/target authorityの移行作業 | `ZB-P0`、`ZB-P7` |

- サンプルデータを更新し、利用者の好みに合う題材・構造・見た目へ見直す
- `miku-project-skills` 側で、上流 `bundle/miku-project.mjs` を `skills/miku-project/runtime/miku-project.mjs` として受け取る手順と smoke test 観点を反映する
- WBS workbook と `miku-project-sample.xlsx` のタイトル行で、フォントサイズ指定をどこまで使うか整理する
- `Mermaid` 出力は Markdown / 設計資料向けに残しつつ、見た目を制御しやすい `WBS SVG` 描画を別系統で追加するか検討する
- `WBS SVG` について、今の既定である `近接ラベル` 表示だけを残し、左側にテキストを描画する `一覧ラベル` モードは将来的に廃止したい
- 完了済みの大分割領域を再度触る前に、本当に新しい責務混在があるか確認する
  - `core-api*`、`msproject-*`、`project-patch-json*`、`project-xlsx*`、`project-workbook-json*`
  - `excel-io*`、`wbs-svg*`、`wbs-xlsx*`
- 構造変更を再開するときは、区切りごとに `npm run build:full` を回し、`tests/miku-project-cli.test.js` の実行時間も継続確認する
- 作成するテキストファイルについて、BOM 付き / なしを切り替えるスイッチを追加する
- `local-data/` 配下のファイルを、参照用・検証用・生成物で整理する
- `local-data/` に置くべきでない生成物や一時ファイルがないか見直す
- `XLSX Import` の実地回帰観点を明文化し、少なくとも次を継続確認する
  - export した `.xlsx` をそのまま import できる
  - Excel で 1 セル変更した `.xlsx` を import できる
  - 同じファイル名で保存し直した `.xlsx` を連続 import できる
  - 空 editable セルを埋めた変更を import できる
  - `Name / Start / Finish / PercentComplete / PercentWorkComplete / Notes` など主要 editable 列が戻る
  - `Milestone / Summary / Critical` など現在の task 真偽値列が戻る
- `Tasks.Predecessors` について、現状の `predecessorUid` `,` 区切り MVP から、`type / linkLag` など複雑な依存表現をどこまで戻すか整理する
- `.xlsm` について、`xlsx` 相当の workbook として import だけ受ける first cut を追加する
  - macro / VBA project の保持は行わず、今回の利用範囲では落ちてよい前提とする
- workbook import の次段候補として、少なくとも次の列を優先順位つきで整理する
  - 優先候補: `Resources.StandardRate / OvertimeRate / CostPerUse`
  - 優先候補: `Assignments.Start / Finish`
- import 前後で、どの `task / calendar / assignment` がどう変わったかを見やすく確認できる差分可視化を追加する
- 差分適用を前提として、生成AI や外部編集結果を全件置換ではなく部分適用できる運用を強化する
- `Overview` タブの summary / validation / preview の情報密度を見直し、どこを見る画面なのかをより直感的に伝わる構成へ調整する
- `Overview` 画面について、簡易な task 操作機能を追加するか検討する
  - 今は表示専用だが、軽い編集や操作だけはできると便利な可能性がある
  - 一方で責務過多や誤操作の不幸もありうるため、まずは仕様整理から始める
- `Output` タブの生成AI連携と各種 export ボタンの優先度表現を見直し、主操作と補助操作の区別をより明確にする
- `build:xlsx-sample` の所要時間を個別計測し、sample workbook 生成処理の支配要因を確認する
- `docs/spec.md` に残っている実装済み前提との差分を定期的に解消する
- 正本 / 表示用 / import 対象 / export 専用 の扱いを、UI または docs で分かりやすく可視化する
- `.xlsx import` の次段として、どのシート・列を今後 import 対象に広げるか整理する
- タスク実績について、`ActualStart / ActualFinish / ActualWork / RemainingWork / ActualCost / RemainingCost` などを今後どう扱うか整理し、将来的に対応する
- 将来検討: Earned Value (`PV / EV / AC / SPI / CPI` など) をどこまで扱うか整理し、必要なら対応する
- 実績・Earned Value 系は、いきなり広く対応せず、まず最小整理と小さな仕様を作って MVP から段階的に進める
- WBS 用の `ステータス` は `Task.ExtendedAttribute` で扱う前提で、`FieldID / FieldName / 値候補` を設計する
- `TaskStatus` 用 `ExtendedAttribute` を `miku-project-sample.xlsx` と `WBS workbook` のどちらまで見せるか決める
- `TaskStatus` 用 `ExtendedAttribute` の値候補と、`PercentComplete` / `Active` との関係を整理する
- 画面検索ではなく、条件指定にもとづく task の部分 export / scoped export を強化できるか整理する
- `phase_detail_view scoped` の延長として、phase 単位の入出力をうまく取り回す方法を整理し、使い勝手のよい導線を検討する
- 画面では `Calendars / Exceptions` を read-only 確認に留める前提で、`XLSX Import` 側の `WeekDays / Exceptions / WorkWeeks` 編集導線をどこまで整えるか整理する
- `Calendar / Baseline / TimephasedData / ExtendedAttributes` をどの順で扱うか優先順位を決める
- validation について、warning の重要度分け、修正候補のヒント、入力由来別の注意をどこまで出すか整理する
- `miku-project-sample.xlsx` の `Project` シートで、構造忠実方針を崩さない範囲の見た目調整を続ける
- `miku-project-sample.xlsx` の `Resources / Assignments / NonWorkingDays` で、強調色が過剰にならない最終バランスを調整する
- `WBS` の `プロジェクト情報` / `凡例` などと、`Project` シートの `Basic Info` に入っているドット編みかけ表現を除去する
- WBS workbook の表示改善を継続する
- SVG 出力について、プロジェクト名の位置を少し上にできるか調整する
- SVG 出力の phase の線を、今より少し太くするか検討する
  - 現状は細く感じるため、可読性と全体バランスを見ながら調整可否を確認する
- SVG のガントチャートについて、前後関係を dependency connector として表示するか検討する
  - 参考イメージは、task bar の背面に置く細い connector line とする
  - 完全な自由曲線ではなく、横→縦→横の直交配線を角丸でつなぐ表現を first candidate とする
  - 始点と終点では、bar 端から少し離した短い水平セグメントを持たせる
  - 長い区間は水平・垂直を基本にし、角だけを小さめの半径で丸める
  - bar / milestone / label より背面に描き、dependency は補助情報として主張しすぎないようにする
  - 線幅は task bar より細く、色は薄めにして、最後だけ小さい矢印で向きを示す
  - first cut では `FS` だけ表示するか、`SS / FF / SF` まで含めるかを整理する
  - connector の交差、密集時の可読性、`Daily / Weekly` の両方で成立するかを確認する
- `Daily` 表示の日ごとの横幅を、もう少し狭くできるか検討する
  - まずは変更仕様の整理から始め、可読性、文字詰まり、祝日/週境界の見え方、`Weekly / Monthly` とのバランスを確認する
- WBS workbook の見た目改善と、構造忠実 workbook との責務分離を保つ
- WBS について、完了タスクの表示 / 非表示を切り替えるオプションを追加する
- 将来検討: WBS workbook について、表示専用列と Excel 再利用向けの機械利用列（hidden 列）の分離が必要か整理する
- 低優先度: 週別または日別の `24h` 表記タイムチャートを追加するか検討する
  - イメージは `4直3交代` のシフト表に近い表示とする
  - まずは仕様検討から始め、対象データ、表示粒度、稼働日/非稼働日との関係、`WBS` 系出力との責務分離を整理する
- WBS Markdown の `プロジェクト情報` / `サマリ` / `WBS ツリー` / `WBS テーブル` をどう出すか sample ベースで固める
- `project summary markdown` のような、WBS 以外の Markdown 出力拡張を検討する
- `phase summary markdown` のような scoped Markdown 出力を追加するか検討する
- `WBS記述書 Markdown` 出力を追加し、task ごとの説明を別 Markdown として保存できるようにする
- `WBS記述書` 用 `Task.ExtendedAttribute` の最小項目として `TaskPurpose / TaskDeliverable / TaskOutOfScope / TaskDoneDefinition / TaskOwner` を扱う
- `WBS記述書 Markdown` では、長文補足を `Task.Notes` から出す
- `WBS記述書 Markdown` の sample 出力を作成し、1 task 1 節構成で読みやすいか確認する
