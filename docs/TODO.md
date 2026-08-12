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
- [x] `ZB-P3.9` [conformance corpus v1](miku-project-conformance-corpus-v1.md)と`testdata/conformance/v1/`に共通fixture / golden、21 workflow / harness case、31 schema / binding adversarial case、比較modeを設計する。runtime manifest/asset/source failure、複数diagnostic集約、command別I/O/effect、Projection/state binding、expected plan不一致、`COMMITTED`後にresultを受け取れないunknown outcomeを含める
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
- [ ] `ZB-P4.6` approvalで束縛したC1 apply、exclusive artifact publication、read-only verification、committed artifact set入力を実装する（詳細は[実施計画のP4.6](miku-project-zero-base-implementation-plan-v20260810.md#p46の実行計画)）
  - [x] `ZB-P4.6.1` `plan-change` resultとapprovalをstrict JSONで読み、current projectからdiff/output planを再計算して`RB-001`〜`RB-006`、runtime、destination、loss/unsupportedをdirectory予約前に再検証した。approval schema不正、binding不一致、stale state、destination raceを区別し、internal preparationからpublicationへはまだ接続しない
  - [x] `ZB-P4.6.2` provenanceをrevalidated input/change/output/runtimeから純粋生成し、13 transformation、normalization、空のloss/unsupported、`RB-007`、schema、BOMなし末尾LF一件のcanonical bytes/raw digestを固定した。XML再decode、state/digest、target/before/after、normalizationの相互照合と、observationの決定的sort/重複排除を実装し、publisherへはまだ接続しない
  - [x] `ZB-P4.6.3` read-only verifierを先に実装し、root/memberのlstat、marker、exact三member、canonical XML/provenance/schema/output digest/state digestから`absent / incomplete / committed / corrupt / 判定不能(null)`を分類した。expected plan指定時はruntime、destination、change/output digest、target/before/after、normalizationを`RB-008`で照合し、mismatchでもcommitted実測bindingsを保持する。verifierは一切のrepair/writeを行わない
  - [x] `ZB-P4.6.4` destinationをnon-recursive exclusive createし、二memberのexclusive write/closeとmarker前検証、空の`COMMITTED`のexclusive create、marker後再検証を行うpublisherを実装した。marker前だけ追跡済みregular memberと空directoryをbest-effort cleanupし、race既存path・想定外entry・marker boundary後setは削除しない。write / cleanup / post-marker failureを区別する内部state machineを注入testで固定した
  - [x] `ZB-P4.6.5` fixed verified bindingの`apply-change` serviceへ再検証→actual apply/post-validate→provenance→publicationを接続し、post-marker verifierのcommitted descriptorだけをsuccess payloadへ載せた。result fileを先にexclusive予約し、reservation失敗時は四入力もdestinationも未読/未確定のままstdout resultへ退避する。marker前write failureは`absent / succeeded cleanup`、result delivery不明時は再applyせず`verify-artifact`で回復する
  - [ ] `ZB-P4.6.6` `--project <directory>`をcommitted verifier経由で`inspect / validate / plan-change / apply-change`へ接続し、incomplete/corrupt setを部分利用せずexternal XMLと同じsemantic pipelineへ渡す
  - [ ] `ZB-P4.6.7` `CA-CHANGE-001`、`CA-DEST-EXISTS-001`、`CA-BINDING-001`、`CA-CLEANUP-AGGREGATE-001`、`CVF-ABSENT-001`、`CVF-INCOMPLETE-001`、`CVF-CORRUPT-001`、`CVF-COMMITTED-001`、`CVF-EXPECTED-PLAN-MISMATCH-001`、`CU-UNKNOWN-OUTCOME-001`と`BC-APPROVAL-DIVERGENCE-001`、`BC-PLAN-BINDINGS-VALID-001`、`BC-APPLY-PATH-DIVERGENCE-001`、`BC-VERIFY-PATH-DIVERGENCE-001`、`BC-VERIFY-STATE-DIVERGENCE-001`をfailure injection付きでmaterializeし、determinism、legacy回帰、fast/full/build:fullを通す
  - [ ] `ZB-P4.6` 完了確認: loss/unsupportedやmarkerなし/corrupt setを成功にせず、cleanup権限・effects・diagnosticsがcommit point前後で一意であり、public runtimeはP4.9までfail closedのままである
- [ ] `ZB-P4.7` 既存coreの再利用部分を新conformance fixturesで検証する
- [ ] `ZB-P4.8` 選択scopeの残りを一sliceずつ追加し、capability matrixを更新する
- [ ] `ZB-P4.9` versioned single `.mjs`、sources、runtime manifest、SHA-256を生成する
- [ ] `ZB-P4.10` repository外のclean temporary directoryでbundle smokeを実行する
- [ ] `Gate G4` Node参照実装、共通conformance、決定性、安全なpublication、bundle/manifestを検証し、人が承認する

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
