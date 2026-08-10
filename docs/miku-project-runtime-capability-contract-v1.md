---
title: miku-project runtime capability contract v1
description: Gate G3で承認された、Node/Java共通capability profile、runtime固有extension、選択・fallback境界。
topics:
  - miku-project
  - cli
  - runtime
  - capability
  - specification
category: specification
status: approved
audience:
  - maintainer
  - developer
  - agent
created: 2026-08-10
updated: 2026-08-11
sources:
  - type: local-file
    role: primary
    path: docs/miku-project-cli-contract-v1.md
    label: Gate G3承認済みCLI契約
    checked: 2026-08-11
  - type: local-file
    role: primary
    path: docs/miku-project-format-and-loss-contract-v1.md
    label: Gate G2承認済みformat and loss contract
    checked: 2026-08-10
  - type: local-file
    role: primary
    path: docs/miku-project-change-contract-v1.md
    label: Gate G2承認済みchange contract
    checked: 2026-08-10
  - type: local-file
    role: planning
    path: docs/miku-project-zero-base-implementation-plan-v20260810.md
    label: ゼロベース新仕様適合計画
    checked: 2026-08-10
---

# miku-project runtime capability contract v1

## 文書の位置づけ

これは`ZB-P3.8`の成果物であり、Node CLIとJava CLIが共有するv1 capability、runtime固有extensionの境界、runtime選択時の機械判定を定義する。[CLI contract v1](miku-project-cli-contract-v1.md) と合わせてGate G3の承認済み正本とする。

capabilityのmachine-readableな格納先、artifact/source digest、fixture suite versionとのbindingは[runtime manifest contract v1](miku-project-runtime-manifest-contract-v1.md)と[runtime manifest JSON Schema v1](schemas/miku-project-runtime-manifest-v1.schema.json)で固定する。この文書ではcapability ID、core profile、判定規則を固定し、manifest側はasset配置と配布時検証を固定する。

## 決定

- Node CLIをv1の参照実装とするが、Nodeであること自体をcapabilityとして扱わない。
- Node/Javaのいずれも、共有CLIをv1適合runtimeとして名乗るには`miku-project-cli-core/v1`の全capabilityを実装し、同じconformance corpusに合格しなければならない。
- core capabilityの部分実装を、v1適合runtimeとして公開しない。開発途中のartifactはrelease manifestを持つ適合artifactとして扱わない。
- v1の共有`miku-project` CLI surfaceでは、Node固有・Java固有のcommand、option、result field、diagnostic意味、format、publication挙動を許可しない。
- v1のruntime extension setはNode/Javaとも空とする。将来固有機能が必要になった場合は、既存core semanticsを変更しないversioned extension contractとして別途承認する。
- runtimeの静的capability宣言と、特定の入力・path・filesystemで実行可能かという動的preflight結果を分離する。

この決定により、Javaは「主要commandの一部だけを持つ代替runtime」ではなく、固定済みcore profileへ適合したときに利用可能になる。Nodeは先にreleaseできるが、Node-onlyの挙動を共通契約へ暗黙昇格させない。

## capability ID規則

capability IDはASCII lowercaseの固定文字列とし、次の形を使う。

```text
miku-project.capability.<area>.<name>/v<positive-integer>
```

- IDは意味とversionを一体で識別する。runtime version、product contract version、artifact schema versionの代替にしない。
- 同じIDをNode/Javaで使う場合、その入力、成功条件、failure class、副作用、出力の意味は同じでなければならない。
- 既存IDの意味を破壊的に変更しない。変更が必要なら`/v2`を新設する。
- alias、大文字小文字の揺れ、runtime familyを含むcore IDを設けない。
- 配列へ格納するときはUnicode code pointによる昇順とし、重複を許さない。

## core capability profile

`miku-project-cli-core/v1`は次の九つをすべて含む閉じたprofileである。

| capability ID | 保証する能力 |
| --- | --- |
| `miku-project.capability.apply-change.set-task-percent-complete/v1` | 承認済みC1 requestを再検証し、leaf task一件の進捗変更を生成する |
| `miku-project.capability.format.ms-project-xml-subset.read/v1` | `miku-project-ms-project-xml-subset/v1`をdecode・validateする |
| `miku-project.capability.format.ms-project-xml-subset.write/v1` | semantic stateを同profileへ決定的にencodeし、再decode equivalenceを検証する |
| `miku-project.capability.inspect.project-overview/v1` | valid stateから`project_overview` Projectionを生成する |
| `miku-project.capability.inspect.task-change-context/v1` | validなleaf taskから`task_change_context` Projectionを生成する |
| `miku-project.capability.plan-change.set-task-percent-complete/v1` | C1 requestのdry-run、semantic diff、output planを生成する |
| `miku-project.capability.publication.exclusive-directory-commit-marker/v1` | exclusive directory/file createと`COMMITTED`によるlogical publicationを実装し、対応可否をpreflightする |
| `miku-project.capability.validate.project/v1` | format、semantic invariant、unsupported dataを検証する |
| `miku-project.capability.verify-artifact/v1` | artifact setを`absent / incomplete / committed / corrupt`へ分類し、bindingとdigestを検証する |

`--help`、`--version`、result/diagnostic schema、exit code、determinism、I/O transportはprofileから任意選択する機能ではなく、v1 CLI contract全体への適合条件である。そのため個別capability IDにはしない。

## commandごとの要求capability

callerまたはAgent Skillsは、command名だけでruntimeを選ばず、引数で選んだpurpose、operation、input roleを含めて次の集合を満たすruntimeを選ぶ。

| command | 常に必要 | 条件により追加で必要 |
| --- | --- | --- |
| `validate` | `validate.project`、XML subset read | committed artifact setを読む場合は`verify-artifact` |
| `inspect` | `validate.project`、XML subset read | purposeに応じて`inspect.project-overview`または`inspect.task-change-context`。committed artifact setを読む場合は`verify-artifact` |
| `plan-change` | `validate.project`、XML subset read/write、`plan-change.set-task-percent-complete`、publication protocol | なし。destination固有preflightは別途必須 |
| `apply-change` | `validate.project`、XML subset read/write、`plan-change.set-task-percent-complete`、`apply-change.set-task-percent-complete`、publication protocol | current projectがartifact setなら`verify-artifact`。destination固有preflightは別途必須 |
| `verify-artifact` | `verify-artifact`、XML subset read、publication protocol | `--expect-plan-result`指定時も新しいcapabilityは増えない |

表中の短縮名は前節の`miku-project.capability.*` IDを指す。`apply-change`がplanning capabilityも要求するのは、approvalされたplanを再計算・再検証し、別の意味や出力へ差し替えないためである。

## runtime manifestでの表現

[runtime manifest contract v1](miku-project-runtime-manifest-contract-v1.md)は次のcapability構造を`compatibility.capabilities`へ格納する。完全なmanifest構造とNode/Java exampleは同契約を参照する。

```json
{
  "capabilities": {
    "catalog_version": "1",
    "profiles": ["miku-project-cli-core/v1"],
    "provided": [
      "miku-project.capability.apply-change.set-task-percent-complete/v1",
      "miku-project.capability.format.ms-project-xml-subset.read/v1",
      "miku-project.capability.format.ms-project-xml-subset.write/v1",
      "miku-project.capability.inspect.project-overview/v1",
      "miku-project.capability.inspect.task-change-context/v1",
      "miku-project.capability.plan-change.set-task-percent-complete/v1",
      "miku-project.capability.publication.exclusive-directory-commit-marker/v1",
      "miku-project.capability.validate.project/v1",
      "miku-project.capability.verify-artifact/v1"
    ],
    "extensions": []
  }
}
```

- `profiles`に`miku-project-cli-core/v1`を記録するには、`provided`が上記九件と完全一致しなければならない。
- v1では`extensions`は空配列だけを許可する。Node/Javaの違いは`runtime.family`、artifact、source、version、digestへ記録し、架空のextensionで表現しない。
- JSON Schemaは`provided`を既知capability IDの重複なし部分集合として構造検証する。未知IDや型不正はmanifest invalidである。core capability欠落、canonical順不一致、profileと集合の不一致は後続のcompatibility検証で`runtime.capability-missing`として拒否し、v1適合runtimeとして起動しない。
- release artifact自身の自己申告だけで適合を証明しない。profile宣言に加えて、対応するfixture suite versionのconformance結果をrelease gateで要求する。
- resultとoutput plan/provenanceはruntime manifest digest、capability profile、fixture suite versionへbindingする。巨大なcapability一覧をresultごとに複製しない。

## 静的capabilityと動的preflight

`miku-project.capability.publication.exclusive-directory-commit-marker/v1`は、そのruntimeが必要なAPI、検査、failure処理を実装するという静的な約束である。任意のfilesystemやdestinationでexclusive createを保証できるという意味ではない。

`plan-change`と`apply-change`は選択済みruntimeでdestination parentを検査し、少なくとも次を`supported / unsupported / unknown`のいずれかに判定する。

- directoryをnon-recursiveかつexclusiveに新規作成できる
- 通常fileをexclusiveに新規作成できる
- symlinkや非通常fileを拒否できる
- file handleをcloseした後にmember、size、digestを再検査できる

`unsupported`または`unknown`ではfail closedにし、承認可能なoutput planやcommitted artifactを生成しない。manifestの静的capability不足は`runtime.capability-missing`、選択済みruntimeはprotocolを実装するがdestination環境が非対応の場合は`publication.capability-unsupported`を使う。

## runtime固有extension

v1の共有CLIではNode-only / Java-only extensionを採用しない。とくに次を禁止する。

- 同じcommand/optionでruntimeごとに異なる意味を持たせる
- Javaだけが受理するv1 input field、Nodeだけが出力するv1 result fieldを追加する
- core validationを緩める、unsupported dataを黙って保持・破棄する
- 固有extensionをcore capability不足の代替として扱う
- runtime familyや「最新版」という理由だけで未宣言機能を推測する

将来extensionを採用する場合は`miku-project.runtime-extension.<family>.<name>/vN`のIDを使い、入力、出力、副作用、diagnostics、互換性、conformance fixtureを持つ別契約として登録する。extensionを使わないcore invocationの結果を変えてはならない。未知extensionは自動実行せず、そのextensionを明示的に要求したworkflowだけが対応runtimeを選ぶ。

共有v1 surfaceに載せない実験的なJava/Node機能は、別entrypointまたは別artifactとして隔離し、`miku-project-cli-core/v1`適合の根拠にしない。

## runtime選択とfallback境界

- runtime選択はcommand開始前に行い、product contract、capability catalog/profile、artifact digestを検証する。
- Nodeが参照実装であることは開発順の判断であり、各操作で常にNodeを優先するというruntime選択規則ではない。具体的な優先順はAgent Skillsの`ZB-P6.4`で決める。
- 選択済みruntimeに必要capabilityがない場合、CLIを試行して人間向けmessageからfallbackしない。manifest検査で候補から除外する。
- 一つのcommand実行中に別runtimeへ切り替えない。失敗後に別runtimeを使う場合は新しいinvocationとして扱う。
- `plan-change`後にruntimeを変更する場合、output plan、XML byte digest、runtime bindingが変わり得るため、新runtimeでplanを作り直し、人間確認とapprovalもやり直す。
- `apply-change`はapprovalにbindされたruntimeと完全一致しなければならず、NodeからJava、JavaからNodeへのfallbackを行わない。

## cross-runtime target matrix

これは実装済み状態ではなく、各release gateが要求するv1 targetである。`planned`を適合済みと読んではならない。

| capability | Node G4 | Java G5 | 比較 |
| --- | --- | --- | --- |
| apply-change / set task percent | required・planned | required・planned | 同じsemantic diff、post-state、publication outcome |
| XML subset read | required・planned | required・planned | 同じvalid/rejected分類とsemantic state |
| XML subset write | required・planned | required・planned | semantic equivalent。契約で指定した箇所だけbyte比較 |
| inspect / project overview | required・planned | required・planned | golden Projectionと一致 |
| inspect / task change context | required・planned | required・planned | golden Projectionと一致 |
| plan-change / set task percent | required・planned | required・planned | semantic artifactは一致。runtime-bound output planのbyte digestはruntime内で決定的 |
| publication protocol | required・planned | required・planned | 同じstate分類、commit point、diagnostic意味 |
| validate project | required・planned | required・planned | 同じrule ID、diagnostic code、status |
| verify artifact | required・planned | required・planned | 同じ`absent / incomplete / committed / corrupt`分類 |
| runtime extensions | none | none | v1では空集合 |

実装後は`planned`を、conformance evidenceに基づく`pass / fail / not-tested`へ更新する。`unsupported`はcore profileのrelease値として許可しない。

## 用語の衝突を避ける

Projectionに記録する`capability.unsupported_data`は、読み取ったproject dataの対応状況を人・Agentへ示すdomain summaryである。この文書の**runtime capability**やruntime選択には使わない。runtime選択の正本はruntime manifestの`capabilities`である。

## P3.8 review checklist

- [x] Node参照実装とruntime capabilityを別概念として定義した
- [x] `miku-project-cli-core/v1`の必須capabilityが閉じた集合になっている
- [x] Node/Javaの部分実装をv1適合runtimeとして扱わない
- [x] v1のNode-only / Java-only extensionが空集合である
- [x] 静的runtime capabilityとdestination固有preflightを分離した
- [x] manifest不足とfilesystem非対応のdiagnosticを区別した
- [x] runtime変更時にplanとapprovalを作り直す境界を定義した
- [x] [conformance corpus v1](miku-project-conformance-corpus-v1.md)へcommand/capability別conformance対象を渡した
- [x] runtime manifest schemaへcapability profileの論理構造とbinding要件が反映されている
