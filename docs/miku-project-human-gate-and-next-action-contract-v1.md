---
title: miku-project human gate and next action contract v1
description: Gate G3で承認された、非対話CLI、human gate、retry・abort・safe next actionの機械契約。
topics:
  - miku-project
  - cli
  - agent
  - approval
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
    path: docs/miku-project-cli-result-contract-v1.md
    label: Gate G3承認済みresult、diagnostics、retryability契約
    checked: 2026-08-11
  - type: local-file
    role: primary
    path: docs/miku-project-change-contract-v1.md
    label: Gate G2承認済みのchange、approval、publication契約
    checked: 2026-08-10
---

# miku-project human gate and next action contract v1

## 文書の位置づけ

これは`ZB-P3.12`の成果物であり、CLIの非対話実行、意味変更前のhuman gate、失敗後のretry / abort / safe next actionを定義する。[CLI result JSON Schema v1](schemas/miku-project-cli-result-v1.schema.json)の`next_action`がmachine-readableな正本であり、この文書はその導出規則とcallerの責務を固定する。

人、shell script、CI、Agent Skills、AI Agentは同じresult fieldだけを読み、message、stderr文言、TTYの有無、会話履歴、作業directoryの暗黙状態から次の操作を推測しない。

## 決定

- 五つのworkflow commandはすべて非対話である。stdinは明示optionで選ばれた一つのartifact入力にだけ使い、prompt、password、確認応答には使わない。
- TTY接続時と非TTY時でcommand semantics、default、status、exit code、artifact publicationを変えない。
- `--yes`、`--force`、「最新planを使う」、会話中の承認を採用する、といったhuman gate短絡経路をv1に設けない。
- CLIはapprovalを生成せず、人のidentityや権限を認証しない。callerがdiffとoutput planを提示し、明示承認後にだけapproval artifactをmaterializeする。
- `next_action`は安全に許される次の行動classであり、実行命令や権限付与ではない。特にretry可能という値は、無制限な自動retryを許可しない。
- validなresultを受信できない場合、message解析による回復を行わない。`apply-change`を起動済みなら同じapplyを再試行せず、指定済みdestinationを`verify-artifact`で確認する。

## resultの`next_action`

全workflow resultは次のobjectを一件持つ。

```json
{
  "next_action": {
    "action": "request-human-approval",
    "command": null,
    "source_retryability": null
  }
}
```

| `action` | `command` | source | 意味 |
| --- | --- | --- | --- |
| `complete` | `null` | success | 要求operationは完了。契約が自動的に要求する後続commandはない |
| `request-human-approval` | `null` | `plan-change` success | diffとoutput planを人へ提示する。承認前はCLI commandを続けない |
| `revise-invocation-or-input` | `null` | `after-input-change` | optionまたは入力artifactを修正するまで停止する |
| `repair-environment` | `null` | `after-environment-change` | permission、容量、runtime、filesystem等を直すまで停止する |
| `replan-and-request-human-approval` | `plan-change` | `after-replan-and-approval` | 古いplan/approvalを再利用せず、current stateから再planして再承認する |
| `verify-artifact` | `verify-artifact` | `apply-change` success | 成功artifactを独立確認する。schema外のunknown outcomeも同じread-only commandで検査する |
| `abort-and-investigate` | `null` | `not-retryable` | 自動続行を中止し、artifact/runtimeを人が調査する |

`command = null`は「callerが自由に次commandを選んでよい」という意味ではない。CLI commandを自動開始できないことを示す。`request-human-approval`で`apply-change`を直接指定しないのは、承認artifactがまだ存在せず、承認拒否という正当な終了もあるためである。

### successからの導出

- `plan-change / succeeded`だけは`request-human-approval`とする。
- `apply-change / succeeded`は`verify-artifact`とし、生成したdestinationと元plan bindingを独立して確認する。
- その他の`succeeded`は`complete`とする。上位workflowが別commandを続けるかはcallerの目的であり、CLIが暗黙pipelineを指示しない。
- `verify-artifact / succeeded`はcommitted setの検証完了を意味し、同じapplyを再実行しない。

### failureからの導出

`rejected / usage-error / runtime-error`では、全diagnosticの`retryability`を次の優先順で集約する。

1. `not-retryable`
2. `after-replan-and-approval`
3. `after-environment-change`
4. `after-input-change`

最上位値を`source_retryability`へ記録し、上表の`action / command`へ一対一に写像する。

JSON Schemaは一つの`source_retryability`と`action / command`の組を検証する。diagnostics配列全体からの最保守値集約は、Node/Java共通conformance testでも検証する。

## human gate

human gateは`plan-change / succeeded`の受信後、approval artifact作成前に一度だけ置く。callerは少なくとも次を同じ確認単位で人へ提示する。

- 対象projectのidentityとbase state digest
- change requestのoperation、対象task UID、before / after値
- semantic diffの全entryとdigest
- output format、canonical destination、artifact member、publication protocol
- preflight normalizationの全entry。loss / unsupportedが空であること
- runtime family/version、artifact/manifest digest、capability profile、fixture suite version
- 「入力と既存pathは変更せず、新規destinationだけを生成する」という副作用境界

表示上の要約や翻訳を加えてよいが、意味変更entry、normalization、destination、runtime bindingを省略して承認を得てはならない。人に提示した値はplan result内の値と一致させ、messageから再構成しない。

human gateの結果は次の三つである。

| 人の判断 | callerの行動 |
| --- | --- |
| 承認 | plan resultのbase/request/diff/output plan digestをそのまま束縛した`miku_project_change_approval/v1`を作り、同じ明示入力群で`apply-change`を一度呼ぶ |
| 拒否 | approvalを作らずworkflowを終了する。`approved = false`をapplyへ渡して記録を代行させない |
| 修正要求 | 新しいchange requestを作って`plan-change`からやり直す。旧planや旧approvalを再利用しない |

approvalはhuman decisionの監査台帳ではなく、確認済みtupleとの機械bindingである。identity認証、署名、組織上の承認権限、監査保存期間はcallerまたは上位systemの責務とし、CLIが保証したと主張しない。

## retryと中止

- v1には、条件を何も変えないautomatic retryを許すactionはない。retryする場合も、入力、環境、plan/approvalのいずれかをresultどおり変更するか、先にartifactをverifyする。
- `revise-invocation-or-input`はprojectを自動修復する許可ではない。修正後はdigestが変わるため、既存request/plan/approvalのbindingを再評価する。
- `repair-environment`は同じdestinationを再利用する許可ではない。destination directoryが作成済みならpublication stateを確認し、新しいplan/destinationが必要か判断する。
- `replan-and-request-human-approval`では古いapprovalを破棄扱いにし、runtime、current state、request、destinationのいずれかが同じに見えても新しいoutput plan digestへ再承認を得る。
- `verify-artifact`では元applyのdestinationを使い、可能なら`--expect-plan-result`も渡す。committedなら成果を回収し、absent / incomplete / corruptなら現在のapply attemptを中止・調査する。利用者が変更をなお望む場合も、調査後に別のplanning workflowとして開始し、旧approvalを再利用しない。
- `abort-and-investigate`ではCLIによるcleanup、修復、alias fallback、別runtime fallbackを自動実行しない。

## valid resultを受信できない場合

launcher failure、signal、stdout切断、result file受信失敗などでschema-validなresultがない場合、`next_action`も存在しない。

- `inspect / validate / plan-change / verify-artifact`は成功を推測せず中止する。caller policyで再実行する前に、外部副作用がなかったことを確認する。
- `apply-change`を起動していないことが確実なら中止する。
- `apply-change`を起動済み、または起動有無が不明なら同じapplyを再実行しない。callerが指定したdestinationに対して独立した`verify-artifact`を実行する。
- runtime manifest/asset integrityを確立できない場合は同じruntimeでprojectを読まず、配布元の固定digestからruntimeを修復する。

unknown outcomeの`verify-artifact`はschema-validな前回resultから導出するのではなく、resultを受信できなかったcallerが適用するschema外failure規則である。schema外failureを成功resultへ捏造しない。

## P3.12 review checklist

- [x] 全workflow commandがTTY非依存の非対話commandである
- [x] `--yes`、prompt、会話履歴、暗黙の最新planでhuman gateを短絡できない
- [x] `plan-change` successだけがhuman gateを要求し、CLIはapprovalを生成しない
- [x] 人へ提示するdiff、normalization、destination、runtime bindingの最小集合が定義されている
- [x] 承認、拒否、修正要求の三経路でapprovalの作成・再利用規則が一意である
- [x] successと全retryabilityから`next_action`を一意に導出できる
- [x] unknown outcomeではapply再試行より`verify-artifact`を優先する
- [x] retry可能性を自動実行権限や無制限retryと混同しない
- [x] valid resultがないschema外failureの安全な回復条件が定義されている
