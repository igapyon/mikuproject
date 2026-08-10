---
title: miku-project change contract v1
description: Gate G2で承認された、C1のchange request、diff、human gate、commit markerによる論理publication、provenance。
topics:
  - miku-project
  - cli
  - agent-skills
  - specification
  - change-management
category: specification
status: approved
audience:
  - maintainer
  - developer
  - agent
created: 2026-08-10
updated: 2026-08-10
sources:
  - type: local-file
    role: primary
    path: docs/miku-project-semantic-contract-v1.md
    label: Gate G1承認済みsemantic contract
    checked: 2026-08-10
  - type: local-file
    role: primary
    path: docs/miku-project-format-and-loss-contract-v1.md
    label: G2 format and loss contract
    checked: 2026-08-10
  - type: local-file
    role: scenario
    path: docs/miku-project-zero-base-scenarios-v1.md
    label: G0承認済みC1シナリオ
    checked: 2026-08-10
---

# miku-project change contract v1

## 文書の位置づけ

これは `Gate G2` で承認されたchange contractである。C1の「leaf task一件の`percentComplete`更新」を、AI Agent、人、shell script、CIがhidden stateなしに安全に扱えるartifact protocolとして定義する。G3で確定した具体的なcommand、flag、I/O/publication境界は [CLI contract v1](miku-project-cli-contract-v1.md)、result/diagnostics/statusは [CLI result and diagnostics contract v1](miku-project-cli-result-contract-v1.md) に従う。

## G2の決定

- C1にwhole-state replacement、merge、自由JSON編集は存在しない。新しい外部XMLを読むことはR1のread-only処理であり、既存stateを置換する変更operationではない。
- v1のchange requestはちょうど一件の`set_task_percent_complete` operationを持つ。selectorはtask UIDだけであり、summary taskや同値更新は受け入れない。
- applyの前に、CLIはrequestを現在のsemantic stateへvalidateし、semantic diffを生成する。さらにplanned stateをMS Project XMLへpreflight encode/redecodeし、normalization、loss、unsupported、出力先を`output plan`として確定する。
- 人によるhuman gateはCLIの外側で行う。applyにはbase state、request、diff、output planへ束縛したapproval artifactを必須にし、人が意味変更と公開先の両方を確認したことをworkflow上の前提とする。
- applyはrequest作成時のstateではなく、apply時点で再読取り・再検証したstateに対して行う。current stateからplanned state、semantic diff、output planを再生成し、state digest、precondition、request digest、diff digest、output plan digestのすべてがapprovalと一致しなければ、再計画・再承認を要求する。
- apply成功後はpost-apply validateとXML encode/redecode equivalenceをもう一度行う。いずれかに失敗すればartifact setを公開しない。

## digestとbinding

| digest | 対象 | 役割 |
| --- | --- | --- |
| `state_digest` | `miku_project_semantic_state/v1` | requestが作られた完全な意味状態を束縛する |
| `change_request_digest` | `miku_project_change_request/v1` | diffとapprovalが同一requestを対象とすることを証明する |
| `semantic_diff_digest` | `miku_project_semantic_diff/v1` | human gateで確認した予定変更とapplyを束縛する |
| `output_plan_digest` | `miku_project_output_plan/v1` | 出力形式、artifact set公開先、preflight normalizationをapprovalへ束縛する |

すべてのdigest algorithmは`sha-256`とする。semantic state/request/diffは[conformance corpus v1](miku-project-conformance-corpus-v1.md)のcanonical JSON serializationを用い、同じ意味artifactがNode/Javaで同じdigestになることをG3のconformance対象とする。

change request、semantic diff、output plan、approval、provenanceのmachine-readableなrequired fieldとclosed shapeは[artifact JSON Schema v1](schemas/miku-project-artifacts-v1.schema.json)を正本とする。別artifact間のdigest、runtime、destination bindingは[result contractの`RB-001`〜`RB-012`](miku-project-cli-result-contract-v1.md#artifact-schema%E3%81%A8cross-artifact-binding)を適用し、schema適合だけで承認・apply可能とは判断しない。

output planは実際にpreflight生成したXML byte digest、destination、runtime bindingを含むため、runtime間で同一digestになることを要求しない。選択済みruntimeが同じ入力とoptionから決定的に再生成したoutput planとの一致を要求する。human gateはruntime選択とpreflightの後に置き、別runtimeへ切り替える場合は新しいoutput planとapprovalを作る。runtime bindingはfamily、version、artifact digest、manifest digest、capability profile、fixture suite versionからなり、値と検証規則は[runtime manifest contract v1](miku-project-runtime-manifest-contract-v1.md)を正本とする。

## change request schema

### validな最小例

```json
{
  "kind": "miku_project_change_request",
  "schema_version": "1",
  "semantic_contract_version": "1",
  "base": {
    "state_digest": {
      "algorithm": "sha-256",
      "value": "<64-lowercase-hex>"
    }
  },
  "operations": [
    {
      "kind": "set_task_percent_complete",
      "target": { "task_uid": "2" },
      "preconditions": { "expected_percent_complete": 0 },
      "value": { "percent_complete": 50 }
    }
  ]
}
```

| field | 規則 |
| --- | --- |
| `kind` / `schema_version` / `semantic_contract_version` | 上記の固定値を必須とする。未知versionはerror。 |
| `base.state_digest` | 必須。apply時に再計算したcurrent state digestと完全一致しなければならない。 |
| `operations` | 長さ1のarrayだけを許可する。順序による複数operation解釈はしない。 |
| `target.task_uid` | stable identity token。存在し、leaf taskでなければならない。 |
| `expected_percent_complete` | current valueとの完全一致を必須とする0..100の整数。 |
| `value.percent_complete` | 0..100の整数かつexpected valueと異なる値。 |

次は`invalid`または`unsupported-error`であり、warning・ignored change・暗黙補完へ落とさない。

| 入力 | 結果 |
| --- | --- |
| `state_digest`なし、unknown kind/version、operationが0件または2件以上 | `invalid` |
| target不存在、summary、expected value不一致、同値更新、範囲外整数 | `invalid` |
| `set_task_percent_complete`以外のoperation、field追加・削除・移動、dependency/resource/assignment/calendar/project変更 | `unsupported-error` |
| G1でunsupportedのdataを含むstateへのrequest | `unsupported-error` |

## semantic diff、output plan、approval artifact

validなrequestをcurrent stateにdry-run適用したときだけ、CLIはsemantic diffを生成する。diffは最低限、base/proposed state digest、change request digest、変更対象UID、変更前後の`percentComplete`、non-target field/collectionがpreservedである検証結果、semantic state段階のloss/normalization/unsupportedの記録を持つ。XML encodeで生じるnormalizationは、後続のoutput planへ記録する。

```json
{
  "kind": "miku_project_semantic_diff",
  "schema_version": "1",
  "semantic_contract_version": "1",
  "base_state_digest": { "algorithm": "sha-256", "value": "<64-lowercase-hex>" },
  "proposed_state_digest": { "algorithm": "sha-256", "value": "<64-lowercase-hex>" },
  "change_request_digest": { "algorithm": "sha-256", "value": "<64-lowercase-hex>" },
  "changes": [
    {
      "kind": "set_task_percent_complete",
      "task_uid": "2",
      "before": 0,
      "after": 50
    }
  ],
  "preservation": { "semantic_equivalent_except_changes": true },
  "provenance": { "losses": [], "normalizations": [], "unsupported": [] }
}
```

diff生成後、CLIはplanned stateをpreflight encode/redecodeし、次のoutput planを生成する。`destination.path`はcallerが`--destination`へ渡した値を、既存parentのreal pathと未使用basenameからcanonical absolute pathへした値であり、人へ提示する。v1は`write_mode = "create-new-directory"`だけを許可し、既存path、input artifact setと同じpathまたはその子孫、direct symlink、解決不能pathを拒否する。

```json
{
  "kind": "miku_project_output_plan",
  "schema_version": "1",
  "semantic_contract_version": "1",
  "base_state_digest": { "algorithm": "sha-256", "value": "<64-lowercase-hex>" },
  "change_request_digest": { "algorithm": "sha-256", "value": "<64-lowercase-hex>" },
  "semantic_diff_digest": { "algorithm": "sha-256", "value": "<64-lowercase-hex>" },
  "runtime": {
    "family": "<node-or-java>",
    "version": "<runtime-version>",
    "artifact_digest": { "algorithm": "sha-256", "value": "<64-lowercase-hex>" },
    "manifest_digest": { "algorithm": "sha-256", "value": "<64-lowercase-hex>" },
    "capability_profile": "miku-project-cli-core/v1",
    "fixture_suite_version": "1"
  },
  "output": {
    "format": "ms-project-xml",
    "format_profile": "miku-project-ms-project-xml-subset/v1",
    "adapter": "ms-project-xml-adapter/v1",
    "artifact_set": "miku_project_artifact_set/v1",
    "destination": { "path": "<absolute-canonical-destination>", "write_mode": "create-new-directory" },
    "publication": {
      "strategy": "exclusive-directory-commit-marker/v1",
      "directory_create": "exclusive",
      "commit_marker": { "path": "COMMITTED", "create_mode": "exclusive-empty-file" },
      "runtime_filesystem_supported": true
    },
    "members": [
      { "role": "external_project", "path": "project.xml" },
      { "role": "provenance", "path": "provenance.json" },
      { "role": "commit_marker", "path": "COMMITTED", "size": 0 }
    ]
  },
  "preflight": {
    "proposed_state_digest": { "algorithm": "sha-256", "value": "<64-lowercase-hex>" },
    "project_artifact_digest": { "algorithm": "sha-256", "value": "<64-lowercase-hex>" },
    "normalizations": [
      {
        "code": "xml.link-lag-legacy-zero-normalized",
        "path": "dependencies[predecessor_uid=1,successor_uid=2,type=FS].lag",
        "before": "PT0H0M0S",
        "after": "LinkLag=0;LagFormat=3"
      },
      {
        "code": "xml.supported-child-order-canonicalized",
        "path": "Project",
        "before": ["Name", "CurrentDate", "StartDate", "FinishDate", "ScheduleFromStart", "CalendarUID", "Calendars", "Tasks", "Resources", "Assignments"],
        "after": ["Name", "ScheduleFromStart", "StartDate", "FinishDate", "CalendarUID", "CurrentDate", "Calendars", "Tasks", "Resources", "Assignments"]
      }
    ],
    "losses": [],
    "unsupported": []
  }
}
```

`preflight.losses`または`preflight.unsupported`が非空、再decode後stateがproposed stateとsemantic equivalentでない、destinationが既存またはunsafe、または選択したruntime/filesystemがdirectoryとfileの排他的な新規作成を保証できない場合は、output planを承認可能な成功artifactとして公開しない。`runtime_filesystem_supported`はcallerの自己申告ではなく、[runtime capability contract v1](miku-project-runtime-capability-contract-v1.md)の静的runtime capabilityとdestination固有preflightからCLIが確定する。normalizationは意味を変えないものだけを列挙し、人がdiffと一緒に確認できるようにする。

human gateの後、callerは次のapproval artifactをCLIへ渡す。これは人のidentityを認証する機構ではなく、CLIが「確認されたdiff/output planと異なるrequest/state/outputをapplyしない」ための機械的bindingである。

```json
{
  "kind": "miku_project_change_approval",
  "schema_version": "1",
  "semantic_contract_version": "1",
  "approved": true,
  "base_state_digest": { "algorithm": "sha-256", "value": "<64-lowercase-hex>" },
  "change_request_digest": { "algorithm": "sha-256", "value": "<64-lowercase-hex>" },
  "semantic_diff_digest": { "algorithm": "sha-256", "value": "<64-lowercase-hex>" },
  "output_plan_digest": { "algorithm": "sha-256", "value": "<64-lowercase-hex>" }
}
```

`approved = false`、digestの欠落、不一致、未知versionはapplyを許可しない。approval artifactは他のrequest、別state、再計画後のdiff、別の出力先へ再利用できない。

CLIはapprovalの作成者が人であることを認証しない。Agent Skillsまたは呼び出し側は、diffとoutput planを人へ提示し、明示的な承認を得た後だけapproval artifactをmaterializeしなければならない。approvalは束縛したtupleに対する再利用可能なattestationであり、hidden ledgerによる単回使用を主張しない。同じapprovalを再送しても、最初の成功後はdestinationが既存になるため`create-new-directory`検証でpublishを拒否する。

## C1 protocol

```text
external XML
  → decode / semantic validate
  → semantic state + state digest
  → purpose別Projection
  → change request
  → request validate + dry-run apply + semantic diff
  → preflight encode / re-decode + output plan
  → human gate（CLI外、diff + output plan）+ approval artifact
  → stateを再decode / digest再照合
  → dry-run apply / diff / output planを再生成してdigest再照合
  → planned stateへapply / post-apply validate
  → destination directoryを排他的に新規作成
  → project.xmlをencode / re-decode equivalence検証
  → provenance.jsonを生成し、member / schema / digestを再検証
  → 空のCOMMITTEDを排他的に新規作成して論理publish
```

| 段階 | 入力 | 公開してよい出力 | 失敗時 |
| --- | --- | --- | --- |
| R1 inspect / validate | external XML | 構造化した検査情報。validならProjection候補 | inputを変更しない。unsupported/invalidを明示 |
| Projection | valid semantic state | purpose別Projection | 全stateや成功編集用Projectionを公開しない |
| request validate / diff | state + change request | valid結果とsemantic diff | planned state、外部XMLを公開しない |
| output preflight | planned state + destination | output plan | loss/unsupported/unsafe destinationなら承認へ進めない |
| human gate | diff + output plan | approval artifact | 中止ならapplyしない。CLI自体は人を認証しない |
| apply / publish | current XML + request + diff + output plan + approval | committed artifact set directory | 失敗時はsuccess artifactを返さない。markerなしのincomplete directoryが残り得る |

## commit markerと非破壊規則

- CLIは入力external artifact、change request、diff、output plan、approvalを変更しない。
- apply前に、current stateからplanned state、semantic diff、output planを再生成し、current state digest、request digest、diff digest、output plan digest、approval binding、G1 validationをすべて再確認する。diff生成後にsourceまたはdestination状態が変わっていればapplyしない。
- applyは、存在しないfinal destinationをnon-recursiveかつexclusiveなdirectory createで予約する。既存path、symlink、同時実行による競合はerrorとし、既存entryを削除・置換・再利用しない。directory createの成功後は、その実行だけがdirectory内を生成対象として扱う。hostileなlocal processによる改変への防御はv1のtrust boundary外だが、member type、symlink、schema、digestの検証は省略しない。
- CLIは予約したdirectory内へ`project.xml`と`provenance.json`を生成し、fileをcloseした後にXML再decode equivalence、provenance schema、member一覧、digestを再検証する。この段階のdirectoryはincompleteであり、成功artifactとして返さない。
- 最後のcommit操作として、通常fileかつ0 byteの`COMMITTED`をexclusive createする。markerが既存、symlink、非0 byte、またはexclusive createが保証できない場合はerrorとする。marker作成後に三memberとdigestを再確認してからsuccess resultを返す。
- directoryの排他的作成とmarkerの排他的作成を保証できないruntime/filesystemはv1非対応とし、preflightでfail closedにする。静的・動的capabilityの境界は[runtime capability contract v1](miku-project-runtime-capability-contract-v1.md)、Node/Javaの具体APIは各実装段階で固定する。
- marker作成前に失敗した場合、CLIは自身が今回予約したdirectoryだけをbest-effort cleanupする。cleanupできなければmarkerなしのincomplete directoryを残し、structured diagnosticsにpathとcleanup statusを記録する。incomplete directoryをartifact output欄へ入れず、後続実行で暗黙削除・再利用しない。
- `COMMITTED`が存在しても、memberが三file以外、markerが0 byteの通常fileでない、provenance/XML/schema/digest検証が失敗する場合はcorruptであり利用禁止とする。部分利用、自動修復、markerの作り直しを行わない。
- ここでいうatomicityは、exclusive createしたcommit markerを境界にした**論理publication**である。三fileのfilesystem上の同時可視性や、電源断後のdurabilityまではv1で約束しない。G3ではhandle closeまでを要求し、file/directoryの`fsync`はv1保証外とした。中断後は`verify-artifact`で再判定する。
- v1 C1には上書きoptionを設けない。既存artifact setの採用・削除・置換はcallerの別行為であり、CLIは入力XMLも既存outputも変更しない。

## provenance、normalization、unsupported

成功したartifact setは次の`provenance.json`を必須memberとして持つ。

```json
{
  "kind": "miku_project_provenance",
  "schema_version": "1",
  "semantic_contract_version": "1",
  "format_profile": "miku-project-ms-project-xml-subset/v1",
  "adapter": "ms-project-xml-adapter/v1",
  "artifact_set": {
    "kind": "miku_project_artifact_set",
    "schema_version": "1",
    "publication_protocol": "exclusive-directory-commit-marker/v1",
    "commit_marker": "COMMITTED"
  },
  "runtime": {
    "family": "<node-or-java>",
    "version": "<runtime-version>",
    "artifact_digest": { "algorithm": "sha-256", "value": "<64-lowercase-hex>" },
    "manifest_digest": { "algorithm": "sha-256", "value": "<64-lowercase-hex>" },
    "capability_profile": "miku-project-cli-core/v1",
    "fixture_suite_version": "1"
  },
  "input": {
    "format": "ms-project-xml",
    "format_profile": "miku-project-ms-project-xml-subset/v1",
    "artifact_digest": { "algorithm": "sha-256", "value": "<64-lowercase-hex>" },
    "state_digest": { "algorithm": "sha-256", "value": "<64-lowercase-hex>" }
  },
  "change": {
    "change_request_digest": { "algorithm": "sha-256", "value": "<64-lowercase-hex>" },
    "semantic_diff_digest": { "algorithm": "sha-256", "value": "<64-lowercase-hex>" },
    "output_plan_digest": { "algorithm": "sha-256", "value": "<64-lowercase-hex>" },
    "target_task_uid": "2",
    "before_percent_complete": 0,
    "after_percent_complete": 50
  },
  "output": {
    "format": "ms-project-xml",
    "format_profile": "miku-project-ms-project-xml-subset/v1",
    "path": "project.xml",
    "artifact_digest": { "algorithm": "sha-256", "value": "<64-lowercase-hex>" },
    "state_digest": { "algorithm": "sha-256", "value": "<64-lowercase-hex>" }
  },
  "transformations": ["decode", "validate", "dry-run-apply", "diff", "preflight-encode", "preflight-redecode", "approval-check", "apply", "reserve-output-directory", "encode", "redecode-validate", "write-provenance", "commit-marker"],
  "normalizations": [
    {
      "code": "xml.link-lag-legacy-zero-normalized",
      "path": "dependencies[predecessor_uid=1,successor_uid=2,type=FS].lag",
      "before": "PT0H0M0S",
      "after": "LinkLag=0;LagFormat=3"
    },
    {
      "code": "xml.supported-child-order-canonicalized",
      "path": "Project",
      "before": ["Name", "CurrentDate", "StartDate", "FinishDate", "ScheduleFromStart", "CalendarUID", "Calendars", "Tasks", "Resources", "Assignments"],
      "after": ["Name", "ScheduleFromStart", "StartDate", "FinishDate", "CalendarUID", "CurrentDate", "Calendars", "Tasks", "Resources", "Assignments"]
    }
  ],
  "losses": [],
  "unsupported": []
}
```

このschemaの各categoryは次を記録する。

| category | recordする内容 |
| --- | --- |
| contract | semantic contract version、artifact schema version、format profile ID、adapter ID、publication protocol、commit marker、runtime family/version、artifact/manifest digest、capability profile、fixture suite version |
| input | role、format ID、digest、source state digest（該当時） |
| transformation | decode、validate、project、dry-run apply、diff、preflight encode/redecode、approval check、apply、output directory reservation、encode、redecode validation、provenance write、commit markerの順序。incomplete directory内のprovenanceは監査記録として信頼せず、committed setだけがmarker実行済みの記録として有効 |
| change | change request digest、semantic diff digest、output plan digest、target UID、before/after value |
| normalization | `code / path / before / after`を持つ、XML字句、container、outline、orderなど意味を変えない変更。配列は`code`、`path`順に決定的に並べる |
| loss | v1成功経路では空array。値があれば成功としてpublishしない |
| unsupported | 検出したdomain/field/path。値があればC1はsuccessにならない |

unknown fieldを無視する、allowlist外operationを部分適用する、unsupported dataを削除して継続する、digest不一致をhuman messageで補う、といった経路は存在しない。

## G2 review checklist

2026-08-10のG2文書横断reviewと人による`Gate G2`承認で、次を確認した。

- [x] semantic stateがinternal IRであり、external XML・workbook JSON・Projectionを正本にしていない
- [x] 全artifactにrole、寿命、schema version、producer/consumerがある
- [x] XML read/write、Projection、change request、diff、output plan、approval、provenanceのloss/unsupported規則が対応づいている
- [x] C1がwhole-state replacementを受け入れず、request/diff/output plan/approval/current stateをdigestで束縛する
- [x] human gateがsemantic diff、normalization、出力先を確認でき、CLIが人を認証しないtrust boundaryも明記されている
- [x] human gate後のapplyが再validate・新規directoryのexclusive reservation・commit markerによる論理publish・non-destructive outputを要求する
- [x] directory/fileのexclusive createを保証できないruntime/filesystemをpreflightでfail closedにする
- [x] incomplete/corrupt/committedの判定と、失敗時に残り得るincomplete directoryの扱いが一意である
- [x] provenance schema、artifact set directory、commit markerによる論理的な公開境界が一意である
- [x] approvalの単回使用をhidden stateなしに強制するという矛盾がない
- [x] Agentが会話履歴や人間向けmessageの解析なしに次のartifactを選べる
