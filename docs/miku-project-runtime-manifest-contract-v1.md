---
title: miku-project runtime manifest contract v1
description: Gate G3で承認された、runtime/source artifact、契約互換性、capability、fixture suite、SHA-256のprovenance。
topics:
  - miku-project
  - cli
  - runtime
  - provenance
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
    path: docs/miku-project-runtime-capability-contract-v1.md
    label: Gate G3承認済みruntime capability契約
    checked: 2026-08-11
  - type: local-file
    role: primary
    path: docs/miku-project-conformance-corpus-v1.md
    label: Gate G3承認済みconformance corpus
    checked: 2026-08-11
  - type: local-file
    role: primary
    path: docs/miku-project-cli-result-contract-v1.md
    label: Gate G3承認済みresult/diagnostics契約
    checked: 2026-08-11
  - type: local-file
    role: current-state
    path: scripts/build-cli-bundle.mjs
    label: 現行single-MJS/source archive buildの再利用証拠
    checked: 2026-08-10
---

# miku-project runtime manifest contract v1

## 文書の位置づけ

これは`ZB-P3.10`の成果物であり、Node/Java runtime releaseの互換契約、実行asset、source archive、capability、conformance corpus、由来、SHA-256を一つのmachine-readable manifestへ束縛する。

正本schemaは [runtime manifest JSON Schema v1](schemas/miku-project-runtime-manifest-v1.schema.json) である。manifestの固定file名は`runtime-manifest.json`とし、Node/Javaの各runtime bundle directoryに一件だけ置く。

## 目的

版番号の一致だけではruntime互換性を判断しない。次を別々の値として記録し、Agent Skills、shell、CIが人間向けfilenameや「最新版」という推測に頼らず検証できるようにする。

- product release version
- product / semantic / format / change / CLI contract version
- semantic / exchange artifact、result、diagnostic schemaとdiagnostic catalog version
- runtime family、役割、runtime version
- capability catalog、core profile、provided capability、extension
- fixture suite versionとcorpus digest
- 実行assetとsource archiveの正確なbasename、media type、size、SHA-256
- contract sourceとruntime sourceのrepository、revision、tag
- Javaが適合対象とした固定Node参照runtimeのversionとmanifest digest

## bundle layout

NodeとJavaは別のbundle directoryを持つ。consumerはdirectory内の固定名manifestだけを起点にし、glob、directory走査順、mtime、lexicographicな「最新filename」を使わない。

```text
runtime/node/
├── runtime-manifest.json
├── miku-project-node-1.0.0.mjs
└── miku-project-node-1.0.0-sources.tgz

runtime/java/
├── runtime-manifest.json
├── miku-project-java-1.0.0.jar
└── miku-project-java-1.0.0-sources.tgz
```

- manifestのartifact pathは同じdirectory直下のbasenameだけを許可する。absolute path、separator、`.`、`..`、symlinkを許可しない。
- `artifacts.executable`だけを実行する。`artifacts.sources`はreview、license、traceability用であり、fallback実行や動的compileへ使わない。
- Node executableはsingle `.mjs`、Java executableはstandalone fat `.jar`、sourceは決定論的`.tgz`を要求する。
- manifest、asset、sourceは通常fileでなければならない。

## manifestの論理構造

例は次を参照する。

- [Node reference runtime manifest example](examples/runtime-manifest-v1/node-runtime-manifest.example.json)
- [Java conforming runtime manifest example](examples/runtime-manifest-v1/java-runtime-manifest.example.json)

主要fieldの意味は次のとおりである。

| field | 意味 |
| --- | --- |
| `product.release_version` | 利用者に公開するmiku-project release version。runtime versionとは別値 |
| `product.*_contract_version` | runtimeが実装する承認済み契約version |
| `product.artifact_schema` | semantic state、Projection、request、diff、plan、approval、provenanceのschema ID `miku_project_artifacts/v1` |
| `runtime.family` | `node / java` |
| `runtime.role` | Nodeは`reference`、Javaは`conforming` |
| `runtime.version` | runtime artifact自身のSemVer |
| `runtime.launcher` | manifestが指すassetを起動する固定方式。`node`または`java-jar` |
| `compatibility.capabilities` | catalog既知IDの宣言集合と空extension。schemaは既知subsetを受理し、core profile適合判定が九件のcanonical集合・順序を要求する |
| `compatibility.conformance` | fixture suite versionと、使用したcorpus全体のdigest |
| `artifacts.executable` | 実際に起動する一fileのbasename、media type、size、digest |
| `artifacts.sources` | executableへ対応するsource archive |
| `source.contract` | product contractを得たmiku-project repository revision |
| `source.runtime` | executable/source archiveをbuildしたrepository revision |
| `reference_runtime` | Javaが適合したNode manifest。Nodeでは`null` |

manifestはbuild timestamp、hostname、absolute build path、runner IDを持たない。同じsource、toolchain、入力から同じartifactを作れる場合、manifest byte列も決定的でなければならない。

## versionの関係

- `product.release_version = 1.0.0`と`runtime.version = 1.0.0`が同じでも、それだけで互換とは判定しない。
- product contract、schema、capability profile、fixture suiteの全値がcaller要求と一致し、digest検証が成功して初めて候補runtimeになる。
- NodeとJavaのruntime versionは独立して更新できる。Java manifestは`reference_runtime.manifest_digest`で、適合対象にした固定Node releaseを明示する。
- Node manifestの`reference_runtime`は`null`である。Node runtime自身を再帰参照しない。
- `source.contract.tag`は`product.release_version`、Nodeの`source.runtime.tag`はNode runtime version、Javaの`source.runtime.tag`はJava runtime versionに対応させる。schemaだけでは値間同値を完全検査できないためrelease validationで照合する。
- executable名はNodeで`miku-project-node-<runtime.version>.mjs`、Javaで`miku-project-java-<runtime.version>.jar`、source名は対応するbasenameのversion部分を共有した`miku-project-<family>-<runtime.version>-sources.tgz`とする。JSON Schemaは安全なbasenameと拡張子を検査し、versionとの文字列一致はrelease validationで検査する。

## fixture corpus digest

`compatibility.conformance.corpus_digest`は`testdata/conformance/v1/`の内容を次の方法で束縛する。

1. directory以下の通常fileを再帰列挙し、symlinkと非通常fileを拒否する。
2. rootからのPOSIX relative pathをUnicode code point昇順に並べる。
3. 各fileのraw byte SHA-256をlowercase hexで計算する。
4. 各fileについて`<hex><two spaces><relative-path><LF>`をUTF-8で連結する。
5. 連結byte列のSHA-256をcorpus digestとする。

corpus内へcorpus digest fileを置かず、自己参照を避ける。P4/P5でsemantic catalogをmaterializeして内容が変われば、fixture suite versionを維持する場合でもcorpus digestは変わる。release manifestは実際にtestしたdigestだけを記録する。

## manifest生成と外側のpin

生成順は次で固定する。

1. executableとsource archiveを決定論的に生成する。
2. 両fileのsizeとSHA-256を計算する。
3. conformance corpus digestを計算し、その内容でrelease testを実行する。
4. source repository revision/tagと全compatibility fieldを入れてmanifestを生成する。
5. manifestを[conformance corpusのcanonical JSON規則](miku-project-conformance-corpus-v1.md)でserializeし、末尾LF一件を付ける。
6. manifest file全体のSHA-256を計算し、Release checksumまたはAgent Skills側のlockへ記録する。

manifest自身へmanifest digestを入れない。manifestが自分自身をhashすると循環するためである。信頼境界は次の二段階になる。

```text
Skills lock / Release checksum
  └─ runtime-manifest.json のSHA-256
       ├─ executableのbasename / size / SHA-256
       ├─ sourcesのbasename / size / SHA-256
       ├─ contract / capability / fixture corpus
       └─ source revision / tag
```

manifestとartifactを同時に改ざんしてdigestを合わせても、外側でpinしたmanifest digestと一致しないため拒否できる。外側のpinを持たず、manifestの自己申告だけをprovenance検証と呼ばない。

## self-validationとconsumer validation

検証責務はruntime自身とconsumerの二層に分ける。

- runtime自身はadjacent manifestを起点にschema、version/capability、executable/sourceのpath・size・digestを検査し、project inputを読む前にbundle内部のbindingを確立する。
- Agent Skillsやrelease smokeのようにtrust anchorを持つconsumerは、runtime起動前にmanifest raw digestをSkills lockまたはRelease checksumと照合し、起動後にresult bindingも照合する。
- direct shell利用者はRelease checksumを配布元から取得して外側で照合できる。照合しない実行はbundle内部の整合性を確認できても、配布元authenticityを証明したことにはならない。

各operation開始前の検証順は次である。

1. 選択policyが指した固定`runtime-manifest.json`を開く。manifest自体が通常fileかつ非symlinkであることを確認する。
2. trust anchorを持つconsumerは、Skills lockまたはRelease checksumが持つmanifest SHA-256とraw byte digestを比較する。runtime自身はraw digestを計算し、result binding用に保持する。
3. UTF-8、BOM、JSON duplicate key、runtime manifest schemaを検証する。
4. product/contract/schema/catalog/capability/fixture suiteがworkflow要求と一致することを確認する。
5. manifest directoryとbasenameからexecutable/source pathを解決し、directory外へ出ないこと、通常file、非symlink、size一致、family/versionから導出した固定filenameとの一致を確認する。
6. executableとsourceのSHA-256を毎operation開始前に検証する。
7. `runtime.launcher`と`artifacts.executable.path`だけからcommandを組み立てる。
8. CLI resultのruntime bindingが、検証済みmanifestと完全一致することを確認する。

network download、source checkout、package manager、`vendor/`、PATH上の同名command、runtime directory内の別versionを暗黙利用しない。候補runtimeを変更する場合は新しいmanifest検証として最初から行う。

## CLI resultとartifact binding

CLI resultの`runtime`は次を持つ。

```json
{
  "binding_status": "verified",
  "family": "node",
  "version": "1.0.0",
  "artifact_digest": { "algorithm": "sha-256", "value": "<64-lowercase-hex>" },
  "manifest_digest": { "algorithm": "sha-256", "value": "<64-lowercase-hex>" },
  "capability_profile": "miku-project-cli-core/v1",
  "fixture_suite_version": "1"
}
```

- `verified`では上記全fieldを必須かつnon-nullにする。
- `verified`はruntime自身がbundle内部のmanifest/asset bindingを検証したことを示す。外部trust anchorによる配布元authenticityまでCLIが自己証明する値ではなく、Agent Skills等のconsumerは外側のmanifest pinとの照合結果を別途成立させる。
- manifest/assetを検証できずstructured runtime errorを返す場合だけ`binding_status = unverified`を許可し、digest/profile/fixture fieldは確定できた値または`null`にする。
- `succeeded`とdomain/validationによる`rejected`は必ず`verified`である。未検証runtimeでprojectを読まない。
- output planとprovenanceにもfamily/version/artifact digestに加えmanifest digest、capability profile、fixture suite versionを記録する。approvalはこれらを含むoutput plan digestへ束縛し、runtime bindingを間接的かつ改変不能に束縛する。
- `plan-change`後にmanifest digestが変われば、同じfamily/version/asset digestに見えても別runtime bindingであり、再plan・再承認する。

## control operationとworkflow operation

`--help`と`--version`はembedded metadataだけで応答できるcontrol operationとし、manifest欠落時にも診断目的で利用できる。五つのworkflow commandはproject inputを読む前にmanifest/executable/capability bindingを検証する。

invocation grammarを確定できない`cli` usage errorは`unverified`を許可する。workflow commandを確定できた後のusage errorはmanifestを検証して`verified`で返すことを基本とする。ただしmanifest自体が不正ならusage errorより`runtime.manifest-invalid`を優先し、project inputやdestinationへ触れない。

## failure mapping

| condition | diagnostic | status / exit | 副作用 |
| --- | --- | --- | --- |
| manifest missing、JSON/schema/field不整合、未知capability ID、path escape | `runtime.manifest-invalid` | runtime-error / 3 | project input未読、destinationなし |
| executableのsize/digest不一致 | `runtime.artifact-digest-mismatch` | runtime-error / 3 | 同上 |
| core capability不足、canonical順不一致、fixture/profile要求不一致 | `runtime.capability-missing` | runtime-error / 3 | 同上 |
| source archive不在・size/digest不一致 | `runtime.manifest-invalid` | runtime-error / 3 | release/Skillsでは起動しない |
| manifestは有効だがdestination filesystemがprotocol非対応 | `publication.capability-unsupported` | rejected / 1 | output plan/committed artifactなし |

source archiveは実行に不要だが、v1 release bundleのprovenance必須memberである。Skillsが意図的にsourceを同梱しない軽量bundleを将来設ける場合は、上流release manifestをそのまま書き換えず、別のbundle lockが「runtime assetだけを受領した」ことを表現する。

## P3.10 review checklist

- [x] product release、contract、runtime、fixture suite、capabilityを別versionとして記録する
- [x] executableとsourceを別role、basename、media type、size、SHA-256で固定する
- [x] family/runtime versionとexecutable/source filenameの一致をrelease validationで検査する
- [x] Node referenceとJava conforming runtimeの関係をmanifest digestで表す
- [x] runtime directoryのnewest探索、glob、PATH fallbackを禁止する
- [x] manifest digestを外側でpinし、manifest自己hashの循環を避ける
- [x] capabilityとfixture corpusをmanifestへ束縛する
- [x] operation開始前のmanifest、asset、source digest検証順を定義する
- [x] runtimeのbundle自己検証とconsumerの外部trust anchor検証を区別する
- [x] result、output plan、provenanceへ同じruntime bindingを渡す
- [x] manifest不正、asset改変、capability不足、filesystem非対応のdiagnosticを区別する
- [x] Node/Java両exampleが同じschemaで検証できる
