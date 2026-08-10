# miku-project conformance corpus v1

このdirectoryは、Node参照実装とJava適合runtimeが共有するv1 conformance corpusである。設計、比較方式、caseの完了条件は [`docs/miku-project-conformance-corpus-v1.md`](../../../docs/miku-project-conformance-corpus-v1.md) を正本とする。

- `suite-index.json`: case ID、入力、期待status、比較方式の機械可読index
- `contract-cases.json`: JSON Schemaと`RB-001`〜`RB-012`を攻撃するmachine-readable mutation case
- `fixtures/project/`: runtimeへ直接渡すexternal project fixture
- `fixtures/change/`: digestやruntime bindingをharnessがmaterializeする入力template
- `golden/semantic/`: runtime非依存の期待semantic state

この段階では製品runtimeやtest runnerを同梱しない。`suite-index.json`の`materialization_phase = "P4/P5"`であるcaseは、契約済みのcaseであって実装済みtestではない。NodeはG4、JavaはG5までに同じ21 workflow / harness caseと31 schema / binding caseを実行可能にする。
