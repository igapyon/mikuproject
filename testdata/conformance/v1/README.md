# miku-project conformance corpus v1

このdirectoryは、Node参照実装とJava適合runtimeが共有するv1 conformance corpusである。設計、比較方式、caseの完了条件は [`docs/miku-project-conformance-corpus-v1.md`](../../../docs/miku-project-conformance-corpus-v1.md) を正本とする。

- `suite-index.json`: case ID、入力、期待status、比較方式の機械可読index
- `contract-cases.json`: JSON Schemaと`RB-001`〜`RB-012`を攻撃するmachine-readable mutation case
- `fixtures/project/`: runtimeへ直接渡すexternal project fixture
- `fixtures/change/`: digestやruntime bindingをharnessがmaterializeする入力template
- `golden/semantic/`: runtime非依存の期待semantic state
- `golden/projection/`: runtime非依存でexact JSON比較する期待Projection

このdirectoryは製品runtimeやtest runnerを同梱しない。`suite-index.json`の`materialization_phase = "P4/P5"`は、caseを実装する計画phaseであって、実装済みtestかどうかを一律に表す値ではない。P4.8時点でNode reference testはhierarchy C1の八caseをmaterialize済みである。NodeはGate G4、JavaはGate G5までに、現行30 workflow / harness caseと31 schema / binding caseを同じ比較規則で実行可能にする。
