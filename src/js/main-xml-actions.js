/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const mainXmlActions = {
        async exportCurrentMermaid(input) {
            if (!input.currentModel) {
                input.setStatus("内部モデルがありません");
                return;
            }
            const mermaidText = input.exportMermaidGantt(input.currentModel);
            input.setMermaidText(mermaidText);
            await input.renderSvgPreview();
            if (!input.silent) {
                input.setStatus("内部モデルから Mermaid gantt を生成し、native SVG preview を更新しました");
                input.showToast("Mermaid を生成しました");
            }
            input.setActiveTab();
        },
        loadSampleXml(input) {
            const model = input.loadSampleXml({
                document: input.document,
                readSampleXml: input.readSampleXml,
                readSampleProjectDraftView: input.readSampleProjectDraftView,
                writeXmlText: input.writeXmlText,
                writeProjectDraftText: input.writeProjectDraftText,
                importSampleXml: input.importSampleXml,
                validateProjectModel: input.validateProjectModel,
                markXmlDirty: input.markXmlDirty,
                applyModelState: input.applyModelState,
                updateSvgButton: input.updateSvgButton,
                setStatus: input.setStatus,
                setActiveTab: input.setActiveTab
            });
            input.setCurrentModel(model);
            input.setXmlSourceDirty(true);
        },
        async importXmlFromFile(input) {
            const xmlText = await input.file.text();
            input.writeXmlText(xmlText);
            input.markXmlDirty();
            const model = input.importMsProjectXml(xmlText);
            input.setCurrentModel(model);
            input.setXmlSourceDirty(false);
            const issues = input.validateProjectModel(model);
            input.applyModelState({ model, issues });
            input.completeTransform(issues.length > 0 ? `XML ファイルを読み込んで解析しました。検証で ${issues.length} 件の問題があります` : "XML ファイルを読み込んで解析しました", "XML を読み込んで解析しました");
            await input.exportCurrentMermaid({ silent: true });
        },
        runRoundTripCheck(input) {
            let model = input.currentModel;
            if (!model) {
                input.parseCurrentXml();
                model = input.getCurrentModel();
                if (!model) {
                    return;
                }
            }
            const validationIssues = input.assertRoundTripStable({
                currentModel: model,
                exportMsProjectXml: input.exportMsProjectXml,
                importMsProjectXml: input.importMsProjectXml,
                validateProjectModel: input.validateProjectModel,
                normalizeProjectModel: input.normalizeProjectModel
            });
            input.renderValidationIssues(validationIssues);
            input.setStatus("再読込テストに成功しました");
            input.showToast("再読込テスト成功");
            input.setActiveTab();
        }
    };
    globalThis.__mikuprojectMainXmlActions = mainXmlActions;
})();
