/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const mainSamples = {
        async copyAiPrompt(deps) {
            const promptText = deps.getAiPromptText();
            if (!promptText) {
                throw new Error("生成AIプロンプトが見つかりません");
            }
            await deps.copyTextToClipboard(promptText);
            deps.showToast("生成AIプロンプトをクリップボードにコピーしました");
            deps.setStatus("生成AIプロンプトをクリップボードにコピーしました");
        },
        loadSampleXml(deps) {
            const sampleXml = deps.readSampleXml();
            deps.writeXmlText(sampleXml);
            const model = deps.importSampleXml(sampleXml);
            deps.markXmlDirty();
            deps.applyModelState({
                model,
                issues: deps.validateProjectModel(model)
            });
            deps.updateSvgButton();
            deps.setStatus("サンプル XML を読み込みました");
            deps.setActiveTab("input");
            return model;
        },
        loadProjectDraftSample(deps) {
            const sampleDraftText = JSON.stringify(deps.readSampleProjectDraftView(), null, 2);
            deps.writeProjectDraftText(sampleDraftText);
            deps.setStatus("サンプル project_draft_view を読み込みました");
            deps.setActiveTab("input");
        }
    };
    globalThis.__mikuprojectMainSamples = mainSamples;
})();
