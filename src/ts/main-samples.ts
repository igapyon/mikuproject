/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  type MainSamplesDeps = {
    document: Document;
    getAiPromptText: () => string;
    copyTextToClipboard: (text: string) => Promise<void>;
    showToast: (message: string) => void;
    setStatus: (message: string) => void;
    readSampleXml: () => string;
    readSampleProjectDraftView: () => unknown;
    writeXmlText: (xmlText: string) => void;
    writeProjectDraftText: (text: string) => void;
    importSampleXml: (xmlText: string) => ProjectModel;
    validateProjectModel: (model: ProjectModel) => ValidationIssue[];
    markXmlDirty: () => void;
    applyModelState: (input: { model: ProjectModel; issues: ValidationIssue[] }) => void;
    updateSvgButton: () => void;
    setActiveTab: (tabId: "input" | "transform" | "output") => void;
  };

  const mainSamples = {
    async copyAiPrompt(deps: MainSamplesDeps): Promise<void> {
      const promptText = deps.getAiPromptText();
      if (!promptText) {
        throw new Error("生成AIプロンプトが見つかりません");
      }
      await deps.copyTextToClipboard(promptText);
      deps.showToast("生成AIプロンプトをクリップボードにコピーしました");
      deps.setStatus("生成AIプロンプトをクリップボードにコピーしました");
    },

    loadSampleXml(deps: MainSamplesDeps): ProjectModel {
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

    loadProjectDraftSample(deps: MainSamplesDeps): void {
      const sampleDraftText = JSON.stringify(deps.readSampleProjectDraftView(), null, 2);
      deps.writeProjectDraftText(sampleDraftText);
      deps.setStatus("サンプル project_draft_view を読み込みました");
      deps.setActiveTab("input");
    }
  };

  (globalThis as typeof globalThis & {
    __mikuprojectMainSamples?: typeof mainSamples;
  }).__mikuprojectMainSamples = mainSamples;
})();
