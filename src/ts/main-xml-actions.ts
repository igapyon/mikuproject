/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  const mainXmlActions = {
    async exportCurrentMermaid(input: {
      currentModel: ProjectModel | null;
      exportMermaidGantt: (model: ProjectModel) => string;
      setMermaidText: (text: string) => void;
      renderSvgPreview: () => Promise<void>;
      setStatus: (message: string) => void;
      showToast: (message: string) => void;
      setActiveTab: () => void;
      silent?: boolean;
    }): Promise<void> {
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

    loadSampleXml(input: {
      loadSampleXml: (deps: {
        document: Document;
        readSampleXml: () => string;
        readSampleProjectDraftView: () => unknown;
        writeXmlText: (xmlText: string) => void;
        writeProjectDraftText: (text: string) => void;
        importSampleXml: (xmlText: string) => ProjectModel;
        validateProjectModel: (model: ProjectModel) => ValidationIssue[];
        markXmlDirty: () => void;
        applyModelState: (input: { model: ProjectModel; issues: ValidationIssue[] }) => void;
        updateSvgButton: () => void;
        setStatus: (message: string) => void;
        setActiveTab: (tabId: "input" | "transform" | "output") => void;
      }) => ProjectModel;
      document: Document;
      readSampleXml: () => string;
      readSampleProjectDraftView: () => unknown;
      writeXmlText: (xmlText: string) => void;
      writeProjectDraftText: (text: string) => void;
      importSampleXml: (xmlText: string) => ProjectModel;
      validateProjectModel: (model: ProjectModel) => ValidationIssue[];
      markXmlDirty: () => void;
      applyModelState: (input: { model: ProjectModel; issues: ValidationIssue[] }) => void;
      updateSvgButton: () => void;
      setStatus: (message: string) => void;
      setActiveTab: (tabId: "input" | "transform" | "output") => void;
      setCurrentModel: (model: ProjectModel) => void;
      setXmlSourceDirty: (dirty: boolean) => void;
    }): void {
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

    async importXmlFromFile(input: {
      file: File;
      writeXmlText: (xmlText: string) => void;
      markXmlDirty: () => void;
      importMsProjectXml: (xmlText: string) => ProjectModel;
      validateProjectModel: (model: ProjectModel) => ValidationIssue[];
      setCurrentModel: (model: ProjectModel) => void;
      setXmlSourceDirty: (dirty: boolean) => void;
      applyModelState: (input: { model: ProjectModel; issues: ValidationIssue[] }) => void;
      completeTransform: (statusMessage: string, toastMessage: string) => void;
      exportCurrentMermaid: (options?: { silent?: boolean }) => Promise<void>;
    }): Promise<void> {
      const xmlText = await input.file.text();
      input.writeXmlText(xmlText);
      input.markXmlDirty();
      const model = input.importMsProjectXml(xmlText);
      input.setCurrentModel(model);
      input.setXmlSourceDirty(false);
      const issues = input.validateProjectModel(model);
      input.applyModelState({ model, issues });
      input.completeTransform(
        issues.length > 0 ? `XML ファイルを読み込んで解析しました。検証で ${issues.length} 件の問題があります` : "XML ファイルを読み込んで解析しました",
        "XML を読み込んで解析しました"
      );
      await input.exportCurrentMermaid({ silent: true });
    },

    runRoundTripCheck(input: {
      currentModel: ProjectModel | null;
      parseCurrentXml: () => void;
      getCurrentModel: () => ProjectModel | null;
      assertRoundTripStable: (input: {
        currentModel: ProjectModel;
        exportMsProjectXml: (model: ProjectModel) => string;
        importMsProjectXml: (xmlText: string) => ProjectModel;
        validateProjectModel: (model: ProjectModel) => ValidationIssue[];
        normalizeProjectModel: (model: ProjectModel) => ProjectModel;
      }) => ValidationIssue[];
      exportMsProjectXml: (model: ProjectModel) => string;
      importMsProjectXml: (xmlText: string) => ProjectModel;
      validateProjectModel: (model: ProjectModel) => ValidationIssue[];
      normalizeProjectModel: (model: ProjectModel) => ProjectModel;
      renderValidationIssues: (issues: ValidationIssue[]) => void;
      setStatus: (message: string) => void;
      showToast: (message: string) => void;
      setActiveTab: () => void;
    }): void {
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

  (globalThis as typeof globalThis & {
    __mikuprojectMainXmlActions?: typeof mainXmlActions;
  }).__mikuprojectMainXmlActions = mainXmlActions;
})();
