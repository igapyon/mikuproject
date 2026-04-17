/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  const mainModel = {
    syncXmlTextFromModel(input: {
      model: ProjectModel;
      exportMsProjectXml: (model: ProjectModel) => string;
      writeXmlText: (xmlText: string) => void;
      clearXmlSourceDirty: () => void;
      refreshXmlSaveState: () => void;
    }): string {
      const xmlText = input.exportMsProjectXml(input.model);
      input.writeXmlText(xmlText);
      input.clearXmlSourceDirty();
      input.refreshXmlSaveState();
      return xmlText;
    },

    ensureCurrentModel(input: {
      currentModel: ProjectModel | null;
      readXmlText: () => string;
      importMsProjectXml: (xmlText: string) => ProjectModel;
      clearXmlSourceDirty: () => void;
    }): ProjectModel {
      if (input.currentModel) {
        return input.currentModel;
      }
      const xmlText = input.readXmlText().trim();
      if (!xmlText) {
        throw new Error("内部モデルがありません");
      }
      const model = input.importMsProjectXml(xmlText);
      input.clearXmlSourceDirty();
      return model;
    },

    parseCurrentXml(input: {
      readXmlText: () => string;
      importMsProjectXml: (xmlText: string) => ProjectModel;
      validateProjectModel: (model: ProjectModel) => ValidationIssue[];
      clearXmlSourceDirty: () => void;
    }): { model: ProjectModel | null; issues: ValidationIssue[]; empty: boolean } {
      const xmlText = input.readXmlText().trim();
      if (!xmlText) {
        return {
          model: null,
          issues: [],
          empty: true
        };
      }
      const model = input.importMsProjectXml(xmlText);
      input.clearXmlSourceDirty();
      return {
        model,
        issues: input.validateProjectModel(model),
        empty: false
      };
    }
  };

  (globalThis as typeof globalThis & {
    __mikuprojectMainModel?: typeof mainModel;
  }).__mikuprojectMainModel = mainModel;
})();
