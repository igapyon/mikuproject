/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  const mikuprojectXml = (globalThis as typeof globalThis & {
    __mikuprojectXml?: {
      SAMPLE_XML: string;
      SAMPLE_PROJECT_DRAFT_VIEW: unknown;
      parseXmlDocument: (xmlText: string) => XMLDocument;
      importMsProjectXml: (xmlText: string) => ProjectModel;
      importCsvParentId: (csvText: string) => ProjectModel;
      exportMsProjectXml: (model: ProjectModel) => string;
      exportCsvParentId: (model: ProjectModel) => string;
      normalizeProjectModel: (model: ProjectModel) => ProjectModel;
      validateProjectModel: (model: ProjectModel) => ValidationIssue[];
    };
  }).__mikuprojectXml;

  if (!mikuprojectXml) {
    throw new Error("mikuproject XML module is not loaded");
  }

  const mikuprojectCoreApiMsprojectAi = (globalThis as typeof globalThis & {
    __mikuprojectCoreApiMsprojectAi?: {
      aiViews: unknown;
      mermaid: unknown;
    };
  }).__mikuprojectCoreApiMsprojectAi;

  if (!mikuprojectCoreApiMsprojectAi) {
    throw new Error("mikuproject core api msproject ai module is not loaded");
  }

  globalThis.__mikuprojectCoreApiMsproject = {
    samples: {
      getSampleXml: () => mikuprojectXml.SAMPLE_XML,
      getSampleProjectDraftView: () => mikuprojectXml.SAMPLE_PROJECT_DRAFT_VIEW
    },
    projectModel: {
      normalize: mikuprojectXml.normalizeProjectModel,
      validate: mikuprojectXml.validateProjectModel
    },
    msProject: {
      parseXmlDocument: mikuprojectXml.parseXmlDocument,
      importFromXml: mikuprojectXml.importMsProjectXml,
      exportToXml: mikuprojectXml.exportMsProjectXml,
      importFromCsvParentId: mikuprojectXml.importCsvParentId,
      exportToCsvParentId: mikuprojectXml.exportCsvParentId
    },
    aiViews: mikuprojectCoreApiMsprojectAi.aiViews,
    mermaid: mikuprojectCoreApiMsprojectAi.mermaid
  };
})();
