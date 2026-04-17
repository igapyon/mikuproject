/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  const mikuprojectCoreApiImport = (globalThis as typeof globalThis & {
    __mikuprojectCoreApiImport?: {
      getAiJsonSpecText: () => string;
      getAiJsonSpec: () => {
        id: "mikuproject-ai-json-spec";
        version: string;
        text: string;
      };
      detectAiJsonDocumentKind: (documentLike: unknown) => "workbook_json" | "project_draft_view" | "patch_json" | undefined;
      parseAiJsonText: (sourceText: string) => {
        sourceText: string;
        jsonText: string;
        document: unknown;
        kind: "workbook_json" | "project_draft_view" | "patch_json" | undefined;
      };
      importAiJsonDocument: (documentLike: unknown, options?: { baseModel?: ProjectModel }) => unknown;
      importAiJsonText: (sourceText: string, options?: { baseModel?: ProjectModel }) => unknown;
      importExternal: (input: unknown) => unknown;
    };
  }).__mikuprojectCoreApiImport;

  if (!mikuprojectCoreApiImport) {
    throw new Error("mikuproject core api import module is not loaded");
  }

  const mikuprojectCoreApiMsproject = (globalThis as typeof globalThis & {
    __mikuprojectCoreApiMsproject?: {
      samples: unknown;
      projectModel: unknown;
      msProject: unknown;
      aiViews: unknown;
    };
  }).__mikuprojectCoreApiMsproject;

  if (!mikuprojectCoreApiMsproject) {
    throw new Error("mikuproject core api msproject module is not loaded");
  }

  const mikuprojectCoreApiWorkbook = (globalThis as typeof globalThis & {
    __mikuprojectCoreApiWorkbook?: {
      workbookJson: unknown;
      xlsx: unknown;
      patchJson: unknown;
    };
  }).__mikuprojectCoreApiWorkbook;

  if (!mikuprojectCoreApiWorkbook) {
    throw new Error("mikuproject core api workbook module is not loaded");
  }

  const mikuprojectCoreApiReportPublic = (globalThis as typeof globalThis & {
    __mikuprojectCoreApiReportPublic?: {
      report: unknown;
    };
  }).__mikuprojectCoreApiReportPublic;

  if (!mikuprojectCoreApiReportPublic) {
    throw new Error("mikuproject core api report public module is not loaded");
  }

  (globalThis as typeof globalThis & {
    __mikuprojectCoreApiRegistry?: {
      getAiJsonSpecText: typeof mikuprojectCoreApiImport.getAiJsonSpecText;
      getAiJsonSpec: typeof mikuprojectCoreApiImport.getAiJsonSpec;
      detectAiJsonDocumentKind: typeof mikuprojectCoreApiImport.detectAiJsonDocumentKind;
      parseAiJsonText: typeof mikuprojectCoreApiImport.parseAiJsonText;
      importAiJsonDocument: typeof mikuprojectCoreApiImport.importAiJsonDocument;
      importAiJsonText: typeof mikuprojectCoreApiImport.importAiJsonText;
      importExternal: typeof mikuprojectCoreApiImport.importExternal;
      samples: unknown;
      projectModel: unknown;
      msProject: unknown;
      aiViews: unknown;
      workbookJson: unknown;
      xlsx: unknown;
      patchJson: unknown;
      report: unknown;
    };
  }).__mikuprojectCoreApiRegistry = {
    getAiJsonSpecText: mikuprojectCoreApiImport.getAiJsonSpecText,
    getAiJsonSpec: mikuprojectCoreApiImport.getAiJsonSpec,
    detectAiJsonDocumentKind: mikuprojectCoreApiImport.detectAiJsonDocumentKind,
    parseAiJsonText: mikuprojectCoreApiImport.parseAiJsonText,
    importAiJsonDocument: mikuprojectCoreApiImport.importAiJsonDocument,
    importAiJsonText: mikuprojectCoreApiImport.importAiJsonText,
    importExternal: mikuprojectCoreApiImport.importExternal,
    samples: mikuprojectCoreApiMsproject.samples,
    projectModel: mikuprojectCoreApiMsproject.projectModel,
    msProject: mikuprojectCoreApiMsproject.msProject,
    aiViews: mikuprojectCoreApiMsproject.aiViews,
    workbookJson: mikuprojectCoreApiWorkbook.workbookJson,
    xlsx: mikuprojectCoreApiWorkbook.xlsx,
    patchJson: mikuprojectCoreApiWorkbook.patchJson,
    report: mikuprojectCoreApiReportPublic.report
  };
})();
