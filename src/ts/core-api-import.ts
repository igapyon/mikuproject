/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  type AiJsonDocumentKind = "workbook_json" | "project_draft_view" | "patch_json";

  const mikuprojectCoreApiAiJson = (globalThis as typeof globalThis & {
    __mikuprojectCoreApiAiJson?: {
      getAiJsonSpecText: () => string;
      getAiJsonSpec: () => {
        id: "mikuproject-ai-json-spec";
        version: string;
        text: string;
      };
      detectAiJsonDocumentKind: (documentLike: unknown) => AiJsonDocumentKind | undefined;
      parseAiJsonText: (sourceText: string) => {
        sourceText: string;
        jsonText: string;
        document: unknown;
        kind: AiJsonDocumentKind | undefined;
      };
      importAiJsonDocument: (
        documentLike: unknown,
        options?: { baseModel?: ProjectModel }
      ) => unknown;
      importAiJsonText: (
        sourceText: string,
        options?: { baseModel?: ProjectModel }
      ) => unknown;
    };
  }).__mikuprojectCoreApiAiJson;

  if (!mikuprojectCoreApiAiJson) {
    throw new Error("mikuproject core api ai json module is not loaded");
  }

  const mikuprojectCoreApiExternalImport = (globalThis as typeof globalThis & {
    __mikuprojectCoreApiExternalImport?: {
      importExternal: (input: {
        source:
          | { format: "ms_project_xml"; text: string }
          | { format: "xlsx"; bytes: Uint8Array }
          | { format: "workbook_json"; document: unknown }
          | { format: "project_draft_view"; document: unknown }
          | { format: "patch_json"; document: unknown };
        mode: "replace" | "merge" | "patch";
        baseModel?: ProjectModel;
      }) => unknown;
    };
  }).__mikuprojectCoreApiExternalImport;

  if (!mikuprojectCoreApiExternalImport) {
    throw new Error("mikuproject core api external import module is not loaded");
  }

  globalThis.__mikuprojectCoreApiImport = {
    getAiJsonSpecText: mikuprojectCoreApiAiJson.getAiJsonSpecText,
    getAiJsonSpec: mikuprojectCoreApiAiJson.getAiJsonSpec,
    detectAiJsonDocumentKind: mikuprojectCoreApiAiJson.detectAiJsonDocumentKind,
    parseAiJsonText: mikuprojectCoreApiAiJson.parseAiJsonText,
    importAiJsonDocument: mikuprojectCoreApiAiJson.importAiJsonDocument,
    importAiJsonText: mikuprojectCoreApiAiJson.importAiJsonText,
    importExternal: mikuprojectCoreApiExternalImport.importExternal
  };
})();
