/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  const mikuprojectCoreApiRegistry = (globalThis as typeof globalThis & {
    __mikuprojectCoreApiRegistry?: {
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
      samples: unknown;
      projectModel: unknown;
      msProject: unknown;
      aiViews: unknown;
      workbookJson: unknown;
      xlsx: unknown;
      patchJson: unknown;
      report: unknown;
    };
  }).__mikuprojectCoreApiRegistry;

  if (!mikuprojectCoreApiRegistry) {
    throw new Error("mikuproject core api registry module is not loaded");
  }

  (globalThis as typeof globalThis & {
    __mikuprojectCoreApiPublic?: {
      version: 1;
      getAiJsonSpecText: typeof mikuprojectCoreApiRegistry.getAiJsonSpecText;
      getAiJsonSpec: typeof mikuprojectCoreApiRegistry.getAiJsonSpec;
      detectAiJsonDocumentKind: typeof mikuprojectCoreApiRegistry.detectAiJsonDocumentKind;
      parseAiJsonText: typeof mikuprojectCoreApiRegistry.parseAiJsonText;
      importAiJsonDocument: typeof mikuprojectCoreApiRegistry.importAiJsonDocument;
      importAiJsonText: typeof mikuprojectCoreApiRegistry.importAiJsonText;
      importExternal: typeof mikuprojectCoreApiRegistry.importExternal;
      samples: unknown;
      projectModel: unknown;
      msProject: unknown;
      aiViews: unknown;
      workbookJson: unknown;
      xlsx: unknown;
      patchJson: unknown;
      report: unknown;
    };
  }).__mikuprojectCoreApiPublic = {
    version: 1,
    ...mikuprojectCoreApiRegistry
  };
})();
