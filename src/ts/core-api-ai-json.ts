/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  type AiJsonDocumentKind = "workbook_json" | "project_draft_view" | "patch_json";

  const mikuprojectAiJsonUtil = (globalThis as typeof globalThis & {
    __mikuprojectAiJsonUtil?: {
      extractLastJsonBlock: (value: string) => string;
    };
  }).__mikuprojectAiJsonUtil;

  if (!mikuprojectAiJsonUtil) {
    throw new Error("mikuproject AI JSON util module is not loaded");
  }

  const mikuprojectAiJsonSpec = (globalThis as typeof globalThis & {
    __mikuprojectAiJsonSpec?: {
      getAiJsonSpecText: () => string;
      getAiJsonSpec: () => {
        id: "mikuproject-ai-json-spec";
        version: string;
        text: string;
      };
    };
  }).__mikuprojectAiJsonSpec;

  if (!mikuprojectAiJsonSpec) {
    throw new Error("mikuproject AI JSON spec module is not loaded");
  }

  const mikuprojectCoreApiAiJsonImport = (globalThis as typeof globalThis & {
    __mikuprojectCoreApiAiJsonImport?: {
      detectAiJsonDocumentKind: (documentLike: unknown) => AiJsonDocumentKind | undefined;
      importAiJsonDocument: (
        documentLike: unknown,
        options?: { baseModel?: ProjectModel }
      ) => unknown;
    };
  }).__mikuprojectCoreApiAiJsonImport;

  if (!mikuprojectCoreApiAiJsonImport) {
    throw new Error("mikuproject core api ai json import module is not loaded");
  }

  function getAiJsonSpecText(): string {
    return mikuprojectAiJsonSpec.getAiJsonSpecText();
  }

  function getAiJsonSpec(): {
    id: "mikuproject-ai-json-spec";
    version: string;
    text: string;
  } {
    return mikuprojectAiJsonSpec.getAiJsonSpec();
  }

  function parseAiJsonText(sourceText: string): {
    sourceText: string;
    jsonText: string;
    document: unknown;
    kind: AiJsonDocumentKind | undefined;
  } {
    const jsonText = mikuprojectAiJsonUtil.extractLastJsonBlock(sourceText);
    const document = JSON.parse(jsonText);
    return {
      sourceText,
      jsonText,
      document,
      kind: mikuprojectCoreApiAiJsonImport.detectAiJsonDocumentKind(document)
    };
  }

  function importAiJsonText(sourceText: string, options: {
    baseModel?: ProjectModel;
  } = {}): {
    sourceText: string;
    jsonText: string;
    document: unknown;
    kind: AiJsonDocumentKind | undefined;
    result: ReturnType<typeof mikuprojectCoreApiAiJsonImport.importAiJsonDocument>;
  } {
    const parsed = parseAiJsonText(sourceText);
    return {
      ...parsed,
      result: mikuprojectCoreApiAiJsonImport.importAiJsonDocument(parsed.document, options)
    };
  }

  (globalThis as typeof globalThis & {
    __mikuprojectCoreApiAiJson?: {
      getAiJsonSpecText: typeof getAiJsonSpecText;
      getAiJsonSpec: typeof getAiJsonSpec;
      detectAiJsonDocumentKind: typeof mikuprojectCoreApiAiJsonImport.detectAiJsonDocumentKind;
      parseAiJsonText: typeof parseAiJsonText;
      importAiJsonDocument: typeof mikuprojectCoreApiAiJsonImport.importAiJsonDocument;
      importAiJsonText: typeof importAiJsonText;
    };
  }).__mikuprojectCoreApiAiJson = {
    getAiJsonSpecText,
    getAiJsonSpec,
    detectAiJsonDocumentKind: mikuprojectCoreApiAiJsonImport.detectAiJsonDocumentKind,
    parseAiJsonText,
    importAiJsonDocument: mikuprojectCoreApiAiJsonImport.importAiJsonDocument,
    importAiJsonText
  };
})();
