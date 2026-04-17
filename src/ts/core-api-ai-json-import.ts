/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  type ImportChange = {
    scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
    uid: string;
    label: string;
    field: string;
    before: string | number | boolean | undefined;
    after: string | number | boolean;
  };

  type ScopedWarning = {
    message: string;
    scope?: "project" | "tasks" | "resources" | "assignments" | "calendars";
    uid?: string;
    label?: string;
  };

  type AiJsonDocumentKind = "workbook_json" | "project_draft_view" | "patch_json";

  const mikuprojectCoreApiMsproject = (globalThis as typeof globalThis & {
    __mikuprojectCoreApiMsproject?: {
      aiViews: {
        importProjectDraftView: (draft: unknown) => ProjectModel;
      };
    };
  }).__mikuprojectCoreApiMsproject;

  if (!mikuprojectCoreApiMsproject) {
    throw new Error("mikuproject core api msproject module is not loaded");
  }

  const mikuprojectCoreApiWorkbook = (globalThis as typeof globalThis & {
    __mikuprojectCoreApiWorkbook?: {
      workbookJson: {
        importAsProjectModel: (documentLike: unknown) => {
          model: ProjectModel;
          warnings: Array<{ message: string }>;
        };
        importIntoProjectModel: (documentLike: unknown, baseModel: ProjectModel) => {
          model: ProjectModel;
          changes: ImportChange[];
          warnings: Array<{ message: string }>;
        };
      };
      patchJson: {
        applyToProjectModel: (documentLike: unknown, baseModel: ProjectModel) => {
          model: ProjectModel;
          changes: ImportChange[];
          warnings: ScopedWarning[];
        };
      };
    };
  }).__mikuprojectCoreApiWorkbook;

  if (!mikuprojectCoreApiWorkbook) {
    throw new Error("mikuproject core api workbook module is not loaded");
  }

  const mikuprojectAiJsonUtil = (globalThis as typeof globalThis & {
    __mikuprojectAiJsonUtil?: {
      detectJsonDocumentKind: (documentLike: unknown) => AiJsonDocumentKind | undefined;
    };
  }).__mikuprojectAiJsonUtil;

  if (!mikuprojectAiJsonUtil) {
    throw new Error("mikuproject AI JSON util module is not loaded");
  }

  function detectAiJsonDocumentKind(documentLike: unknown): AiJsonDocumentKind | undefined {
    return mikuprojectAiJsonUtil.detectJsonDocumentKind(documentLike);
  }

  function importAiJsonDocument(documentLike: unknown, options: {
    baseModel?: ProjectModel;
  } = {}):
    | {
      kind: "project_draft_view";
      mode: "replace";
      model: ProjectModel;
      warnings: [];
    }
    | {
      kind: "workbook_json";
      mode: "replace";
      model: ProjectModel;
      warnings: Array<{ message: string }>;
    }
    | {
      kind: "workbook_json";
      mode: "merge";
      model: ProjectModel;
      changes: ImportChange[];
      warnings: Array<{ message: string }>;
    }
    | {
      kind: "patch_json";
      mode: "patch";
      model: ProjectModel;
      changes: ImportChange[];
      warnings: ScopedWarning[];
    } {
    const kind = detectAiJsonDocumentKind(documentLike);
    if (kind === "project_draft_view") {
      return {
        kind,
        mode: "replace",
        model: mikuprojectCoreApiMsproject.aiViews.importProjectDraftView(documentLike),
        warnings: []
      };
    }

    if (kind === "workbook_json") {
      if (options.baseModel) {
        return {
          kind,
          mode: "merge",
          ...mikuprojectCoreApiWorkbook.workbookJson.importIntoProjectModel(documentLike, options.baseModel)
        };
      }
      return {
        kind,
        mode: "replace",
        ...mikuprojectCoreApiWorkbook.workbookJson.importAsProjectModel(documentLike)
      };
    }

    if (kind === "patch_json") {
      if (!options.baseModel) {
        throw new Error("Patch JSON の適用には baseModel が必要です");
      }
      return {
        kind,
        mode: "patch",
        ...mikuprojectCoreApiWorkbook.patchJson.applyToProjectModel(documentLike, options.baseModel)
      };
    }

    throw new Error("AI JSON の format / view_type を判別できません");
  }

  (globalThis as typeof globalThis & {
    __mikuprojectCoreApiAiJsonImport?: {
      detectAiJsonDocumentKind: typeof detectAiJsonDocumentKind;
      importAiJsonDocument: typeof importAiJsonDocument;
    };
  }).__mikuprojectCoreApiAiJsonImport = {
    detectAiJsonDocumentKind,
    importAiJsonDocument
  };
})();
