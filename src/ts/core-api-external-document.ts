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

  const mikuprojectCoreApiAiJson = (globalThis as typeof globalThis & {
    __mikuprojectCoreApiAiJson?: {
      importAiJsonDocument: (
        documentLike: unknown,
        options?: { baseModel?: ProjectModel }
      ) =>
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
        };
    };
  }).__mikuprojectCoreApiAiJson;

  if (!mikuprojectCoreApiAiJson) {
    throw new Error("mikuproject core api ai json module is not loaded");
  }

  function importDocument(
    format: AiJsonDocumentKind,
    document: unknown,
    mode: "replace" | "merge" | "patch",
    baseModel?: ProjectModel
  ) {
    if (format === "workbook_json") {
      if (mode === "patch") {
        throw new Error("workbook JSON は patch import に対応していません");
      }
      if (mode === "merge" && !baseModel) {
        throw new Error("workbook JSON merge import には baseModel が必要です");
      }
      return mikuprojectCoreApiAiJson.importAiJsonDocument(document, mode === "merge" ? { baseModel } : {});
    }

    if (format === "project_draft_view") {
      if (mode !== "replace") {
        throw new Error("project_draft_view は replace import のみ対応です");
      }
      return mikuprojectCoreApiAiJson.importAiJsonDocument(document);
    }

    if (format === "patch_json") {
      if (mode !== "patch") {
        throw new Error("patch JSON は patch import のみ対応です");
      }
      return mikuprojectCoreApiAiJson.importAiJsonDocument(document, { baseModel });
    }

    throw new Error(`未対応の import format です: ${format}`);
  }

  (globalThis as typeof globalThis & {
    __mikuprojectCoreApiExternalDocument?: {
      importDocument: typeof importDocument;
    };
  }).__mikuprojectCoreApiExternalDocument = {
    importDocument
  };
})();
