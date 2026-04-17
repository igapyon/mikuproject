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

  const mikuprojectProjectWorkbookJson = (globalThis as typeof globalThis & {
    __mikuprojectProjectWorkbookJson?: {
      exportProjectWorkbookJson: (model: ProjectModel) => unknown;
      importProjectWorkbookJsonAsProjectModel: (documentLike: unknown) => {
        model: ProjectModel;
        warnings: Array<{ message: string }>;
      };
      importProjectWorkbookJson: (documentLike: unknown, baseModel: ProjectModel) => {
        model: ProjectModel;
        changes: ImportChange[];
        warnings: Array<{ message: string }>;
      };
      validateWorkbookJsonDocument: (documentLike: unknown) => {
        document: unknown;
        warnings: Array<{ message: string }>;
      };
    };
  }).__mikuprojectProjectWorkbookJson;

  if (!mikuprojectProjectWorkbookJson) {
    throw new Error("mikuproject Project Workbook JSON module is not loaded");
  }

  const mikuprojectCoreApiWorkbookXlsx = (globalThis as typeof globalThis & {
    __mikuprojectCoreApiWorkbookXlsx?: unknown;
  }).__mikuprojectCoreApiWorkbookXlsx;

  if (!mikuprojectCoreApiWorkbookXlsx) {
    throw new Error("mikuproject core api workbook xlsx module is not loaded");
  }

  const mikuprojectProjectPatchJson = (globalThis as typeof globalThis & {
    __mikuprojectProjectPatchJson?: {
      importProjectPatchJson: (documentLike: unknown, baseModel: ProjectModel) => {
        model: ProjectModel;
        changes: ImportChange[];
        warnings: ScopedWarning[];
      };
      validatePatchDocument: (documentLike: unknown) => {
        document: unknown;
        warnings: ScopedWarning[];
      };
    };
  }).__mikuprojectProjectPatchJson;

  if (!mikuprojectProjectPatchJson) {
    throw new Error("mikuproject Project Patch JSON module is not loaded");
  }

  (globalThis as typeof globalThis & {
    __mikuprojectCoreApiWorkbook?: {
      workbookJson: {
        exportDocument: typeof mikuprojectProjectWorkbookJson.exportProjectWorkbookJson;
        validateDocument: typeof mikuprojectProjectWorkbookJson.validateWorkbookJsonDocument;
        importAsProjectModel: typeof mikuprojectProjectWorkbookJson.importProjectWorkbookJsonAsProjectModel;
        importIntoProjectModel: typeof mikuprojectProjectWorkbookJson.importProjectWorkbookJson;
      };
      xlsx: typeof mikuprojectCoreApiWorkbookXlsx;
      patchJson: {
        validateDocument: typeof mikuprojectProjectPatchJson.validatePatchDocument;
        applyToProjectModel: typeof mikuprojectProjectPatchJson.importProjectPatchJson;
      };
    };
  }).__mikuprojectCoreApiWorkbook = {
    workbookJson: {
      exportDocument: mikuprojectProjectWorkbookJson.exportProjectWorkbookJson,
      validateDocument: mikuprojectProjectWorkbookJson.validateWorkbookJsonDocument,
      importAsProjectModel: mikuprojectProjectWorkbookJson.importProjectWorkbookJsonAsProjectModel,
      importIntoProjectModel: mikuprojectProjectWorkbookJson.importProjectWorkbookJson
    },
    xlsx: mikuprojectCoreApiWorkbookXlsx,
    patchJson: {
      validateDocument: mikuprojectProjectPatchJson.validatePatchDocument,
      applyToProjectModel: mikuprojectProjectPatchJson.importProjectPatchJson
    }
  };
})();
