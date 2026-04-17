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

  type WorkbookJsonWarning = {
    message: string;
  };

  type WorkbookJsonRow = Record<string, string | number | boolean | null>;

  type WorkbookJsonDocument = {
    format: "mikuproject_workbook_json";
    version: 1;
    sheets: {
      Project?: WorkbookJsonRow[];
      Tasks?: WorkbookJsonRow[];
      Resources?: WorkbookJsonRow[];
      Assignments?: WorkbookJsonRow[];
      Calendars?: WorkbookJsonRow[];
      NonWorkingDays?: WorkbookJsonRow[];
      [sheetName: string]: WorkbookJsonRow[] | undefined;
    };
  };

  const projectWorkbookJsonExport = (globalThis as typeof globalThis & {
    __mikuprojectProjectWorkbookJsonExport?: {
      exportProjectWorkbookJson: (model: ProjectModel) => WorkbookJsonDocument;
    };
  }).__mikuprojectProjectWorkbookJsonExport;

  if (!projectWorkbookJsonExport) {
    throw new Error("mikuproject Project Workbook JSON export module is not loaded");
  }

  const projectWorkbookJsonImport = (globalThis as typeof globalThis & {
    __mikuprojectProjectWorkbookJsonImport?: {
      importProjectWorkbookJson: (documentLike: unknown, baseModel: ProjectModel) => {
        model: ProjectModel;
        changes: ImportChange[];
        warnings: WorkbookJsonWarning[];
      };
      importProjectWorkbookJsonAsProjectModel: (documentLike: unknown) => {
        model: ProjectModel;
        warnings: WorkbookJsonWarning[];
      };
      validateWorkbookJsonDocument: (documentLike: unknown) => {
        document: WorkbookJsonDocument;
        warnings: WorkbookJsonWarning[];
      };
    };
  }).__mikuprojectProjectWorkbookJsonImport;

  if (!projectWorkbookJsonImport) {
    throw new Error("mikuproject Project Workbook JSON import module is not loaded");
  }

  (globalThis as typeof globalThis & {
    __mikuprojectProjectWorkbookJson?: {
      exportProjectWorkbookJson: (model: ProjectModel) => WorkbookJsonDocument;
      importProjectWorkbookJsonAsProjectModel: (documentLike: unknown) => {
        model: ProjectModel;
        warnings: WorkbookJsonWarning[];
      };
      importProjectWorkbookJson: (documentLike: unknown, baseModel: ProjectModel) => {
        model: ProjectModel;
        changes: ImportChange[];
        warnings: WorkbookJsonWarning[];
      };
      validateWorkbookJsonDocument: (documentLike: unknown) => {
        document: WorkbookJsonDocument;
        warnings: WorkbookJsonWarning[];
      };
    };
  }).__mikuprojectProjectWorkbookJson = {
    exportProjectWorkbookJson: (model: ProjectModel): WorkbookJsonDocument =>
      projectWorkbookJsonExport.exportProjectWorkbookJson(model),
    importProjectWorkbookJsonAsProjectModel: (documentLike: unknown): {
      model: ProjectModel;
      warnings: WorkbookJsonWarning[];
    } => projectWorkbookJsonImport.importProjectWorkbookJsonAsProjectModel(documentLike),
    importProjectWorkbookJson: (documentLike: unknown, baseModel: ProjectModel): {
      model: ProjectModel;
      changes: ImportChange[];
      warnings: WorkbookJsonWarning[];
    } => projectWorkbookJsonImport.importProjectWorkbookJson(documentLike, baseModel),
    validateWorkbookJsonDocument: (documentLike: unknown): {
      document: WorkbookJsonDocument;
      warnings: WorkbookJsonWarning[];
    } => projectWorkbookJsonImport.validateWorkbookJsonDocument(documentLike)
  };
})();
