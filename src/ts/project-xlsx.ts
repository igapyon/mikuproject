/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  const projectXlsxImport = (globalThis as typeof globalThis & {
    __mikuprojectProjectXlsxImport?: {
      importProjectWorkbook: (workbook: XlsxWorkbookLike, baseModel: ProjectModel) => ProjectModel;
      importProjectWorkbookAsProjectModel: (workbook: XlsxWorkbookLike) => ProjectModel;
      importProjectWorkbookDetailed: (workbook: XlsxWorkbookLike, baseModel: ProjectModel) => {
        model: ProjectModel;
        changes: ImportChange[];
      };
    };
  }).__mikuprojectProjectXlsxImport;

  if (!projectXlsxImport) {
    throw new Error("mikuproject Project XLSX import module is not loaded");
  }

  const projectXlsxExport = (globalThis as typeof globalThis & {
    __mikuprojectProjectXlsxExport?: {
      exportProjectWorkbook: (model: ProjectModel) => XlsxWorkbookLike;
    };
  }).__mikuprojectProjectXlsxExport;

  if (!projectXlsxExport) {
    throw new Error("mikuproject Project XLSX export module is not loaded");
  }

  type XlsxCellLike = {
    value?: string | number | boolean;
    numberFormat?: "general" | "integer" | "decimal" | "date" | "datetime" | "percent";
    horizontalAlign?: "left" | "center" | "right";
    bold?: boolean;
    fontSize?: number;
    fillColor?: string;
    border?: "thin";
  };

  type XlsxWorkbookLike = {
    sheets: Array<{
      name: string;
      columns?: Array<{ width?: number }>;
      mergedRanges?: string[];
      dataValidations?: Array<{
        type: "list";
        sqref: string;
        formula1: string;
        allowBlank?: boolean;
      }>;
      rows: Array<{
        height?: number;
        cells: XlsxCellLike[];
      }>;
    }>;
  };

  type ImportChange = {
    scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
    uid: string;
    label: string;
    field: string;
    before: string | number | boolean | undefined;
    after: string | number | boolean;
  };

  (globalThis as typeof globalThis & {
    __mikuprojectProjectXlsx?: {
      exportProjectWorkbook: (model: ProjectModel) => XlsxWorkbookLike;
      importProjectWorkbook: (workbook: XlsxWorkbookLike, baseModel: ProjectModel) => ProjectModel;
      importProjectWorkbookAsProjectModel: (workbook: XlsxWorkbookLike) => ProjectModel;
      importProjectWorkbookDetailed: (workbook: XlsxWorkbookLike, baseModel: ProjectModel) => {
        model: ProjectModel;
        changes: ImportChange[];
      };
    };
  }).__mikuprojectProjectXlsx = {
    exportProjectWorkbook: (model: ProjectModel): XlsxWorkbookLike =>
      projectXlsxExport.exportProjectWorkbook(model),
    importProjectWorkbook: (workbook: XlsxWorkbookLike, baseModel: ProjectModel): ProjectModel =>
      projectXlsxImport.importProjectWorkbook(workbook, baseModel),
    importProjectWorkbookAsProjectModel: (workbook: XlsxWorkbookLike): ProjectModel =>
      projectXlsxImport.importProjectWorkbookAsProjectModel(workbook),
    importProjectWorkbookDetailed: (workbook: XlsxWorkbookLike, baseModel: ProjectModel): {
      model: ProjectModel;
      changes: ImportChange[];
    } => projectXlsxImport.importProjectWorkbookDetailed(workbook, baseModel)
  };
})();
