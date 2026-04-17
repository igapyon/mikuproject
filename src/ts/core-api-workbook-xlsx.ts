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

  const mikuprojectProjectXlsx = (globalThis as typeof globalThis & {
    __mikuprojectProjectXlsx?: {
      exportProjectWorkbook: (model: ProjectModel) => XlsxWorkbookLike;
      importProjectWorkbook: (workbook: XlsxWorkbookLike, baseModel: ProjectModel) => ProjectModel;
      importProjectWorkbookAsProjectModel: (workbook: XlsxWorkbookLike) => ProjectModel;
      importProjectWorkbookDetailed?: (workbook: XlsxWorkbookLike, baseModel: ProjectModel) => {
        model: ProjectModel;
        changes: ImportChange[];
      };
    };
  }).__mikuprojectProjectXlsx;

  if (!mikuprojectProjectXlsx) {
    throw new Error("mikuproject Project XLSX module is not loaded");
  }

  const mikuprojectExcelIo = (globalThis as typeof globalThis & {
    __mikuprojectExcelIo?: {
      XlsxWorkbookCodec: new () => {
        exportWorkbook: (workbook: XlsxWorkbookLike) => Uint8Array;
        importWorkbook: (bytes: Uint8Array) => XlsxWorkbookLike;
      };
    };
  }).__mikuprojectExcelIo;

  if (!mikuprojectExcelIo) {
    throw new Error("mikuproject Excel IO module is not loaded");
  }

  (globalThis as typeof globalThis & {
    __mikuprojectCoreApiWorkbookXlsx?: {
      decodeWorkbook: (bytes: Uint8Array) => XlsxWorkbookLike;
      encodeWorkbook: (workbook: XlsxWorkbookLike) => Uint8Array;
      exportWorkbook: (model: ProjectModel) => XlsxWorkbookLike;
      importAsProjectModel: (workbook: XlsxWorkbookLike) => ProjectModel;
      importIntoProjectModel: (workbook: XlsxWorkbookLike, baseModel: ProjectModel) => ProjectModel;
      importIntoProjectModelDetailed: (
        workbook: XlsxWorkbookLike,
        baseModel: ProjectModel
      ) => { model: ProjectModel; changes: ImportChange[] } | undefined;
    };
  }).__mikuprojectCoreApiWorkbookXlsx = {
    decodeWorkbook: (bytes: Uint8Array) => new mikuprojectExcelIo.XlsxWorkbookCodec().importWorkbook(bytes),
    encodeWorkbook: (workbook: XlsxWorkbookLike) => new mikuprojectExcelIo.XlsxWorkbookCodec().exportWorkbook(workbook),
    exportWorkbook: mikuprojectProjectXlsx.exportProjectWorkbook,
    importAsProjectModel: mikuprojectProjectXlsx.importProjectWorkbookAsProjectModel,
    importIntoProjectModel: mikuprojectProjectXlsx.importProjectWorkbook,
    importIntoProjectModelDetailed: (
      workbook: XlsxWorkbookLike,
      baseModel: ProjectModel
    ): { model: ProjectModel; changes: ImportChange[] } | undefined => {
      if (typeof mikuprojectProjectXlsx.importProjectWorkbookDetailed !== "function") {
        return undefined;
      }
      return mikuprojectProjectXlsx.importProjectWorkbookDetailed(workbook, baseModel);
    }
  };
})();
