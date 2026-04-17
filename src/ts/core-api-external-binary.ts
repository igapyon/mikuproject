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

  const mikuprojectCoreApiMsproject = (globalThis as typeof globalThis & {
    __mikuprojectCoreApiMsproject?: {
      msProject: {
        importFromXml: (xmlText: string) => ProjectModel;
      };
    };
  }).__mikuprojectCoreApiMsproject;

  if (!mikuprojectCoreApiMsproject) {
    throw new Error("mikuproject core api msproject module is not loaded");
  }

  const mikuprojectCoreApiWorkbook = (globalThis as typeof globalThis & {
    __mikuprojectCoreApiWorkbook?: {
      xlsx: {
        decodeWorkbook: (bytes: Uint8Array) => XlsxWorkbookLike;
        importAsProjectModel: (workbook: XlsxWorkbookLike) => ProjectModel;
        importIntoProjectModel: (workbook: XlsxWorkbookLike, baseModel: ProjectModel) => ProjectModel;
        importIntoProjectModelDetailed?: (
          workbook: XlsxWorkbookLike,
          baseModel: ProjectModel
        ) => {
          model: ProjectModel;
          changes: ImportChange[];
        };
      };
    };
  }).__mikuprojectCoreApiWorkbook;

  if (!mikuprojectCoreApiWorkbook) {
    throw new Error("mikuproject core api workbook module is not loaded");
  }

  function importMsProjectXml(sourceText: string, mode: "replace") {
    if (mode !== "replace") {
      throw new Error("MS Project XML は replace import のみ対応です");
    }
    return {
      kind: "ms_project_xml" as const,
      mode,
      model: mikuprojectCoreApiMsproject.msProject.importFromXml(sourceText),
      warnings: [] as []
    };
  }

  function importXlsx(sourceBytes: Uint8Array, mode: "replace" | "merge", baseModel?: ProjectModel) {
    if (mode === "patch") {
      throw new Error("XLSX は replace または merge import のみ対応です");
    }
    const workbook = mikuprojectCoreApiWorkbook.xlsx.decodeWorkbook(sourceBytes);
    if (mode === "replace") {
      return {
        kind: "xlsx" as const,
        mode,
        model: mikuprojectCoreApiWorkbook.xlsx.importAsProjectModel(workbook),
        warnings: [] as []
      };
    }
    if (!baseModel) {
      throw new Error("XLSX merge import には baseModel が必要です");
    }
    if (typeof mikuprojectCoreApiWorkbook.xlsx.importIntoProjectModelDetailed === "function") {
      return {
        kind: "xlsx" as const,
        mode,
        ...mikuprojectCoreApiWorkbook.xlsx.importIntoProjectModelDetailed(workbook, baseModel),
        warnings: [] as []
      };
    }
    return {
      kind: "xlsx" as const,
      mode,
      model: mikuprojectCoreApiWorkbook.xlsx.importIntoProjectModel(workbook, baseModel),
      changes: [] as ImportChange[],
      warnings: [] as []
    };
  }

  (globalThis as typeof globalThis & {
    __mikuprojectCoreApiExternalBinary?: {
      importMsProjectXml: typeof importMsProjectXml;
      importXlsx: typeof importXlsx;
    };
  }).__mikuprojectCoreApiExternalBinary = {
    importMsProjectXml,
    importXlsx
  };
})();
