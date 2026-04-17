/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  const workbookSchema = (globalThis as typeof globalThis & {
    __mikuprojectProjectWorkbookSchema?: {
      HEADER_ROW_INDEX: number;
      DATA_ROW_START_INDEX: number;
      PROJECT_FIELD_ORDER: readonly string[];
      SHEET_HEADERS: {
        Tasks: readonly string[];
        Resources: readonly string[];
        Assignments: readonly string[];
        Calendars: readonly string[];
        NonWorkingDays: readonly string[];
      };
    };
  }).__mikuprojectProjectWorkbookSchema;

  if (!workbookSchema) {
    throw new Error("mikuproject Project Workbook Schema module is not loaded");
  }

  const {
    HEADER_ROW_INDEX,
    DATA_ROW_START_INDEX,
    PROJECT_FIELD_ORDER,
    SHEET_HEADERS
  } = workbookSchema;

  type XlsxCellLike = {
    value?: string | number | boolean;
  };

  type XlsxWorkbookLike = {
    sheets: Array<{
      name: string;
      rows: Array<{
        cells: XlsxCellLike[];
      }>;
    }>;
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

  const projectXlsx = (globalThis as typeof globalThis & {
    __mikuprojectProjectXlsx?: {
      exportProjectWorkbook: (model: ProjectModel) => XlsxWorkbookLike;
    };
  }).__mikuprojectProjectXlsx;

  if (!projectXlsx) {
    throw new Error("mikuproject Project XLSX module is not loaded");
  }

  function exportProjectSheetRows(workbook: XlsxWorkbookLike): WorkbookJsonRow[] {
    const sheet = workbook.sheets.find((item) => item.name === "Project");
    if (!sheet) {
      return [];
    }
    const rows: WorkbookJsonRow[] = [];
    for (const row of sheet.rows.slice(DATA_ROW_START_INDEX)) {
      const field = toJsonScalar(row.cells[0]?.value);
      if (typeof field !== "string" || !PROJECT_FIELD_ORDER.includes(field as typeof PROJECT_FIELD_ORDER[number])) {
        continue;
      }
      rows.push({
        Field: field,
        Value: toJsonScalar(row.cells[1]?.value)
      });
    }
    return rows;
  }

  function exportTabularSheetRows(
    workbook: XlsxWorkbookLike,
    sheetName: keyof typeof SHEET_HEADERS
  ): WorkbookJsonRow[] {
    const sheet = workbook.sheets.find((item) => item.name === sheetName);
    if (!sheet) {
      return [];
    }
    const headers = readHeaderRow(sheet);
    return sheet.rows.slice(DATA_ROW_START_INDEX).map((row) => {
      const item: WorkbookJsonRow = {};
      headers.forEach((header, index) => {
        item[header] = toJsonScalar(row.cells[index]?.value);
      });
      return item;
    });
  }

  function readHeaderRow(sheet: XlsxWorkbookLike["sheets"][number]): string[] {
    return (sheet.rows[HEADER_ROW_INDEX]?.cells || [])
      .map((cell) => (typeof cell.value === "string" ? cell.value : ""))
      .filter((value) => value !== "");
  }

  function toJsonScalar(value: string | number | boolean | undefined): string | number | boolean | null {
    if (value === undefined) {
      return null;
    }
    return value;
  }

  globalThis.__mikuprojectProjectWorkbookJsonExport = {
    exportProjectWorkbookJson(model: ProjectModel): WorkbookJsonDocument {
      const workbook = projectXlsx.exportProjectWorkbook(model);
      return {
        format: "mikuproject_workbook_json",
        version: 1,
        sheets: {
          Project: exportProjectSheetRows(workbook),
          Tasks: exportTabularSheetRows(workbook, "Tasks"),
          Resources: exportTabularSheetRows(workbook, "Resources"),
          Assignments: exportTabularSheetRows(workbook, "Assignments"),
          Calendars: exportTabularSheetRows(workbook, "Calendars"),
          NonWorkingDays: exportTabularSheetRows(workbook, "NonWorkingDays")
        }
      };
    }
  };
})();
