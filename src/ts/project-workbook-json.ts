/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  const workbookSchema = (globalThis as typeof globalThis & {
    __mikuprojectProjectWorkbookSchema?: {
      SHEET_NAMES: readonly ["Project", "Tasks", "Resources", "Assignments", "Calendars", "NonWorkingDays"];
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
    SHEET_NAMES,
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

  const projectXlsx = (globalThis as typeof globalThis & {
    __mikuprojectProjectXlsx?: {
      exportProjectWorkbook: (model: ProjectModel) => XlsxWorkbookLike;
      importProjectWorkbookDetailed: (workbook: XlsxWorkbookLike, baseModel: ProjectModel) => {
        model: ProjectModel;
        changes: ImportChange[];
      };
    };
  }).__mikuprojectProjectXlsx;

  if (!projectXlsx) {
    throw new Error("mikuproject Project XLSX module is not loaded");
  }

  function exportProjectWorkbookJson(model: ProjectModel): WorkbookJsonDocument {
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

  function importProjectWorkbookJson(documentLike: unknown, baseModel: ProjectModel): {
    model: ProjectModel;
    changes: ImportChange[];
    warnings: WorkbookJsonWarning[];
  } {
    const validation = validateWorkbookJsonDocument(documentLike);
    const document = validation.document;
    const workbook = {
      sheets: [
        buildProjectSheet(document.sheets.Project || []),
        buildTabularSheet("Tasks", document.sheets.Tasks || [], SHEET_HEADERS.Tasks),
        buildTabularSheet("Resources", document.sheets.Resources || [], SHEET_HEADERS.Resources),
        buildTabularSheet("Assignments", document.sheets.Assignments || [], SHEET_HEADERS.Assignments),
        buildTabularSheet("Calendars", document.sheets.Calendars || [], SHEET_HEADERS.Calendars),
        buildTabularSheet("NonWorkingDays", document.sheets.NonWorkingDays || [], SHEET_HEADERS.NonWorkingDays)
      ]
    };
    const result = projectXlsx.importProjectWorkbookDetailed(workbook, baseModel);
    return {
      ...result,
      warnings: validation.warnings
    };
  }

  function importProjectWorkbookJsonAsProjectModel(documentLike: unknown): {
    model: ProjectModel;
    warnings: WorkbookJsonWarning[];
  } {
    const validation = validateWorkbookJsonDocument(documentLike);
    const document = validation.document;
    const workbook = {
      sheets: [
        buildProjectSheet(document.sheets.Project || []),
        buildTabularSheet("Tasks", document.sheets.Tasks || [], SHEET_HEADERS.Tasks),
        buildTabularSheet("Resources", document.sheets.Resources || [], SHEET_HEADERS.Resources),
        buildTabularSheet("Assignments", document.sheets.Assignments || [], SHEET_HEADERS.Assignments),
        buildTabularSheet("Calendars", document.sheets.Calendars || [], SHEET_HEADERS.Calendars),
        buildTabularSheet("NonWorkingDays", document.sheets.NonWorkingDays || [], SHEET_HEADERS.NonWorkingDays)
      ]
    };
    return {
      model: projectXlsx.importProjectWorkbookAsProjectModel(workbook),
      warnings: validation.warnings
    };
  }

  function validateWorkbookJsonDocument(documentLike: unknown): {
    document: WorkbookJsonDocument;
    warnings: WorkbookJsonWarning[];
  } {
    if (!documentLike || typeof documentLike !== "object") {
      throw new Error("workbook JSON がオブジェクトではありません");
    }
    const document = documentLike as Partial<WorkbookJsonDocument>;
    if (document.format !== "mikuproject_workbook_json") {
      throw new Error("format が mikuproject_workbook_json ではありません");
    }
    if (document.version !== 1) {
      throw new Error("version は 1 である必要があります");
    }
    if (!document.sheets || typeof document.sheets !== "object") {
      throw new Error("sheets がありません");
    }
    const warnings: WorkbookJsonWarning[] = [];
    for (const [sheetName, rows] of Object.entries(document.sheets)) {
      if (!SHEET_NAMES.includes(sheetName as typeof SHEET_NAMES[number])) {
        warnings.push({ message: `未知の sheet は無視します: ${sheetName}` });
        continue;
      }
      if (!Array.isArray(rows)) {
        throw new Error(`sheets.${sheetName} は配列である必要があります`);
      }
      for (const [rowIndex, row] of rows.entries()) {
        if (!row || typeof row !== "object" || Array.isArray(row)) {
          throw new Error(`sheets.${sheetName} にオブジェクトではない行があります`);
        }
        for (const key of Object.keys(row)) {
          if (!isKnownColumn(sheetName as typeof SHEET_NAMES[number], key)) {
            warnings.push({ message: `未知の列は無視します: ${sheetName}[${rowIndex}].${key}` });
          }
        }
      }
    }
    return {
      document: document as WorkbookJsonDocument,
      warnings
    };
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

  function exportTabularSheetRows(workbook: XlsxWorkbookLike, sheetName: keyof typeof SHEET_HEADERS): WorkbookJsonRow[] {
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

  function buildProjectSheet(rows: WorkbookJsonRow[]): XlsxWorkbookLike["sheets"][number] {
    const valueByField = new Map<string, string | number | boolean | undefined>();
    for (const row of rows) {
      const field = typeof row.Field === "string" ? row.Field : "";
      if (!PROJECT_FIELD_ORDER.includes(field as typeof PROJECT_FIELD_ORDER[number])) {
        continue;
      }
      valueByField.set(field, toWorkbookScalar(row.Value));
    }
    return {
      name: "Project",
      rows: [
        { cells: [{ value: "Project" }, {}] },
        { cells: [{ value: "Basic Info" }, {}] },
        { cells: [{ value: "Field" }, { value: "Value" }] },
        ...PROJECT_FIELD_ORDER.map((field) => ({
          cells: [
            { value: field },
            { value: valueByField.get(field) }
          ]
        }))
      ]
    };
  }

  function buildTabularSheet(
    sheetName: keyof typeof SHEET_HEADERS,
    rows: WorkbookJsonRow[],
    headers: readonly string[]
  ): XlsxWorkbookLike["sheets"][number] {
    return {
      name: sheetName,
      rows: [
        { cells: [{ value: sheetName }] },
        { cells: [{ value: `${sheetName} List` }] },
        { cells: headers.map((header) => ({ value: header })) },
        ...rows.map((row) => ({
          cells: headers.map((header) => ({
            value: toWorkbookScalar(row[header])
          }))
        }))
      ]
    };
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

  function toWorkbookScalar(value: string | number | boolean | null | undefined): string | number | boolean | undefined {
    if (value === null || value === undefined) {
      return undefined;
    }
    return value;
  }

  function isKnownColumn(sheetName: typeof SHEET_NAMES[number], key: string): boolean {
    if (sheetName === "Project") {
      return key === "Field" || key === "Value";
    }
    return (SHEET_HEADERS[sheetName as keyof typeof SHEET_HEADERS] || []).includes(key);
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
    exportProjectWorkbookJson,
    importProjectWorkbookJsonAsProjectModel,
    importProjectWorkbookJson,
    validateWorkbookJsonDocument
  };
})();
