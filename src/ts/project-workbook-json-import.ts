/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  const workbookSchema = (globalThis as typeof globalThis & {
    __mikuprojectProjectWorkbookSchema?: {
      HEADER_ROW_INDEX: number;
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

  const { HEADER_ROW_INDEX, PROJECT_FIELD_ORDER, SHEET_HEADERS } = workbookSchema;

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
      importProjectWorkbookDetailed: (workbook: XlsxWorkbookLike, baseModel: ProjectModel) => {
        model: ProjectModel;
        changes: ImportChange[];
      };
      importProjectWorkbookAsProjectModel: (workbook: XlsxWorkbookLike) => ProjectModel;
    };
  }).__mikuprojectProjectXlsx;

  if (!projectXlsx) {
    throw new Error("mikuproject Project XLSX module is not loaded");
  }

  const projectWorkbookJsonValidate = (globalThis as typeof globalThis & {
    __mikuprojectProjectWorkbookJsonValidate?: {
      validateWorkbookJsonDocument: (documentLike: unknown) => {
        document: WorkbookJsonDocument;
        warnings: WorkbookJsonWarning[];
      };
    };
  }).__mikuprojectProjectWorkbookJsonValidate;

  if (!projectWorkbookJsonValidate) {
    throw new Error("mikuproject Project Workbook JSON validate module is not loaded");
  }
  function importProjectWorkbookJson(documentLike: unknown, baseModel: ProjectModel): {
    model: ProjectModel;
    changes: ImportChange[];
    warnings: WorkbookJsonWarning[];
  } {
    const validation = projectWorkbookJsonValidate.validateWorkbookJsonDocument(documentLike);
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
    const validation = projectWorkbookJsonValidate.validateWorkbookJsonDocument(documentLike);
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
    return projectWorkbookJsonValidate.validateWorkbookJsonDocument(documentLike);
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

  (globalThis as typeof globalThis & {
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
  }).__mikuprojectProjectWorkbookJsonImport = {
    importProjectWorkbookJson,
    importProjectWorkbookJsonAsProjectModel,
    validateWorkbookJsonDocument
  };
})();
