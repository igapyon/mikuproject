/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  const workbookSchema = (globalThis as typeof globalThis & {
    __mikuprojectProjectWorkbookSchema?: {
      SHEET_NAMES: readonly ["Project", "Tasks", "Resources", "Assignments", "Calendars", "NonWorkingDays"];
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

  const { SHEET_NAMES, PROJECT_FIELD_ORDER, SHEET_HEADERS } = workbookSchema;

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

  function isKnownColumn(sheetName: typeof SHEET_NAMES[number], key: string): boolean {
    if (sheetName === "Project") {
      return key === "Field" || key === "Value";
    }
    if (sheetName === "Project" && key === "Field") {
      return PROJECT_FIELD_ORDER.includes(key as typeof PROJECT_FIELD_ORDER[number]);
    }
    return (SHEET_HEADERS[sheetName as keyof typeof SHEET_HEADERS] || []).includes(key);
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

  globalThis.__mikuprojectProjectWorkbookJsonValidate = {
    validateWorkbookJsonDocument
  };
})();
