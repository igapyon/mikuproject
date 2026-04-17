/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  const workbookSchema = (globalThis as typeof globalThis & {
    __mikuprojectProjectWorkbookSchema?: {
      PROJECT_FIELD_ORDER: readonly string[];
    };
  }).__mikuprojectProjectWorkbookSchema;

  if (!workbookSchema) {
    throw new Error("mikuproject Project Workbook Schema module is not loaded");
  }

  const { PROJECT_FIELD_ORDER } = workbookSchema;

  type XlsxCellLike = {
    value?: string | number | boolean;
    numberFormat?: "general" | "integer" | "decimal" | "date" | "datetime" | "percent";
    horizontalAlign?: "left" | "center" | "right";
    bold?: boolean;
    fontSize?: number;
    fillColor?: string;
    border?: "thin";
  };

  const projectXlsxExportUtil = (globalThis as typeof globalThis & {
    __mikuprojectProjectXlsxExportUtil?: {
      SHEET_THEMES: {
        project: { section: string; header: string; label: string };
      };
      titleRow: (title: string, fillColor?: string) => { height: number; cells: XlsxCellLike[] };
      headerRow: (labels: string[], fillColor?: string) => { height: number; cells: XlsxCellLike[] };
      keyValueRow: (label: string, value: string | number | boolean | undefined, labelFill?: string) => { cells: XlsxCellLike[] };
      readProjectFieldValue: (project: ProjectModel["project"], field: string) => string | number | boolean | undefined;
      buildBooleanDataValidations: (ranges: Array<string | undefined>) => Array<{
        type: "list";
        sqref: string;
        formula1: string;
        allowBlank?: boolean;
      }> | undefined;
      buildSingleCellReference: (columnIndex: number, rowNumber: number | undefined) => string | undefined;
    };
  }).__mikuprojectProjectXlsxExportUtil;

  if (!projectXlsxExportUtil) {
    throw new Error("mikuproject Project XLSX export util module is not loaded");
  }

  function buildProjectSheet(model: ProjectModel) {
    const project = model.project;
    const rows = [
      projectXlsxExportUtil.titleRow("Project", projectXlsxExportUtil.SHEET_THEMES.project.section),
      projectXlsxExportUtil.titleRow("Basic Info", projectXlsxExportUtil.SHEET_THEMES.project.section),
      projectXlsxExportUtil.headerRow(["Field", "Value"], projectXlsxExportUtil.SHEET_THEMES.project.header),
      ...PROJECT_FIELD_ORDER.slice(0, 8).map((field) => projectXlsxExportUtil.keyValueRow(field, projectXlsxExportUtil.readProjectFieldValue(project, field), projectXlsxExportUtil.SHEET_THEMES.project.label)),
      projectXlsxExportUtil.titleRow("Settings", projectXlsxExportUtil.SHEET_THEMES.project.section),
      ...PROJECT_FIELD_ORDER.slice(8).map((field) => projectXlsxExportUtil.keyValueRow(field, projectXlsxExportUtil.readProjectFieldValue(project, field), projectXlsxExportUtil.SHEET_THEMES.project.label))
    ];

    const rowNumberByField = new Map<string, number>();
    rows.forEach((row, index) => {
      const field = typeof row.cells[0]?.value === "string" ? row.cells[0].value : undefined;
      if (field) {
        rowNumberByField.set(field, index + 1);
      }
    });

    return {
      name: "Project",
      columns: [{ width: 26 }, { width: 42 }],
      mergedRanges: ["A11:B11"],
      dataValidations: projectXlsxExportUtil.buildBooleanDataValidations([
        projectXlsxExportUtil.buildSingleCellReference(1, rowNumberByField.get("ScheduleFromStart"))
      ]),
      rows
    };
  }

  (globalThis as typeof globalThis & {
    __mikuprojectProjectXlsxExportProject?: {
      buildProjectSheet: (model: ProjectModel) => {
        name: string;
        columns?: Array<{ width?: number }>;
        mergedRanges?: string[];
        dataValidations?: Array<{
          type: "list";
          sqref: string;
          formula1: string;
          allowBlank?: boolean;
        }>;
        rows: Array<{ height?: number; cells: XlsxCellLike[] }>;
      };
    };
  }).__mikuprojectProjectXlsxExportProject = {
    buildProjectSheet
  };
})();
