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
      PROJECT_EDITABLE_FIELDS: readonly string[];
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

  const projectXlsxExportUtil = (globalThis as typeof globalThis & {
    __mikuprojectProjectXlsxExportUtil?: {
      SUMMARY_FILL: string;
      MILESTONE_FILL: string;
      CRITICAL_FILL: string;
      SHEET_THEMES: {
        project: { section: string; header: string; label: string };
        tasks: { section: string; header: string; label: string };
        resources: { section: string; header: string; label: string };
        assignments: { section: string; header: string; label: string };
        calendars: { section: string; header: string; label: string };
        nonWorkingDays: { section: string; header: string; label: string };
      };
      titleRow: (title: string, fillColor?: string) => { height: number; cells: XlsxCellLike[] };
      sectionTitleRow: (title: string, columnCount: number, fillColor?: string) => { height: number; cells: XlsxCellLike[] };
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
      buildColumnRange: (columnIndex: number, startRow: number, endRow: number) => string | undefined;
      countCell: (value: string | number | boolean | undefined, rowIndex: number) => XlsxCellLike;
      editableCell: (cellLike: XlsxCellLike) => XlsxCellLike;
      taskNameCell: (task: TaskModel, rowIndex: number) => XlsxCellLike;
      textCell: (value: string | number | boolean | undefined, rowIndex: number) => XlsxCellLike;
      dateTimeCell: (value: string | number | boolean | undefined, rowIndex: number) => XlsxCellLike;
      durationCell: (value: string | number | boolean | undefined, rowIndex: number) => XlsxCellLike;
      percentCell: (value: string | number | boolean | undefined, rowIndex: number) => XlsxCellLike;
      taskFlagCell: (value: string | number | boolean | undefined, rowIndex: number, activeFillColor: string) => XlsxCellLike;
      referenceCell: (value: string | number | boolean | undefined, rowIndex: number) => XlsxCellLike;
      predecessorCell: (value: string | number | boolean | undefined, rowIndex: number) => XlsxCellLike;
      notesCell: (value: string | number | boolean | undefined, rowIndex: number) => XlsxCellLike;
      entityNameCell: (value: string | number | boolean | undefined, rowIndex: number) => XlsxCellLike;
      workCell: (value: string | number | boolean | undefined, rowIndex: number) => XlsxCellLike;
      dateOnlyCell: (value: string | number | boolean | undefined, rowIndex: number) => XlsxCellLike;
      formatExceptionDate: (exception: CalendarExceptionModel) => string | undefined;
      formatExceptionBoundaryDate: (value: string | undefined) => string | undefined;
      buildOptionsSheet: () => {
        name: string;
        columns?: Array<{ width?: number }>;
        mergedRanges?: string[];
        rows: Array<{ height?: number; cells: XlsxCellLike[] }>;
      };
    };
  }).__mikuprojectProjectXlsxExportUtil;

  if (!projectXlsxExportUtil) {
    throw new Error("mikuproject Project XLSX export util module is not loaded");
  }

  const projectXlsxExportEntities = (globalThis as typeof globalThis & {
    __mikuprojectProjectXlsxExportEntities?: {
      buildTasksSheet: (model: ProjectModel) => {
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
      buildResourcesSheet: (model: ProjectModel) => {
        name: string;
        columns?: Array<{ width?: number }>;
        mergedRanges?: string[];
        rows: Array<{ height?: number; cells: XlsxCellLike[] }>;
      };
      buildAssignmentsSheet: (model: ProjectModel) => {
        name: string;
        columns?: Array<{ width?: number }>;
        mergedRanges?: string[];
        rows: Array<{ height?: number; cells: XlsxCellLike[] }>;
      };
    };
  }).__mikuprojectProjectXlsxExportEntities;

  if (!projectXlsxExportEntities) {
    throw new Error("mikuproject Project XLSX export entities module is not loaded");
  }

  const projectXlsxExportProject = (globalThis as typeof globalThis & {
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
  }).__mikuprojectProjectXlsxExportProject;

  if (!projectXlsxExportProject) {
    throw new Error("mikuproject Project XLSX export project module is not loaded");
  }

  const projectXlsxExportCalendars = (globalThis as typeof globalThis & {
    __mikuprojectProjectXlsxExportCalendars?: {
      buildCalendarsSheet: (model: ProjectModel) => {
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
      buildNonWorkingDaysSheet: (model: ProjectModel) => {
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
  }).__mikuprojectProjectXlsxExportCalendars;

  if (!projectXlsxExportCalendars) {
    throw new Error("mikuproject Project XLSX export calendars module is not loaded");
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
  function exportProjectWorkbook(model: ProjectModel): XlsxWorkbookLike {
    const projectSheet = projectXlsxExportProject.buildProjectSheet(model);
    const tasksSheet = projectXlsxExportEntities.buildTasksSheet(model);
    const resourcesSheet = projectXlsxExportEntities.buildResourcesSheet(model);
    const assignmentsSheet = projectXlsxExportEntities.buildAssignmentsSheet(model);
    const calendarsSheet = projectXlsxExportCalendars.buildCalendarsSheet(model);
    const nonWorkingDaysSheet = projectXlsxExportCalendars.buildNonWorkingDaysSheet(model);
    return {
      sheets: [
        projectSheet,
        tasksSheet,
        resourcesSheet,
        assignmentsSheet,
        calendarsSheet,
        nonWorkingDaysSheet,
        projectXlsxExportUtil.buildOptionsSheet()
      ]
    };
  }
  globalThis.__mikuprojectProjectXlsxExport = {
    exportProjectWorkbook
  };
})();
