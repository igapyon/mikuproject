/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  const workbookSchema = (globalThis as typeof globalThis & {
    __mikuprojectProjectWorkbookSchema?: {
      HEADER_ROW_INDEX: number;
      DATA_ROW_START_INDEX: number;
    };
  }).__mikuprojectProjectWorkbookSchema;

  if (!workbookSchema) {
    throw new Error("mikuproject Project Workbook Schema module is not loaded");
  }

  const { HEADER_ROW_INDEX, DATA_ROW_START_INDEX } = workbookSchema;

  const projectXlsxImportUtil = (globalThis as typeof globalThis & {
    __mikuprojectProjectXlsxImportUtil?: {
      cloneProjectModel: (model: ProjectModel) => ProjectModel;
      buildDefaultWorkingTimes: (project: ProjectInfo) => WorkingTimeModel[];
      buildDefaultStandardWeekDays: (project: ProjectInfo) => WeekDayModel[];
      formatExceptionDate: (exception: CalendarExceptionModel) => string | undefined;
      formatExceptionBoundaryDate: (value: string | undefined) => string | undefined;
      normalizeDateOnly: (value: string) => string | undefined;
      normalizeDateTimeInput: (value: string | undefined) => string | undefined;
      normalizeExceptionBoundaryInput: (value: string | undefined, isEndOfDay: boolean) => string | undefined;
      readHeaderMap: (sheet: XlsxWorkbookLike["sheets"][number], headerRowIndex: number) => Map<string, number>;
      readStringCellAt: (cells: XlsxCellLike[], index: number | undefined) => string | undefined;
      readNumberCellAt: (cells: XlsxCellLike[], index: number | undefined) => number | undefined;
      readBooleanCellAt: (cells: XlsxCellLike[], index: number | undefined) => boolean | undefined;
      readStringCell: (cell: XlsxCellLike | undefined) => string | undefined;
      readNumberCell: (cell: XlsxCellLike | undefined) => number | undefined;
      readBooleanCell: (cell: XlsxCellLike | undefined) => boolean | undefined;
      assignIfChanged: <T extends object, K extends keyof T>(
        changes: ImportChange[],
        scope: ImportChange["scope"],
        uid: string,
        label: string,
        target: T,
        key: K,
        field: string,
        value: T[K] | undefined
      ) => void;
      parsePredecessorCell: (value: string | undefined) => PredecessorModel[] | undefined;
      assignPredecessorsIfChanged: (
        changes: ImportChange[],
        uid: string,
        label: string,
        task: TaskModel,
        predecessors: PredecessorModel[] | undefined
      ) => void;
    };
  }).__mikuprojectProjectXlsxImportUtil;

  if (!projectXlsxImportUtil) {
    throw new Error("mikuproject Project XLSX import util module is not loaded");
  }

  const projectXlsxImportCalendars = (globalThis as typeof globalThis & {
    __mikuprojectProjectXlsxImportCalendars?: {
      importCalendarsSheet: (workbook: XlsxWorkbookLike, model: ProjectModel, changes: ImportChange[]) => void;
      importCalendarsSheetAsCalendars: (workbook: XlsxWorkbookLike, project: ProjectInfo) => CalendarModel[];
      importNonWorkingDaysSheet: (workbook: XlsxWorkbookLike, model: ProjectModel, changes: ImportChange[]) => void;
      importNonWorkingDaysSheetAsCalendarExceptions: (workbook: XlsxWorkbookLike, calendars: CalendarModel[]) => void;
    };
  }).__mikuprojectProjectXlsxImportCalendars;

  if (!projectXlsxImportCalendars) {
    throw new Error("mikuproject Project XLSX import calendars module is not loaded");
  }

  const projectXlsxImportEntities = (globalThis as typeof globalThis & {
    __mikuprojectProjectXlsxImportEntities?: {
      importTasksSheet: (workbook: XlsxWorkbookLike, model: ProjectModel, changes: ImportChange[]) => void;
      importTasksSheetAsTasks: (workbook: XlsxWorkbookLike) => TaskModel[];
      importResourcesSheet: (workbook: XlsxWorkbookLike, model: ProjectModel, changes: ImportChange[]) => void;
      importResourcesSheetAsResources: (workbook: XlsxWorkbookLike) => ResourceModel[];
      importAssignmentsSheet: (workbook: XlsxWorkbookLike, model: ProjectModel, changes: ImportChange[]) => void;
      importAssignmentsSheetAsAssignments: (workbook: XlsxWorkbookLike) => AssignmentModel[];
    };
  }).__mikuprojectProjectXlsxImportEntities;

  if (!projectXlsxImportEntities) {
    throw new Error("mikuproject Project XLSX import entities module is not loaded");
  }

  const projectXlsxImportProject = (globalThis as typeof globalThis & {
    __mikuprojectProjectXlsxImportProject?: {
      importProjectSheet: (workbook: XlsxWorkbookLike, model: ProjectModel, changes: ImportChange[]) => void;
      importProjectSheetAsProjectInfo: (workbook: XlsxWorkbookLike) => ProjectInfo;
    };
  }).__mikuprojectProjectXlsxImportProject;

  if (!projectXlsxImportProject) {
    throw new Error("mikuproject Project XLSX import project module is not loaded");
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

  function importProjectWorkbook(workbook: XlsxWorkbookLike, baseModel: ProjectModel): ProjectModel {
    return importProjectWorkbookDetailed(workbook, baseModel).model;
  }

  function importProjectWorkbookAsProjectModel(workbook: XlsxWorkbookLike): ProjectModel {
    const project = projectXlsxImportProject.importProjectSheetAsProjectInfo(workbook);
    const tasks = projectXlsxImportEntities.importTasksSheetAsTasks(workbook);
    const resources = projectXlsxImportEntities.importResourcesSheetAsResources(workbook);
    const assignments = projectXlsxImportEntities.importAssignmentsSheetAsAssignments(workbook);
    const calendars = projectXlsxImportCalendars.importCalendarsSheetAsCalendars(workbook, project);
    projectXlsxImportCalendars.importNonWorkingDaysSheetAsCalendarExceptions(workbook, calendars);
    return {
      project,
      tasks,
      resources,
      assignments,
      calendars
    };
  }

  function importProjectWorkbookDetailed(workbook: XlsxWorkbookLike, baseModel: ProjectModel): {
    model: ProjectModel;
    changes: ImportChange[];
  } {
    const nextModel = projectXlsxImportUtil.cloneProjectModel(baseModel);
    const changes: ImportChange[] = [];
    projectXlsxImportProject.importProjectSheet(workbook, nextModel, changes);
    projectXlsxImportEntities.importTasksSheet(workbook, nextModel, changes);
    projectXlsxImportEntities.importResourcesSheet(workbook, nextModel, changes);
    projectXlsxImportEntities.importAssignmentsSheet(workbook, nextModel, changes);
    projectXlsxImportCalendars.importCalendarsSheet(workbook, nextModel, changes);
    projectXlsxImportCalendars.importNonWorkingDaysSheet(workbook, nextModel, changes);
    return {
      model: nextModel,
      changes
    };
  }

  (globalThis as typeof globalThis & {
    __mikuprojectProjectXlsxImport?: {
      importProjectWorkbook: (workbook: XlsxWorkbookLike, baseModel: ProjectModel) => ProjectModel;
      importProjectWorkbookAsProjectModel: (workbook: XlsxWorkbookLike) => ProjectModel;
      importProjectWorkbookDetailed: (workbook: XlsxWorkbookLike, baseModel: ProjectModel) => {
        model: ProjectModel;
        changes: ImportChange[];
      };
    };
  }).__mikuprojectProjectXlsxImport = {
    importProjectWorkbook,
    importProjectWorkbookAsProjectModel,
    importProjectWorkbookDetailed
  };
})();
