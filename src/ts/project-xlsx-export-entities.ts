/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  const workbookSchema = (globalThis as typeof globalThis & {
    __mikuprojectProjectWorkbookSchema?: {
      DATA_ROW_START_INDEX: number;
      SHEET_HEADERS: {
        Tasks: readonly string[];
        Resources: readonly string[];
        Assignments: readonly string[];
      };
    };
  }).__mikuprojectProjectWorkbookSchema;

  if (!workbookSchema) {
    throw new Error("mikuproject Project Workbook Schema module is not loaded");
  }

  const { DATA_ROW_START_INDEX, SHEET_HEADERS } = workbookSchema;

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
      SUMMARY_FILL: string;
      MILESTONE_FILL: string;
      CRITICAL_FILL: string;
      SHEET_THEMES: {
        tasks: { section: string; header: string; label: string };
        resources: { section: string; header: string; label: string };
        assignments: { section: string; header: string; label: string };
      };
      sectionTitleRow: (title: string, columnCount: number, fillColor?: string) => { height: number; cells: XlsxCellLike[] };
      headerRow: (labels: string[], fillColor?: string) => { height: number; cells: XlsxCellLike[] };
      buildBooleanDataValidations: (ranges: Array<string | undefined>) => Array<{
        type: "list";
        sqref: string;
        formula1: string;
        allowBlank?: boolean;
      }> | undefined;
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
    };
  }).__mikuprojectProjectXlsxExportUtil;

  if (!projectXlsxExportUtil) {
    throw new Error("mikuproject Project XLSX export util module is not loaded");
  }

  function buildTasksSheet(model: ProjectModel) {
    return {
      name: "Tasks",
      columns: [
        { width: 10 }, { width: 8 }, { width: 28 }, { width: 12 },
        { width: 14 }, { width: 12 }, { width: 20 }, { width: 20 },
        { width: 14 }, { width: 16 }, { width: 18 }, { width: 12 },
        { width: 12 }, { width: 12 }, { width: 12 }, { width: 20 },
        { width: 34 }
      ],
      mergedRanges: [],
      dataValidations: projectXlsxExportUtil.buildBooleanDataValidations([
        projectXlsxExportUtil.buildColumnRange(11, DATA_ROW_START_INDEX + 1, DATA_ROW_START_INDEX + model.tasks.length),
        projectXlsxExportUtil.buildColumnRange(12, DATA_ROW_START_INDEX + 1, DATA_ROW_START_INDEX + model.tasks.length),
        projectXlsxExportUtil.buildColumnRange(13, DATA_ROW_START_INDEX + 1, DATA_ROW_START_INDEX + model.tasks.length)
      ]),
      rows: [
        projectXlsxExportUtil.sectionTitleRow("Tasks", 17, projectXlsxExportUtil.SHEET_THEMES.tasks.section),
        projectXlsxExportUtil.sectionTitleRow("Task List", 17, projectXlsxExportUtil.SHEET_THEMES.tasks.section),
        projectXlsxExportUtil.headerRow([...SHEET_HEADERS.Tasks], projectXlsxExportUtil.SHEET_THEMES.tasks.header),
        ...model.tasks.map((task, index) => ({
          cells: [
            projectXlsxExportUtil.countCell(task.uid, index),
            projectXlsxExportUtil.countCell(task.id, index),
            projectXlsxExportUtil.editableCell(projectXlsxExportUtil.taskNameCell(task, index)),
            projectXlsxExportUtil.countCell(task.outlineLevel, index),
            projectXlsxExportUtil.textCell(task.outlineNumber, index),
            projectXlsxExportUtil.textCell(task.wbs, index),
            projectXlsxExportUtil.editableCell(projectXlsxExportUtil.dateTimeCell(task.start, index)),
            projectXlsxExportUtil.editableCell(projectXlsxExportUtil.dateTimeCell(task.finish, index)),
            projectXlsxExportUtil.editableCell(projectXlsxExportUtil.durationCell(task.duration, index)),
            projectXlsxExportUtil.editableCell(projectXlsxExportUtil.percentCell(task.percentComplete, index)),
            projectXlsxExportUtil.editableCell(projectXlsxExportUtil.percentCell(task.percentWorkComplete, index)),
            projectXlsxExportUtil.taskFlagCell(task.milestone, index, projectXlsxExportUtil.MILESTONE_FILL),
            projectXlsxExportUtil.taskFlagCell(task.summary, index, projectXlsxExportUtil.SUMMARY_FILL),
            projectXlsxExportUtil.taskFlagCell(task.critical, index, projectXlsxExportUtil.CRITICAL_FILL),
            projectXlsxExportUtil.editableCell(projectXlsxExportUtil.referenceCell(task.calendarUID, index)),
            projectXlsxExportUtil.editableCell(projectXlsxExportUtil.predecessorCell(task.predecessors.map((item) => item.predecessorUid).join(", "), index)),
            projectXlsxExportUtil.editableCell(projectXlsxExportUtil.notesCell(task.notes, index))
          ]
        }))
      ]
    };
  }

  function buildResourcesSheet(model: ProjectModel) {
    const resourceRows = model.resources.length > 0
      ? model.resources.map((resource, index) => ({
        cells: [
          projectXlsxExportUtil.countCell(resource.uid, index),
          projectXlsxExportUtil.countCell(resource.id, index),
          projectXlsxExportUtil.editableCell(projectXlsxExportUtil.entityNameCell(resource.name, index)),
          projectXlsxExportUtil.countCell(resource.type, index),
          projectXlsxExportUtil.textCell(resource.initials, index),
          projectXlsxExportUtil.editableCell(projectXlsxExportUtil.textCell(resource.group, index)),
          projectXlsxExportUtil.editableCell(projectXlsxExportUtil.countCell(resource.maxUnits, index)),
          projectXlsxExportUtil.editableCell(projectXlsxExportUtil.referenceCell(resource.calendarUID, index)),
          projectXlsxExportUtil.textCell(resource.standardRate, index),
          projectXlsxExportUtil.textCell(resource.overtimeRate, index),
          projectXlsxExportUtil.countCell(resource.costPerUse, index),
          projectXlsxExportUtil.workCell(resource.work, index),
          projectXlsxExportUtil.workCell(resource.actualWork, index),
          projectXlsxExportUtil.workCell(resource.remainingWork, index)
        ]
      }))
      : [{
        cells: [
          projectXlsxExportUtil.countCell(undefined, 0),
          projectXlsxExportUtil.countCell(undefined, 0),
          projectXlsxExportUtil.editableCell(projectXlsxExportUtil.entityNameCell(undefined, 0)),
          projectXlsxExportUtil.countCell(undefined, 0),
          projectXlsxExportUtil.textCell(undefined, 0),
          projectXlsxExportUtil.editableCell(projectXlsxExportUtil.textCell(undefined, 0)),
          projectXlsxExportUtil.editableCell(projectXlsxExportUtil.countCell(undefined, 0)),
          projectXlsxExportUtil.editableCell(projectXlsxExportUtil.referenceCell(undefined, 0)),
          projectXlsxExportUtil.textCell(undefined, 0),
          projectXlsxExportUtil.textCell(undefined, 0),
          projectXlsxExportUtil.countCell(undefined, 0),
          projectXlsxExportUtil.workCell(undefined, 0),
          projectXlsxExportUtil.workCell(undefined, 0),
          projectXlsxExportUtil.workCell(undefined, 0)
        ]
      }];

    return {
      name: "Resources",
      columns: [
        { width: 10 }, { width: 8 }, { width: 24 }, { width: 10 },
        { width: 12 }, { width: 18 }, { width: 12 }, { width: 12 },
        { width: 14 }, { width: 14 }, { width: 12 }, { width: 14 },
        { width: 14 }, { width: 14 }
      ],
      mergedRanges: [],
      rows: [
        projectXlsxExportUtil.sectionTitleRow("Resources", 14, projectXlsxExportUtil.SHEET_THEMES.resources.section),
        projectXlsxExportUtil.sectionTitleRow("Resource List", 14, projectXlsxExportUtil.SHEET_THEMES.resources.section),
        projectXlsxExportUtil.headerRow([...SHEET_HEADERS.Resources], projectXlsxExportUtil.SHEET_THEMES.resources.header),
        ...resourceRows
      ]
    };
  }

  function buildAssignmentsSheet(model: ProjectModel) {
    const taskNameByUid = new Map(model.tasks.map((task) => [task.uid, task.name]));
    const resourceNameByUid = new Map(model.resources.map((resource) => [resource.uid, resource.name]));
    const assignmentRows = model.assignments.length > 0
      ? model.assignments.map((assignment, index) => ({
        cells: [
          projectXlsxExportUtil.countCell(assignment.uid, index),
          projectXlsxExportUtil.referenceCell(assignment.taskUid, index),
          projectXlsxExportUtil.entityNameCell(taskNameByUid.get(assignment.taskUid), index),
          projectXlsxExportUtil.referenceCell(assignment.resourceUid, index),
          projectXlsxExportUtil.entityNameCell(resourceNameByUid.get(assignment.resourceUid), index),
          projectXlsxExportUtil.dateTimeCell(assignment.start, index),
          projectXlsxExportUtil.dateTimeCell(assignment.finish, index),
          projectXlsxExportUtil.editableCell(projectXlsxExportUtil.countCell(assignment.units, index)),
          projectXlsxExportUtil.editableCell(projectXlsxExportUtil.workCell(assignment.work, index)),
          projectXlsxExportUtil.workCell(assignment.actualWork, index),
          projectXlsxExportUtil.workCell(assignment.remainingWork, index),
          projectXlsxExportUtil.editableCell(projectXlsxExportUtil.percentCell(assignment.percentWorkComplete, index))
        ]
      }))
      : [{
        cells: [
          projectXlsxExportUtil.countCell(undefined, 0),
          projectXlsxExportUtil.referenceCell(undefined, 0),
          projectXlsxExportUtil.entityNameCell(undefined, 0),
          projectXlsxExportUtil.referenceCell(undefined, 0),
          projectXlsxExportUtil.entityNameCell(undefined, 0),
          projectXlsxExportUtil.dateTimeCell(undefined, 0),
          projectXlsxExportUtil.dateTimeCell(undefined, 0),
          projectXlsxExportUtil.editableCell(projectXlsxExportUtil.countCell(undefined, 0)),
          projectXlsxExportUtil.editableCell(projectXlsxExportUtil.workCell(undefined, 0)),
          projectXlsxExportUtil.workCell(undefined, 0),
          projectXlsxExportUtil.workCell(undefined, 0),
          projectXlsxExportUtil.editableCell(projectXlsxExportUtil.percentCell(undefined, 0))
        ]
      }];

    return {
      name: "Assignments",
      columns: [
        { width: 10 }, { width: 10 }, { width: 24 }, { width: 12 },
        { width: 24 }, { width: 20 }, { width: 20 }, { width: 10 },
        { width: 14 }, { width: 14 }, { width: 14 }, { width: 18 }
      ],
      mergedRanges: [],
      rows: [
        projectXlsxExportUtil.sectionTitleRow("Assignments", 12, projectXlsxExportUtil.SHEET_THEMES.assignments.section),
        projectXlsxExportUtil.sectionTitleRow("Assignment List", 12, projectXlsxExportUtil.SHEET_THEMES.assignments.section),
        projectXlsxExportUtil.headerRow([...SHEET_HEADERS.Assignments], projectXlsxExportUtil.SHEET_THEMES.assignments.header),
        ...assignmentRows
      ]
    };
  }

  (globalThis as typeof globalThis & {
    __mikuprojectProjectXlsxExportEntities?: {
      buildTasksSheet: (model: ProjectModel) => {
        name: string;
        columns?: Array<{ width?: number }>;
        mergedRanges?: string[];
        dataValidations?: Array<{ type: "list"; sqref: string; formula1: string; allowBlank?: boolean }>;
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
  }).__mikuprojectProjectXlsxExportEntities = {
    buildTasksSheet,
    buildResourcesSheet,
    buildAssignmentsSheet
  };
})();
