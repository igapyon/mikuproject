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
      readHeaderMap: (sheet: XlsxWorkbookLike["sheets"][number], headerRowIndex: number) => Map<string, number>;
      readStringCellAt: (cells: XlsxCellLike[], index: number | undefined) => string | undefined;
      readNumberCellAt: (cells: XlsxCellLike[], index: number | undefined) => number | undefined;
      readBooleanCellAt: (cells: XlsxCellLike[], index: number | undefined) => boolean | undefined;
      readStringCell: (cell: XlsxCellLike | undefined) => string | undefined;
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
      normalizeDateTimeInput: (value: string | undefined) => string | undefined;
    };
  }).__mikuprojectProjectXlsxImportUtil;

  if (!projectXlsxImportUtil) {
    throw new Error("mikuproject Project XLSX import util module is not loaded");
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

  function importTasksSheet(workbook: XlsxWorkbookLike, model: ProjectModel, changes: ImportChange[]): void {
    const tasksSheet = workbook.sheets.find((sheet) => sheet.name === "Tasks");
    if (!tasksSheet) {
      return;
    }
    const columnIndexByLabel = projectXlsxImportUtil.readHeaderMap(tasksSheet, HEADER_ROW_INDEX);
    const uidColumnIndex = columnIndexByLabel.get("UID");
    if (uidColumnIndex === undefined) {
      return;
    }
    const taskByUid = new Map(model.tasks.map((task) => [task.uid, task]));
    for (const row of tasksSheet.rows.slice(DATA_ROW_START_INDEX)) {
      const uid = projectXlsxImportUtil.readStringCell(row.cells[uidColumnIndex]);
      if (!uid) {
        continue;
      }
      const task = taskByUid.get(uid);
      if (!task) {
        continue;
      }
      const taskLabel = task.name;
      projectXlsxImportUtil.assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "name", "Name", projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Name")));
      projectXlsxImportUtil.assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "start", "Start", projectXlsxImportUtil.normalizeDateTimeInput(projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Start"))));
      projectXlsxImportUtil.assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "finish", "Finish", projectXlsxImportUtil.normalizeDateTimeInput(projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Finish"))));
      projectXlsxImportUtil.assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "duration", "Duration", projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Duration")));
      projectXlsxImportUtil.assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "percentComplete", "PercentComplete", projectXlsxImportUtil.readNumberCellAt(row.cells, columnIndexByLabel.get("PercentComplete")));
      projectXlsxImportUtil.assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "percentWorkComplete", "PercentWorkComplete", projectXlsxImportUtil.readNumberCellAt(row.cells, columnIndexByLabel.get("PercentWorkComplete")));
      projectXlsxImportUtil.assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "milestone", "Milestone", projectXlsxImportUtil.readBooleanCellAt(row.cells, columnIndexByLabel.get("Milestone")));
      projectXlsxImportUtil.assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "summary", "Summary", projectXlsxImportUtil.readBooleanCellAt(row.cells, columnIndexByLabel.get("Summary")));
      projectXlsxImportUtil.assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "critical", "Critical", projectXlsxImportUtil.readBooleanCellAt(row.cells, columnIndexByLabel.get("Critical")));
      projectXlsxImportUtil.assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "calendarUID", "CalendarUID", projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("CalendarUID")));
      projectXlsxImportUtil.assignPredecessorsIfChanged(changes, task.uid, taskLabel, task, projectXlsxImportUtil.parsePredecessorCell(projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Predecessors"))));
      projectXlsxImportUtil.assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "notes", "Notes", projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Notes")));
    }
  }

  function importTasksSheetAsTasks(workbook: XlsxWorkbookLike): TaskModel[] {
    const tasksSheet = workbook.sheets.find((sheet) => sheet.name === "Tasks");
    if (!tasksSheet) {
      return [];
    }
    const columnIndexByLabel = projectXlsxImportUtil.readHeaderMap(tasksSheet, HEADER_ROW_INDEX);
    const uidColumnIndex = columnIndexByLabel.get("UID");
    if (uidColumnIndex === undefined) {
      return [];
    }
    const tasks: TaskModel[] = [];
    for (const [rowIndex, row] of tasksSheet.rows.slice(DATA_ROW_START_INDEX).entries()) {
      const uid = projectXlsxImportUtil.readStringCell(row.cells[uidColumnIndex]);
      if (!uid) {
        continue;
      }
      tasks.push({
        uid,
        id: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("ID")) || String(rowIndex + 1),
        name: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Name")) || uid,
        outlineLevel: projectXlsxImportUtil.readNumberCellAt(row.cells, columnIndexByLabel.get("OutlineLevel")) || 1,
        outlineNumber: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("OutlineNumber")) || String(rowIndex + 1),
        wbs: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("WBS")) || undefined,
        start: projectXlsxImportUtil.normalizeDateTimeInput(projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Start"))) || "",
        finish: projectXlsxImportUtil.normalizeDateTimeInput(projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Finish"))) || "",
        duration: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Duration")) || "",
        milestone: projectXlsxImportUtil.readBooleanCellAt(row.cells, columnIndexByLabel.get("Milestone")) ?? false,
        summary: projectXlsxImportUtil.readBooleanCellAt(row.cells, columnIndexByLabel.get("Summary")) ?? false,
        critical: projectXlsxImportUtil.readBooleanCellAt(row.cells, columnIndexByLabel.get("Critical")) ?? undefined,
        percentComplete: projectXlsxImportUtil.readNumberCellAt(row.cells, columnIndexByLabel.get("PercentComplete")) || 0,
        percentWorkComplete: projectXlsxImportUtil.readNumberCellAt(row.cells, columnIndexByLabel.get("PercentWorkComplete")) ?? undefined,
        calendarUID: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("CalendarUID")) || undefined,
        notes: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Notes")) || undefined,
        predecessors: projectXlsxImportUtil.parsePredecessorCell(projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Predecessors"))) || [],
        extendedAttributes: [],
        baselines: [],
        timephasedData: []
      });
    }
    return tasks;
  }

  function importResourcesSheet(workbook: XlsxWorkbookLike, model: ProjectModel, changes: ImportChange[]): void {
    const resourcesSheet = workbook.sheets.find((sheet) => sheet.name === "Resources");
    if (!resourcesSheet) {
      return;
    }
    const columnIndexByLabel = projectXlsxImportUtil.readHeaderMap(resourcesSheet, HEADER_ROW_INDEX);
    const uidColumnIndex = columnIndexByLabel.get("UID");
    if (uidColumnIndex === undefined) {
      return;
    }
    const resourceByUid = new Map(model.resources.map((resource) => [resource.uid, resource]));
    for (const row of resourcesSheet.rows.slice(DATA_ROW_START_INDEX)) {
      const uid = projectXlsxImportUtil.readStringCell(row.cells[uidColumnIndex]);
      if (!uid) {
        continue;
      }
      const resource = resourceByUid.get(uid);
      if (!resource) {
        continue;
      }
      const resourceLabel = resource.name;
      projectXlsxImportUtil.assignIfChanged(changes, "resources", resource.uid, resourceLabel, resource, "name", "Name", projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Name")));
      projectXlsxImportUtil.assignIfChanged(changes, "resources", resource.uid, resourceLabel, resource, "group", "Group", projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Group")));
      projectXlsxImportUtil.assignIfChanged(changes, "resources", resource.uid, resourceLabel, resource, "maxUnits", "MaxUnits", projectXlsxImportUtil.readNumberCellAt(row.cells, columnIndexByLabel.get("MaxUnits")));
      projectXlsxImportUtil.assignIfChanged(changes, "resources", resource.uid, resourceLabel, resource, "calendarUID", "CalendarUID", projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("CalendarUID")));
    }
  }

  function importResourcesSheetAsResources(workbook: XlsxWorkbookLike): ResourceModel[] {
    const resourcesSheet = workbook.sheets.find((sheet) => sheet.name === "Resources");
    if (!resourcesSheet) {
      return [];
    }
    const columnIndexByLabel = projectXlsxImportUtil.readHeaderMap(resourcesSheet, HEADER_ROW_INDEX);
    const uidColumnIndex = columnIndexByLabel.get("UID");
    if (uidColumnIndex === undefined) {
      return [];
    }
    const resources: ResourceModel[] = [];
    for (const [rowIndex, row] of resourcesSheet.rows.slice(DATA_ROW_START_INDEX).entries()) {
      const uid = projectXlsxImportUtil.readStringCell(row.cells[uidColumnIndex]);
      if (!uid) {
        continue;
      }
      resources.push({
        uid,
        id: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("ID")) || String(rowIndex + 1),
        name: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Name")) || uid,
        type: projectXlsxImportUtil.readNumberCellAt(row.cells, columnIndexByLabel.get("Type")) ?? undefined,
        initials: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Initials")) || undefined,
        group: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Group")) || undefined,
        maxUnits: projectXlsxImportUtil.readNumberCellAt(row.cells, columnIndexByLabel.get("MaxUnits")) ?? undefined,
        calendarUID: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("CalendarUID")) || undefined,
        standardRate: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("StandardRate")) || undefined,
        overtimeRate: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("OvertimeRate")) || undefined,
        costPerUse: projectXlsxImportUtil.readNumberCellAt(row.cells, columnIndexByLabel.get("CostPerUse")) ?? undefined,
        work: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Work")) || undefined,
        actualWork: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("ActualWork")) || undefined,
        remainingWork: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("RemainingWork")) || undefined,
        extendedAttributes: [],
        baselines: [],
        timephasedData: []
      });
    }
    return resources;
  }

  function importAssignmentsSheet(workbook: XlsxWorkbookLike, model: ProjectModel, changes: ImportChange[]): void {
    const assignmentsSheet = workbook.sheets.find((sheet) => sheet.name === "Assignments");
    if (!assignmentsSheet) {
      return;
    }
    const columnIndexByLabel = projectXlsxImportUtil.readHeaderMap(assignmentsSheet, HEADER_ROW_INDEX);
    const uidColumnIndex = columnIndexByLabel.get("UID");
    if (uidColumnIndex === undefined) {
      return;
    }
    const assignmentByUid = new Map(model.assignments.map((assignment) => [assignment.uid, assignment]));
    for (const row of assignmentsSheet.rows.slice(DATA_ROW_START_INDEX)) {
      const uid = projectXlsxImportUtil.readStringCell(row.cells[uidColumnIndex]);
      if (!uid) {
        continue;
      }
      const assignment = assignmentByUid.get(uid);
      if (!assignment) {
        continue;
      }
      const assignmentLabel = `TaskUID=${assignment.taskUid}`;
      projectXlsxImportUtil.assignIfChanged(changes, "assignments", assignment.uid, assignmentLabel, assignment, "units", "Units", projectXlsxImportUtil.readNumberCellAt(row.cells, columnIndexByLabel.get("Units")));
      projectXlsxImportUtil.assignIfChanged(changes, "assignments", assignment.uid, assignmentLabel, assignment, "work", "Work", projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Work")));
      projectXlsxImportUtil.assignIfChanged(changes, "assignments", assignment.uid, assignmentLabel, assignment, "percentWorkComplete", "PercentWorkComplete", projectXlsxImportUtil.readNumberCellAt(row.cells, columnIndexByLabel.get("PercentWorkComplete")));
    }
  }

  function importAssignmentsSheetAsAssignments(workbook: XlsxWorkbookLike): AssignmentModel[] {
    const assignmentsSheet = workbook.sheets.find((sheet) => sheet.name === "Assignments");
    if (!assignmentsSheet) {
      return [];
    }
    const columnIndexByLabel = projectXlsxImportUtil.readHeaderMap(assignmentsSheet, HEADER_ROW_INDEX);
    const uidColumnIndex = columnIndexByLabel.get("UID");
    if (uidColumnIndex === undefined) {
      return [];
    }
    const assignments: AssignmentModel[] = [];
    for (const row of assignmentsSheet.rows.slice(DATA_ROW_START_INDEX)) {
      const uid = projectXlsxImportUtil.readStringCell(row.cells[uidColumnIndex]);
      if (!uid) {
        continue;
      }
      const taskUid = projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("TaskUID"));
      const resourceUid = projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("ResourceUID"));
      if (!taskUid || !resourceUid) {
        continue;
      }
      assignments.push({
        uid,
        taskUid,
        resourceUid,
        start: projectXlsxImportUtil.normalizeDateTimeInput(projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Start"))) || undefined,
        finish: projectXlsxImportUtil.normalizeDateTimeInput(projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Finish"))) || undefined,
        units: projectXlsxImportUtil.readNumberCellAt(row.cells, columnIndexByLabel.get("Units")) ?? undefined,
        work: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Work")) || undefined,
        actualWork: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("ActualWork")) || undefined,
        remainingWork: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("RemainingWork")) || undefined,
        percentWorkComplete: projectXlsxImportUtil.readNumberCellAt(row.cells, columnIndexByLabel.get("PercentWorkComplete")) ?? undefined,
        extendedAttributes: [],
        baselines: [],
        timephasedData: []
      });
    }
    return assignments;
  }

  (globalThis as typeof globalThis & {
    __mikuprojectProjectXlsxImportEntities?: {
      importTasksSheet: (workbook: XlsxWorkbookLike, model: ProjectModel, changes: ImportChange[]) => void;
      importTasksSheetAsTasks: (workbook: XlsxWorkbookLike) => TaskModel[];
      importResourcesSheet: (workbook: XlsxWorkbookLike, model: ProjectModel, changes: ImportChange[]) => void;
      importResourcesSheetAsResources: (workbook: XlsxWorkbookLike) => ResourceModel[];
      importAssignmentsSheet: (workbook: XlsxWorkbookLike, model: ProjectModel, changes: ImportChange[]) => void;
      importAssignmentsSheetAsAssignments: (workbook: XlsxWorkbookLike) => AssignmentModel[];
    };
  }).__mikuprojectProjectXlsxImportEntities = {
    importTasksSheet,
    importTasksSheetAsTasks,
    importResourcesSheet,
    importResourcesSheetAsResources,
    importAssignmentsSheet,
    importAssignmentsSheetAsAssignments
  };
})();
