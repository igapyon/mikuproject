(() => {
  type XlsxCellLike = {
    value?: string | number | boolean;
    numberFormat?: "general" | "integer" | "decimal" | "date" | "datetime" | "percent";
    horizontalAlign?: "left" | "center" | "right";
    bold?: boolean;
    fillColor?: string;
    border?: "thin";
  };

  type XlsxWorkbookLike = {
    sheets: Array<{
      name: string;
      columns?: Array<{ width?: number }>;
      mergedRanges?: string[];
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

  const HEADER_FILL = "#D9EAF7";

  function exportProjectWorkbook(model: ProjectModel): XlsxWorkbookLike {
    return {
      sheets: [
        buildProjectSheet(model),
        buildTasksSheet(model),
        buildResourcesSheet(model),
        buildAssignmentsSheet(model),
        buildCalendarsSheet(model)
      ]
    };
  }

  function importProjectWorkbook(workbook: XlsxWorkbookLike, baseModel: ProjectModel): ProjectModel {
    return importProjectWorkbookDetailed(workbook, baseModel).model;
  }

  function importProjectWorkbookDetailed(workbook: XlsxWorkbookLike, baseModel: ProjectModel): {
    model: ProjectModel;
    changes: ImportChange[];
  } {
    const nextModel = cloneProjectModel(baseModel);
    const changes: ImportChange[] = [];
    importProjectSheet(workbook, nextModel, changes);
    importTasksSheet(workbook, nextModel, changes);
    importResourcesSheet(workbook, nextModel, changes);
    importAssignmentsSheet(workbook, nextModel, changes);
    importCalendarsSheet(workbook, nextModel, changes);
    return {
      model: nextModel,
      changes
    };
  }

  function importProjectSheet(workbook: XlsxWorkbookLike, model: ProjectModel, changes: ImportChange[]): void {
    const projectSheet = workbook.sheets.find((sheet) => sheet.name === "Project");
    if (!projectSheet) {
      return;
    }
    const valueByField = new Map<string, XlsxCellLike | undefined>();
    for (const row of projectSheet.rows.slice(3)) {
      const field = readStringCell(row.cells[0]);
      if (!field) {
        continue;
      }
      valueByField.set(field, row.cells[1]);
    }
    const projectLabel = model.project.name;
    assignIfChanged(changes, "project", "project", projectLabel, model.project, "name", "Name", readStringCell(valueByField.get("Name")));
    assignIfChanged(changes, "project", "project", projectLabel, model.project, "title", "Title", readStringCell(valueByField.get("Title")));
    assignIfChanged(changes, "project", "project", projectLabel, model.project, "author", "Author", readStringCell(valueByField.get("Author")));
    assignIfChanged(changes, "project", "project", projectLabel, model.project, "company", "Company", readStringCell(valueByField.get("Company")));
    assignIfChanged(changes, "project", "project", projectLabel, model.project, "startDate", "StartDate", readStringCell(valueByField.get("StartDate")));
    assignIfChanged(changes, "project", "project", projectLabel, model.project, "finishDate", "FinishDate", readStringCell(valueByField.get("FinishDate")));
    assignIfChanged(changes, "project", "project", projectLabel, model.project, "currentDate", "CurrentDate", readStringCell(valueByField.get("CurrentDate")));
    assignIfChanged(changes, "project", "project", projectLabel, model.project, "statusDate", "StatusDate", readStringCell(valueByField.get("StatusDate")));
    assignIfChanged(changes, "project", "project", projectLabel, model.project, "calendarUID", "CalendarUID", readStringCell(valueByField.get("CalendarUID")));
    assignIfChanged(changes, "project", "project", projectLabel, model.project, "minutesPerDay", "MinutesPerDay", readNumberCell(valueByField.get("MinutesPerDay")));
    assignIfChanged(changes, "project", "project", projectLabel, model.project, "minutesPerWeek", "MinutesPerWeek", readNumberCell(valueByField.get("MinutesPerWeek")));
    assignIfChanged(changes, "project", "project", projectLabel, model.project, "daysPerMonth", "DaysPerMonth", readNumberCell(valueByField.get("DaysPerMonth")));
    assignIfChanged(changes, "project", "project", projectLabel, model.project, "scheduleFromStart", "ScheduleFromStart", readBooleanCell(valueByField.get("ScheduleFromStart")));
  }

  function buildProjectSheet(model: ProjectModel) {
    const project = model.project;
    const rows = [
      titleRow("Project"),
      titleRow("Basic Info"),
      headerRow(["Field", "Value"]),
      keyValueRow("Name", project.name),
      keyValueRow("Title", project.title),
      keyValueRow("Author", project.author),
      keyValueRow("Company", project.company),
      keyValueRow("StartDate", project.startDate),
      keyValueRow("FinishDate", project.finishDate),
      keyValueRow("CurrentDate", project.currentDate),
      keyValueRow("StatusDate", project.statusDate),
      titleRow("Settings"),
      keyValueRow("CalendarUID", project.calendarUID),
      keyValueRow("MinutesPerDay", project.minutesPerDay),
      keyValueRow("MinutesPerWeek", project.minutesPerWeek),
      keyValueRow("DaysPerMonth", project.daysPerMonth),
      keyValueRow("ScheduleFromStart", project.scheduleFromStart),
      keyValueRow("OutlineCodes", project.outlineCodes.length),
      keyValueRow("WBSMasks", project.wbsMasks.length),
      keyValueRow("ExtendedAttributes", project.extendedAttributes.length)
    ];

    return {
      name: "Project",
      columns: [{ width: 24 }, { width: 32 }],
      mergedRanges: ["A1:B1", "A2:B2", "A11:B11"],
      rows
    };
  }

  function buildTasksSheet(model: ProjectModel) {
    return {
      name: "Tasks",
      columns: [
        { width: 10 }, { width: 8 }, { width: 28 }, { width: 12 },
        { width: 14 }, { width: 12 }, { width: 20 }, { width: 20 },
        { width: 14 }, { width: 16 }, { width: 18 }, { width: 12 },
        { width: 12 }, { width: 12 }, { width: 12 }, { width: 18 }
      ],
      mergedRanges: ["A1:P1", "A2:P2"],
      rows: [
        sectionTitleRow("Tasks", 16),
        sectionTitleRow("Task List", 16),
        headerRow([
          "UID", "ID", "Name", "OutlineLevel", "OutlineNumber", "WBS",
          "Start", "Finish", "Duration", "PercentComplete", "PercentWorkComplete",
          "Milestone", "Summary", "Critical", "CalendarUID", "Predecessors"
        ]),
        ...model.tasks.map((task) => ({
          cells: [
            cell(task.uid),
            cell(task.id),
            cell(task.name),
            cell(task.outlineLevel),
            cell(task.outlineNumber),
            cell(task.wbs),
            cell(task.start),
            cell(task.finish),
            cell(task.duration),
            cell(task.percentComplete),
            cell(task.percentWorkComplete),
            cell(task.milestone),
            cell(task.summary),
            cell(task.critical),
            cell(task.calendarUID),
            cell(task.predecessors.map((item) => item.predecessorUid).join(", "))
          ]
        }))
      ]
    };
  }

  function buildResourcesSheet(model: ProjectModel) {
    return {
      name: "Resources",
      columns: [
        { width: 10 }, { width: 8 }, { width: 24 }, { width: 10 },
        { width: 12 }, { width: 18 }, { width: 12 }, { width: 12 },
        { width: 14 }, { width: 14 }, { width: 12 }, { width: 14 },
        { width: 14 }, { width: 14 }
      ],
      mergedRanges: ["A1:N1", "A2:N2"],
      rows: [
        sectionTitleRow("Resources", 14),
        sectionTitleRow("Resource List", 14),
        headerRow([
          "UID", "ID", "Name", "Type", "Initials", "Group", "MaxUnits",
          "CalendarUID", "StandardRate", "OvertimeRate", "CostPerUse",
          "Work", "ActualWork", "RemainingWork"
        ]),
        ...model.resources.map((resource) => ({
          cells: [
            cell(resource.uid),
            cell(resource.id),
            cell(resource.name),
            cell(resource.type),
            cell(resource.initials),
            cell(resource.group),
            cell(resource.maxUnits),
            cell(resource.calendarUID),
            cell(resource.standardRate),
            cell(resource.overtimeRate),
            cell(resource.costPerUse),
            cell(resource.work),
            cell(resource.actualWork),
            cell(resource.remainingWork)
          ]
        }))
      ]
    };
  }

  function buildAssignmentsSheet(model: ProjectModel) {
    const taskNameByUid = new Map(model.tasks.map((task) => [task.uid, task.name]));
    const resourceNameByUid = new Map(model.resources.map((resource) => [resource.uid, resource.name]));

    return {
      name: "Assignments",
      columns: [
        { width: 10 }, { width: 10 }, { width: 24 }, { width: 12 },
        { width: 24 }, { width: 20 }, { width: 20 }, { width: 10 },
        { width: 14 }, { width: 14 }, { width: 14 }, { width: 18 }
      ],
      mergedRanges: ["A1:L1", "A2:L2"],
      rows: [
        sectionTitleRow("Assignments", 12),
        sectionTitleRow("Assignment List", 12),
        headerRow([
          "UID", "TaskUID", "TaskName", "ResourceUID", "ResourceName", "Start",
          "Finish", "Units", "Work", "ActualWork", "RemainingWork", "PercentWorkComplete"
        ]),
        ...model.assignments.map((assignment) => ({
          cells: [
            cell(assignment.uid),
            cell(assignment.taskUid),
            cell(taskNameByUid.get(assignment.taskUid)),
            cell(assignment.resourceUid),
            cell(resourceNameByUid.get(assignment.resourceUid)),
            cell(assignment.start),
            cell(assignment.finish),
            cell(assignment.units),
            cell(assignment.work),
            cell(assignment.actualWork),
            cell(assignment.remainingWork),
            cell(assignment.percentWorkComplete)
          ]
        }))
      ]
    };
  }

  function buildCalendarsSheet(model: ProjectModel) {
    return {
      name: "Calendars",
      columns: [
        { width: 10 }, { width: 24 }, { width: 14 }, { width: 16 },
        { width: 10 }, { width: 12 }, { width: 10 }
      ],
      mergedRanges: ["A1:G1", "A2:G2"],
      rows: [
        sectionTitleRow("Calendars", 7),
        sectionTitleRow("Calendar List", 7),
        headerRow([
          "UID", "Name", "IsBaseCalendar", "BaseCalendarUID", "WeekDays", "Exceptions", "WorkWeeks"
        ]),
        ...model.calendars.map((calendar) => ({
          cells: [
            cell(calendar.uid),
            cell(calendar.name),
            cell(calendar.isBaseCalendar),
            cell(calendar.baseCalendarUID),
            cell(calendar.weekDays.length),
            cell(calendar.exceptions.length),
            cell(calendar.workWeeks.length)
          ]
        }))
      ]
    };
  }

  function headerRow(labels: string[]) {
    return {
      height: 24,
      cells: labels.map((label) => ({
        value: label,
        bold: true,
        fillColor: HEADER_FILL,
        border: "thin",
        horizontalAlign: "center"
      }))
    };
  }

  function titleRow(title: string) {
    return {
      height: 28,
      cells: [
        {
          value: title,
          bold: true,
          fillColor: HEADER_FILL,
          border: "thin",
          horizontalAlign: "center"
        },
        {}
      ]
    };
  }

  function sectionTitleRow(title: string, columnCount: number) {
    return {
      height: 26,
      cells: [
        {
          value: title,
          bold: true,
          fillColor: HEADER_FILL,
          border: "thin",
          horizontalAlign: "center"
        },
        ...Array.from({ length: Math.max(0, columnCount - 1) }, () => ({}))
      ]
    };
  }

  function keyValueRow(label: string, value: string | number | boolean | undefined) {
    return {
      cells: [
        {
          value: label,
          bold: true,
          fillColor: HEADER_FILL,
          border: "thin"
        },
        cell(value)
      ]
    };
  }

  function cell(value: string | number | boolean | undefined): XlsxCellLike {
    if (value === undefined) {
      return {};
    }
    return {
      value,
      border: "thin"
    };
  }

  function cloneProjectModel(model: ProjectModel): ProjectModel {
    return JSON.parse(JSON.stringify(model)) as ProjectModel;
  }

  function importTasksSheet(workbook: XlsxWorkbookLike, model: ProjectModel, changes: ImportChange[]): void {
    const tasksSheet = workbook.sheets.find((sheet) => sheet.name === "Tasks");
    if (!tasksSheet) {
      return;
    }
    const columnIndexByLabel = readHeaderMap(tasksSheet, 2);
    const uidColumnIndex = columnIndexByLabel.get("UID");
    if (uidColumnIndex === undefined) {
      return;
    }
    const taskByUid = new Map(model.tasks.map((task) => [task.uid, task]));
    for (const row of tasksSheet.rows.slice(3)) {
      const uid = readStringCell(row.cells[uidColumnIndex]);
      if (!uid) {
        continue;
      }
      const task = taskByUid.get(uid);
      if (!task) {
        continue;
      }
      const taskLabel = task.name;
      assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "name", "Name", readStringCellAt(row.cells, columnIndexByLabel.get("Name")));
      assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "start", "Start", readStringCellAt(row.cells, columnIndexByLabel.get("Start")));
      assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "finish", "Finish", readStringCellAt(row.cells, columnIndexByLabel.get("Finish")));
      assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "percentComplete", "PercentComplete", readNumberCellAt(row.cells, columnIndexByLabel.get("PercentComplete")));
      assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "percentWorkComplete", "PercentWorkComplete", readNumberCellAt(row.cells, columnIndexByLabel.get("PercentWorkComplete")));
    }
  }

  function importResourcesSheet(workbook: XlsxWorkbookLike, model: ProjectModel, changes: ImportChange[]): void {
    const resourcesSheet = workbook.sheets.find((sheet) => sheet.name === "Resources");
    if (!resourcesSheet) {
      return;
    }
    const columnIndexByLabel = readHeaderMap(resourcesSheet, 2);
    const uidColumnIndex = columnIndexByLabel.get("UID");
    if (uidColumnIndex === undefined) {
      return;
    }
    const resourceByUid = new Map(model.resources.map((resource) => [resource.uid, resource]));
    for (const row of resourcesSheet.rows.slice(3)) {
      const uid = readStringCell(row.cells[uidColumnIndex]);
      if (!uid) {
        continue;
      }
      const resource = resourceByUid.get(uid);
      if (!resource) {
        continue;
      }
      const resourceLabel = resource.name;
      assignIfChanged(changes, "resources", resource.uid, resourceLabel, resource, "name", "Name", readStringCellAt(row.cells, columnIndexByLabel.get("Name")));
      assignIfChanged(changes, "resources", resource.uid, resourceLabel, resource, "group", "Group", readStringCellAt(row.cells, columnIndexByLabel.get("Group")));
      assignIfChanged(changes, "resources", resource.uid, resourceLabel, resource, "maxUnits", "MaxUnits", readNumberCellAt(row.cells, columnIndexByLabel.get("MaxUnits")));
    }
  }

  function importAssignmentsSheet(workbook: XlsxWorkbookLike, model: ProjectModel, changes: ImportChange[]): void {
    const assignmentsSheet = workbook.sheets.find((sheet) => sheet.name === "Assignments");
    if (!assignmentsSheet) {
      return;
    }
    const columnIndexByLabel = readHeaderMap(assignmentsSheet, 2);
    const uidColumnIndex = columnIndexByLabel.get("UID");
    if (uidColumnIndex === undefined) {
      return;
    }
    const assignmentByUid = new Map(model.assignments.map((assignment) => [assignment.uid, assignment]));
    for (const row of assignmentsSheet.rows.slice(3)) {
      const uid = readStringCell(row.cells[uidColumnIndex]);
      if (!uid) {
        continue;
      }
      const assignment = assignmentByUid.get(uid);
      if (!assignment) {
        continue;
      }
      const assignmentLabel = `TaskUID=${assignment.taskUid}`;
      assignIfChanged(changes, "assignments", assignment.uid, assignmentLabel, assignment, "units", "Units", readNumberCellAt(row.cells, columnIndexByLabel.get("Units")));
      assignIfChanged(changes, "assignments", assignment.uid, assignmentLabel, assignment, "work", "Work", readStringCellAt(row.cells, columnIndexByLabel.get("Work")));
      assignIfChanged(changes, "assignments", assignment.uid, assignmentLabel, assignment, "percentWorkComplete", "PercentWorkComplete", readNumberCellAt(row.cells, columnIndexByLabel.get("PercentWorkComplete")));
    }
  }

  function importCalendarsSheet(workbook: XlsxWorkbookLike, model: ProjectModel, changes: ImportChange[]): void {
    const calendarsSheet = workbook.sheets.find((sheet) => sheet.name === "Calendars");
    if (!calendarsSheet) {
      return;
    }
    const columnIndexByLabel = readHeaderMap(calendarsSheet, 2);
    const uidColumnIndex = columnIndexByLabel.get("UID");
    if (uidColumnIndex === undefined) {
      return;
    }
    const calendarByUid = new Map(model.calendars.map((calendar) => [calendar.uid, calendar]));
    for (const row of calendarsSheet.rows.slice(3)) {
      const uid = readStringCell(row.cells[uidColumnIndex]);
      if (!uid) {
        continue;
      }
      const calendar = calendarByUid.get(uid);
      if (!calendar) {
        continue;
      }
      const calendarLabel = calendar.name;
      assignIfChanged(changes, "calendars", calendar.uid, calendarLabel, calendar, "name", "Name", readStringCellAt(row.cells, columnIndexByLabel.get("Name")));
      assignIfChanged(changes, "calendars", calendar.uid, calendarLabel, calendar, "isBaseCalendar", "IsBaseCalendar", readBooleanCellAt(row.cells, columnIndexByLabel.get("IsBaseCalendar")));
      assignIfChanged(changes, "calendars", calendar.uid, calendarLabel, calendar, "baseCalendarUID", "BaseCalendarUID", readStringCellAt(row.cells, columnIndexByLabel.get("BaseCalendarUID")));
    }
  }

  function readHeaderMap(sheet: XlsxWorkbookLike["sheets"][number], headerRowIndex: number): Map<string, number> {
    const headerRow = sheet.rows[headerRowIndex];
    const columnIndexByLabel = new Map<string, number>();
    if (!headerRow) {
      return columnIndexByLabel;
    }
    headerRow.cells.forEach((cell, index) => {
      if (typeof cell.value === "string" && cell.value) {
        columnIndexByLabel.set(cell.value, index);
      }
    });
    return columnIndexByLabel;
  }

  function readStringCellAt(cells: XlsxCellLike[], index: number | undefined): string | undefined {
    if (index === undefined) {
      return undefined;
    }
    return readStringCell(cells[index]);
  }

  function readNumberCellAt(cells: XlsxCellLike[], index: number | undefined): number | undefined {
    if (index === undefined) {
      return undefined;
    }
    return readNumberCell(cells[index]);
  }

  function readBooleanCellAt(cells: XlsxCellLike[], index: number | undefined): boolean | undefined {
    if (index === undefined) {
      return undefined;
    }
    return readBooleanCell(cells[index]);
  }

  function readStringCell(cell: XlsxCellLike | undefined): string | undefined {
    if (!cell || cell.value === undefined) {
      return undefined;
    }
    if (typeof cell.value === "string") {
      return cell.value;
    }
    if (typeof cell.value === "number" || typeof cell.value === "boolean") {
      return String(cell.value);
    }
    return undefined;
  }

  function readNumberCell(cell: XlsxCellLike | undefined): number | undefined {
    if (!cell || cell.value === undefined) {
      return undefined;
    }
    if (typeof cell.value === "number" && Number.isFinite(cell.value)) {
      return cell.value;
    }
    if (typeof cell.value === "string" && cell.value.trim() !== "") {
      const parsed = Number(cell.value);
      return Number.isFinite(parsed) ? parsed : undefined;
    }
    return undefined;
  }

  function readBooleanCell(cell: XlsxCellLike | undefined): boolean | undefined {
    if (!cell || cell.value === undefined) {
      return undefined;
    }
    if (typeof cell.value === "boolean") {
      return cell.value;
    }
    if (typeof cell.value === "number") {
      return cell.value !== 0;
    }
    if (typeof cell.value === "string") {
      if (cell.value === "true" || cell.value === "TRUE" || cell.value === "1") {
        return true;
      }
      if (cell.value === "false" || cell.value === "FALSE" || cell.value === "0") {
        return false;
      }
    }
    return undefined;
  }

  function assignIfChanged<T extends object, K extends keyof T>(
    changes: ImportChange[],
    scope: ImportChange["scope"],
    uid: string,
    label: string,
    target: T,
    key: K,
    field: string,
    value: T[K] | undefined
  ): void {
    if (value === undefined) {
      return;
    }
    const before = target[key];
    if (before === value) {
      return;
    }
    target[key] = value;
    changes.push({
      scope,
      uid,
      label,
      field,
      before: before as string | number | boolean | undefined,
      after: value as string | number | boolean
    });
  }

  (globalThis as typeof globalThis & {
    __mikuprojectProjectXlsx?: {
      exportProjectWorkbook: (model: ProjectModel) => XlsxWorkbookLike;
      importProjectWorkbook: (workbook: XlsxWorkbookLike, baseModel: ProjectModel) => ProjectModel;
      importProjectWorkbookDetailed: (workbook: XlsxWorkbookLike, baseModel: ProjectModel) => {
        model: ProjectModel;
        changes: ImportChange[];
      };
    };
  }).__mikuprojectProjectXlsx = {
    exportProjectWorkbook,
    importProjectWorkbook,
    importProjectWorkbookDetailed
  };
})();
