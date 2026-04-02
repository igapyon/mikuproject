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
    PROJECT_EDITABLE_FIELDS,
    SHEET_HEADERS
  } = workbookSchema;

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

  const HEADER_FILL = "#D9EAF7";
  const SECTION_FILL = "#BFD7EA";
  const LABEL_FILL = "#EDF5FB";
  const ALT_ROW_FILL = "#F9FBFD";
  const DATE_FILL = "#FFF4E8";
  const PERCENT_FILL = "#FCECF3";
  const REFERENCE_FILL = "#EEF7F4";
  const COUNT_FILL = "#F2F5F8";
  const EDITABLE_FILL = "#FDE7C7";
  const DURATION_FILL = "#FBF6ED";
  const NOTES_FILL = "#FFFBEA";
  const NAME_FILL = "#FAF6FF";
  const WORK_FILL = "#F1F8FD";
  const BOOLEAN_TRUE_LABEL = "○";
  const BOOLEAN_FALSE_LABEL = "ー";
  const OPTIONS_SHEET_NAME = "Options";
  const SUMMARY_FILL = "#E6F2E0";
  const MILESTONE_FILL = "#FFF0CF";
  const CRITICAL_FILL = "#F8DDE6";
  const SHEET_THEMES = {
    project: { section: "#BFD7EA", header: "#D9EAF7", label: "#EDF5FB" },
    tasks: { section: "#D4E0EC", header: "#E6EDF4", label: "#F2F6FA" },
    resources: { section: "#C8E3D8", header: "#DDF0E8", label: "#EFF8F4" },
    assignments: { section: "#D7D2EC", header: "#E7E3F5", label: "#F2F0FA" },
    calendars: { section: "#D7E3C4", header: "#E7F0DA", label: "#F2F7EA" },
    nonWorkingDays: { section: "#E9C7D5", header: "#F4DDE6", label: "#FBEEF3" }
  } as const;

  function exportProjectWorkbook(model: ProjectModel): XlsxWorkbookLike {
    const projectSheet = buildProjectSheet(model);
    const tasksSheet = buildTasksSheet(model);
    const resourcesSheet = buildResourcesSheet(model);
    const assignmentsSheet = buildAssignmentsSheet(model);
    const calendarsSheet = buildCalendarsSheet(model);
    const nonWorkingDaysSheet = buildNonWorkingDaysSheet(model);
    return {
      sheets: [
        projectSheet,
        tasksSheet,
        resourcesSheet,
        assignmentsSheet,
        calendarsSheet,
        nonWorkingDaysSheet,
        buildOptionsSheet()
      ]
    };
  }

  function importProjectWorkbook(workbook: XlsxWorkbookLike, baseModel: ProjectModel): ProjectModel {
    return importProjectWorkbookDetailed(workbook, baseModel).model;
  }

  function importProjectWorkbookAsProjectModel(workbook: XlsxWorkbookLike): ProjectModel {
    const project = importProjectSheetAsProjectInfo(workbook);
    const tasks = importTasksSheetAsTasks(workbook);
    const resources = importResourcesSheetAsResources(workbook);
    const assignments = importAssignmentsSheetAsAssignments(workbook);
    const calendars = importCalendarsSheetAsCalendars(workbook, project);
    importNonWorkingDaysSheetAsCalendarExceptions(workbook, calendars);
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
    const nextModel = cloneProjectModel(baseModel);
    const changes: ImportChange[] = [];
    importProjectSheet(workbook, nextModel, changes);
    importTasksSheet(workbook, nextModel, changes);
    importResourcesSheet(workbook, nextModel, changes);
    importAssignmentsSheet(workbook, nextModel, changes);
    importCalendarsSheet(workbook, nextModel, changes);
    importNonWorkingDaysSheet(workbook, nextModel, changes);
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
    for (const row of projectSheet.rows.slice(DATA_ROW_START_INDEX)) {
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
    assignIfChanged(changes, "project", "project", projectLabel, model.project, "startDate", "StartDate", normalizeDateTimeInput(readStringCell(valueByField.get("StartDate"))));
    assignIfChanged(changes, "project", "project", projectLabel, model.project, "finishDate", "FinishDate", normalizeDateTimeInput(readStringCell(valueByField.get("FinishDate"))));
    assignIfChanged(changes, "project", "project", projectLabel, model.project, "currentDate", "CurrentDate", normalizeDateTimeInput(readStringCell(valueByField.get("CurrentDate"))));
    assignIfChanged(changes, "project", "project", projectLabel, model.project, "statusDate", "StatusDate", normalizeDateTimeInput(readStringCell(valueByField.get("StatusDate"))));
    assignIfChanged(changes, "project", "project", projectLabel, model.project, "calendarUID", "CalendarUID", readStringCell(valueByField.get("CalendarUID")));
    assignIfChanged(changes, "project", "project", projectLabel, model.project, "minutesPerDay", "MinutesPerDay", readNumberCell(valueByField.get("MinutesPerDay")));
    assignIfChanged(changes, "project", "project", projectLabel, model.project, "minutesPerWeek", "MinutesPerWeek", readNumberCell(valueByField.get("MinutesPerWeek")));
    assignIfChanged(changes, "project", "project", projectLabel, model.project, "daysPerMonth", "DaysPerMonth", readNumberCell(valueByField.get("DaysPerMonth")));
    assignIfChanged(changes, "project", "project", projectLabel, model.project, "scheduleFromStart", "ScheduleFromStart", readBooleanCell(valueByField.get("ScheduleFromStart")));
  }

  function importProjectSheetAsProjectInfo(workbook: XlsxWorkbookLike): ProjectInfo {
    const projectSheet = workbook.sheets.find((sheet) => sheet.name === "Project");
    const valueByField = new Map<string, XlsxCellLike | undefined>();
    for (const row of projectSheet?.rows.slice(DATA_ROW_START_INDEX) || []) {
      const field = readStringCell(row.cells[0]);
      if (!field) {
        continue;
      }
      valueByField.set(field, row.cells[1]);
    }
    const name = readStringCell(valueByField.get("Name")) || "Imported Project";
    return {
      name,
      title: readStringCell(valueByField.get("Title")) || undefined,
      author: readStringCell(valueByField.get("Author")) || undefined,
      company: readStringCell(valueByField.get("Company")) || undefined,
      startDate: normalizeDateTimeInput(readStringCell(valueByField.get("StartDate"))) || "",
      finishDate: normalizeDateTimeInput(readStringCell(valueByField.get("FinishDate"))) || "",
      currentDate: normalizeDateTimeInput(readStringCell(valueByField.get("CurrentDate"))) || undefined,
      statusDate: normalizeDateTimeInput(readStringCell(valueByField.get("StatusDate"))) || undefined,
      calendarUID: readStringCell(valueByField.get("CalendarUID")) || undefined,
      minutesPerDay: readNumberCell(valueByField.get("MinutesPerDay")),
      minutesPerWeek: readNumberCell(valueByField.get("MinutesPerWeek")),
      daysPerMonth: readNumberCell(valueByField.get("DaysPerMonth")),
      scheduleFromStart: readBooleanCell(valueByField.get("ScheduleFromStart")) ?? true,
      outlineCodes: [],
      wbsMasks: [],
      extendedAttributes: []
    };
  }

  function buildProjectSheet(model: ProjectModel) {
    const project = model.project;
    const rows = [
      titleRow("Project", SHEET_THEMES.project.section),
      titleRow("Basic Info", SHEET_THEMES.project.section),
      headerRow(["Field", "Value"], SHEET_THEMES.project.header),
      ...PROJECT_FIELD_ORDER.slice(0, 8).map((field) => keyValueRow(field, readProjectFieldValue(project, field), SHEET_THEMES.project.label)),
      titleRow("Settings", SHEET_THEMES.project.section),
      ...PROJECT_FIELD_ORDER.slice(8).map((field) => keyValueRow(field, readProjectFieldValue(project, field), SHEET_THEMES.project.label))
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
      dataValidations: buildBooleanDataValidations([
        buildSingleCellReference(1, rowNumberByField.get("ScheduleFromStart"))
      ]),
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
        { width: 12 }, { width: 12 }, { width: 12 }, { width: 20 },
        { width: 34 }
      ],
      mergedRanges: [],
      dataValidations: buildBooleanDataValidations([
        buildColumnRange(11, DATA_ROW_START_INDEX + 1, DATA_ROW_START_INDEX + model.tasks.length),
        buildColumnRange(12, DATA_ROW_START_INDEX + 1, DATA_ROW_START_INDEX + model.tasks.length),
        buildColumnRange(13, DATA_ROW_START_INDEX + 1, DATA_ROW_START_INDEX + model.tasks.length)
      ]),
      rows: [
        sectionTitleRow("Tasks", 17, SHEET_THEMES.tasks.section),
        sectionTitleRow("Task List", 17, SHEET_THEMES.tasks.section),
        headerRow([...SHEET_HEADERS.Tasks], SHEET_THEMES.tasks.header),
        ...model.tasks.map((task, index) => ({
          cells: [
            countCell(task.uid, index),
            countCell(task.id, index),
            editableCell(taskNameCell(task, index)),
            countCell(task.outlineLevel, index),
            textCell(task.outlineNumber, index),
            textCell(task.wbs, index),
            editableCell(dateTimeCell(task.start, index)),
            editableCell(dateTimeCell(task.finish, index)),
            editableCell(durationCell(task.duration, index)),
            editableCell(percentCell(task.percentComplete, index)),
            editableCell(percentCell(task.percentWorkComplete, index)),
            taskFlagCell(task.milestone, index, MILESTONE_FILL),
            taskFlagCell(task.summary, index, SUMMARY_FILL),
            taskFlagCell(task.critical, index, CRITICAL_FILL),
            editableCell(referenceCell(task.calendarUID, index)),
            editableCell(predecessorCell(task.predecessors.map((item) => item.predecessorUid).join(", "), index)),
            editableCell(notesCell(task.notes, index))
          ]
        }))
      ]
    };
  }

  function buildResourcesSheet(model: ProjectModel) {
    const resourceRows = model.resources.length > 0
      ? model.resources.map((resource, index) => ({
        cells: [
          countCell(resource.uid, index),
          countCell(resource.id, index),
          editableCell(entityNameCell(resource.name, index)),
          countCell(resource.type, index),
          textCell(resource.initials, index),
          editableCell(textCell(resource.group, index)),
          editableCell(countCell(resource.maxUnits, index)),
          editableCell(referenceCell(resource.calendarUID, index)),
          textCell(resource.standardRate, index),
          textCell(resource.overtimeRate, index),
          countCell(resource.costPerUse, index),
          workCell(resource.work, index),
          workCell(resource.actualWork, index),
          workCell(resource.remainingWork, index)
        ]
      }))
      : [{
        cells: [
          countCell(undefined, 0),
          countCell(undefined, 0),
          editableCell(entityNameCell(undefined, 0)),
          countCell(undefined, 0),
          textCell(undefined, 0),
          editableCell(textCell(undefined, 0)),
          editableCell(countCell(undefined, 0)),
          editableCell(referenceCell(undefined, 0)),
          textCell(undefined, 0),
          textCell(undefined, 0),
          countCell(undefined, 0),
          workCell(undefined, 0),
          workCell(undefined, 0),
          workCell(undefined, 0)
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
        sectionTitleRow("Resources", 14, SHEET_THEMES.resources.section),
        sectionTitleRow("Resource List", 14, SHEET_THEMES.resources.section),
        headerRow([...SHEET_HEADERS.Resources], SHEET_THEMES.resources.header),
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
          countCell(assignment.uid, index),
          referenceCell(assignment.taskUid, index),
          entityNameCell(taskNameByUid.get(assignment.taskUid), index),
          referenceCell(assignment.resourceUid, index),
          entityNameCell(resourceNameByUid.get(assignment.resourceUid), index),
          dateTimeCell(assignment.start, index),
          dateTimeCell(assignment.finish, index),
          editableCell(countCell(assignment.units, index)),
          editableCell(workCell(assignment.work, index)),
          workCell(assignment.actualWork, index),
          workCell(assignment.remainingWork, index),
          editableCell(percentCell(assignment.percentWorkComplete, index))
        ]
      }))
      : [{
        cells: [
          countCell(undefined, 0),
          referenceCell(undefined, 0),
          entityNameCell(undefined, 0),
          referenceCell(undefined, 0),
          entityNameCell(undefined, 0),
          dateTimeCell(undefined, 0),
          dateTimeCell(undefined, 0),
          editableCell(countCell(undefined, 0)),
          editableCell(workCell(undefined, 0)),
          workCell(undefined, 0),
          workCell(undefined, 0),
          editableCell(percentCell(undefined, 0))
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
        sectionTitleRow("Assignments", 12, SHEET_THEMES.assignments.section),
        sectionTitleRow("Assignment List", 12, SHEET_THEMES.assignments.section),
        headerRow([...SHEET_HEADERS.Assignments], SHEET_THEMES.assignments.header),
        ...assignmentRows
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
      mergedRanges: [],
      dataValidations: buildBooleanDataValidations([
        buildColumnRange(2, DATA_ROW_START_INDEX + 1, DATA_ROW_START_INDEX + model.calendars.length)
      ]),
      rows: [
        sectionTitleRow("Calendars", 7, SHEET_THEMES.calendars.section),
        sectionTitleRow("Calendar List", 7, SHEET_THEMES.calendars.section),
        headerRow([...SHEET_HEADERS.Calendars], SHEET_THEMES.calendars.header),
        ...model.calendars.map((calendar, index) => ({
          cells: [
            countCell(calendar.uid, index),
            editableCell(entityNameCell(calendar.name, index)),
            editableCell(countCell(calendar.isBaseCalendar, index)),
            editableCell(referenceCell(calendar.baseCalendarUID, index)),
            countCell(calendar.weekDays.length, index),
            countCell(calendar.exceptions.length, index),
            countCell(calendar.workWeeks.length, index)
          ]
        }))
      ]
    };
  }

  function buildNonWorkingDaysSheet(model: ProjectModel) {
    const rows = model.calendars.flatMap((calendar) => calendar.exceptions.map((exception, index) => ({
      cells: [
        countCell(calendar.uid, index),
        countCell(index + 1, index),
        textCell(calendar.name, index),
        editableCell(entityNameCell(exception.name, index)),
        editableCell(dateOnlyCell(formatExceptionDate(exception), index)),
        editableCell(dateOnlyCell(formatExceptionBoundaryDate(exception.fromDate), index)),
        editableCell(dateOnlyCell(formatExceptionBoundaryDate(exception.toDate), index)),
        editableCell(countCell(exception.dayWorking, index))
      ]
    })));

    return {
      name: "NonWorkingDays",
      columns: [
        { width: 12 }, { width: 10 }, { width: 22 }, { width: 24 },
        { width: 14 }, { width: 22 }, { width: 22 }, { width: 12 }
      ],
      mergedRanges: [],
      dataValidations: buildBooleanDataValidations([
        buildColumnRange(7, DATA_ROW_START_INDEX + 1, DATA_ROW_START_INDEX + rows.length)
      ]),
      rows: [
        sectionTitleRow("NonWorkingDays", 8, SHEET_THEMES.nonWorkingDays.section),
        sectionTitleRow("Calendar Exceptions", 8, SHEET_THEMES.nonWorkingDays.section),
        headerRow([...SHEET_HEADERS.NonWorkingDays], SHEET_THEMES.nonWorkingDays.header),
        ...rows
      ]
    };
  }

  function headerRow(labels: string[], fillColor = HEADER_FILL) {
    return {
      height: 24,
      cells: labels.map((label) => ({
        value: label,
        bold: true,
        fillColor,
        border: "thin",
        horizontalAlign: "center"
      }))
    };
  }

  function titleRow(title: string, fillColor = SECTION_FILL) {
    return {
      height: 28,
      cells: [
        {
          value: title,
          bold: true,
          fontSize: 16,
          fillColor,
          horizontalAlign: "left"
        },
        {
          fillColor
        }
      ]
    };
  }

  function sectionTitleRow(title: string, columnCount: number, fillColor = SECTION_FILL) {
    return {
      height: 26,
      cells: [
        {
          value: title,
          bold: true,
          fontSize: 14,
          fillColor,
          horizontalAlign: "left"
        },
        ...Array.from({ length: Math.max(0, columnCount - 1) }, () => ({
          fillColor
        }))
      ]
    };
  }

  function keyValueRow(label: string, value: string | number | boolean | undefined, labelFill = LABEL_FILL) {
    return {
      cells: [
        {
          value: label,
          bold: true,
          fillColor: labelFill,
          border: "thin"
        },
        keyValueCell(label, value)
      ]
    };
  }

  function cell(value: string | number | boolean | undefined): XlsxCellLike {
    if (value === undefined) {
      return {};
    }
    return {
      value: stringifyCellValue(value),
      border: "thin"
    };
  }

  function stringifyCellValue(value: string | number | boolean): string {
    if (typeof value === "boolean") {
      return value ? BOOLEAN_TRUE_LABEL : BOOLEAN_FALSE_LABEL;
    }
    return typeof value === "string" ? value : String(value);
  }

  function keyValueCell(label: string, value: string | number | boolean | undefined): XlsxCellLike {
    if (isEditableProjectLabel(label)) {
      return editableCell(buildProjectValueCell(label, value));
    }
    return buildProjectValueCell(label, value);
  }

  function buildProjectValueCell(label: string, value: string | number | boolean | undefined): XlsxCellLike {
    if (isDateTimeLabel(label)) {
      return {
        ...cell(formatDateTimeDisplay(value)),
        fillColor: DATE_FILL
      };
    }
    if (label === "Name" || label === "Title") {
      return {
        ...cell(value),
        fillColor: NAME_FILL,
        bold: true
      };
    }
    if (label === "Author" || label === "Company") {
      return {
        ...cell(value),
        fillColor: NAME_FILL
      };
    }
    if (label === "CalendarUID") {
      return {
        ...cell(value),
        fillColor: REFERENCE_FILL,
        horizontalAlign: "center"
      };
    }
    if (label === "ScheduleFromStart") {
      return {
        ...cell(value),
        fillColor: COUNT_FILL,
        horizontalAlign: "center"
      };
    }
    return {
      ...cell(value),
      fillColor: isNumericSummaryLabel(label) ? COUNT_FILL : undefined
    };
  }

  function textCell(value: string | number | boolean | undefined, rowIndex: number): XlsxCellLike {
    return styledCell(value, rowIndex);
  }

  function taskNameCell(task: TaskModel, rowIndex: number): XlsxCellLike {
    const fillColor = task.summary ? SUMMARY_FILL : (task.critical ? CRITICAL_FILL : undefined);
    return {
      ...styledCell(task.name, rowIndex, { fillColor }),
      bold: task.summary || task.milestone
    };
  }

  function taskFlagCell(value: string | number | boolean | undefined, rowIndex: number, activeFillColor: string): XlsxCellLike {
    const displayValue = value === undefined ? undefined : stringifyCellValue(value);
    return {
      value: displayValue,
      border: "thin",
      fillColor: EDITABLE_FILL,
      horizontalAlign: "center"
    };
  }

  function referenceCell(value: string | number | boolean | undefined, rowIndex: number): XlsxCellLike {
    return styledCell(value, rowIndex, {
      fillColor: REFERENCE_FILL,
      horizontalAlign: "center"
    });
  }

  function countCell(value: string | number | boolean | undefined, rowIndex: number): XlsxCellLike {
    return styledCell(value, rowIndex, {
      fillColor: COUNT_FILL,
      horizontalAlign: "center"
    });
  }

  function percentCell(value: string | number | boolean | undefined, rowIndex: number): XlsxCellLike {
    return styledCell(value, rowIndex, {
      fillColor: PERCENT_FILL,
      horizontalAlign: "center"
    });
  }

  function durationCell(value: string | number | boolean | undefined, rowIndex: number): XlsxCellLike {
    return styledCell(value, rowIndex, {
      fillColor: DURATION_FILL,
      horizontalAlign: "center"
    });
  }

  function predecessorCell(value: string | number | boolean | undefined, rowIndex: number): XlsxCellLike {
    return styledCell(value, rowIndex, {
      fillColor: REFERENCE_FILL
    });
  }

  function notesCell(value: string | number | boolean | undefined, rowIndex: number): XlsxCellLike {
    return styledCell(value, rowIndex, {
      fillColor: NOTES_FILL
    });
  }

  function entityNameCell(value: string | number | boolean | undefined, rowIndex: number): XlsxCellLike {
    return styledCell(value, rowIndex, {
      fillColor: NAME_FILL
    });
  }

  function workCell(value: string | number | boolean | undefined, rowIndex: number): XlsxCellLike {
    return styledCell(value, rowIndex, {
      fillColor: WORK_FILL
    });
  }

  function editableCell(cellLike: XlsxCellLike): XlsxCellLike {
    return {
      ...cellLike,
      border: cellLike.border || "thin",
      fillColor: EDITABLE_FILL
    };
  }

  function dateTimeCell(value: string | number | boolean | undefined, rowIndex: number): XlsxCellLike {
    return styledCell(formatDateTimeDisplay(value), rowIndex, {
      fillColor: DATE_FILL
    });
  }

  function dateOnlyCell(value: string | number | boolean | undefined, rowIndex: number): XlsxCellLike {
    return styledCell(value, rowIndex, {
      fillColor: DATE_FILL,
      horizontalAlign: "center"
    });
  }

  function styledCell(
    value: string | number | boolean | undefined,
    rowIndex: number,
    overrides: Partial<XlsxCellLike> = {}
  ): XlsxCellLike {
    const base = cell(value);
    if (base.value === undefined) {
      return base;
    }
    return {
      ...base,
      fillColor: overrides.fillColor || (rowIndex % 2 === 0 ? ALT_ROW_FILL : undefined),
      horizontalAlign: overrides.horizontalAlign,
      numberFormat: overrides.numberFormat
    };
  }

  function formatDateTimeDisplay(value: string | number | boolean | undefined): string | number | boolean | undefined {
    if (typeof value !== "string") {
      return value;
    }
    return value.replace("T", " ");
  }

  function isDateTimeLabel(label: string): boolean {
    return ["StartDate", "FinishDate", "CurrentDate", "StatusDate"].includes(label);
  }

  function isNumericSummaryLabel(label: string): boolean {
    return ["OutlineCodes", "WBSMasks", "ExtendedAttributes", "MinutesPerDay", "MinutesPerWeek", "DaysPerMonth"].includes(label);
  }

  function isEditableProjectLabel(label: string): boolean {
    return PROJECT_EDITABLE_FIELDS.includes(label);
  }

  function readProjectFieldValue(project: ProjectModel["project"], field: string): string | number | boolean | undefined {
    switch (field) {
      case "Name":
        return project.name;
      case "Title":
        return project.title;
      case "Author":
        return project.author;
      case "Company":
        return project.company;
      case "StartDate":
        return project.startDate;
      case "FinishDate":
        return project.finishDate;
      case "CurrentDate":
        return project.currentDate;
      case "StatusDate":
        return project.statusDate;
      case "CalendarUID":
        return project.calendarUID;
      case "MinutesPerDay":
        return project.minutesPerDay;
      case "MinutesPerWeek":
        return project.minutesPerWeek;
      case "DaysPerMonth":
        return project.daysPerMonth;
      case "ScheduleFromStart":
        return project.scheduleFromStart;
      case "OutlineCodes":
        return project.outlineCodes.length;
      case "WBSMasks":
        return project.wbsMasks.length;
      case "ExtendedAttributes":
        return project.extendedAttributes.length;
      default:
        return undefined;
    }
  }

  function cloneProjectModel(model: ProjectModel): ProjectModel {
    return JSON.parse(JSON.stringify(model)) as ProjectModel;
  }

  function importTasksSheet(workbook: XlsxWorkbookLike, model: ProjectModel, changes: ImportChange[]): void {
    const tasksSheet = workbook.sheets.find((sheet) => sheet.name === "Tasks");
    if (!tasksSheet) {
      return;
    }
    const columnIndexByLabel = readHeaderMap(tasksSheet, HEADER_ROW_INDEX);
    const uidColumnIndex = columnIndexByLabel.get("UID");
    if (uidColumnIndex === undefined) {
      return;
    }
    const taskByUid = new Map(model.tasks.map((task) => [task.uid, task]));
    for (const row of tasksSheet.rows.slice(DATA_ROW_START_INDEX)) {
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
      assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "start", "Start", normalizeDateTimeInput(readStringCellAt(row.cells, columnIndexByLabel.get("Start"))));
      assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "finish", "Finish", normalizeDateTimeInput(readStringCellAt(row.cells, columnIndexByLabel.get("Finish"))));
      assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "duration", "Duration", readStringCellAt(row.cells, columnIndexByLabel.get("Duration")));
      assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "percentComplete", "PercentComplete", readNumberCellAt(row.cells, columnIndexByLabel.get("PercentComplete")));
      assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "percentWorkComplete", "PercentWorkComplete", readNumberCellAt(row.cells, columnIndexByLabel.get("PercentWorkComplete")));
      assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "milestone", "Milestone", readBooleanCellAt(row.cells, columnIndexByLabel.get("Milestone")));
      assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "summary", "Summary", readBooleanCellAt(row.cells, columnIndexByLabel.get("Summary")));
      assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "critical", "Critical", readBooleanCellAt(row.cells, columnIndexByLabel.get("Critical")));
      assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "calendarUID", "CalendarUID", readStringCellAt(row.cells, columnIndexByLabel.get("CalendarUID")));
      assignPredecessorsIfChanged(changes, task.uid, taskLabel, task, parsePredecessorCell(readStringCellAt(row.cells, columnIndexByLabel.get("Predecessors"))));
      assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "notes", "Notes", readStringCellAt(row.cells, columnIndexByLabel.get("Notes")));
    }
  }

  function importTasksSheetAsTasks(workbook: XlsxWorkbookLike): TaskModel[] {
    const tasksSheet = workbook.sheets.find((sheet) => sheet.name === "Tasks");
    if (!tasksSheet) {
      return [];
    }
    const columnIndexByLabel = readHeaderMap(tasksSheet, HEADER_ROW_INDEX);
    const uidColumnIndex = columnIndexByLabel.get("UID");
    if (uidColumnIndex === undefined) {
      return [];
    }
    const tasks: TaskModel[] = [];
    for (const [rowIndex, row] of tasksSheet.rows.slice(DATA_ROW_START_INDEX).entries()) {
      const uid = readStringCell(row.cells[uidColumnIndex]);
      if (!uid) {
        continue;
      }
      tasks.push({
        uid,
        id: readStringCellAt(row.cells, columnIndexByLabel.get("ID")) || String(rowIndex + 1),
        name: readStringCellAt(row.cells, columnIndexByLabel.get("Name")) || uid,
        outlineLevel: readNumberCellAt(row.cells, columnIndexByLabel.get("OutlineLevel")) || 1,
        outlineNumber: readStringCellAt(row.cells, columnIndexByLabel.get("OutlineNumber")) || String(rowIndex + 1),
        wbs: readStringCellAt(row.cells, columnIndexByLabel.get("WBS")) || undefined,
        start: normalizeDateTimeInput(readStringCellAt(row.cells, columnIndexByLabel.get("Start"))) || "",
        finish: normalizeDateTimeInput(readStringCellAt(row.cells, columnIndexByLabel.get("Finish"))) || "",
        duration: readStringCellAt(row.cells, columnIndexByLabel.get("Duration")) || "",
        milestone: readBooleanCellAt(row.cells, columnIndexByLabel.get("Milestone")) ?? false,
        summary: readBooleanCellAt(row.cells, columnIndexByLabel.get("Summary")) ?? false,
        critical: readBooleanCellAt(row.cells, columnIndexByLabel.get("Critical")) ?? undefined,
        percentComplete: readNumberCellAt(row.cells, columnIndexByLabel.get("PercentComplete")) || 0,
        percentWorkComplete: readNumberCellAt(row.cells, columnIndexByLabel.get("PercentWorkComplete")) ?? undefined,
        calendarUID: readStringCellAt(row.cells, columnIndexByLabel.get("CalendarUID")) || undefined,
        notes: readStringCellAt(row.cells, columnIndexByLabel.get("Notes")) || undefined,
        predecessors: parsePredecessorCell(readStringCellAt(row.cells, columnIndexByLabel.get("Predecessors"))) || [],
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
    const columnIndexByLabel = readHeaderMap(resourcesSheet, HEADER_ROW_INDEX);
    const uidColumnIndex = columnIndexByLabel.get("UID");
    if (uidColumnIndex === undefined) {
      return;
    }
    const resourceByUid = new Map(model.resources.map((resource) => [resource.uid, resource]));
    for (const row of resourcesSheet.rows.slice(DATA_ROW_START_INDEX)) {
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
      assignIfChanged(changes, "resources", resource.uid, resourceLabel, resource, "calendarUID", "CalendarUID", readStringCellAt(row.cells, columnIndexByLabel.get("CalendarUID")));
    }
  }

  function importResourcesSheetAsResources(workbook: XlsxWorkbookLike): ResourceModel[] {
    const resourcesSheet = workbook.sheets.find((sheet) => sheet.name === "Resources");
    if (!resourcesSheet) {
      return [];
    }
    const columnIndexByLabel = readHeaderMap(resourcesSheet, HEADER_ROW_INDEX);
    const uidColumnIndex = columnIndexByLabel.get("UID");
    if (uidColumnIndex === undefined) {
      return [];
    }
    const resources: ResourceModel[] = [];
    for (const [rowIndex, row] of resourcesSheet.rows.slice(DATA_ROW_START_INDEX).entries()) {
      const uid = readStringCell(row.cells[uidColumnIndex]);
      if (!uid) {
        continue;
      }
      resources.push({
        uid,
        id: readStringCellAt(row.cells, columnIndexByLabel.get("ID")) || String(rowIndex + 1),
        name: readStringCellAt(row.cells, columnIndexByLabel.get("Name")) || uid,
        type: readNumberCellAt(row.cells, columnIndexByLabel.get("Type")) ?? undefined,
        initials: readStringCellAt(row.cells, columnIndexByLabel.get("Initials")) || undefined,
        group: readStringCellAt(row.cells, columnIndexByLabel.get("Group")) || undefined,
        maxUnits: readNumberCellAt(row.cells, columnIndexByLabel.get("MaxUnits")) ?? undefined,
        calendarUID: readStringCellAt(row.cells, columnIndexByLabel.get("CalendarUID")) || undefined,
        standardRate: readStringCellAt(row.cells, columnIndexByLabel.get("StandardRate")) || undefined,
        overtimeRate: readStringCellAt(row.cells, columnIndexByLabel.get("OvertimeRate")) || undefined,
        costPerUse: readNumberCellAt(row.cells, columnIndexByLabel.get("CostPerUse")) ?? undefined,
        work: readStringCellAt(row.cells, columnIndexByLabel.get("Work")) || undefined,
        actualWork: readStringCellAt(row.cells, columnIndexByLabel.get("ActualWork")) || undefined,
        remainingWork: readStringCellAt(row.cells, columnIndexByLabel.get("RemainingWork")) || undefined,
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
    const columnIndexByLabel = readHeaderMap(assignmentsSheet, HEADER_ROW_INDEX);
    const uidColumnIndex = columnIndexByLabel.get("UID");
    if (uidColumnIndex === undefined) {
      return;
    }
    const assignmentByUid = new Map(model.assignments.map((assignment) => [assignment.uid, assignment]));
    for (const row of assignmentsSheet.rows.slice(DATA_ROW_START_INDEX)) {
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

  function importAssignmentsSheetAsAssignments(workbook: XlsxWorkbookLike): AssignmentModel[] {
    const assignmentsSheet = workbook.sheets.find((sheet) => sheet.name === "Assignments");
    if (!assignmentsSheet) {
      return [];
    }
    const columnIndexByLabel = readHeaderMap(assignmentsSheet, HEADER_ROW_INDEX);
    const uidColumnIndex = columnIndexByLabel.get("UID");
    if (uidColumnIndex === undefined) {
      return [];
    }
    const assignments: AssignmentModel[] = [];
    for (const row of assignmentsSheet.rows.slice(DATA_ROW_START_INDEX)) {
      const uid = readStringCell(row.cells[uidColumnIndex]);
      if (!uid) {
        continue;
      }
      const taskUid = readStringCellAt(row.cells, columnIndexByLabel.get("TaskUID"));
      const resourceUid = readStringCellAt(row.cells, columnIndexByLabel.get("ResourceUID"));
      if (!taskUid || !resourceUid) {
        continue;
      }
      assignments.push({
        uid,
        taskUid,
        resourceUid,
        start: normalizeDateTimeInput(readStringCellAt(row.cells, columnIndexByLabel.get("Start"))) || undefined,
        finish: normalizeDateTimeInput(readStringCellAt(row.cells, columnIndexByLabel.get("Finish"))) || undefined,
        units: readNumberCellAt(row.cells, columnIndexByLabel.get("Units")) ?? undefined,
        work: readStringCellAt(row.cells, columnIndexByLabel.get("Work")) || undefined,
        actualWork: readStringCellAt(row.cells, columnIndexByLabel.get("ActualWork")) || undefined,
        remainingWork: readStringCellAt(row.cells, columnIndexByLabel.get("RemainingWork")) || undefined,
        percentWorkComplete: readNumberCellAt(row.cells, columnIndexByLabel.get("PercentWorkComplete")) ?? undefined,
        extendedAttributes: [],
        baselines: [],
        timephasedData: []
      });
    }
    return assignments;
  }

  function importCalendarsSheet(workbook: XlsxWorkbookLike, model: ProjectModel, changes: ImportChange[]): void {
    const calendarsSheet = workbook.sheets.find((sheet) => sheet.name === "Calendars");
    if (!calendarsSheet) {
      return;
    }
    const columnIndexByLabel = readHeaderMap(calendarsSheet, HEADER_ROW_INDEX);
    const uidColumnIndex = columnIndexByLabel.get("UID");
    if (uidColumnIndex === undefined) {
      return;
    }
    const calendarByUid = new Map(model.calendars.map((calendar) => [calendar.uid, calendar]));
    for (const row of calendarsSheet.rows.slice(DATA_ROW_START_INDEX)) {
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

  function importCalendarsSheetAsCalendars(workbook: XlsxWorkbookLike, project: ProjectInfo): CalendarModel[] {
    const calendarsSheet = workbook.sheets.find((sheet) => sheet.name === "Calendars");
    const calendars: CalendarModel[] = [];
    const columnIndexByLabel = calendarsSheet ? readHeaderMap(calendarsSheet, HEADER_ROW_INDEX) : new Map<string, number>();
    const uidColumnIndex = columnIndexByLabel.get("UID");
    if (calendarsSheet && uidColumnIndex !== undefined) {
      for (const row of calendarsSheet.rows.slice(DATA_ROW_START_INDEX)) {
        const uid = readStringCell(row.cells[uidColumnIndex]);
        if (!uid) {
          continue;
        }
        calendars.push({
          uid,
          name: readStringCellAt(row.cells, columnIndexByLabel.get("Name")) || uid,
          isBaseCalendar: readBooleanCellAt(row.cells, columnIndexByLabel.get("IsBaseCalendar")) ?? false,
          baseCalendarUID: readStringCellAt(row.cells, columnIndexByLabel.get("BaseCalendarUID")) || undefined,
          weekDays: [],
          exceptions: [],
          workWeeks: []
        });
      }
    }
    if (calendars.length === 0) {
      calendars.push({
        uid: project.calendarUID || "1",
        name: "Standard",
        isBaseCalendar: true,
        weekDays: buildDefaultStandardWeekDays(project),
        exceptions: [],
        workWeeks: []
      });
      if (!project.calendarUID) {
        project.calendarUID = calendars[0].uid;
      }
    }
    return calendars;
  }

  function importNonWorkingDaysSheet(workbook: XlsxWorkbookLike, model: ProjectModel, changes: ImportChange[]): void {
    const nonWorkingDaysSheet = workbook.sheets.find((sheet) => sheet.name === "NonWorkingDays");
    if (!nonWorkingDaysSheet) {
      return;
    }
    const columnIndexByLabel = readHeaderMap(nonWorkingDaysSheet, HEADER_ROW_INDEX);
    const calendarUidColumnIndex = columnIndexByLabel.get("CalendarUID");
    const indexColumnIndex = columnIndexByLabel.get("Index");
    if (calendarUidColumnIndex === undefined || indexColumnIndex === undefined) {
      return;
    }
    const calendarByUid = new Map(model.calendars.map((calendar) => [calendar.uid, calendar]));
    for (const row of nonWorkingDaysSheet.rows.slice(DATA_ROW_START_INDEX)) {
      const calendarUid = readStringCell(row.cells[calendarUidColumnIndex]);
      const indexValue = readNumberCell(row.cells[indexColumnIndex]);
      if (!calendarUid || !indexValue) {
        continue;
      }
      const calendar = calendarByUid.get(calendarUid);
      if (!calendar) {
        continue;
      }
      const exception = calendar.exceptions[indexValue - 1];
      if (!exception) {
        continue;
      }
      const exceptionLabel = exception.name || `Exception ${indexValue}`;
      assignIfChanged(changes, "calendars", calendar.uid, calendar.name, exception, "name", `Exception${indexValue}.Name`, readStringCellAt(row.cells, columnIndexByLabel.get("Name")));

      const dateValue = readStringCellAt(row.cells, columnIndexByLabel.get("Date"));
      if (dateValue) {
        const normalizedDate = normalizeDateOnly(dateValue);
        if (normalizedDate) {
          assignIfChanged(changes, "calendars", calendar.uid, calendar.name, exception, "fromDate", `Exception${indexValue}.FromDate`, `${normalizedDate}T00:00:00`);
          assignIfChanged(changes, "calendars", calendar.uid, calendar.name, exception, "toDate", `Exception${indexValue}.ToDate`, `${normalizedDate}T23:59:59`);
        }
      } else {
        assignIfChanged(changes, "calendars", calendar.uid, calendar.name, exception, "fromDate", `Exception${indexValue}.FromDate`, normalizeExceptionBoundaryInput(readStringCellAt(row.cells, columnIndexByLabel.get("FromDate")), false));
        assignIfChanged(changes, "calendars", calendar.uid, calendar.name, exception, "toDate", `Exception${indexValue}.ToDate`, normalizeExceptionBoundaryInput(readStringCellAt(row.cells, columnIndexByLabel.get("ToDate")), true));
      }
      assignIfChanged(changes, "calendars", calendar.uid, calendar.name, exception, "dayWorking", `Exception${indexValue}.DayWorking`, readBooleanCellAt(row.cells, columnIndexByLabel.get("DayWorking")));
    }
  }

  function importNonWorkingDaysSheetAsCalendarExceptions(workbook: XlsxWorkbookLike, calendars: CalendarModel[]): void {
    const nonWorkingDaysSheet = workbook.sheets.find((sheet) => sheet.name === "NonWorkingDays");
    if (!nonWorkingDaysSheet) {
      return;
    }
    const columnIndexByLabel = readHeaderMap(nonWorkingDaysSheet, HEADER_ROW_INDEX);
    const calendarUidColumnIndex = columnIndexByLabel.get("CalendarUID");
    const indexColumnIndex = columnIndexByLabel.get("Index");
    if (calendarUidColumnIndex === undefined || indexColumnIndex === undefined) {
      return;
    }
    const calendarByUid = new Map(calendars.map((calendar) => [calendar.uid, calendar]));
    for (const row of nonWorkingDaysSheet.rows.slice(DATA_ROW_START_INDEX)) {
      const calendarUid = readStringCell(row.cells[calendarUidColumnIndex]);
      const indexValue = readNumberCell(row.cells[indexColumnIndex]);
      if (!calendarUid || !indexValue) {
        continue;
      }
      const calendar = calendarByUid.get(calendarUid);
      if (!calendar) {
        continue;
      }
      const dateValue = readStringCellAt(row.cells, columnIndexByLabel.get("Date"));
      const name = readStringCellAt(row.cells, columnIndexByLabel.get("Name")) || undefined;
      const dayWorking = readBooleanCellAt(row.cells, columnIndexByLabel.get("DayWorking")) ?? false;
      let fromDate: string | undefined;
      let toDate: string | undefined;
      if (dateValue) {
        const normalizedDate = normalizeDateOnly(dateValue);
        if (normalizedDate) {
          fromDate = `${normalizedDate}T00:00:00`;
          toDate = `${normalizedDate}T23:59:59`;
        }
      } else {
        fromDate = normalizeExceptionBoundaryInput(readStringCellAt(row.cells, columnIndexByLabel.get("FromDate")), false) || undefined;
        toDate = normalizeExceptionBoundaryInput(readStringCellAt(row.cells, columnIndexByLabel.get("ToDate")), true) || undefined;
      }
      calendar.exceptions[indexValue - 1] = {
        name,
        fromDate,
        toDate,
        dayWorking,
        workingTimes: []
      };
    }
    for (const calendar of calendars) {
      calendar.exceptions = calendar.exceptions.filter(Boolean);
    }
  }

  function buildDefaultWorkingTimes(project: ProjectInfo): WorkingTimeModel[] {
    const start = project.defaultStartTime || "09:00:00";
    const finish = project.defaultFinishTime || "18:00:00";
    if (start < "12:00:00" && finish > "13:00:00") {
      return [
        { fromTime: start, toTime: "12:00:00" },
        { fromTime: "13:00:00", toTime: finish }
      ];
    }
    return [{ fromTime: start, toTime: finish }];
  }

  function buildDefaultStandardWeekDays(project: ProjectInfo): WeekDayModel[] {
    const workingTimes = buildDefaultWorkingTimes(project);
    return Array.from({ length: 7 }, (_, index) => {
      const dayType = index + 1;
      const dayWorking = dayType !== 1 && dayType !== 7;
      return {
        dayType,
        dayWorking,
        workingTimes: dayWorking ? workingTimes.map((item) => ({ ...item })) : []
      };
    });
  }

  function formatExceptionDate(exception: CalendarExceptionModel): string | undefined {
    const fromDate = exception.fromDate?.slice(0, 10);
    const toDate = exception.toDate?.slice(0, 10);
    if (!fromDate || !toDate) {
      return undefined;
    }
    return fromDate === toDate ? fromDate : undefined;
  }

  function formatExceptionBoundaryDate(value: string | undefined): string | undefined {
    return value?.slice(0, 10);
  }

  function normalizeDateOnly(value: string): string | undefined {
    const trimmed = value.trim();
    const match = trimmed.match(/^(\d{4}-\d{2}-\d{2})/);
    return match ? match[1] : undefined;
  }

  function normalizeDateTimeInput(value: string | undefined): string | undefined {
    if (!value) {
      return value;
    }
    const trimmed = value.trim();
    if (/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/.test(trimmed)) {
      return trimmed.replace(" ", "T");
    }
    return trimmed;
  }

  function normalizeExceptionBoundaryInput(value: string | undefined, isEndOfDay: boolean): string | undefined {
    const normalizedDate = value ? normalizeDateOnly(value) : undefined;
    if (normalizedDate && normalizedDate === value?.trim()) {
      return `${normalizedDate}${isEndOfDay ? "T23:59:59" : "T00:00:00"}`;
    }
    return normalizeDateTimeInput(value);
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
      if (cell.value === "true" || cell.value === "TRUE" || cell.value === "1" || cell.value === BOOLEAN_TRUE_LABEL || cell.value === "〇" || cell.value === "○") {
        return true;
      }
      if (cell.value === "false" || cell.value === "FALSE" || cell.value === "0" || cell.value === BOOLEAN_FALSE_LABEL || cell.value === "-" || cell.value === "－") {
        return false;
      }
    }
    return undefined;
  }

  function isTrueLike(value: string | number | boolean | undefined): boolean {
    return value === true || value === 1 || value === "1" || value === BOOLEAN_TRUE_LABEL || value === "○" || value === "〇";
  }

  function buildOptionsSheet() {
    return {
      name: OPTIONS_SHEET_NAME,
      columns: [{ width: 18 }, { width: 14 }],
      mergedRanges: [],
      rows: [
        headerRow(["BooleanChoice", "Meaning"], HEADER_FILL),
        { cells: [textCell(BOOLEAN_TRUE_LABEL, 0), textCell("true", 0)] },
        { cells: [textCell(BOOLEAN_FALSE_LABEL, 1), textCell("false", 1)] }
      ]
    };
  }

  function buildBooleanDataValidations(ranges: Array<string | undefined>) {
    const sqref = ranges.filter((value): value is string => Boolean(value)).join(" ");
    if (!sqref) {
      return undefined;
    }
    return [{
      type: "list" as const,
      sqref,
      formula1: `${OPTIONS_SHEET_NAME}!$A$2:$A$3`,
      allowBlank: true
    }];
  }

  function buildColumnRange(columnIndex: number, startRow: number, endRow: number): string | undefined {
    if (endRow < startRow) {
      return undefined;
    }
    const columnName = encodeColumnName(columnIndex);
    return `${columnName}${startRow}:${columnName}${endRow}`;
  }

  function buildSingleCellReference(columnIndex: number, rowNumber: number | undefined): string | undefined {
    if (!rowNumber) {
      return undefined;
    }
    return `${encodeColumnName(columnIndex)}${rowNumber}`;
  }

  function encodeColumnName(columnIndex: number): string {
    let current = columnIndex + 1;
    let name = "";
    while (current > 0) {
      const remainder = (current - 1) % 26;
      name = String.fromCharCode(65 + remainder) + name;
      current = Math.floor((current - 1) / 26);
    }
    return name;
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

  function parsePredecessorCell(value: string | undefined): PredecessorModel[] | undefined {
    if (value === undefined) {
      return undefined;
    }
    const normalized = value
      .split(",")
      .map((item) => item.trim())
      .filter(Boolean);
    return normalized.map((predecessorUid) => ({ predecessorUid }));
  }

  function formatPredecessorList(predecessors: PredecessorModel[]): string {
    return predecessors.map((item) => item.predecessorUid).join(", ");
  }

  function assignPredecessorsIfChanged(
    changes: ImportChange[],
    uid: string,
    label: string,
    task: TaskModel,
    predecessors: PredecessorModel[] | undefined
  ): void {
    if (predecessors === undefined) {
      return;
    }
    const beforeText = formatPredecessorList(task.predecessors);
    const afterText = formatPredecessorList(predecessors);
    if (beforeText === afterText) {
      return;
    }
    task.predecessors = predecessors;
    changes.push({
      scope: "tasks",
      uid,
      label,
      field: "Predecessors",
      before: beforeText || undefined,
      after: afterText
    });
  }

  (globalThis as typeof globalThis & {
    __mikuprojectProjectXlsx?: {
      exportProjectWorkbook: (model: ProjectModel) => XlsxWorkbookLike;
      importProjectWorkbook: (workbook: XlsxWorkbookLike, baseModel: ProjectModel) => ProjectModel;
      importProjectWorkbookAsProjectModel: (workbook: XlsxWorkbookLike) => ProjectModel;
      importProjectWorkbookDetailed: (workbook: XlsxWorkbookLike, baseModel: ProjectModel) => {
        model: ProjectModel;
        changes: ImportChange[];
      };
    };
  }).__mikuprojectProjectXlsx = {
    exportProjectWorkbook,
    importProjectWorkbook,
    importProjectWorkbookAsProjectModel,
    importProjectWorkbookDetailed
  };
})();
