/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  const workbookSchema = (globalThis as typeof globalThis & {
    __mikuprojectProjectWorkbookSchema?: {
      DATA_ROW_START_INDEX: number;
    };
  }).__mikuprojectProjectWorkbookSchema;

  if (!workbookSchema) {
    throw new Error("mikuproject Project Workbook Schema module is not loaded");
  }

  const { DATA_ROW_START_INDEX } = workbookSchema;

  const projectXlsxImportUtil = (globalThis as typeof globalThis & {
    __mikuprojectProjectXlsxImportUtil?: {
      normalizeDateTimeInput: (value: string | undefined) => string | undefined;
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

  function readProjectValueByField(workbook: XlsxWorkbookLike): Map<string, XlsxCellLike | undefined> {
    const projectSheet = workbook.sheets.find((sheet) => sheet.name === "Project");
    const valueByField = new Map<string, XlsxCellLike | undefined>();
    for (const row of projectSheet?.rows.slice(DATA_ROW_START_INDEX) || []) {
      const field = projectXlsxImportUtil.readStringCell(row.cells[0]);
      if (!field) {
        continue;
      }
      valueByField.set(field, row.cells[1]);
    }
    return valueByField;
  }

  function importProjectSheet(workbook: XlsxWorkbookLike, model: ProjectModel, changes: ImportChange[]): void {
    const projectSheet = workbook.sheets.find((sheet) => sheet.name === "Project");
    if (!projectSheet) {
      return;
    }
    const valueByField = readProjectValueByField(workbook);
    const projectLabel = model.project.name;
    projectXlsxImportUtil.assignIfChanged(changes, "project", "project", projectLabel, model.project, "name", "Name", projectXlsxImportUtil.readStringCell(valueByField.get("Name")));
    projectXlsxImportUtil.assignIfChanged(changes, "project", "project", projectLabel, model.project, "title", "Title", projectXlsxImportUtil.readStringCell(valueByField.get("Title")));
    projectXlsxImportUtil.assignIfChanged(changes, "project", "project", projectLabel, model.project, "author", "Author", projectXlsxImportUtil.readStringCell(valueByField.get("Author")));
    projectXlsxImportUtil.assignIfChanged(changes, "project", "project", projectLabel, model.project, "company", "Company", projectXlsxImportUtil.readStringCell(valueByField.get("Company")));
    projectXlsxImportUtil.assignIfChanged(changes, "project", "project", projectLabel, model.project, "startDate", "StartDate", projectXlsxImportUtil.normalizeDateTimeInput(projectXlsxImportUtil.readStringCell(valueByField.get("StartDate"))));
    projectXlsxImportUtil.assignIfChanged(changes, "project", "project", projectLabel, model.project, "finishDate", "FinishDate", projectXlsxImportUtil.normalizeDateTimeInput(projectXlsxImportUtil.readStringCell(valueByField.get("FinishDate"))));
    projectXlsxImportUtil.assignIfChanged(changes, "project", "project", projectLabel, model.project, "currentDate", "CurrentDate", projectXlsxImportUtil.normalizeDateTimeInput(projectXlsxImportUtil.readStringCell(valueByField.get("CurrentDate"))));
    projectXlsxImportUtil.assignIfChanged(changes, "project", "project", projectLabel, model.project, "statusDate", "StatusDate", projectXlsxImportUtil.normalizeDateTimeInput(projectXlsxImportUtil.readStringCell(valueByField.get("StatusDate"))));
    projectXlsxImportUtil.assignIfChanged(changes, "project", "project", projectLabel, model.project, "calendarUID", "CalendarUID", projectXlsxImportUtil.readStringCell(valueByField.get("CalendarUID")));
    projectXlsxImportUtil.assignIfChanged(changes, "project", "project", projectLabel, model.project, "minutesPerDay", "MinutesPerDay", projectXlsxImportUtil.readNumberCell(valueByField.get("MinutesPerDay")));
    projectXlsxImportUtil.assignIfChanged(changes, "project", "project", projectLabel, model.project, "minutesPerWeek", "MinutesPerWeek", projectXlsxImportUtil.readNumberCell(valueByField.get("MinutesPerWeek")));
    projectXlsxImportUtil.assignIfChanged(changes, "project", "project", projectLabel, model.project, "daysPerMonth", "DaysPerMonth", projectXlsxImportUtil.readNumberCell(valueByField.get("DaysPerMonth")));
    projectXlsxImportUtil.assignIfChanged(changes, "project", "project", projectLabel, model.project, "scheduleFromStart", "ScheduleFromStart", projectXlsxImportUtil.readBooleanCell(valueByField.get("ScheduleFromStart")));
  }

  function importProjectSheetAsProjectInfo(workbook: XlsxWorkbookLike): ProjectInfo {
    const valueByField = readProjectValueByField(workbook);
    const name = projectXlsxImportUtil.readStringCell(valueByField.get("Name")) || "Imported Project";
    return {
      name,
      title: projectXlsxImportUtil.readStringCell(valueByField.get("Title")) || undefined,
      author: projectXlsxImportUtil.readStringCell(valueByField.get("Author")) || undefined,
      company: projectXlsxImportUtil.readStringCell(valueByField.get("Company")) || undefined,
      startDate: projectXlsxImportUtil.normalizeDateTimeInput(projectXlsxImportUtil.readStringCell(valueByField.get("StartDate"))) || "",
      finishDate: projectXlsxImportUtil.normalizeDateTimeInput(projectXlsxImportUtil.readStringCell(valueByField.get("FinishDate"))) || "",
      currentDate: projectXlsxImportUtil.normalizeDateTimeInput(projectXlsxImportUtil.readStringCell(valueByField.get("CurrentDate"))) || undefined,
      statusDate: projectXlsxImportUtil.normalizeDateTimeInput(projectXlsxImportUtil.readStringCell(valueByField.get("StatusDate"))) || undefined,
      calendarUID: projectXlsxImportUtil.readStringCell(valueByField.get("CalendarUID")) || undefined,
      minutesPerDay: projectXlsxImportUtil.readNumberCell(valueByField.get("MinutesPerDay")),
      minutesPerWeek: projectXlsxImportUtil.readNumberCell(valueByField.get("MinutesPerWeek")),
      daysPerMonth: projectXlsxImportUtil.readNumberCell(valueByField.get("DaysPerMonth")),
      scheduleFromStart: projectXlsxImportUtil.readBooleanCell(valueByField.get("ScheduleFromStart")) ?? true,
      outlineCodes: [],
      wbsMasks: [],
      extendedAttributes: []
    };
  }

  (globalThis as typeof globalThis & {
    __mikuprojectProjectXlsxImportProject?: {
      importProjectSheet: (workbook: XlsxWorkbookLike, model: ProjectModel, changes: ImportChange[]) => void;
      importProjectSheetAsProjectInfo: (workbook: XlsxWorkbookLike) => ProjectInfo;
    };
  }).__mikuprojectProjectXlsxImportProject = {
    importProjectSheet,
    importProjectSheetAsProjectInfo
  };
})();
