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

  const projectXlsxImportUtil = (globalThis as typeof globalThis & {
    __mikuprojectProjectXlsxImportUtil?: {
      buildDefaultStandardWeekDays: (project: ProjectInfo) => WeekDayModel[];
      readHeaderMap: (sheet: XlsxWorkbookLike["sheets"][number], headerRowIndex: number) => Map<string, number>;
      readStringCellAt: (cells: XlsxCellLike[], index: number | undefined) => string | undefined;
      readNumberCellAt: (cells: XlsxCellLike[], index: number | undefined) => number | undefined;
      readBooleanCellAt: (cells: XlsxCellLike[], index: number | undefined) => boolean | undefined;
      readStringCell: (cell: XlsxCellLike | undefined) => string | undefined;
      readNumberCell: (cell: XlsxCellLike | undefined) => number | undefined;
      normalizeDateOnly: (value: string) => string | undefined;
      normalizeExceptionBoundaryInput: (value: string | undefined, isEndOfDay: boolean) => string | undefined;
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

  function importCalendarsSheet(workbook: XlsxWorkbookLike, model: ProjectModel, changes: ImportChange[]): void {
    const calendarsSheet = workbook.sheets.find((sheet) => sheet.name === "Calendars");
    if (!calendarsSheet) {
      return;
    }
    const columnIndexByLabel = projectXlsxImportUtil.readHeaderMap(calendarsSheet, HEADER_ROW_INDEX);
    const uidColumnIndex = columnIndexByLabel.get("UID");
    if (uidColumnIndex === undefined) {
      return;
    }
    const calendarByUid = new Map(model.calendars.map((calendar) => [calendar.uid, calendar]));
    for (const row of calendarsSheet.rows.slice(DATA_ROW_START_INDEX)) {
      const uid = projectXlsxImportUtil.readStringCell(row.cells[uidColumnIndex]);
      if (!uid) {
        continue;
      }
      const calendar = calendarByUid.get(uid);
      if (!calendar) {
        continue;
      }
      const calendarLabel = calendar.name;
      projectXlsxImportUtil.assignIfChanged(changes, "calendars", calendar.uid, calendarLabel, calendar, "name", "Name", projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Name")));
      projectXlsxImportUtil.assignIfChanged(changes, "calendars", calendar.uid, calendarLabel, calendar, "isBaseCalendar", "IsBaseCalendar", projectXlsxImportUtil.readBooleanCellAt(row.cells, columnIndexByLabel.get("IsBaseCalendar")));
      projectXlsxImportUtil.assignIfChanged(changes, "calendars", calendar.uid, calendarLabel, calendar, "baseCalendarUID", "BaseCalendarUID", projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("BaseCalendarUID")));
    }
  }

  function importCalendarsSheetAsCalendars(workbook: XlsxWorkbookLike, project: ProjectInfo): CalendarModel[] {
    const calendarsSheet = workbook.sheets.find((sheet) => sheet.name === "Calendars");
    const calendars: CalendarModel[] = [];
    const columnIndexByLabel = calendarsSheet ? projectXlsxImportUtil.readHeaderMap(calendarsSheet, HEADER_ROW_INDEX) : new Map<string, number>();
    const uidColumnIndex = columnIndexByLabel.get("UID");
    if (calendarsSheet && uidColumnIndex !== undefined) {
      for (const row of calendarsSheet.rows.slice(DATA_ROW_START_INDEX)) {
        const uid = projectXlsxImportUtil.readStringCell(row.cells[uidColumnIndex]);
        if (!uid) {
          continue;
        }
        calendars.push({
          uid,
          name: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Name")) || uid,
          isBaseCalendar: projectXlsxImportUtil.readBooleanCellAt(row.cells, columnIndexByLabel.get("IsBaseCalendar")) ?? false,
          baseCalendarUID: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("BaseCalendarUID")) || undefined,
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
        weekDays: projectXlsxImportUtil.buildDefaultStandardWeekDays(project),
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
    const columnIndexByLabel = projectXlsxImportUtil.readHeaderMap(nonWorkingDaysSheet, HEADER_ROW_INDEX);
    const calendarUidColumnIndex = columnIndexByLabel.get("CalendarUID");
    const indexColumnIndex = columnIndexByLabel.get("Index");
    if (calendarUidColumnIndex === undefined || indexColumnIndex === undefined) {
      return;
    }
    const calendarByUid = new Map(model.calendars.map((calendar) => [calendar.uid, calendar]));
    for (const row of nonWorkingDaysSheet.rows.slice(DATA_ROW_START_INDEX)) {
      const calendarUid = projectXlsxImportUtil.readStringCell(row.cells[calendarUidColumnIndex]);
      const indexValue = projectXlsxImportUtil.readNumberCell(row.cells[indexColumnIndex]);
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
      projectXlsxImportUtil.assignIfChanged(changes, "calendars", calendar.uid, calendar.name, exception, "name", `Exception${indexValue}.Name`, projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Name")));
      const dateValue = projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Date"));
      if (dateValue) {
        const normalizedDate = projectXlsxImportUtil.normalizeDateOnly(dateValue);
        if (normalizedDate) {
          projectXlsxImportUtil.assignIfChanged(changes, "calendars", calendar.uid, calendar.name, exception, "fromDate", `Exception${indexValue}.FromDate`, `${normalizedDate}T00:00:00`);
          projectXlsxImportUtil.assignIfChanged(changes, "calendars", calendar.uid, calendar.name, exception, "toDate", `Exception${indexValue}.ToDate`, `${normalizedDate}T23:59:59`);
        }
      } else {
        projectXlsxImportUtil.assignIfChanged(changes, "calendars", calendar.uid, calendar.name, exception, "fromDate", `Exception${indexValue}.FromDate`, projectXlsxImportUtil.normalizeExceptionBoundaryInput(projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("FromDate")), false));
        projectXlsxImportUtil.assignIfChanged(changes, "calendars", calendar.uid, calendar.name, exception, "toDate", `Exception${indexValue}.ToDate`, projectXlsxImportUtil.normalizeExceptionBoundaryInput(projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("ToDate")), true));
      }
      projectXlsxImportUtil.assignIfChanged(changes, "calendars", calendar.uid, calendar.name, exception, "dayWorking", `Exception${indexValue}.DayWorking`, projectXlsxImportUtil.readBooleanCellAt(row.cells, columnIndexByLabel.get("DayWorking")));
    }
  }

  function importNonWorkingDaysSheetAsCalendarExceptions(workbook: XlsxWorkbookLike, calendars: CalendarModel[]): void {
    const nonWorkingDaysSheet = workbook.sheets.find((sheet) => sheet.name === "NonWorkingDays");
    if (!nonWorkingDaysSheet) {
      return;
    }
    const columnIndexByLabel = projectXlsxImportUtil.readHeaderMap(nonWorkingDaysSheet, HEADER_ROW_INDEX);
    const calendarUidColumnIndex = columnIndexByLabel.get("CalendarUID");
    const indexColumnIndex = columnIndexByLabel.get("Index");
    if (calendarUidColumnIndex === undefined || indexColumnIndex === undefined) {
      return;
    }
    const calendarByUid = new Map(calendars.map((calendar) => [calendar.uid, calendar]));
    for (const row of nonWorkingDaysSheet.rows.slice(DATA_ROW_START_INDEX)) {
      const calendarUid = projectXlsxImportUtil.readStringCell(row.cells[calendarUidColumnIndex]);
      const indexValue = projectXlsxImportUtil.readNumberCell(row.cells[indexColumnIndex]);
      if (!calendarUid || !indexValue) {
        continue;
      }
      const calendar = calendarByUid.get(calendarUid);
      if (!calendar) {
        continue;
      }
      const dateValue = projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Date"));
      const name = projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Name")) || undefined;
      const dayWorking = projectXlsxImportUtil.readBooleanCellAt(row.cells, columnIndexByLabel.get("DayWorking")) ?? false;
      let fromDate: string | undefined;
      let toDate: string | undefined;
      if (dateValue) {
        const normalizedDate = projectXlsxImportUtil.normalizeDateOnly(dateValue);
        if (normalizedDate) {
          fromDate = `${normalizedDate}T00:00:00`;
          toDate = `${normalizedDate}T23:59:59`;
        }
      } else {
        fromDate = projectXlsxImportUtil.normalizeExceptionBoundaryInput(projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("FromDate")), false) || undefined;
        toDate = projectXlsxImportUtil.normalizeExceptionBoundaryInput(projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("ToDate")), true) || undefined;
      }
      calendar.exceptions[indexValue - 1] = { name, fromDate, toDate, dayWorking, workingTimes: [] };
    }
    for (const calendar of calendars) {
      calendar.exceptions = calendar.exceptions.filter(Boolean);
    }
  }

  globalThis.__mikuprojectProjectXlsxImportCalendars = {
    importCalendarsSheet,
    importCalendarsSheetAsCalendars,
    importNonWorkingDaysSheet,
    importNonWorkingDaysSheetAsCalendarExceptions
  };
})();
