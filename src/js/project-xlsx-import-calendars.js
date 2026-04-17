/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const workbookSchema = globalThis.__mikuprojectProjectWorkbookSchema;
    if (!workbookSchema) {
        throw new Error("mikuproject Project Workbook Schema module is not loaded");
    }
    const { HEADER_ROW_INDEX, DATA_ROW_START_INDEX } = workbookSchema;
    const projectXlsxImportUtil = globalThis.__mikuprojectProjectXlsxImportUtil;
    if (!projectXlsxImportUtil) {
        throw new Error("mikuproject Project XLSX import util module is not loaded");
    }
    function importCalendarsSheet(workbook, model, changes) {
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
    function importCalendarsSheetAsCalendars(workbook, project) {
        var _a;
        const calendarsSheet = workbook.sheets.find((sheet) => sheet.name === "Calendars");
        const calendars = [];
        const columnIndexByLabel = calendarsSheet ? projectXlsxImportUtil.readHeaderMap(calendarsSheet, HEADER_ROW_INDEX) : new Map();
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
                    isBaseCalendar: (_a = projectXlsxImportUtil.readBooleanCellAt(row.cells, columnIndexByLabel.get("IsBaseCalendar"))) !== null && _a !== void 0 ? _a : false,
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
    function importNonWorkingDaysSheet(workbook, model, changes) {
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
            }
            else {
                projectXlsxImportUtil.assignIfChanged(changes, "calendars", calendar.uid, calendar.name, exception, "fromDate", `Exception${indexValue}.FromDate`, projectXlsxImportUtil.normalizeExceptionBoundaryInput(projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("FromDate")), false));
                projectXlsxImportUtil.assignIfChanged(changes, "calendars", calendar.uid, calendar.name, exception, "toDate", `Exception${indexValue}.ToDate`, projectXlsxImportUtil.normalizeExceptionBoundaryInput(projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("ToDate")), true));
            }
            projectXlsxImportUtil.assignIfChanged(changes, "calendars", calendar.uid, calendar.name, exception, "dayWorking", `Exception${indexValue}.DayWorking`, projectXlsxImportUtil.readBooleanCellAt(row.cells, columnIndexByLabel.get("DayWorking")));
        }
    }
    function importNonWorkingDaysSheetAsCalendarExceptions(workbook, calendars) {
        var _a;
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
            const dayWorking = (_a = projectXlsxImportUtil.readBooleanCellAt(row.cells, columnIndexByLabel.get("DayWorking"))) !== null && _a !== void 0 ? _a : false;
            let fromDate;
            let toDate;
            if (dateValue) {
                const normalizedDate = projectXlsxImportUtil.normalizeDateOnly(dateValue);
                if (normalizedDate) {
                    fromDate = `${normalizedDate}T00:00:00`;
                    toDate = `${normalizedDate}T23:59:59`;
                }
            }
            else {
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
