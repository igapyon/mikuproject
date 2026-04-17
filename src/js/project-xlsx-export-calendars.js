/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const workbookSchema = globalThis.__mikuprojectProjectWorkbookSchema;
    if (!workbookSchema) {
        throw new Error("mikuproject Project Workbook Schema module is not loaded");
    }
    const { DATA_ROW_START_INDEX, SHEET_HEADERS } = workbookSchema;
    const projectXlsxExportUtil = globalThis.__mikuprojectProjectXlsxExportUtil;
    if (!projectXlsxExportUtil) {
        throw new Error("mikuproject Project XLSX export util module is not loaded");
    }
    function buildCalendarsSheet(model) {
        return {
            name: "Calendars",
            columns: [
                { width: 10 }, { width: 24 }, { width: 14 }, { width: 16 },
                { width: 10 }, { width: 12 }, { width: 10 }
            ],
            mergedRanges: [],
            dataValidations: projectXlsxExportUtil.buildBooleanDataValidations([
                projectXlsxExportUtil.buildColumnRange(2, DATA_ROW_START_INDEX + 1, DATA_ROW_START_INDEX + model.calendars.length)
            ]),
            rows: [
                projectXlsxExportUtil.sectionTitleRow("Calendars", 7, projectXlsxExportUtil.SHEET_THEMES.calendars.section),
                projectXlsxExportUtil.sectionTitleRow("Calendar List", 7, projectXlsxExportUtil.SHEET_THEMES.calendars.section),
                projectXlsxExportUtil.headerRow([...SHEET_HEADERS.Calendars], projectXlsxExportUtil.SHEET_THEMES.calendars.header),
                ...model.calendars.map((calendar, index) => ({
                    cells: [
                        projectXlsxExportUtil.countCell(calendar.uid, index),
                        projectXlsxExportUtil.editableCell(projectXlsxExportUtil.entityNameCell(calendar.name, index)),
                        projectXlsxExportUtil.editableCell(projectXlsxExportUtil.countCell(calendar.isBaseCalendar, index)),
                        projectXlsxExportUtil.editableCell(projectXlsxExportUtil.referenceCell(calendar.baseCalendarUID, index)),
                        projectXlsxExportUtil.countCell(calendar.weekDays.length, index),
                        projectXlsxExportUtil.countCell(calendar.exceptions.length, index),
                        projectXlsxExportUtil.countCell(calendar.workWeeks.length, index)
                    ]
                }))
            ]
        };
    }
    function buildNonWorkingDaysSheet(model) {
        const rows = model.calendars.flatMap((calendar) => calendar.exceptions.map((exception, index) => ({
            cells: [
                projectXlsxExportUtil.countCell(calendar.uid, index),
                projectXlsxExportUtil.countCell(index + 1, index),
                projectXlsxExportUtil.textCell(calendar.name, index),
                projectXlsxExportUtil.editableCell(projectXlsxExportUtil.entityNameCell(exception.name, index)),
                projectXlsxExportUtil.editableCell(projectXlsxExportUtil.dateOnlyCell(projectXlsxExportUtil.formatExceptionDate(exception), index)),
                projectXlsxExportUtil.editableCell(projectXlsxExportUtil.dateOnlyCell(projectXlsxExportUtil.formatExceptionBoundaryDate(exception.fromDate), index)),
                projectXlsxExportUtil.editableCell(projectXlsxExportUtil.dateOnlyCell(projectXlsxExportUtil.formatExceptionBoundaryDate(exception.toDate), index)),
                projectXlsxExportUtil.editableCell(projectXlsxExportUtil.countCell(exception.dayWorking, index))
            ]
        })));
        return {
            name: "NonWorkingDays",
            columns: [
                { width: 12 }, { width: 10 }, { width: 22 }, { width: 24 },
                { width: 14 }, { width: 22 }, { width: 22 }, { width: 12 }
            ],
            mergedRanges: [],
            dataValidations: projectXlsxExportUtil.buildBooleanDataValidations([
                projectXlsxExportUtil.buildColumnRange(7, DATA_ROW_START_INDEX + 1, DATA_ROW_START_INDEX + rows.length)
            ]),
            rows: [
                projectXlsxExportUtil.sectionTitleRow("NonWorkingDays", 8, projectXlsxExportUtil.SHEET_THEMES.nonWorkingDays.section),
                projectXlsxExportUtil.sectionTitleRow("Calendar Exceptions", 8, projectXlsxExportUtil.SHEET_THEMES.nonWorkingDays.section),
                projectXlsxExportUtil.headerRow([...SHEET_HEADERS.NonWorkingDays], projectXlsxExportUtil.SHEET_THEMES.nonWorkingDays.header),
                ...rows
            ]
        };
    }
    globalThis.__mikuprojectProjectXlsxExportCalendars = {
        buildCalendarsSheet,
        buildNonWorkingDaysSheet
    };
})();
