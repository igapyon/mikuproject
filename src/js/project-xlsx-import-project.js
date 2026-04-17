/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const workbookSchema = globalThis.__mikuprojectProjectWorkbookSchema;
    if (!workbookSchema) {
        throw new Error("mikuproject Project Workbook Schema module is not loaded");
    }
    const { DATA_ROW_START_INDEX } = workbookSchema;
    const projectXlsxImportUtil = globalThis.__mikuprojectProjectXlsxImportUtil;
    if (!projectXlsxImportUtil) {
        throw new Error("mikuproject Project XLSX import util module is not loaded");
    }
    function readProjectValueByField(workbook) {
        const projectSheet = workbook.sheets.find((sheet) => sheet.name === "Project");
        const valueByField = new Map();
        for (const row of (projectSheet === null || projectSheet === void 0 ? void 0 : projectSheet.rows.slice(DATA_ROW_START_INDEX)) || []) {
            const field = projectXlsxImportUtil.readStringCell(row.cells[0]);
            if (!field) {
                continue;
            }
            valueByField.set(field, row.cells[1]);
        }
        return valueByField;
    }
    function importProjectSheet(workbook, model, changes) {
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
    function importProjectSheetAsProjectInfo(workbook) {
        var _a;
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
            scheduleFromStart: (_a = projectXlsxImportUtil.readBooleanCell(valueByField.get("ScheduleFromStart"))) !== null && _a !== void 0 ? _a : true,
            outlineCodes: [],
            wbsMasks: [],
            extendedAttributes: []
        };
    }
    globalThis.__mikuprojectProjectXlsxImportProject = {
        importProjectSheet,
        importProjectSheetAsProjectInfo
    };
})();
