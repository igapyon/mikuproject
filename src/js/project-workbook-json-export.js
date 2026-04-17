/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const workbookSchema = globalThis.__mikuprojectProjectWorkbookSchema;
    if (!workbookSchema) {
        throw new Error("mikuproject Project Workbook Schema module is not loaded");
    }
    const { HEADER_ROW_INDEX, DATA_ROW_START_INDEX, PROJECT_FIELD_ORDER, SHEET_HEADERS } = workbookSchema;
    const projectXlsx = globalThis.__mikuprojectProjectXlsx;
    if (!projectXlsx) {
        throw new Error("mikuproject Project XLSX module is not loaded");
    }
    function exportProjectSheetRows(workbook) {
        var _a, _b;
        const sheet = workbook.sheets.find((item) => item.name === "Project");
        if (!sheet) {
            return [];
        }
        const rows = [];
        for (const row of sheet.rows.slice(DATA_ROW_START_INDEX)) {
            const field = toJsonScalar((_a = row.cells[0]) === null || _a === void 0 ? void 0 : _a.value);
            if (typeof field !== "string" || !PROJECT_FIELD_ORDER.includes(field)) {
                continue;
            }
            rows.push({
                Field: field,
                Value: toJsonScalar((_b = row.cells[1]) === null || _b === void 0 ? void 0 : _b.value)
            });
        }
        return rows;
    }
    function exportTabularSheetRows(workbook, sheetName) {
        const sheet = workbook.sheets.find((item) => item.name === sheetName);
        if (!sheet) {
            return [];
        }
        const headers = readHeaderRow(sheet);
        return sheet.rows.slice(DATA_ROW_START_INDEX).map((row) => {
            const item = {};
            headers.forEach((header, index) => {
                var _a;
                item[header] = toJsonScalar((_a = row.cells[index]) === null || _a === void 0 ? void 0 : _a.value);
            });
            return item;
        });
    }
    function readHeaderRow(sheet) {
        var _a;
        return (((_a = sheet.rows[HEADER_ROW_INDEX]) === null || _a === void 0 ? void 0 : _a.cells) || [])
            .map((cell) => (typeof cell.value === "string" ? cell.value : ""))
            .filter((value) => value !== "");
    }
    function toJsonScalar(value) {
        if (value === undefined) {
            return null;
        }
        return value;
    }
    globalThis.__mikuprojectProjectWorkbookJsonExport = {
        exportProjectWorkbookJson(model) {
            const workbook = projectXlsx.exportProjectWorkbook(model);
            return {
                format: "mikuproject_workbook_json",
                version: 1,
                sheets: {
                    Project: exportProjectSheetRows(workbook),
                    Tasks: exportTabularSheetRows(workbook, "Tasks"),
                    Resources: exportTabularSheetRows(workbook, "Resources"),
                    Assignments: exportTabularSheetRows(workbook, "Assignments"),
                    Calendars: exportTabularSheetRows(workbook, "Calendars"),
                    NonWorkingDays: exportTabularSheetRows(workbook, "NonWorkingDays")
                }
            };
        }
    };
})();
