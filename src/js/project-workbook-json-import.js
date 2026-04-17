/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const workbookSchema = globalThis.__mikuprojectProjectWorkbookSchema;
    if (!workbookSchema) {
        throw new Error("mikuproject Project Workbook Schema module is not loaded");
    }
    const { HEADER_ROW_INDEX, PROJECT_FIELD_ORDER, SHEET_HEADERS } = workbookSchema;
    const projectXlsx = globalThis.__mikuprojectProjectXlsx;
    if (!projectXlsx) {
        throw new Error("mikuproject Project XLSX module is not loaded");
    }
    const projectWorkbookJsonValidate = globalThis.__mikuprojectProjectWorkbookJsonValidate;
    if (!projectWorkbookJsonValidate) {
        throw new Error("mikuproject Project Workbook JSON validate module is not loaded");
    }
    function importProjectWorkbookJson(documentLike, baseModel) {
        const validation = projectWorkbookJsonValidate.validateWorkbookJsonDocument(documentLike);
        const document = validation.document;
        const workbook = {
            sheets: [
                buildProjectSheet(document.sheets.Project || []),
                buildTabularSheet("Tasks", document.sheets.Tasks || [], SHEET_HEADERS.Tasks),
                buildTabularSheet("Resources", document.sheets.Resources || [], SHEET_HEADERS.Resources),
                buildTabularSheet("Assignments", document.sheets.Assignments || [], SHEET_HEADERS.Assignments),
                buildTabularSheet("Calendars", document.sheets.Calendars || [], SHEET_HEADERS.Calendars),
                buildTabularSheet("NonWorkingDays", document.sheets.NonWorkingDays || [], SHEET_HEADERS.NonWorkingDays)
            ]
        };
        const result = projectXlsx.importProjectWorkbookDetailed(workbook, baseModel);
        return {
            ...result,
            warnings: validation.warnings
        };
    }
    function importProjectWorkbookJsonAsProjectModel(documentLike) {
        const validation = projectWorkbookJsonValidate.validateWorkbookJsonDocument(documentLike);
        const document = validation.document;
        const workbook = {
            sheets: [
                buildProjectSheet(document.sheets.Project || []),
                buildTabularSheet("Tasks", document.sheets.Tasks || [], SHEET_HEADERS.Tasks),
                buildTabularSheet("Resources", document.sheets.Resources || [], SHEET_HEADERS.Resources),
                buildTabularSheet("Assignments", document.sheets.Assignments || [], SHEET_HEADERS.Assignments),
                buildTabularSheet("Calendars", document.sheets.Calendars || [], SHEET_HEADERS.Calendars),
                buildTabularSheet("NonWorkingDays", document.sheets.NonWorkingDays || [], SHEET_HEADERS.NonWorkingDays)
            ]
        };
        return {
            model: projectXlsx.importProjectWorkbookAsProjectModel(workbook),
            warnings: validation.warnings
        };
    }
    function validateWorkbookJsonDocument(documentLike) {
        return projectWorkbookJsonValidate.validateWorkbookJsonDocument(documentLike);
    }
    function buildProjectSheet(rows) {
        const valueByField = new Map();
        for (const row of rows) {
            const field = typeof row.Field === "string" ? row.Field : "";
            if (!PROJECT_FIELD_ORDER.includes(field)) {
                continue;
            }
            valueByField.set(field, toWorkbookScalar(row.Value));
        }
        return {
            name: "Project",
            rows: [
                { cells: [{ value: "Project" }, {}] },
                { cells: [{ value: "Basic Info" }, {}] },
                { cells: [{ value: "Field" }, { value: "Value" }] },
                ...PROJECT_FIELD_ORDER.map((field) => ({
                    cells: [
                        { value: field },
                        { value: valueByField.get(field) }
                    ]
                }))
            ]
        };
    }
    function buildTabularSheet(sheetName, rows, headers) {
        return {
            name: sheetName,
            rows: [
                { cells: [{ value: sheetName }] },
                { cells: [{ value: `${sheetName} List` }] },
                { cells: headers.map((header) => ({ value: header })) },
                ...rows.map((row) => ({
                    cells: headers.map((header) => ({
                        value: toWorkbookScalar(row[header])
                    }))
                }))
            ]
        };
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
    function toWorkbookScalar(value) {
        if (value === null || value === undefined) {
            return undefined;
        }
        return value;
    }
    globalThis.__mikuprojectProjectWorkbookJsonImport = {
        importProjectWorkbookJson,
        importProjectWorkbookJsonAsProjectModel,
        validateWorkbookJsonDocument
    };
})();
