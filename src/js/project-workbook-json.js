/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const workbookSchema = globalThis.__mikuprojectProjectWorkbookSchema;
    if (!workbookSchema) {
        throw new Error("mikuproject Project Workbook Schema module is not loaded");
    }
    const { SHEET_NAMES, HEADER_ROW_INDEX, DATA_ROW_START_INDEX, PROJECT_FIELD_ORDER, SHEET_HEADERS } = workbookSchema;
    const projectXlsx = globalThis.__mikuprojectProjectXlsx;
    if (!projectXlsx) {
        throw new Error("mikuproject Project XLSX module is not loaded");
    }
    function exportProjectWorkbookJson(model) {
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
    function importProjectWorkbookJson(documentLike, baseModel) {
        const validation = validateWorkbookJsonDocument(documentLike);
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
        const validation = validateWorkbookJsonDocument(documentLike);
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
        if (!documentLike || typeof documentLike !== "object") {
            throw new Error("workbook JSON がオブジェクトではありません");
        }
        const document = documentLike;
        if (document.format !== "mikuproject_workbook_json") {
            throw new Error("format が mikuproject_workbook_json ではありません");
        }
        if (document.version !== 1) {
            throw new Error("version は 1 である必要があります");
        }
        if (!document.sheets || typeof document.sheets !== "object") {
            throw new Error("sheets がありません");
        }
        const warnings = [];
        for (const [sheetName, rows] of Object.entries(document.sheets)) {
            if (!SHEET_NAMES.includes(sheetName)) {
                warnings.push({ message: `未知の sheet は無視します: ${sheetName}` });
                continue;
            }
            if (!Array.isArray(rows)) {
                throw new Error(`sheets.${sheetName} は配列である必要があります`);
            }
            for (const [rowIndex, row] of rows.entries()) {
                if (!row || typeof row !== "object" || Array.isArray(row)) {
                    throw new Error(`sheets.${sheetName} にオブジェクトではない行があります`);
                }
                for (const key of Object.keys(row)) {
                    if (!isKnownColumn(sheetName, key)) {
                        warnings.push({ message: `未知の列は無視します: ${sheetName}[${rowIndex}].${key}` });
                    }
                }
            }
        }
        return {
            document: document,
            warnings
        };
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
    function isKnownColumn(sheetName, key) {
        if (sheetName === "Project") {
            return key === "Field" || key === "Value";
        }
        return (SHEET_HEADERS[sheetName] || []).includes(key);
    }
    globalThis.__mikuprojectProjectWorkbookJson = {
        exportProjectWorkbookJson,
        importProjectWorkbookJsonAsProjectModel,
        importProjectWorkbookJson,
        validateWorkbookJsonDocument
    };
})();
