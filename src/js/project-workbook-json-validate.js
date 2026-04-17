/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const workbookSchema = globalThis.__mikuprojectProjectWorkbookSchema;
    if (!workbookSchema) {
        throw new Error("mikuproject Project Workbook Schema module is not loaded");
    }
    const { SHEET_NAMES, PROJECT_FIELD_ORDER, SHEET_HEADERS } = workbookSchema;
    function isKnownColumn(sheetName, key) {
        if (sheetName === "Project") {
            return key === "Field" || key === "Value";
        }
        if (sheetName === "Project" && key === "Field") {
            return PROJECT_FIELD_ORDER.includes(key);
        }
        return (SHEET_HEADERS[sheetName] || []).includes(key);
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
    globalThis.__mikuprojectProjectWorkbookJsonValidate = {
        validateWorkbookJsonDocument
    };
})();
