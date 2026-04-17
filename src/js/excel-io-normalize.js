/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const INVALID_SHEET_NAME_PATTERN = /[:\\/?*\[\]]/;
    const NUMBER_FORMATS = ["general", "integer", "decimal", "date", "datetime", "percent", "text"];
    const HORIZONTAL_ALIGNS = ["left", "center", "right"];
    const VERTICAL_ALIGNS = ["top", "center", "bottom"];
    const BORDER_STYLES = ["thin"];
    function normalizeWorkbook(workbook) {
        if (!workbook || !Array.isArray(workbook.sheets) || workbook.sheets.length === 0) {
            throw new Error("Workbook must contain at least one sheet");
        }
        const seenNames = new Set();
        return {
            sheets: workbook.sheets.map((sheet) => {
                if (!sheet || typeof sheet.name !== "string") {
                    throw new Error("Each sheet must have a valid sheet name");
                }
                validateSheetName(sheet.name);
                const canonicalName = sheet.name.toLocaleLowerCase();
                if (seenNames.has(canonicalName)) {
                    throw new Error(`Duplicate sheet name is not allowed: ${sheet.name}`);
                }
                seenNames.add(canonicalName);
                return {
                    name: sheet.name,
                    columns: Array.isArray(sheet.columns) ? sheet.columns.map((column) => normalizeColumn(column)) : undefined,
                    freezePane: normalizeFreezePane(sheet.freezePane),
                    mergedRanges: Array.isArray(sheet.mergedRanges) ? sheet.mergedRanges.map((range) => normalizeMergedRange(range)) : undefined,
                    dataValidations: Array.isArray(sheet.dataValidations)
                        ? sheet.dataValidations.map((dataValidation) => normalizeDataValidation(dataValidation))
                        : undefined,
                    rows: Array.isArray(sheet.rows)
                        ? sheet.rows.map((row) => ({
                            height: normalizeOptionalPositiveNumber(row === null || row === void 0 ? void 0 : row.height, "Row height"),
                            cells: Array.isArray(row === null || row === void 0 ? void 0 : row.cells)
                                ? row.cells.map((cell) => normalizeCell(cell))
                                : []
                        }))
                        : []
                };
            })
        };
    }
    function normalizeColumn(column) {
        if (!column) {
            return {};
        }
        if (column.hidden !== undefined && typeof column.hidden !== "boolean") {
            throw new Error("Column hidden must be boolean");
        }
        return {
            width: normalizeOptionalPositiveNumber(column.width, "Column width"),
            hidden: column.hidden === true ? true : undefined
        };
    }
    function normalizeDataValidation(dataValidation) {
        if (!dataValidation) {
            throw new Error("Data validation must be defined");
        }
        if (dataValidation.type !== "list") {
            throw new Error(`Unsupported data validation type: ${String(dataValidation.type)}`);
        }
        const sqref = normalizeSqref(dataValidation.sqref);
        if (!dataValidation.formula1 || typeof dataValidation.formula1 !== "string") {
            throw new Error("Data validation formula1 must be a non-empty string");
        }
        return {
            type: "list",
            sqref,
            formula1: dataValidation.formula1,
            allowBlank: dataValidation.allowBlank === true ? true : undefined
        };
    }
    function normalizeFreezePane(freezePane) {
        if (!freezePane) {
            return undefined;
        }
        const rowSplit = normalizeOptionalPositiveInteger(freezePane.rowSplit, "Freeze pane rowSplit");
        const colSplit = normalizeOptionalPositiveInteger(freezePane.colSplit, "Freeze pane colSplit");
        if (rowSplit === undefined && colSplit === undefined) {
            return undefined;
        }
        return {
            rowSplit,
            colSplit
        };
    }
    function normalizeCell(cell) {
        if (!cell) {
            return {};
        }
        if (cell.value !== undefined && typeof cell.value !== "string" && typeof cell.value !== "number" && typeof cell.value !== "boolean") {
            throw new Error("Cell value must be string, number, or boolean");
        }
        if (cell.formula !== undefined && typeof cell.formula !== "string") {
            throw new Error("Cell formula must be a string");
        }
        if (cell.numberFormat !== undefined && !NUMBER_FORMATS.includes(cell.numberFormat)) {
            throw new Error(`Unsupported cell number format: ${cell.numberFormat}`);
        }
        if (cell.horizontalAlign !== undefined && !HORIZONTAL_ALIGNS.includes(cell.horizontalAlign)) {
            throw new Error(`Unsupported cell horizontal align: ${cell.horizontalAlign}`);
        }
        if (cell.verticalAlign !== undefined && !VERTICAL_ALIGNS.includes(cell.verticalAlign)) {
            throw new Error(`Unsupported cell vertical align: ${cell.verticalAlign}`);
        }
        if (cell.border !== undefined && !BORDER_STYLES.includes(cell.border)) {
            throw new Error(`Unsupported cell border: ${cell.border}`);
        }
        if (cell.fontSize !== undefined) {
            normalizeOptionalPositiveNumber(cell.fontSize, "Cell fontSize");
        }
        if (cell.fillColor !== undefined) {
            assertColor(cell.fillColor);
        }
        return {
            value: cell.value,
            formula: cell.formula,
            numberFormat: cell.numberFormat,
            horizontalAlign: cell.horizontalAlign,
            verticalAlign: cell.verticalAlign,
            wrapText: cell.wrapText === true ? true : undefined,
            bold: cell.bold === true ? true : undefined,
            fontSize: cell.fontSize,
            fillColor: cell.fillColor ? normalizeColor(cell.fillColor) : undefined,
            border: cell.border
        };
    }
    function normalizeMergedRange(range) {
        if (typeof range !== "string") {
            throw new Error("Merged range must be a string");
        }
        const trimmed = range.trim().toUpperCase();
        if (!/^[A-Z]+\d+:[A-Z]+\d+$/.test(trimmed)) {
            throw new Error(`Invalid merged range: ${range}`);
        }
        return trimmed;
    }
    function normalizeSqref(sqref) {
        if (typeof sqref !== "string") {
            throw new Error("Data validation sqref must be a string");
        }
        const normalized = sqref
            .trim()
            .toUpperCase()
            .split(/\s+/)
            .filter(Boolean)
            .map((part) => normalizeSqrefPart(part))
            .join(" ");
        if (!normalized) {
            throw new Error("Data validation sqref must not be empty");
        }
        return normalized;
    }
    function normalizeSqrefPart(part) {
        if (/^[A-Z]+\d+$/.test(part)) {
            return part;
        }
        if (/^[A-Z]+\d+:[A-Z]+\d+$/.test(part)) {
            return part;
        }
        throw new Error(`Invalid data validation sqref: ${part}`);
    }
    function normalizeOptionalPositiveNumber(value, label) {
        if (value === undefined) {
            return undefined;
        }
        if (!Number.isFinite(value) || value <= 0) {
            throw new Error(`${label} must be a finite positive number`);
        }
        return value;
    }
    function normalizeOptionalPositiveInteger(value, label) {
        if (value === undefined) {
            return undefined;
        }
        if (!Number.isInteger(value) || value <= 0) {
            throw new Error(`${label} must be a positive integer`);
        }
        return value;
    }
    function validateSheetName(name) {
        if (!name || !name.trim()) {
            throw new Error("Sheet name must not be empty");
        }
        if (name.length > 31) {
            throw new Error(`Sheet name is too long: ${name}`);
        }
        if (INVALID_SHEET_NAME_PATTERN.test(name)) {
            throw new Error(`Sheet name contains invalid characters: ${name}`);
        }
        if (name.startsWith("'") || name.endsWith("'")) {
            throw new Error(`Sheet name must not start or end with apostrophe: ${name}`);
        }
    }
    function assertColor(color) {
        if (!/^#?[0-9a-fA-F]{6}$/.test(color)) {
            throw new Error(`Unsupported color format: ${color}`);
        }
    }
    function normalizeColor(color) {
        const hex = color.startsWith("#") ? color.slice(1) : color;
        return `FF${hex.toUpperCase()}`;
    }
    function denormalizeColor(color) {
        if (!color) {
            return undefined;
        }
        const normalized = color.toUpperCase();
        if (/^[0-9A-F]{8}$/.test(normalized)) {
            return `#${normalized.slice(2)}`;
        }
        if (/^[0-9A-F]{6}$/.test(normalized)) {
            return `#${normalized}`;
        }
        return undefined;
    }
    globalThis.__mikuprojectExcelIoNormalize = {
        normalizeWorkbook,
        normalizeMergedRange,
        validateSheetName,
        denormalizeColor
    };
})();
