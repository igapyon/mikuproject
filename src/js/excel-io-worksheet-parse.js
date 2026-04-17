/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const DEFAULT_STYLE = { numberFormat: "general" };
    function parseWorksheetXml(name, xmlText, styleBook, sharedStrings, helpers) {
        const document = helpers.parseXmlDocument(xmlText);
        const rowElements = Array.from(document.getElementsByTagNameNS("http://schemas.openxmlformats.org/spreadsheetml/2006/main", "row"));
        return {
            name,
            columns: parseWorksheetColumns(document, helpers),
            freezePane: parseWorksheetFreezePane(document, helpers),
            mergedRanges: parseWorksheetMergedRanges(document, helpers),
            rows: rowElements.map((rowElement) => ({
                height: helpers.parseOptionalNumber(rowElement.getAttribute("ht")),
                cells: parseWorksheetRowCells(rowElement, styleBook, sharedStrings, helpers)
            }))
        };
    }
    function parseWorksheetColumns(document, helpers) {
        const colElements = Array.from(document.getElementsByTagNameNS("http://schemas.openxmlformats.org/spreadsheetml/2006/main", "col"));
        if (colElements.length === 0) {
            return undefined;
        }
        const columns = [];
        for (const colElement of colElements) {
            const min = Number(colElement.getAttribute("min") || "0");
            const max = Number(colElement.getAttribute("max") || "0");
            const width = helpers.parseOptionalNumber(colElement.getAttribute("width"));
            const hidden = colElement.getAttribute("hidden") === "1" ? true : undefined;
            for (let index = min; index <= max; index += 1) {
                columns[index - 1] = { width, hidden };
            }
        }
        while (columns.length > 0 && !columns[columns.length - 1]) {
            columns.pop();
        }
        return columns.length > 0 ? columns.map((column) => column || {}) : undefined;
    }
    function parseWorksheetMergedRanges(document, helpers) {
        const mergeCellElements = Array.from(document.getElementsByTagNameNS("http://schemas.openxmlformats.org/spreadsheetml/2006/main", "mergeCell"));
        if (mergeCellElements.length === 0) {
            return undefined;
        }
        return mergeCellElements
            .map((element) => helpers.normalizeMergedRange(element.getAttribute("ref") || ""))
            .filter(Boolean);
    }
    function parseWorksheetFreezePane(document, helpers) {
        const paneElement = document.getElementsByTagNameNS("http://schemas.openxmlformats.org/spreadsheetml/2006/main", "pane")[0];
        if (!paneElement || paneElement.getAttribute("state") !== "frozen") {
            return undefined;
        }
        const rowSplit = helpers.parseOptionalNumber(paneElement.getAttribute("ySplit"));
        const colSplit = helpers.parseOptionalNumber(paneElement.getAttribute("xSplit"));
        if (rowSplit === undefined && colSplit === undefined) {
            return undefined;
        }
        return {
            rowSplit,
            colSplit
        };
    }
    function parseWorksheetRowCells(rowElement, styleBook, sharedStrings, helpers) {
        const cells = [];
        const cellElements = Array.from(rowElement.getElementsByTagNameNS("http://schemas.openxmlformats.org/spreadsheetml/2006/main", "c")).filter((element) => element.parentElement === rowElement);
        for (const cellElement of cellElements) {
            const reference = cellElement.getAttribute("r") || "";
            const columnIndex = decodeColumnReference(reference);
            while (cells.length < columnIndex) {
                cells.push({});
            }
            cells.push(parseWorksheetCell(cellElement, styleBook, sharedStrings, helpers));
        }
        return cells;
    }
    function parseWorksheetCell(cellElement, styleBook, sharedStrings, helpers) {
        const type = cellElement.getAttribute("t") || "";
        const styleIndex = Number(cellElement.getAttribute("s") || "0");
        const formulaElement = helpers.findDirectChild(cellElement, "f");
        const valueElement = helpers.findDirectChild(cellElement, "v");
        const inlineStringElement = helpers.findDirectChild(cellElement, "is");
        let value;
        if (type === "inlineStr") {
            const textElement = inlineStringElement ? helpers.findDirectChild(inlineStringElement, "t") : null;
            value = textElement ? (textElement.textContent || "") : "";
        }
        else if (type === "s") {
            const sharedStringIndex = Number((valueElement === null || valueElement === void 0 ? void 0 : valueElement.textContent) || "0");
            value = Number.isFinite(sharedStringIndex) ? (sharedStrings[sharedStringIndex] || "") : "";
        }
        else if (type === "b") {
            value = (valueElement === null || valueElement === void 0 ? void 0 : valueElement.textContent) === "1";
        }
        else if (type === "str") {
            value = (valueElement === null || valueElement === void 0 ? void 0 : valueElement.textContent) || "";
        }
        else if (valueElement) {
            const rawValue = valueElement.textContent || "";
            value = rawValue === "" ? "" : Number(rawValue);
        }
        const style = styleBook[styleIndex] || DEFAULT_STYLE;
        const cell = {};
        if (style.numberFormat !== "general") {
            cell.numberFormat = style.numberFormat;
        }
        if (style.horizontalAlign) {
            cell.horizontalAlign = style.horizontalAlign;
        }
        if (style.verticalAlign) {
            cell.verticalAlign = style.verticalAlign;
        }
        if (style.wrapText) {
            cell.wrapText = true;
        }
        if (style.bold) {
            cell.bold = true;
        }
        if (style.fontSize !== undefined) {
            cell.fontSize = style.fontSize;
        }
        if (style.fillColor) {
            cell.fillColor = helpers.denormalizeColor(style.fillColor);
        }
        if (style.border) {
            cell.border = style.border;
        }
        if (formulaElement) {
            cell.formula = formulaElement.textContent || "";
        }
        if (value !== undefined) {
            cell.value = value;
        }
        return cell;
    }
    function parseSharedStringsXml(xmlText, helpers) {
        if (!xmlText) {
            return [];
        }
        const document = helpers.parseXmlDocument(xmlText);
        const stringItems = Array.from(document.getElementsByTagNameNS("http://schemas.openxmlformats.org/spreadsheetml/2006/main", "si"));
        return stringItems.map((item) => extractSharedStringText(item, helpers));
    }
    function extractSharedStringText(itemElement, helpers) {
        const directText = helpers.findDirectChild(itemElement, "t");
        if (directText) {
            return directText.textContent || "";
        }
        const richTextRuns = Array.from(itemElement.getElementsByTagNameNS("http://schemas.openxmlformats.org/spreadsheetml/2006/main", "r"));
        if (richTextRuns.length === 0) {
            return "";
        }
        return richTextRuns
            .map((run) => { var _a; return ((_a = helpers.findDirectChild(run, "t")) === null || _a === void 0 ? void 0 : _a.textContent) || ""; })
            .join("");
    }
    function decodeColumnReference(reference) {
        const match = /^([A-Z]+)\d+$/i.exec(reference);
        if (!match) {
            throw new Error(`Invalid cell reference: ${reference}`);
        }
        const letters = match[1].toUpperCase();
        let value = 0;
        for (const character of letters) {
            value = (value * 26) + (character.charCodeAt(0) - 64);
        }
        return value - 1;
    }
    globalThis.__mikuprojectExcelIoWorksheetParse = {
        parseWorksheetXml,
        parseSharedStringsXml
    };
})();
