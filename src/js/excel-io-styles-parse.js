/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const BORDER_STYLES = ["thin"];
    const DEFAULT_STYLE = { numberFormat: "general" };
    function parseStylesXml(xmlText, helpers) {
        if (!xmlText) {
            return [DEFAULT_STYLE];
        }
        const document = helpers.parseXmlDocument(xmlText);
        const fonts = parseFonts(document, helpers);
        const fills = parseFills(document, helpers);
        const borders = parseBorders(document, helpers);
        const xfElements = Array.from(document.getElementsByTagNameNS("http://schemas.openxmlformats.org/spreadsheetml/2006/main", "xf")).filter((element) => { var _a; return ((_a = element.parentElement) === null || _a === void 0 ? void 0 : _a.localName) === "cellXfs"; });
        if (xfElements.length === 0) {
            return [DEFAULT_STYLE];
        }
        return xfElements.map((xfElement) => {
            var _a, _b, _c, _d;
            const numFmtId = Number(xfElement.getAttribute("numFmtId") || "0");
            const fontId = Number(xfElement.getAttribute("fontId") || "0");
            const fillId = Number(xfElement.getAttribute("fillId") || "0");
            const borderId = Number(xfElement.getAttribute("borderId") || "0");
            const alignmentElement = helpers.findDirectChild(xfElement, "alignment");
            const horizontalAlign = alignmentElement === null || alignmentElement === void 0 ? void 0 : alignmentElement.getAttribute("horizontal");
            const verticalAlign = alignmentElement === null || alignmentElement === void 0 ? void 0 : alignmentElement.getAttribute("vertical");
            return {
                numberFormat: parseNumberFormatId(numFmtId),
                horizontalAlign: horizontalAlign || undefined,
                verticalAlign: verticalAlign || undefined,
                wrapText: (alignmentElement === null || alignmentElement === void 0 ? void 0 : alignmentElement.getAttribute("wrapText")) === "1" ? true : undefined,
                bold: ((_a = fonts[fontId]) === null || _a === void 0 ? void 0 : _a.bold) ? true : undefined,
                fontSize: (_b = fonts[fontId]) === null || _b === void 0 ? void 0 : _b.fontSize,
                fillColor: (_c = fills[fillId]) === null || _c === void 0 ? void 0 : _c.fillColor,
                border: (_d = borders[borderId]) === null || _d === void 0 ? void 0 : _d.border
            };
        });
    }
    function parseFonts(document, helpers) {
        return Array.from(document.getElementsByTagNameNS("http://schemas.openxmlformats.org/spreadsheetml/2006/main", "font")).map((fontElement) => {
            var _a;
            return ({
                bold: helpers.findDirectChild(fontElement, "b") ? true : undefined,
                fontSize: helpers.parseOptionalNumber(((_a = helpers.findDirectChild(fontElement, "sz")) === null || _a === void 0 ? void 0 : _a.getAttribute("val")) || null)
            });
        });
    }
    function parseFills(document, helpers) {
        return Array.from(document.getElementsByTagNameNS("http://schemas.openxmlformats.org/spreadsheetml/2006/main", "fill")).map((fillElement) => {
            const patternFill = helpers.findDirectChild(fillElement, "patternFill");
            const patternType = patternFill === null || patternFill === void 0 ? void 0 : patternFill.getAttribute("patternType");
            const fgColor = patternFill ? helpers.findDirectChild(patternFill, "fgColor") : null;
            return {
                patternType: patternType || undefined,
                fillColor: (fgColor === null || fgColor === void 0 ? void 0 : fgColor.getAttribute("rgb")) || undefined
            };
        });
    }
    function parseBorders(document, helpers) {
        return Array.from(document.getElementsByTagNameNS("http://schemas.openxmlformats.org/spreadsheetml/2006/main", "border")).map((borderElement) => {
            const left = helpers.findDirectChild(borderElement, "left");
            const style = left === null || left === void 0 ? void 0 : left.getAttribute("style");
            return {
                border: style && BORDER_STYLES.includes(style) ? style : undefined
            };
        });
    }
    function parseNumberFormatId(numFmtId) {
        switch (numFmtId) {
            case 1:
                return "integer";
            case 2:
                return "decimal";
            case 10:
                return "percent";
            case 14:
                return "date";
            case 22:
                return "datetime";
            default:
                return "general";
        }
    }
    globalThis.__mikuprojectExcelIoStylesParse = {
        parseStylesXml
    };
})();
