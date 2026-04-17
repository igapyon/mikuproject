/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const textEncoder = new TextEncoder();
    const textDecoder = new TextDecoder();
    function resolveXmlDomEnvironment() {
        const configured = globalThis.__mikuprojectXmlDom;
        if (configured && typeof configured.DOMParser === "function") {
            return configured;
        }
        return {
            DOMParser: globalThis.DOMParser
        };
    }
    function buildInlineStringTextXml(value) {
        const sanitizedValue = sanitizeXmlText(value);
        const preserveWhitespace = /^[\s]/.test(sanitizedValue)
            || /[\s]$/.test(sanitizedValue)
            || sanitizedValue.includes("\n")
            || sanitizedValue.includes("\r")
            || sanitizedValue.includes("\t");
        const preserveAttribute = preserveWhitespace ? ` xml:space="preserve"` : "";
        return `<t${preserveAttribute}>${escapeXml(sanitizedValue)}</t>`;
    }
    function parseOptionalNumber(value) {
        if (!value) {
            return undefined;
        }
        return Number(value);
    }
    function findDirectChild(element, localName) {
        for (const childNode of Array.from(element.childNodes)) {
            if (childNode.nodeType !== Node.ELEMENT_NODE) {
                continue;
            }
            const childElement = childNode;
            if (childElement.localName === localName) {
                return childElement;
            }
        }
        return null;
    }
    function parseXmlDocument(xmlText) {
        const environment = resolveXmlDomEnvironment();
        if (typeof environment.DOMParser !== "function") {
            throw new Error("XML DOMParser is not available");
        }
        const document = new environment.DOMParser().parseFromString(xmlText, "application/xml");
        if (document.getElementsByTagName("parsererror")[0]) {
            throw new Error("Failed to parse XML document");
        }
        return document;
    }
    function decodeRequiredEntry(entries, name) {
        const bytes = entries[name];
        if (!bytes) {
            throw new Error(`Required ZIP entry is missing: ${name}`);
        }
        return decodeUtf8(bytes);
    }
    function encodeUtf8(value) {
        return textEncoder.encode(value);
    }
    function decodeUtf8(bytes) {
        return textDecoder.decode(bytes);
    }
    function escapeXml(value) {
        return sanitizeXmlText(value)
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&apos;");
    }
    function sanitizeXmlText(value) {
        return value.replace(/[^\u0009\u000A\u000D\u0020-\uD7FF\uE000-\uFFFD]/g, "");
    }
    function encodeColumnName(columnIndex) {
        let current = columnIndex + 1;
        let result = "";
        while (current > 0) {
            const remainder = (current - 1) % 26;
            result = String.fromCharCode(65 + remainder) + result;
            current = Math.floor((current - 1) / 26);
        }
        return result;
    }
    function encodeCellReference(rowIndex, columnIndex) {
        if (rowIndex <= 0 && columnIndex <= 0) {
            return "";
        }
        return `${encodeColumnName(columnIndex)}${rowIndex + 1}`;
    }
    function resolveActivePane(freezePane) {
        if (freezePane.rowSplit && freezePane.colSplit) {
            return "bottomRight";
        }
        if (freezePane.rowSplit) {
            return "bottomLeft";
        }
        return "topRight";
    }
    globalThis.__mikuprojectExcelIoUtil = {
        buildInlineStringTextXml,
        parseOptionalNumber,
        findDirectChild,
        parseXmlDocument,
        decodeRequiredEntry,
        encodeUtf8,
        decodeUtf8,
        escapeXml,
        sanitizeXmlText,
        encodeColumnName,
        encodeCellReference,
        resolveActivePane
    };
})();
