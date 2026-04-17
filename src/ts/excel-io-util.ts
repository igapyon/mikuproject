/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  const textEncoder = new TextEncoder();
  const textDecoder = new TextDecoder();

  function resolveXmlDomEnvironment(): { DOMParser?: typeof DOMParser } {
    const configured = (globalThis as typeof globalThis & {
      __mikuprojectXmlDom?: {
        DOMParser?: typeof DOMParser;
      };
    }).__mikuprojectXmlDom;
    if (configured && typeof configured.DOMParser === "function") {
      return configured;
    }
    return {
      DOMParser: globalThis.DOMParser
    };
  }

  function buildInlineStringTextXml(value: string): string {
    const sanitizedValue = sanitizeXmlText(value);
    const preserveWhitespace = /^[\s]/.test(sanitizedValue)
      || /[\s]$/.test(sanitizedValue)
      || sanitizedValue.includes("\n")
      || sanitizedValue.includes("\r")
      || sanitizedValue.includes("\t");
    const preserveAttribute = preserveWhitespace ? ` xml:space="preserve"` : "";
    return `<t${preserveAttribute}>${escapeXml(sanitizedValue)}</t>`;
  }

  function parseOptionalNumber(value: string | null): number | undefined {
    if (!value) {
      return undefined;
    }
    return Number(value);
  }

  function findDirectChild(element: Element, localName: string): Element | null {
    for (const childNode of Array.from(element.childNodes)) {
      if (childNode.nodeType !== Node.ELEMENT_NODE) {
        continue;
      }
      const childElement = childNode as Element;
      if (childElement.localName === localName) {
        return childElement;
      }
    }
    return null;
  }

  function parseXmlDocument(xmlText: string): XMLDocument {
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

  function decodeRequiredEntry(entries: Record<string, Uint8Array>, name: string): string {
    const bytes = entries[name];
    if (!bytes) {
      throw new Error(`Required ZIP entry is missing: ${name}`);
    }
    return decodeUtf8(bytes);
  }

  function encodeUtf8(value: string): Uint8Array {
    return textEncoder.encode(value);
  }

  function decodeUtf8(bytes: Uint8Array): string {
    return textDecoder.decode(bytes);
  }

  function escapeXml(value: string): string {
    return sanitizeXmlText(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }

  function sanitizeXmlText(value: string): string {
    return value.replace(/[^\u0009\u000A\u000D\u0020-\uD7FF\uE000-\uFFFD]/g, "");
  }

  function encodeColumnName(columnIndex: number): string {
    let current = columnIndex + 1;
    let result = "";
    while (current > 0) {
      const remainder = (current - 1) % 26;
      result = String.fromCharCode(65 + remainder) + result;
      current = Math.floor((current - 1) / 26);
    }
    return result;
  }

  function encodeCellReference(rowIndex: number, columnIndex: number): string {
    if (rowIndex <= 0 && columnIndex <= 0) {
      return "";
    }
    return `${encodeColumnName(columnIndex)}${rowIndex + 1}`;
  }

  function resolveActivePane(freezePane: { rowSplit?: number; colSplit?: number }): string {
    if (freezePane.rowSplit && freezePane.colSplit) {
      return "bottomRight";
    }
    if (freezePane.rowSplit) {
      return "bottomLeft";
    }
    return "topRight";
  }

  (globalThis as typeof globalThis & {
    __mikuprojectExcelIoUtil?: {
      buildInlineStringTextXml: (value: string) => string;
      parseOptionalNumber: (value: string | null) => number | undefined;
      findDirectChild: (element: Element, localName: string) => Element | null;
      parseXmlDocument: (xmlText: string) => XMLDocument;
      decodeRequiredEntry: (entries: Record<string, Uint8Array>, name: string) => string;
      encodeUtf8: (value: string) => Uint8Array;
      decodeUtf8: (bytes: Uint8Array) => string;
      escapeXml: (value: string) => string;
      sanitizeXmlText: (value: string) => string;
      encodeColumnName: (columnIndex: number) => string;
      encodeCellReference: (rowIndex: number, columnIndex: number) => string;
      resolveActivePane: (freezePane: { rowSplit?: number; colSplit?: number }) => string;
    };
  }).__mikuprojectExcelIoUtil = {
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
