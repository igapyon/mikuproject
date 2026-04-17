/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  type XlsxPrimitiveValue = string | number | boolean;
  type XlsxNumberFormat = "general" | "integer" | "decimal" | "date" | "datetime" | "percent" | "text";
  type XlsxHorizontalAlign = "left" | "center" | "right";
  type XlsxVerticalAlign = "top" | "center" | "bottom";
  type XlsxBorderStyle = "thin";

  type XlsxCellModel = {
    value?: XlsxPrimitiveValue;
    formula?: string;
    numberFormat?: XlsxNumberFormat;
    horizontalAlign?: XlsxHorizontalAlign;
    verticalAlign?: XlsxVerticalAlign;
    wrapText?: boolean;
    bold?: boolean;
    fontSize?: number;
    fillColor?: string;
    border?: XlsxBorderStyle;
  };

  type XlsxRowModel = {
    height?: number;
    cells: XlsxCellModel[];
  };

  type XlsxColumnModel = {
    width?: number;
    hidden?: boolean;
  };

  type XlsxFreezePaneModel = {
    rowSplit?: number;
    colSplit?: number;
  };

  type XlsxSheetModel = {
    name: string;
    columns?: XlsxColumnModel[];
    freezePane?: XlsxFreezePaneModel;
    mergedRanges?: string[];
    rows: XlsxRowModel[];
  };

  type XlsxWorkbookModel = {
    sheets: XlsxSheetModel[];
  };

  type StyleDescriptor = {
    numberFormat: XlsxNumberFormat;
    horizontalAlign?: XlsxHorizontalAlign;
    verticalAlign?: XlsxVerticalAlign;
    wrapText?: boolean;
    bold?: boolean;
    fontSize?: number;
    fillColor?: string;
    border?: XlsxBorderStyle;
  };

  function parseWorkbookEntries(
    entries: Record<string, Uint8Array>,
    helpers: {
      decodeRequiredEntry: (entries: Record<string, Uint8Array>, name: string) => string;
      decodeUtf8: (bytes: Uint8Array) => string;
      parseXmlDocument: (xmlText: string) => XMLDocument;
      findDirectChild: (element: Element, localName: string) => Element | null;
      parseOptionalNumber: (value: string | null) => number | undefined;
      validateSheetName: (name: string) => void;
      normalizeMergedRange: (range: string) => string;
      denormalizeColor: (color: string | undefined) => string | undefined;
      parseStylesXml: (
        xmlText: string | null,
        helpers: {
          parseXmlDocument: (xmlText: string) => XMLDocument;
          findDirectChild: (element: Element, localName: string) => Element | null;
          parseOptionalNumber: (value: string | null) => number | undefined;
        }
      ) => StyleDescriptor[];
      parseSharedStringsXml: (
        xmlText: string | null,
        helpers: {
          parseXmlDocument: (xmlText: string) => XMLDocument;
          findDirectChild: (element: Element, localName: string) => Element | null;
        }
      ) => string[];
      parseWorksheetXml: (
        name: string,
        xmlText: string,
        styleBook: StyleDescriptor[],
        sharedStrings: string[],
        helpers: {
          parseXmlDocument: (xmlText: string) => XMLDocument;
          parseOptionalNumber: (value: string | null) => number | undefined;
          findDirectChild: (element: Element, localName: string) => Element | null;
          normalizeMergedRange: (range: string) => string;
          denormalizeColor: (color: string | undefined) => string | undefined;
        }
      ) => XlsxSheetModel;
    }
  ): XlsxWorkbookModel {
    const workbookXml = helpers.decodeRequiredEntry(entries, "xl/workbook.xml");
    const workbookRelsXml = helpers.decodeRequiredEntry(entries, "xl/_rels/workbook.xml.rels");
    const stylesXml = entries["xl/styles.xml"] ? helpers.decodeUtf8(entries["xl/styles.xml"]) : null;
    const sharedStringsXml = entries["xl/sharedStrings.xml"] ? helpers.decodeUtf8(entries["xl/sharedStrings.xml"]) : null;
    const workbookDocument = helpers.parseXmlDocument(workbookXml);
    const relationshipsDocument = helpers.parseXmlDocument(workbookRelsXml);
    const styleBook = helpers.parseStylesXml(stylesXml, {
      parseXmlDocument: helpers.parseXmlDocument,
      findDirectChild: helpers.findDirectChild,
      parseOptionalNumber: helpers.parseOptionalNumber
    });
    const sharedStrings = helpers.parseSharedStringsXml(sharedStringsXml, {
      parseXmlDocument: helpers.parseXmlDocument,
      findDirectChild: helpers.findDirectChild
    });
    const relationshipMap = new Map<string, string>();
    const relationshipElements = Array.from(relationshipsDocument.getElementsByTagNameNS(
      "http://schemas.openxmlformats.org/package/2006/relationships",
      "Relationship"
    ));

    for (const relationshipElement of relationshipElements) {
      const id = relationshipElement.getAttribute("Id");
      const target = relationshipElement.getAttribute("Target");
      if (id && target) {
        relationshipMap.set(id, normalizeWorkbookTarget(target));
      }
    }

    const sheetElements = Array.from(workbookDocument.getElementsByTagNameNS(
      "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "sheet"
    ));

    return {
      sheets: sheetElements.map((sheetElement) => {
        const name = sheetElement.getAttribute("name") || "";
        helpers.validateSheetName(name);
        const relationshipId = sheetElement.getAttributeNS(
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
          "id"
        ) || sheetElement.getAttribute("r:id");
        if (!relationshipId) {
          throw new Error(`Worksheet relationship id is missing for sheet: ${name}`);
        }
        const target = relationshipMap.get(relationshipId);
        if (!target) {
          throw new Error(`Worksheet relationship target is missing for sheet: ${name}`);
        }
        const worksheetXml = helpers.decodeRequiredEntry(entries, target);
        return helpers.parseWorksheetXml(name, worksheetXml, styleBook, sharedStrings, {
          parseXmlDocument: helpers.parseXmlDocument,
          parseOptionalNumber: helpers.parseOptionalNumber,
          findDirectChild: helpers.findDirectChild,
          normalizeMergedRange: helpers.normalizeMergedRange,
          denormalizeColor: helpers.denormalizeColor
        });
      })
    };
  }

  function normalizeWorkbookTarget(target: string): string {
    return target.startsWith("xl/") ? target : `xl/${target.replace(/^\.\//, "")}`;
  }

  (globalThis as typeof globalThis & {
    __mikuprojectExcelIoWorkbookParse?: {
      parseWorkbookEntries: (
        entries: Record<string, Uint8Array>,
        helpers: {
          decodeRequiredEntry: (entries: Record<string, Uint8Array>, name: string) => string;
          decodeUtf8: (bytes: Uint8Array) => string;
          parseXmlDocument: (xmlText: string) => XMLDocument;
          findDirectChild: (element: Element, localName: string) => Element | null;
          parseOptionalNumber: (value: string | null) => number | undefined;
          validateSheetName: (name: string) => void;
          normalizeMergedRange: (range: string) => string;
          denormalizeColor: (color: string | undefined) => string | undefined;
          parseStylesXml: (
            xmlText: string | null,
            helpers: {
              parseXmlDocument: (xmlText: string) => XMLDocument;
              findDirectChild: (element: Element, localName: string) => Element | null;
              parseOptionalNumber: (value: string | null) => number | undefined;
            }
          ) => StyleDescriptor[];
          parseSharedStringsXml: (
            xmlText: string | null,
            helpers: {
              parseXmlDocument: (xmlText: string) => XMLDocument;
              findDirectChild: (element: Element, localName: string) => Element | null;
            }
          ) => string[];
          parseWorksheetXml: (
            name: string,
            xmlText: string,
            styleBook: StyleDescriptor[],
            sharedStrings: string[],
            helpers: {
              parseXmlDocument: (xmlText: string) => XMLDocument;
              parseOptionalNumber: (value: string | null) => number | undefined;
              findDirectChild: (element: Element, localName: string) => Element | null;
              normalizeMergedRange: (range: string) => string;
              denormalizeColor: (color: string | undefined) => string | undefined;
            }
          ) => XlsxSheetModel;
        }
      ) => XlsxWorkbookModel;
    };
  }).__mikuprojectExcelIoWorkbookParse = {
    parseWorkbookEntries
  };
})();
