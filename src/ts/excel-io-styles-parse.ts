/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  type XlsxNumberFormat = "general" | "integer" | "decimal" | "date" | "datetime" | "percent" | "text";
  type XlsxHorizontalAlign = "left" | "center" | "right";
  type XlsxVerticalAlign = "top" | "center" | "bottom";
  type XlsxBorderStyle = "thin";

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

  type FontDescriptor = {
    bold?: boolean;
    fontSize?: number;
  };

  type FillDescriptor = {
    patternType?: "none" | "gray125" | "solid";
    fillColor?: string;
  };

  type BorderDescriptor = {
    border?: XlsxBorderStyle;
  };

  const BORDER_STYLES: XlsxBorderStyle[] = ["thin"];
  const DEFAULT_STYLE: StyleDescriptor = { numberFormat: "general" };

  function parseStylesXml(
    xmlText: string | null,
    helpers: {
      parseXmlDocument: (xmlText: string) => XMLDocument;
      findDirectChild: (element: Element, localName: string) => Element | null;
      parseOptionalNumber: (value: string | null) => number | undefined;
    }
  ): StyleDescriptor[] {
    if (!xmlText) {
      return [DEFAULT_STYLE];
    }
    const document = helpers.parseXmlDocument(xmlText);
    const fonts = parseFonts(document, helpers);
    const fills = parseFills(document, helpers);
    const borders = parseBorders(document, helpers);
    const xfElements = Array.from(document.getElementsByTagNameNS(
      "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "xf"
    )).filter((element) => element.parentElement?.localName === "cellXfs");

    if (xfElements.length === 0) {
      return [DEFAULT_STYLE];
    }

    return xfElements.map((xfElement) => {
      const numFmtId = Number(xfElement.getAttribute("numFmtId") || "0");
      const fontId = Number(xfElement.getAttribute("fontId") || "0");
      const fillId = Number(xfElement.getAttribute("fillId") || "0");
      const borderId = Number(xfElement.getAttribute("borderId") || "0");
      const alignmentElement = helpers.findDirectChild(xfElement, "alignment");
      const horizontalAlign = alignmentElement?.getAttribute("horizontal") as XlsxHorizontalAlign | null;
      const verticalAlign = alignmentElement?.getAttribute("vertical") as XlsxVerticalAlign | null;
      return {
        numberFormat: parseNumberFormatId(numFmtId),
        horizontalAlign: horizontalAlign || undefined,
        verticalAlign: verticalAlign || undefined,
        wrapText: alignmentElement?.getAttribute("wrapText") === "1" ? true : undefined,
        bold: fonts[fontId]?.bold ? true : undefined,
        fontSize: fonts[fontId]?.fontSize,
        fillColor: fills[fillId]?.fillColor,
        border: borders[borderId]?.border
      };
    });
  }

  function parseFonts(
    document: XMLDocument,
    helpers: {
      findDirectChild: (element: Element, localName: string) => Element | null;
      parseOptionalNumber: (value: string | null) => number | undefined;
    }
  ): FontDescriptor[] {
    return Array.from(document.getElementsByTagNameNS(
      "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "font"
    )).map((fontElement) => ({
      bold: helpers.findDirectChild(fontElement, "b") ? true : undefined,
      fontSize: helpers.parseOptionalNumber(helpers.findDirectChild(fontElement, "sz")?.getAttribute("val") || null)
    }));
  }

  function parseFills(
    document: XMLDocument,
    helpers: {
      findDirectChild: (element: Element, localName: string) => Element | null;
    }
  ): FillDescriptor[] {
    return Array.from(document.getElementsByTagNameNS(
      "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "fill"
    )).map((fillElement) => {
      const patternFill = helpers.findDirectChild(fillElement, "patternFill");
      const patternType = patternFill?.getAttribute("patternType") as FillDescriptor["patternType"] | null;
      const fgColor = patternFill ? helpers.findDirectChild(patternFill, "fgColor") : null;
      return {
        patternType: patternType || undefined,
        fillColor: fgColor?.getAttribute("rgb") || undefined
      };
    });
  }

  function parseBorders(
    document: XMLDocument,
    helpers: {
      findDirectChild: (element: Element, localName: string) => Element | null;
    }
  ): BorderDescriptor[] {
    return Array.from(document.getElementsByTagNameNS(
      "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "border"
    )).map((borderElement) => {
      const left = helpers.findDirectChild(borderElement, "left");
      const style = left?.getAttribute("style") as XlsxBorderStyle | null;
      return {
        border: style && BORDER_STYLES.includes(style) ? style : undefined
      };
    });
  }

  function parseNumberFormatId(numFmtId: number): XlsxNumberFormat {
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

  (globalThis as typeof globalThis & {
    __mikuprojectExcelIoStylesParse?: {
      parseStylesXml: (
        xmlText: string | null,
        helpers: {
          parseXmlDocument: (xmlText: string) => XMLDocument;
          findDirectChild: (element: Element, localName: string) => Element | null;
          parseOptionalNumber: (value: string | null) => number | undefined;
        }
      ) => StyleDescriptor[];
    };
  }).__mikuprojectExcelIoStylesParse = {
    parseStylesXml
  };
})();
