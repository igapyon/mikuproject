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

  type XlsxSheetModel = {
    name: string;
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

  type StyleBook = {
    styles: StyleDescriptor[];
    styleIndexByKey: Map<string, number>;
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

  const STYLE_KEY_DELIMITER = "::";
  const DEFAULT_STYLE: StyleDescriptor = { numberFormat: "general" };
  const DEFAULT_FILL_NONE: FillDescriptor = { patternType: "none", fillColor: undefined };
  const DEFAULT_FILL_GRAY125: FillDescriptor = { patternType: "gray125", fillColor: undefined };

  function createStyleBook(workbook: XlsxWorkbookModel): StyleBook {
    const styles: StyleDescriptor[] = [DEFAULT_STYLE];
    const styleIndexByKey = new Map<string, number>([[styleKey(DEFAULT_STYLE), 0]]);

    for (const sheet of workbook.sheets) {
      for (const row of sheet.rows) {
        for (const cell of row.cells) {
          const descriptor = getStyleDescriptor(cell);
          if (!descriptor) {
            continue;
          }
          const key = styleKey(descriptor);
          if (!styleIndexByKey.has(key)) {
            styleIndexByKey.set(key, styles.length);
            styles.push(descriptor);
          }
        }
      }
    }

    return { styles, styleIndexByKey };
  }

  function getStyleDescriptor(cell: XlsxCellModel): StyleDescriptor | null {
    const numberFormat = resolveCellNumberFormat(cell);
    if (numberFormat === "general" && !cell.horizontalAlign && !cell.verticalAlign && !cell.wrapText && !cell.bold && !cell.fontSize && !cell.fillColor && !cell.border) {
      return null;
    }
    return {
      numberFormat,
      horizontalAlign: cell.horizontalAlign,
      verticalAlign: cell.verticalAlign,
      wrapText: cell.wrapText === true ? true : undefined,
      bold: cell.bold === true ? true : undefined,
      fontSize: cell.fontSize,
      fillColor: cell.fillColor,
      border: cell.border
    };
  }

  function styleKey(style: StyleDescriptor): string {
    return [
      style.numberFormat,
      style.horizontalAlign || "",
      style.verticalAlign || "",
      style.wrapText ? "wrap" : "",
      style.bold ? "bold" : "",
      style.fontSize !== undefined ? String(style.fontSize) : "",
      style.fillColor || "",
      style.border || ""
    ].join(STYLE_KEY_DELIMITER);
  }

  function resolveStyleIndex(cell: XlsxCellModel, styleBook: StyleBook): number {
    const descriptor = getStyleDescriptor(cell);
    if (!descriptor) {
      return 0;
    }
    return styleBook.styleIndexByKey.get(styleKey(descriptor)) || 0;
  }

  function buildStylesXml(styles: StyleDescriptor[]): string {
    const fonts = dedupeDescriptors(
      styles.map((style) => ({ bold: style.bold, fontSize: style.fontSize } as FontDescriptor)),
      fontKey,
      { bold: undefined, fontSize: undefined }
    );
    const fills = dedupeFillDescriptors(styles.map((style) => ({ patternType: style.fillColor ? "solid" : "none", fillColor: style.fillColor } as FillDescriptor)));
    const borders = dedupeDescriptors(styles.map((style) => ({ border: style.border } as BorderDescriptor)), borderKey, { border: undefined });

    const styleNodes = styles.map((style) => {
      const numFmtId = mapNumberFormatId(style.numberFormat);
      const fontId = fonts.indexByKey.get(fontKey({ bold: style.bold, fontSize: style.fontSize })) || 0;
      const fillId = fills.indexByKey.get(fillKey({ patternType: style.fillColor ? "solid" : "none", fillColor: style.fillColor })) || 0;
      const borderId = borders.indexByKey.get(borderKey({ border: style.border })) || 0;
      const applyNumberFormat = numFmtId !== 0 ? ` applyNumberFormat="1"` : "";
      const applyAlignment = style.horizontalAlign || style.verticalAlign || style.wrapText ? ` applyAlignment="1"` : "";
      const applyFont = fontId !== 0 ? ` applyFont="1"` : "";
      const applyFill = fillId !== 0 ? ` applyFill="1"` : "";
      const applyBorder = borderId !== 0 ? ` applyBorder="1"` : "";
      const alignmentAttributes = [
        style.horizontalAlign ? ` horizontal="${style.horizontalAlign}"` : "",
        style.verticalAlign ? ` vertical="${style.verticalAlign}"` : "",
        style.wrapText ? ` wrapText="1"` : ""
      ].join("");
      const alignmentNode = alignmentAttributes ? `<alignment${alignmentAttributes}/>` : "";
      return `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"${applyNumberFormat}${applyAlignment}${applyFont}${applyFill}${applyBorder}>${alignmentNode}</xf>`;
    }).join("");

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="0"/>
  <fonts count="${fonts.items.length}">
    ${fonts.items.map(buildFontXml).join("")}
  </fonts>
  <fills count="${fills.items.length}">
    ${fills.items.map(buildFillXml).join("")}
  </fills>
  <borders count="${borders.items.length}">
    ${borders.items.map(buildBorderXml).join("")}
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="${styles.length}">
    ${styleNodes}
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>`;
  }

  function dedupeDescriptors<T>(items: T[], keyFn: (item: T) => string, defaultItem: T): { items: T[]; indexByKey: Map<string, number> } {
    const uniqueItems: T[] = [defaultItem];
    const indexByKey = new Map<string, number>([[keyFn(defaultItem), 0]]);
    for (const item of items) {
      const key = keyFn(item);
      if (!indexByKey.has(key)) {
        indexByKey.set(key, uniqueItems.length);
        uniqueItems.push(item);
      }
    }
    return { items: uniqueItems, indexByKey };
  }

  function fontKey(font: FontDescriptor): string {
    return [font.bold ? "bold" : "", font.fontSize !== undefined ? String(font.fontSize) : ""].join(STYLE_KEY_DELIMITER);
  }

  function fillKey(fill: FillDescriptor): string {
    return [fill.patternType || "none", fill.fillColor || ""].join(STYLE_KEY_DELIMITER);
  }

  function borderKey(border: BorderDescriptor): string {
    return border.border || "";
  }

  function formatNumber(value: number): string {
    if (!Number.isFinite(value)) {
      throw new Error(`Cell number must be finite: ${value}`);
    }
    return String(value);
  }

  function buildFontXml(font: FontDescriptor): string {
    const parts = [
      font.bold ? "<b/>" : "",
      font.fontSize !== undefined ? `<sz val="${escapeXml(formatNumber(font.fontSize))}"/>` : ""
    ].join("");
    return parts ? `<font>${parts}</font>` : `<font/>`;
  }

  function buildFillXml(fill: FillDescriptor): string {
    if (fill.patternType === "gray125") {
      return `<fill><patternFill patternType="gray125"/></fill>`;
    }
    if (!fill.fillColor || fill.patternType === "none") {
      return `<fill><patternFill patternType="none"/></fill>`;
    }
    return `<fill><patternFill patternType="solid"><fgColor rgb="${fill.fillColor}"/><bgColor indexed="64"/></patternFill></fill>`;
  }

  function dedupeFillDescriptors(items: FillDescriptor[]): { items: FillDescriptor[]; indexByKey: Map<string, number> } {
    const uniqueItems: FillDescriptor[] = [DEFAULT_FILL_NONE, DEFAULT_FILL_GRAY125];
    const indexByKey = new Map<string, number>([
      [fillKey(DEFAULT_FILL_NONE), 0],
      [fillKey(DEFAULT_FILL_GRAY125), 1]
    ]);
    for (const item of items) {
      const normalizedItem: FillDescriptor = {
        patternType: item.fillColor ? "solid" : (item.patternType || "none"),
        fillColor: item.fillColor
      };
      const key = fillKey(normalizedItem);
      if (!indexByKey.has(key)) {
        indexByKey.set(key, uniqueItems.length);
        uniqueItems.push(normalizedItem);
      }
    }
    return { items: uniqueItems, indexByKey };
  }

  function buildBorderXml(border: BorderDescriptor): string {
    if (!border.border) {
      return `<border/>`;
    }
    return `<border><left style="${border.border}"/><right style="${border.border}"/><top style="${border.border}"/><bottom style="${border.border}"/><diagonal/></border>`;
  }

  function mapNumberFormatId(numberFormat: XlsxNumberFormat): number {
    switch (numberFormat) {
      case "integer":
        return 1;
      case "decimal":
        return 2;
      case "text":
        return 49;
      case "date":
        return 14;
      case "datetime":
        return 22;
      case "percent":
        return 10;
      case "general":
      default:
        return 0;
    }
  }

  function resolveCellNumberFormat(cell: XlsxCellModel): XlsxNumberFormat {
    if (cell.numberFormat) {
      return cell.numberFormat;
    }
    if (cell.formula === undefined && cell.value !== undefined) {
      return "text";
    }
    return "general";
  }

  function escapeXml(value: string): string {
    return value
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }

  (globalThis as typeof globalThis & {
    __mikuprojectExcelIoStylesBuild?: {
      createStyleBook: (workbook: XlsxWorkbookModel) => StyleBook;
      resolveStyleIndex: (cell: XlsxCellModel, styleBook: StyleBook) => number;
      resolveCellNumberFormat: (cell: XlsxCellModel) => XlsxNumberFormat;
      buildStylesXml: (styles: StyleDescriptor[]) => string;
    };
  }).__mikuprojectExcelIoStylesBuild = {
    createStyleBook,
    resolveStyleIndex,
    resolveCellNumberFormat,
    buildStylesXml
  };
})();
