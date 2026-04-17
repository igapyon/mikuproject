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

  type XlsxDataValidationModel = {
    type: "list";
    sqref: string;
    formula1: string;
    allowBlank?: boolean;
  };

  type XlsxSheetModel = {
    name: string;
    columns?: XlsxColumnModel[];
    freezePane?: XlsxFreezePaneModel;
    mergedRanges?: string[];
    dataValidations?: XlsxDataValidationModel[];
    rows: XlsxRowModel[];
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

  function buildWorksheetXml(
    sheet: XlsxSheetModel,
    styleBook: StyleBook,
    helpers: {
      resolveStyleIndex: (cell: XlsxCellModel, styleBook: StyleBook) => number;
      resolveCellNumberFormat: (cell: XlsxCellModel) => XlsxNumberFormat;
      escapeXml: (value: string) => string;
    }
  ): string {
    const sheetViewsXml = buildSheetViewsXml(sheet.freezePane);
    const colsXml = buildColumnsXml(sheet.columns);
    const mergeCellsXml = buildMergeCellsXml(sheet.mergedRanges);
    const dataValidationsXml = buildDataValidationsXml(sheet.dataValidations, helpers.escapeXml);
    const rows = sheet.rows
      .map((row, rowIndex) => buildWorksheetRowXml(row, rowIndex, styleBook, helpers))
      .filter(Boolean)
      .join("");
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  ${sheetViewsXml}
  ${colsXml}
  <sheetData>${rows}</sheetData>
  ${mergeCellsXml}
  ${dataValidationsXml}
</worksheet>`;
  }

  function buildSheetViewsXml(freezePane: XlsxFreezePaneModel | undefined): string {
    if (!freezePane || (!freezePane.rowSplit && !freezePane.colSplit)) {
      return "";
    }
    const xSplit = freezePane.colSplit ? ` xSplit="${freezePane.colSplit}"` : "";
    const ySplit = freezePane.rowSplit ? ` ySplit="${freezePane.rowSplit}"` : "";
    const topLeftCell = encodeCellReference(freezePane.rowSplit || 0, freezePane.colSplit || 0);
    const topLeftCellAttribute = topLeftCell ? ` topLeftCell="${topLeftCell}"` : "";
    const activePane = resolveActivePane(freezePane);
    return `<sheetViews><sheetView workbookViewId="0"><pane${xSplit}${ySplit}${topLeftCellAttribute} activePane="${activePane}" state="frozen"/></sheetView></sheetViews>`;
  }

  function buildColumnsXml(columns: XlsxColumnModel[] | undefined): string {
    if (!columns || columns.length === 0 || columns.every((column) => column.width === undefined && column.hidden !== true)) {
      return "";
    }
    const cols = columns.map((column, index) => (
      (column.width !== undefined || column.hidden === true)
        ? `<col min="${index + 1}" max="${index + 1}"${column.width !== undefined ? ` width="${formatNumber(column.width)}" customWidth="1"` : ""}${column.hidden === true ? ` hidden="1"` : ""}/>`
        : ""
    )).filter(Boolean).join("");
    return cols ? `<cols>${cols}</cols>` : "";
  }

  function buildMergeCellsXml(mergedRanges: string[] | undefined): string {
    if (!mergedRanges || mergedRanges.length === 0) {
      return "";
    }
    const mergeCells = mergedRanges
      .map((range) => `<mergeCell ref="${range}"/>`)
      .join("");
    return `<mergeCells count="${mergedRanges.length}">${mergeCells}</mergeCells>`;
  }

  function buildDataValidationsXml(
    dataValidations: XlsxDataValidationModel[] | undefined,
    escapeXml: (value: string) => string
  ): string {
    if (!dataValidations || dataValidations.length === 0) {
      return "";
    }
    const nodes = dataValidations.map((dataValidation) => (
      `<dataValidation type="${dataValidation.type}" allowBlank="${dataValidation.allowBlank === true ? "1" : "0"}" showErrorMessage="1" sqref="${escapeXml(dataValidation.sqref)}"><formula1>${escapeXml(dataValidation.formula1)}</formula1></dataValidation>`
    )).join("");
    return `<dataValidations count="${dataValidations.length}">${nodes}</dataValidations>`;
  }

  function buildWorksheetRowXml(
    row: XlsxRowModel,
    rowIndex: number,
    styleBook: StyleBook,
    helpers: {
      resolveStyleIndex: (cell: XlsxCellModel, styleBook: StyleBook) => number;
      resolveCellNumberFormat: (cell: XlsxCellModel) => XlsxNumberFormat;
      escapeXml: (value: string) => string;
    }
  ): string {
    const cells = row.cells
      .map((cell, cellIndex) => buildWorksheetCellXml(cell, rowIndex, cellIndex, styleBook, helpers))
      .filter(Boolean)
      .join("");

    if (!cells) {
      return "";
    }

    const heightAttributes = row.height !== undefined
      ? ` ht="${formatNumber(row.height)}" customHeight="1"`
      : "";
    return `<row r="${rowIndex + 1}"${heightAttributes}>${cells}</row>`;
  }

  function buildWorksheetCellXml(
    cell: XlsxCellModel,
    rowIndex: number,
    cellIndex: number,
    styleBook: StyleBook,
    helpers: {
      resolveStyleIndex: (cell: XlsxCellModel, styleBook: StyleBook) => number;
      resolveCellNumberFormat: (cell: XlsxCellModel) => XlsxNumberFormat;
      escapeXml: (value: string) => string;
    }
  ): string {
    const reference = `${encodeColumnName(cellIndex)}${rowIndex + 1}`;
    const styleIndex = helpers.resolveStyleIndex(cell, styleBook);
    const styleAttribute = styleIndex > 0 ? ` s="${styleIndex}"` : "";
    const resolvedNumberFormat = helpers.resolveCellNumberFormat(cell);

    if (cell.formula !== undefined) {
      const formulaXml = `<f>${helpers.escapeXml(cell.formula)}</f>`;
      const valueXml = buildFormulaValueXml(cell.value, helpers.escapeXml);
      const typeAttribute = getCellTypeAttribute(cell.value, true);
      return `<c r="${reference}"${styleAttribute}${typeAttribute}>${formulaXml}${valueXml}</c>`;
    }

    if (cell.value === undefined) {
      return styleIndex > 0 ? `<c r="${reference}"${styleAttribute}/>` : "";
    }

    if (resolvedNumberFormat === "text") {
      return `<c r="${reference}"${styleAttribute} t="inlineStr"><is>${buildInlineStringTextXml(String(cell.value), helpers.escapeXml)}</is></c>`;
    }
    if (typeof cell.value === "string") {
      return `<c r="${reference}"${styleAttribute} t="inlineStr"><is>${buildInlineStringTextXml(cell.value, helpers.escapeXml)}</is></c>`;
    }
    if (typeof cell.value === "number") {
      return `<c r="${reference}"${styleAttribute}><v>${formatNumber(cell.value)}</v></c>`;
    }
    return `<c r="${reference}"${styleAttribute} t="b"><v>${cell.value ? "1" : "0"}</v></c>`;
  }

  function buildFormulaValueXml(
    value: XlsxPrimitiveValue | undefined,
    escapeXml: (value: string) => string
  ): string {
    if (value === undefined) {
      return "";
    }
    if (typeof value === "string") {
      return `<v>${escapeXml(value)}</v>`;
    }
    if (typeof value === "number") {
      return `<v>${formatNumber(value)}</v>`;
    }
    return `<v>${value ? "1" : "0"}</v>`;
  }

  function getCellTypeAttribute(value: XlsxPrimitiveValue | undefined, hasFormula: boolean): string {
    if (!hasFormula) {
      return "";
    }
    if (typeof value === "string") {
      return ` t="str"`;
    }
    if (typeof value === "boolean") {
      return ` t="b"`;
    }
    return "";
  }

  function formatNumber(value: number): string {
    if (!Number.isFinite(value)) {
      throw new Error(`Cell number must be finite: ${value}`);
    }
    return String(value);
  }

  function buildInlineStringTextXml(value: string, escapeXml: (value: string) => string): string {
    if (value.length === 0) {
      return `<t xml:space="preserve"></t>`;
    }
    const needsPreserve = /^\s|\s$/.test(value) || value.includes("\n");
    const attribute = needsPreserve ? ` xml:space="preserve"` : "";
    return `<t${attribute}>${escapeXml(sanitizeXmlText(value))}</t>`;
  }

  function sanitizeXmlText(value: string): string {
    return value.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");
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
    if (rowIndex < 0 || columnIndex < 0) {
      return "";
    }
    return `${encodeColumnName(columnIndex)}${rowIndex + 1}`;
  }

  function resolveActivePane(freezePane: XlsxFreezePaneModel): string {
    if (freezePane.rowSplit && freezePane.colSplit) {
      return "bottomRight";
    }
    if (freezePane.rowSplit) {
      return "bottomLeft";
    }
    return "topRight";
  }

  (globalThis as typeof globalThis & {
    __mikuprojectExcelIoWorksheetBuild?: {
      buildWorksheetXml: (
        sheet: XlsxSheetModel,
        styleBook: StyleBook,
        helpers: {
          resolveStyleIndex: (cell: XlsxCellModel, styleBook: StyleBook) => number;
          resolveCellNumberFormat: (cell: XlsxCellModel) => XlsxNumberFormat;
          escapeXml: (value: string) => string;
        }
      ) => string;
    };
  }).__mikuprojectExcelIoWorksheetBuild = {
    buildWorksheetXml
  };
})();
