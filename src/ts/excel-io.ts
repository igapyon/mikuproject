/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  type XlsxPrimitiveValue = string | number | boolean;
  type XlsxNumberFormat = "general" | "integer" | "decimal" | "date" | "datetime" | "percent";
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

  type ZipEntry = {
    name: string;
    data: Uint8Array;
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
    fillColor?: string;
  };

  type BorderDescriptor = {
    border?: XlsxBorderStyle;
  };

  const TEXT_ENCODER = new TextEncoder();
  const TEXT_DECODER = new TextDecoder();
  const INVALID_SHEET_NAME_PATTERN = /[:\\/?*\[\]]/;
  const CRC32_TABLE = buildCrc32Table();
  const NUMBER_FORMATS: XlsxNumberFormat[] = ["general", "integer", "decimal", "date", "datetime", "percent"];
  const HORIZONTAL_ALIGNS: XlsxHorizontalAlign[] = ["left", "center", "right"];
  const VERTICAL_ALIGNS: XlsxVerticalAlign[] = ["top", "center", "bottom"];
  const BORDER_STYLES: XlsxBorderStyle[] = ["thin"];
  const STYLE_KEY_DELIMITER = "::";
  const DEFAULT_STYLE: StyleDescriptor = { numberFormat: "general" };

  class XlsxWorkbookCodec {
    exportWorkbook(workbook: XlsxWorkbookModel): Uint8Array {
      const normalizedWorkbook = normalizeWorkbook(workbook);
      const entries = this.createWorkbookEntries(normalizedWorkbook);
      return packZip(entries);
    }

    importWorkbook(bytes: Uint8Array): XlsxWorkbookModel {
      const entries = this.unpackEntries(bytes);
      return parseWorkbookEntries(entries);
    }

    listEntries(bytes: Uint8Array): string[] {
      return Object.keys(this.unpackEntries(bytes)).sort();
    }

    unpackEntries(bytes: Uint8Array): Record<string, Uint8Array> {
      return unpackZip(bytes);
    }

    private createWorkbookEntries(workbook: XlsxWorkbookModel): ZipEntry[] {
      const styleBook = createStyleBook(workbook);
      const worksheetRelationships = workbook.sheets.map((sheet, index) => ({
        relationshipId: `rId${index + 1}`,
        target: `worksheets/sheet${index + 1}.xml`,
        name: sheet.name
      }));
      const worksheetEntries = workbook.sheets.map((sheet, index) => ({
        name: `xl/worksheets/sheet${index + 1}.xml`,
        data: encodeUtf8(buildWorksheetXml(sheet, styleBook))
      }));

      const entries: ZipEntry[] = [
        {
          name: "[Content_Types].xml",
          data: encodeUtf8(buildContentTypesXml(workbook.sheets.length, styleBook.styles.length > 1))
        },
        {
          name: "_rels/.rels",
          data: encodeUtf8(buildRootRelationshipsXml())
        },
        {
          name: "xl/_rels/workbook.xml.rels",
          data: encodeUtf8(buildWorkbookRelationshipsXml(worksheetRelationships, styleBook.styles.length > 1))
        },
        {
          name: "xl/workbook.xml",
          data: encodeUtf8(buildWorkbookXml(worksheetRelationships))
        }
      ];

      if (styleBook.styles.length > 1) {
        entries.push({
          name: "xl/styles.xml",
          data: encodeUtf8(buildStylesXml(styleBook.styles))
        });
      }

      entries.push(...worksheetEntries);
      return entries;
    }
  }

  function normalizeWorkbook(workbook: XlsxWorkbookModel): XlsxWorkbookModel {
    if (!workbook || !Array.isArray(workbook.sheets) || workbook.sheets.length === 0) {
      throw new Error("Workbook must contain at least one sheet");
    }

    const seenNames = new Set<string>();
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
          rows: Array.isArray(sheet.rows)
            ? sheet.rows.map((row) => ({
              height: normalizeOptionalPositiveNumber(row?.height, "Row height"),
              cells: Array.isArray(row?.cells)
                ? row.cells.map((cell) => normalizeCell(cell))
                : []
            }))
            : []
        };
      })
    };
  }

  function normalizeColumn(column: XlsxColumnModel | undefined): XlsxColumnModel {
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

  function normalizeFreezePane(freezePane: XlsxFreezePaneModel | undefined): XlsxFreezePaneModel | undefined {
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

  function normalizeCell(cell: XlsxCellModel | undefined): XlsxCellModel {
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

  function normalizeMergedRange(range: string): string {
    if (typeof range !== "string") {
      throw new Error("Merged range must be a string");
    }
    const trimmed = range.trim().toUpperCase();
    if (!/^[A-Z]+\d+:[A-Z]+\d+$/.test(trimmed)) {
      throw new Error(`Invalid merged range: ${range}`);
    }
    return trimmed;
  }

  function normalizeOptionalPositiveNumber(value: number | undefined, label: string): number | undefined {
    if (value === undefined) {
      return undefined;
    }
    if (!Number.isFinite(value) || value <= 0) {
      throw new Error(`${label} must be a finite positive number`);
    }
    return value;
  }

  function normalizeOptionalPositiveInteger(value: number | undefined, label: string): number | undefined {
    if (value === undefined) {
      return undefined;
    }
    if (!Number.isInteger(value) || value <= 0) {
      throw new Error(`${label} must be a positive integer`);
    }
    return value;
  }

  function validateSheetName(name: string): void {
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

  function assertColor(color: string): void {
    if (!/^#?[0-9a-fA-F]{6}$/.test(color)) {
      throw new Error(`Unsupported color format: ${color}`);
    }
  }

  function normalizeColor(color: string): string {
    const hex = color.startsWith("#") ? color.slice(1) : color;
    return `FF${hex.toUpperCase()}`;
  }

  function denormalizeColor(color: string | undefined): string | undefined {
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

  function buildContentTypesXml(sheetCount: number, includeStyles: boolean): string {
    const worksheetOverrides = Array.from({ length: sheetCount }, (_unused, index) => (
      `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
    )).join("");
    const stylesOverride = includeStyles
      ? `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>`
      : "";

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  ${worksheetOverrides}
  ${stylesOverride}
</Types>`;
  }

  function buildRootRelationshipsXml(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
  }

  function buildWorkbookRelationshipsXml(
    relationships: Array<{ relationshipId: string; target: string }>,
    includeStyles: boolean
  ): string {
    const worksheetNodes = relationships.map((relationship) => (
      `<Relationship Id="${escapeXml(relationship.relationshipId)}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="${escapeXml(relationship.target)}"/>`
    )).join("");
    const stylesNode = includeStyles
      ? `<Relationship Id="rId${relationships.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`
      : "";

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${worksheetNodes}
  ${stylesNode}
</Relationships>`;
  }

  function buildWorkbookXml(
    relationships: Array<{ relationshipId: string; name: string }>
  ): string {
    const sheets = relationships.map((relationship, index) => (
      `<sheet name="${escapeXml(relationship.name)}" sheetId="${index + 1}" r:id="${escapeXml(relationship.relationshipId)}"/>`
    )).join("");

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>${sheets}</sheets>
</workbook>`;
  }

  function buildWorksheetXml(sheet: XlsxSheetModel, styleBook: StyleBook): string {
    const sheetViewsXml = buildSheetViewsXml(sheet.freezePane);
    const colsXml = buildColumnsXml(sheet.columns);
    const mergeCellsXml = buildMergeCellsXml(sheet.mergedRanges);
    const rows = sheet.rows.map((row, rowIndex) => buildWorksheetRowXml(row, rowIndex, styleBook)).filter(Boolean).join("");
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  ${sheetViewsXml}
  ${colsXml}
  <sheetData>${rows}</sheetData>
  ${mergeCellsXml}
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

  function buildWorksheetRowXml(row: XlsxRowModel, rowIndex: number, styleBook: StyleBook): string {
    const cells = row.cells
      .map((cell, cellIndex) => buildWorksheetCellXml(cell, rowIndex, cellIndex, styleBook))
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

  function buildWorksheetCellXml(cell: XlsxCellModel, rowIndex: number, cellIndex: number, styleBook: StyleBook): string {
    const reference = `${encodeColumnName(cellIndex)}${rowIndex + 1}`;
    const styleIndex = resolveStyleIndex(cell, styleBook);
    const styleAttribute = styleIndex > 0 ? ` s="${styleIndex}"` : "";

    if (cell.formula !== undefined) {
      const formulaXml = `<f>${escapeXml(cell.formula)}</f>`;
      const valueXml = buildFormulaValueXml(cell.value);
      const typeAttribute = getCellTypeAttribute(cell.value, true);
      return `<c r="${reference}"${styleAttribute}${typeAttribute}>${formulaXml}${valueXml}</c>`;
    }

    if (cell.value === undefined) {
      return styleIndex > 0 ? `<c r="${reference}"${styleAttribute}/>` : "";
    }

    if (typeof cell.value === "string") {
      return `<c r="${reference}"${styleAttribute} t="inlineStr"><is><t>${escapeXml(cell.value)}</t></is></c>`;
    }
    if (typeof cell.value === "number") {
      return `<c r="${reference}"${styleAttribute}><v>${formatNumber(cell.value)}</v></c>`;
    }
    return `<c r="${reference}"${styleAttribute} t="b"><v>${cell.value ? "1" : "0"}</v></c>`;
  }

  function buildFormulaValueXml(value: XlsxPrimitiveValue | undefined): string {
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
    if (!cell.numberFormat && !cell.horizontalAlign && !cell.verticalAlign && !cell.wrapText && !cell.bold && !cell.fontSize && !cell.fillColor && !cell.border) {
      return null;
    }
    return {
      numberFormat: cell.numberFormat || "general",
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
    const fills = dedupeDescriptors(styles.map((style) => ({ fillColor: style.fillColor } as FillDescriptor)), fillKey, { fillColor: undefined });
    const borders = dedupeDescriptors(styles.map((style) => ({ border: style.border } as BorderDescriptor)), borderKey, { border: undefined });

    const styleNodes = styles.map((style) => {
      const numFmtId = mapNumberFormatId(style.numberFormat);
      const fontId = fonts.indexByKey.get(fontKey({ bold: style.bold, fontSize: style.fontSize })) || 0;
      const fillId = fills.indexByKey.get(fillKey({ fillColor: style.fillColor })) || 0;
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
      const alignmentNode = alignmentAttributes
        ? `<alignment${alignmentAttributes}/>`
        : "";
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
    return [
      font.bold ? "bold" : "",
      font.fontSize !== undefined ? String(font.fontSize) : ""
    ].join(STYLE_KEY_DELIMITER);
  }

  function fillKey(fill: FillDescriptor): string {
    return fill.fillColor || "";
  }

  function borderKey(border: BorderDescriptor): string {
    return border.border || "";
  }

  function buildFontXml(font: FontDescriptor): string {
    const parts = [
      font.bold ? "<b/>" : "",
      font.fontSize !== undefined ? `<sz val="${escapeXml(formatNumber(font.fontSize))}"/>` : ""
    ].join("");
    return parts ? `<font>${parts}</font>` : `<font/>`;
  }

  function buildFillXml(fill: FillDescriptor): string {
    if (!fill.fillColor) {
      return `<fill><patternFill patternType="none"/></fill>`;
    }
    return `<fill><patternFill patternType="solid"><fgColor rgb="${fill.fillColor}"/><bgColor indexed="64"/></patternFill></fill>`;
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

  function parseWorkbookEntries(entries: Record<string, Uint8Array>): XlsxWorkbookModel {
    const workbookXml = decodeRequiredEntry(entries, "xl/workbook.xml");
    const workbookRelsXml = decodeRequiredEntry(entries, "xl/_rels/workbook.xml.rels");
    const stylesXml = entries["xl/styles.xml"] ? decodeUtf8(entries["xl/styles.xml"]) : null;
    const workbookDocument = parseXmlDocument(workbookXml);
    const relationshipsDocument = parseXmlDocument(workbookRelsXml);
    const styleBook = parseStylesXml(stylesXml);
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
        validateSheetName(name);
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
        const worksheetXml = decodeRequiredEntry(entries, target);
        return parseWorksheetXml(name, worksheetXml, styleBook);
      })
    };
  }

  function normalizeWorkbookTarget(target: string): string {
    return target.startsWith("xl/") ? target : `xl/${target.replace(/^\.\//, "")}`;
  }

  function parseWorksheetXml(name: string, xmlText: string, styleBook: StyleDescriptor[]): XlsxSheetModel {
    const document = parseXmlDocument(xmlText);
    const rowElements = Array.from(document.getElementsByTagNameNS(
      "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "row"
    ));

    return {
      name,
      columns: parseWorksheetColumns(document),
      freezePane: parseWorksheetFreezePane(document),
      mergedRanges: parseWorksheetMergedRanges(document),
      rows: rowElements.map((rowElement) => ({
        height: parseOptionalNumber(rowElement.getAttribute("ht")),
        cells: parseWorksheetRowCells(rowElement, styleBook)
      }))
    };
  }

  function parseWorksheetColumns(document: XMLDocument): XlsxColumnModel[] | undefined {
    const colElements = Array.from(document.getElementsByTagNameNS(
      "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "col"
    ));
    if (colElements.length === 0) {
      return undefined;
    }
    const columns: XlsxColumnModel[] = [];
    for (const colElement of colElements) {
      const min = Number(colElement.getAttribute("min") || "0");
      const max = Number(colElement.getAttribute("max") || "0");
      const width = parseOptionalNumber(colElement.getAttribute("width"));
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

  function parseWorksheetMergedRanges(document: XMLDocument): string[] | undefined {
    const mergeCellElements = Array.from(document.getElementsByTagNameNS(
      "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "mergeCell"
    ));
    if (mergeCellElements.length === 0) {
      return undefined;
    }
    return mergeCellElements
      .map((element) => normalizeMergedRange(element.getAttribute("ref") || ""))
      .filter(Boolean);
  }

  function parseWorksheetFreezePane(document: XMLDocument): XlsxFreezePaneModel | undefined {
    const paneElement = document.getElementsByTagNameNS(
      "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "pane"
    )[0];
    if (!paneElement || paneElement.getAttribute("state") !== "frozen") {
      return undefined;
    }
    const rowSplit = parseOptionalNumber(paneElement.getAttribute("ySplit"));
    const colSplit = parseOptionalNumber(paneElement.getAttribute("xSplit"));
    if (rowSplit === undefined && colSplit === undefined) {
      return undefined;
    }
    return {
      rowSplit,
      colSplit
    };
  }

  function parseWorksheetRowCells(rowElement: Element, styleBook: StyleDescriptor[]): XlsxCellModel[] {
    const cells: XlsxCellModel[] = [];
    const cellElements = Array.from(rowElement.getElementsByTagNameNS(
      "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "c"
    )).filter((element) => element.parentElement === rowElement);

    for (const cellElement of cellElements) {
      const reference = cellElement.getAttribute("r") || "";
      const columnIndex = decodeColumnReference(reference);
      while (cells.length < columnIndex) {
        cells.push({});
      }
      cells.push(parseWorksheetCell(cellElement, styleBook));
    }

    return cells;
  }

  function parseWorksheetCell(cellElement: Element, styleBook: StyleDescriptor[]): XlsxCellModel {
    const type = cellElement.getAttribute("t") || "";
    const styleIndex = Number(cellElement.getAttribute("s") || "0");
    const formulaElement = findDirectChild(cellElement, "f");
    const valueElement = findDirectChild(cellElement, "v");
    const inlineStringElement = findDirectChild(cellElement, "is");
    let value: XlsxPrimitiveValue | undefined;

    if (type === "inlineStr") {
      const textElement = inlineStringElement ? findDirectChild(inlineStringElement, "t") : null;
      value = textElement ? (textElement.textContent || "") : "";
    } else if (type === "b") {
      value = valueElement?.textContent === "1";
    } else if (type === "str") {
      value = valueElement?.textContent || "";
    } else if (valueElement) {
      const rawValue = valueElement.textContent || "";
      value = rawValue === "" ? "" : Number(rawValue);
    }

    const style = styleBook[styleIndex] || DEFAULT_STYLE;
    const cell: XlsxCellModel = {};
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
      cell.fillColor = denormalizeColor(style.fillColor);
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

  function parseStylesXml(xmlText: string | null): StyleDescriptor[] {
    if (!xmlText) {
      return [DEFAULT_STYLE];
    }
    const document = parseXmlDocument(xmlText);
    const fonts = parseFonts(document);
    const fills = parseFills(document);
    const borders = parseBorders(document);
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
      const alignmentElement = findDirectChild(xfElement, "alignment");
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

  function parseFonts(document: XMLDocument): FontDescriptor[] {
    return Array.from(document.getElementsByTagNameNS(
      "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "font"
    )).map((fontElement) => ({
      bold: findDirectChild(fontElement, "b") ? true : undefined,
      fontSize: parseOptionalNumber(findDirectChild(fontElement, "sz")?.getAttribute("val") || null)
    }));
  }

  function parseFills(document: XMLDocument): FillDescriptor[] {
    return Array.from(document.getElementsByTagNameNS(
      "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "fill"
    )).map((fillElement) => {
      const patternFill = findDirectChild(fillElement, "patternFill");
      const fgColor = patternFill ? findDirectChild(patternFill, "fgColor") : null;
      return {
        fillColor: fgColor?.getAttribute("rgb") || undefined
      };
    });
  }

  function parseBorders(document: XMLDocument): BorderDescriptor[] {
    return Array.from(document.getElementsByTagNameNS(
      "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "border"
    )).map((borderElement) => {
      const left = findDirectChild(borderElement, "left");
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
    const document = new DOMParser().parseFromString(xmlText, "application/xml");
    if (document.querySelector("parsererror")) {
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
    return TEXT_ENCODER.encode(value);
  }

  function decodeUtf8(bytes: Uint8Array): string {
    return TEXT_DECODER.decode(bytes);
  }

  function escapeXml(value: string): string {
    return value
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
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

  function resolveActivePane(freezePane: XlsxFreezePaneModel): string {
    if (freezePane.rowSplit && freezePane.colSplit) {
      return "bottomRight";
    }
    if (freezePane.rowSplit) {
      return "bottomLeft";
    }
    return "topRight";
  }

  function decodeColumnReference(reference: string): number {
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

  function packZip(entries: ZipEntry[]): Uint8Array {
    const localParts: Uint8Array[] = [];
    const centralParts: Uint8Array[] = [];
    let offset = 0;

    for (const entry of entries) {
      const nameBytes = encodeUtf8(entry.name);
      const crc32 = computeCrc32(entry.data);
      const localHeader = new Uint8Array(30 + nameBytes.length);
      const localView = new DataView(localHeader.buffer);
      localView.setUint32(0, 0x04034b50, true);
      localView.setUint16(4, 20, true);
      localView.setUint16(6, 0, true);
      localView.setUint16(8, 0, true);
      localView.setUint16(10, 0, true);
      localView.setUint16(12, 0, true);
      localView.setUint32(14, crc32, true);
      localView.setUint32(18, entry.data.byteLength, true);
      localView.setUint32(22, entry.data.byteLength, true);
      localView.setUint16(26, nameBytes.length, true);
      localView.setUint16(28, 0, true);
      localHeader.set(nameBytes, 30);

      const centralHeader = new Uint8Array(46 + nameBytes.length);
      const centralView = new DataView(centralHeader.buffer);
      centralView.setUint32(0, 0x02014b50, true);
      centralView.setUint16(4, 20, true);
      centralView.setUint16(6, 20, true);
      centralView.setUint16(8, 0, true);
      centralView.setUint16(10, 0, true);
      centralView.setUint16(12, 0, true);
      centralView.setUint16(14, 0, true);
      centralView.setUint32(16, crc32, true);
      centralView.setUint32(20, entry.data.byteLength, true);
      centralView.setUint32(24, entry.data.byteLength, true);
      centralView.setUint16(28, nameBytes.length, true);
      centralView.setUint16(30, 0, true);
      centralView.setUint16(32, 0, true);
      centralView.setUint16(34, 0, true);
      centralView.setUint16(36, 0, true);
      centralView.setUint32(38, 0, true);
      centralView.setUint32(42, offset, true);
      centralHeader.set(nameBytes, 46);

      localParts.push(localHeader, entry.data);
      centralParts.push(centralHeader);
      offset += localHeader.byteLength + entry.data.byteLength;
    }

    const centralDirectoryOffset = offset;
    const centralDirectorySize = centralParts.reduce((sum, part) => sum + part.byteLength, 0);
    const endOfCentralDirectory = new Uint8Array(22);
    const endView = new DataView(endOfCentralDirectory.buffer);
    endView.setUint32(0, 0x06054b50, true);
    endView.setUint16(4, 0, true);
    endView.setUint16(6, 0, true);
    endView.setUint16(8, entries.length, true);
    endView.setUint16(10, entries.length, true);
    endView.setUint32(12, centralDirectorySize, true);
    endView.setUint32(16, centralDirectoryOffset, true);
    endView.setUint16(20, 0, true);

    return concatUint8Arrays([...localParts, ...centralParts, endOfCentralDirectory]);
  }

  function unpackZip(bytes: Uint8Array): Record<string, Uint8Array> {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const endOffset = findEndOfCentralDirectoryOffset(bytes);
    const totalEntries = view.getUint16(endOffset + 10, true);
    const centralDirectoryOffset = view.getUint32(endOffset + 16, true);
    const entries: Record<string, Uint8Array> = {};
    let pointer = centralDirectoryOffset;

    for (let index = 0; index < totalEntries; index += 1) {
      if (view.getUint32(pointer, true) !== 0x02014b50) {
        throw new Error("Invalid ZIP central directory header");
      }

      const compressionMethod = view.getUint16(pointer + 10, true);
      const compressedSize = view.getUint32(pointer + 20, true);
      const uncompressedSize = view.getUint32(pointer + 24, true);
      const fileNameLength = view.getUint16(pointer + 28, true);
      const extraLength = view.getUint16(pointer + 30, true);
      const commentLength = view.getUint16(pointer + 32, true);
      const localHeaderOffset = view.getUint32(pointer + 42, true);
      const fileName = decodeUtf8(bytes.subarray(pointer + 46, pointer + 46 + fileNameLength));
      const localView = new DataView(bytes.buffer, bytes.byteOffset + localHeaderOffset, bytes.byteLength - localHeaderOffset);
      if (localView.getUint32(0, true) !== 0x04034b50) {
        throw new Error(`Invalid ZIP local header for entry: ${fileName}`);
      }
      const localFileNameLength = localView.getUint16(26, true);
      const localExtraLength = localView.getUint16(28, true);
      const dataOffset = localHeaderOffset + 30 + localFileNameLength + localExtraLength;
      const data = bytes.slice(dataOffset, dataOffset + compressedSize);

      if (compressionMethod !== 0) {
        throw new Error(`Unsupported ZIP compression method for entry ${fileName}: ${compressionMethod}`);
      }
      if (compressedSize !== uncompressedSize) {
        throw new Error(`Stored ZIP entry size mismatch: ${fileName}`);
      }

      entries[fileName] = data;
      pointer += 46 + fileNameLength + extraLength + commentLength;
    }

    return entries;
  }

  function findEndOfCentralDirectoryOffset(bytes: Uint8Array): number {
    for (let index = bytes.byteLength - 22; index >= 0; index -= 1) {
      if (
        bytes[index] === 0x50 &&
        bytes[index + 1] === 0x4b &&
        bytes[index + 2] === 0x05 &&
        bytes[index + 3] === 0x06
      ) {
        return index;
      }
    }
    throw new Error("ZIP end of central directory not found");
  }

  function concatUint8Arrays(parts: Uint8Array[]): Uint8Array {
    const totalLength = parts.reduce((sum, part) => sum + part.byteLength, 0);
    const result = new Uint8Array(totalLength);
    let offset = 0;
    for (const part of parts) {
      result.set(part, offset);
      offset += part.byteLength;
    }
    return result;
  }

  function computeCrc32(bytes: Uint8Array): number {
    let crc = 0xffffffff;
    for (const byte of bytes) {
      crc = (crc >>> 8) ^ CRC32_TABLE[(crc ^ byte) & 0xff];
    }
    return (crc ^ 0xffffffff) >>> 0;
  }

  function buildCrc32Table(): Uint32Array {
    const table = new Uint32Array(256);
    for (let index = 0; index < 256; index += 1) {
      let value = index;
      for (let bit = 0; bit < 8; bit += 1) {
        value = (value & 1) !== 0 ? (0xedb88320 ^ (value >>> 1)) : (value >>> 1);
      }
      table[index] = value >>> 0;
    }
    return table;
  }

  (globalThis as typeof globalThis & {
    __mikuprojectExcelIo?: {
      XlsxWorkbookCodec: typeof XlsxWorkbookCodec;
    };
  }).__mikuprojectExcelIo = {
    XlsxWorkbookCodec
  };
})();
