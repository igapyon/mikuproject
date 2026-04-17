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

  function createWorkbookEntries(
    workbook: XlsxWorkbookModel,
    helpers: {
      createStyleBook: (workbook: XlsxWorkbookModel) => StyleBook;
      buildWorksheetXml: (
        sheet: XlsxSheetModel,
        styleBook: StyleBook,
        helpers: {
          resolveStyleIndex: (cell: XlsxCellModel, styleBook: StyleBook) => number;
          resolveCellNumberFormat: (cell: XlsxCellModel) => XlsxNumberFormat;
          escapeXml: (value: string) => string;
        }
      ) => string;
      resolveStyleIndex: (cell: XlsxCellModel, styleBook: StyleBook) => number;
      resolveCellNumberFormat: (cell: XlsxCellModel) => XlsxNumberFormat;
      buildContentTypesXml: (sheetCount: number, includeStyles: boolean) => string;
      buildRootRelationshipsXml: () => string;
      buildWorkbookRelationshipsXml: (
        relationships: Array<{ relationshipId: string; target: string }>,
        includeStyles: boolean
      ) => string;
      buildWorkbookXml: (
        relationships: Array<{ relationshipId: string; name: string }>
      ) => string;
      buildStylesXml: (styles: StyleDescriptor[]) => string;
      encodeUtf8: (value: string) => Uint8Array;
      escapeXml: (value: string) => string;
    }
  ): ZipEntry[] {
    const styleBook = helpers.createStyleBook(workbook);
    const worksheetRelationships = workbook.sheets.map((sheet, index) => ({
      relationshipId: `rId${index + 1}`,
      target: `worksheets/sheet${index + 1}.xml`,
      name: sheet.name
    }));
    const worksheetEntries = workbook.sheets.map((sheet, index) => ({
      name: `xl/worksheets/sheet${index + 1}.xml`,
      data: helpers.encodeUtf8(helpers.buildWorksheetXml(sheet, styleBook, {
        resolveStyleIndex: helpers.resolveStyleIndex,
        resolveCellNumberFormat: helpers.resolveCellNumberFormat,
        escapeXml: helpers.escapeXml
      }))
    }));

    const entries: ZipEntry[] = [
      {
        name: "[Content_Types].xml",
        data: helpers.encodeUtf8(helpers.buildContentTypesXml(workbook.sheets.length, styleBook.styles.length > 1))
      },
      {
        name: "_rels/.rels",
        data: helpers.encodeUtf8(helpers.buildRootRelationshipsXml())
      },
      {
        name: "xl/_rels/workbook.xml.rels",
        data: helpers.encodeUtf8(helpers.buildWorkbookRelationshipsXml(worksheetRelationships, styleBook.styles.length > 1))
      },
      {
        name: "xl/workbook.xml",
        data: helpers.encodeUtf8(helpers.buildWorkbookXml(worksheetRelationships))
      }
    ];

    if (styleBook.styles.length > 1) {
      entries.push({
        name: "xl/styles.xml",
        data: helpers.encodeUtf8(helpers.buildStylesXml(styleBook.styles))
      });
    }

    entries.push(...worksheetEntries);
    return entries;
  }

  (globalThis as typeof globalThis & {
    __mikuprojectExcelIoWorkbookBuild?: {
      createWorkbookEntries: (
        workbook: XlsxWorkbookModel,
        helpers: {
          createStyleBook: (workbook: XlsxWorkbookModel) => StyleBook;
          buildWorksheetXml: (
            sheet: XlsxSheetModel,
            styleBook: StyleBook,
            helpers: {
              resolveStyleIndex: (cell: XlsxCellModel, styleBook: StyleBook) => number;
              resolveCellNumberFormat: (cell: XlsxCellModel) => XlsxNumberFormat;
              escapeXml: (value: string) => string;
            }
          ) => string;
          resolveStyleIndex: (cell: XlsxCellModel, styleBook: StyleBook) => number;
          resolveCellNumberFormat: (cell: XlsxCellModel) => XlsxNumberFormat;
          buildContentTypesXml: (sheetCount: number, includeStyles: boolean) => string;
          buildRootRelationshipsXml: () => string;
          buildWorkbookRelationshipsXml: (
            relationships: Array<{ relationshipId: string; target: string }>,
            includeStyles: boolean
          ) => string;
          buildWorkbookXml: (
            relationships: Array<{ relationshipId: string; name: string }>
          ) => string;
          buildStylesXml: (styles: StyleDescriptor[]) => string;
          encodeUtf8: (value: string) => Uint8Array;
          escapeXml: (value: string) => string;
        }
      ) => ZipEntry[];
    };
  }).__mikuprojectExcelIoWorkbookBuild = {
    createWorkbookEntries
  };
})();
