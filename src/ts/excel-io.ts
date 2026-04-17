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

  const excelIoUtil = (globalThis as typeof globalThis & {
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
  }).__mikuprojectExcelIoUtil;

  if (!excelIoUtil) {
    throw new Error("mikuproject Excel util module is not loaded");
  }
  const excelIoZip = (globalThis as typeof globalThis & {
    __mikuprojectExcelIoZip?: {
      packZip: (entries: ZipEntry[]) => Uint8Array;
      unpackZip: (bytes: Uint8Array) => Record<string, Uint8Array>;
      unpackZipAsync: (bytes: Uint8Array) => Promise<Record<string, Uint8Array>>;
    };
  }).__mikuprojectExcelIoZip;

  if (!excelIoZip) {
    throw new Error("mikuproject Excel ZIP module is not loaded");
  }
  const excelIoNormalize = (globalThis as typeof globalThis & {
    __mikuprojectExcelIoNormalize?: {
      normalizeWorkbook: (workbook: XlsxWorkbookModel) => XlsxWorkbookModel;
      normalizeMergedRange: (range: string) => string;
      validateSheetName: (name: string) => void;
      denormalizeColor: (color: string | undefined) => string | undefined;
    };
  }).__mikuprojectExcelIoNormalize;

  if (!excelIoNormalize) {
    throw new Error("mikuproject Excel normalize module is not loaded");
  }
  const excelIoPackageXml = (globalThis as typeof globalThis & {
    __mikuprojectExcelIoPackageXml?: {
      buildContentTypesXml: (sheetCount: number, includeStyles: boolean) => string;
      buildRootRelationshipsXml: () => string;
      buildWorkbookRelationshipsXml: (
        relationships: Array<{ relationshipId: string; target: string }>,
        includeStyles: boolean
      ) => string;
      buildWorkbookXml: (
        relationships: Array<{ relationshipId: string; name: string }>
      ) => string;
    };
  }).__mikuprojectExcelIoPackageXml;

  if (!excelIoPackageXml) {
    throw new Error("mikuproject Excel package XML module is not loaded");
  }
  const excelIoWorksheetBuild = (globalThis as typeof globalThis & {
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
  }).__mikuprojectExcelIoWorksheetBuild;

  if (!excelIoWorksheetBuild) {
    throw new Error("mikuproject Excel worksheet build module is not loaded");
  }
  const excelIoStylesBuild = (globalThis as typeof globalThis & {
    __mikuprojectExcelIoStylesBuild?: {
      createStyleBook: (workbook: XlsxWorkbookModel) => StyleBook;
      resolveStyleIndex: (cell: XlsxCellModel, styleBook: StyleBook) => number;
      resolveCellNumberFormat: (cell: XlsxCellModel) => XlsxNumberFormat;
      buildStylesXml: (styles: StyleDescriptor[]) => string;
    };
  }).__mikuprojectExcelIoStylesBuild;

  if (!excelIoStylesBuild) {
    throw new Error("mikuproject Excel styles build module is not loaded");
  }
  const excelIoStylesParse = (globalThis as typeof globalThis & {
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
  }).__mikuprojectExcelIoStylesParse;

  if (!excelIoStylesParse) {
    throw new Error("mikuproject Excel styles parse module is not loaded");
  }
  const excelIoWorksheetParse = (globalThis as typeof globalThis & {
    __mikuprojectExcelIoWorksheetParse?: {
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
      parseSharedStringsXml: (
        xmlText: string | null,
        helpers: {
          parseXmlDocument: (xmlText: string) => XMLDocument;
          findDirectChild: (element: Element, localName: string) => Element | null;
        }
      ) => string[];
    };
  }).__mikuprojectExcelIoWorksheetParse;

  if (!excelIoWorksheetParse) {
    throw new Error("mikuproject Excel worksheet parse module is not loaded");
  }
  const excelIoWorkbookParse = (globalThis as typeof globalThis & {
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
  }).__mikuprojectExcelIoWorkbookParse;

  if (!excelIoWorkbookParse) {
    throw new Error("mikuproject Excel workbook parse module is not loaded");
  }
  const excelIoWorkbookBuild = (globalThis as typeof globalThis & {
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
  }).__mikuprojectExcelIoWorkbookBuild;

  if (!excelIoWorkbookBuild) {
    throw new Error("mikuproject Excel workbook build module is not loaded");
  }

  class XlsxWorkbookCodec {
    exportWorkbook(workbook: XlsxWorkbookModel): Uint8Array {
      const normalizedWorkbook = excelIoNormalize.normalizeWorkbook(workbook);
      const entries = excelIoWorkbookBuild.createWorkbookEntries(normalizedWorkbook, {
        createStyleBook: excelIoStylesBuild.createStyleBook,
        buildWorksheetXml: excelIoWorksheetBuild.buildWorksheetXml,
        resolveStyleIndex: excelIoStylesBuild.resolveStyleIndex,
        resolveCellNumberFormat: excelIoStylesBuild.resolveCellNumberFormat,
        buildContentTypesXml: excelIoPackageXml.buildContentTypesXml,
        buildRootRelationshipsXml: excelIoPackageXml.buildRootRelationshipsXml,
        buildWorkbookRelationshipsXml: excelIoPackageXml.buildWorkbookRelationshipsXml,
        buildWorkbookXml: excelIoPackageXml.buildWorkbookXml,
        buildStylesXml: excelIoStylesBuild.buildStylesXml,
        encodeUtf8: excelIoUtil.encodeUtf8,
        escapeXml: excelIoUtil.escapeXml
      });
      return excelIoZip.packZip(entries);
    }

    importWorkbook(bytes: Uint8Array): XlsxWorkbookModel {
      const entries = this.unpackEntries(bytes);
      return excelIoWorkbookParse.parseWorkbookEntries(entries, {
        decodeRequiredEntry: excelIoUtil.decodeRequiredEntry,
        decodeUtf8: excelIoUtil.decodeUtf8,
        parseXmlDocument: excelIoUtil.parseXmlDocument,
        findDirectChild: excelIoUtil.findDirectChild,
        parseOptionalNumber: excelIoUtil.parseOptionalNumber,
        validateSheetName: excelIoNormalize.validateSheetName,
        normalizeMergedRange: excelIoNormalize.normalizeMergedRange,
        denormalizeColor: excelIoNormalize.denormalizeColor,
        parseStylesXml: excelIoStylesParse.parseStylesXml,
        parseSharedStringsXml: excelIoWorksheetParse.parseSharedStringsXml,
        parseWorksheetXml: excelIoWorksheetParse.parseWorksheetXml
      });
    }

    async importWorkbookAsync(bytes: Uint8Array): Promise<XlsxWorkbookModel> {
      const entries = await this.unpackEntriesAsync(bytes);
      return excelIoWorkbookParse.parseWorkbookEntries(entries, {
        decodeRequiredEntry: excelIoUtil.decodeRequiredEntry,
        decodeUtf8: excelIoUtil.decodeUtf8,
        parseXmlDocument: excelIoUtil.parseXmlDocument,
        findDirectChild: excelIoUtil.findDirectChild,
        parseOptionalNumber: excelIoUtil.parseOptionalNumber,
        validateSheetName: excelIoNormalize.validateSheetName,
        normalizeMergedRange: excelIoNormalize.normalizeMergedRange,
        denormalizeColor: excelIoNormalize.denormalizeColor,
        parseStylesXml: excelIoStylesParse.parseStylesXml,
        parseSharedStringsXml: excelIoWorksheetParse.parseSharedStringsXml,
        parseWorksheetXml: excelIoWorksheetParse.parseWorksheetXml
      });
    }

    listEntries(bytes: Uint8Array): string[] {
      return Object.keys(this.unpackEntries(bytes)).sort();
    }

    async listEntriesAsync(bytes: Uint8Array): Promise<string[]> {
      const entries = await this.unpackEntriesAsync(bytes);
      return Object.keys(entries).sort();
    }

    unpackEntries(bytes: Uint8Array): Record<string, Uint8Array> {
      return excelIoZip.unpackZip(bytes);
    }

    async unpackEntriesAsync(bytes: Uint8Array): Promise<Record<string, Uint8Array>> {
      return excelIoZip.unpackZipAsync(bytes);
    }
  }

  (globalThis as typeof globalThis & {
    __mikuprojectExcelIo?: {
      XlsxWorkbookCodec: typeof XlsxWorkbookCodec;
    };
  }).__mikuprojectExcelIo = {
    XlsxWorkbookCodec
  };
})();
