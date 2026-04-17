/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  type WbsXlsxCellLike = {
    value?: string | number | boolean;
    numberFormat?: "general" | "integer" | "decimal" | "date" | "datetime" | "percent";
    horizontalAlign?: "left" | "center" | "right";
    verticalAlign?: "top" | "center" | "bottom";
    wrapText?: boolean;
    bold?: boolean;
    fontSize?: number;
    fillColor?: string;
    border?: "thin";
  };

  type WbsXlsxWorkbookLike = {
    sheets: Array<{
      name: string;
      columns?: Array<{ width?: number; hidden?: boolean }>;
      mergedRanges?: string[];
      rows: Array<{
        height?: number;
        cells: WbsXlsxCellLike[];
      }>;
    }>;
  };

  type WbsExportOptions = {
    holidayDates?: string[];
    displayDaysBeforeBaseDate?: number;
    displayDaysAfterBaseDate?: number;
    useBusinessDaysForDisplayRange?: boolean;
    useBusinessDaysForProgressBand?: boolean;
  };

  type WbsSheetLayoutHelper = {
    columnName(columnIndex: number): string;
    columnIndex(columnReference: string): number;
    reference(rowNumber: number, columnIndex: number): string;
    parseCellReference(reference: string): {
      reference: string;
      rowNumber: number;
      rowIndex: number;
      columnName: string;
      columnIndex: number;
    };
    range(startReference: string, endReference: string): string;
    describeCell(reference: string): string;
    logCell(reference: string, label?: string, logger?: (...args: unknown[]) => void): string;
  };

  type WbsXlsxFacade = {
    collectWbsHolidayDates(model: ProjectModel): string[];
    exportWbsWorkbook(model: ProjectModel, options?: WbsExportOptions): WbsXlsxWorkbookLike;
    layout: WbsSheetLayoutHelper;
  };

  type CreateWbsXlsxFacadeConfig = {
    collectWbsHolidayDates(model: ProjectModel): string[];
    exportWbsWorkbook(model: ProjectModel, options?: WbsExportOptions): WbsXlsxWorkbookLike;
    layout: WbsSheetLayoutHelper;
  };

  function createWbsXlsxFacade(config: CreateWbsXlsxFacadeConfig): WbsXlsxFacade {
    return {
      collectWbsHolidayDates: (model) => config.collectWbsHolidayDates(model),
      exportWbsWorkbook: (model, options) => config.exportWbsWorkbook(model, options),
      layout: config.layout
    };
  }

  (globalThis as typeof globalThis & {
    __mikuprojectWbsXlsxPublic?: {
      createWbsXlsxFacade: (config: CreateWbsXlsxFacadeConfig) => WbsXlsxFacade;
    };
  }).__mikuprojectWbsXlsxPublic = {
    createWbsXlsxFacade
  };
})();
