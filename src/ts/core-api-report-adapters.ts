/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  type ReportDisplayRangeOptions = {
    holidayDates?: string[];
    displayDaysBeforeBaseDate?: number;
    displayDaysAfterBaseDate?: number;
    useBusinessDaysForDisplayRange?: boolean;
  };

  type ReportProgressOptions = ReportDisplayRangeOptions & {
    useBusinessDaysForProgressBand?: boolean;
  };

  type XlsxCellLike = {
    value?: string | number | boolean;
    numberFormat?: "general" | "integer" | "decimal" | "date" | "datetime" | "percent";
    horizontalAlign?: "left" | "center" | "right";
    bold?: boolean;
    fontSize?: number;
    fillColor?: string;
    border?: "thin";
  };

  type XlsxWorkbookLike = {
    sheets: Array<{
      name: string;
      columns?: Array<{ width?: number }>;
      mergedRanges?: string[];
      dataValidations?: Array<{
        type: "list";
        sqref: string;
        formula1: string;
        allowBlank?: boolean;
      }>;
      rows: Array<{
        height?: number;
        cells: XlsxCellLike[];
      }>;
    }>;
  };

  const mikuprojectCoreApiReport = (globalThis as typeof globalThis & {
    __mikuprojectCoreApiReport?: {
      exportAllReportEntries: (
        model: ProjectModel,
        options?: ReportProgressOptions
      ) => Array<{ name: string; data: Uint8Array }>;
    };
  }).__mikuprojectCoreApiReport;

  if (!mikuprojectCoreApiReport) {
    throw new Error("mikuproject core api report module is not loaded");
  }

  const mikuprojectCoreApiMsproject = (globalThis as typeof globalThis & {
    __mikuprojectCoreApiMsproject?: {
      mermaid: {
        exportGantt: (model: ProjectModel) => string;
      };
    };
  }).__mikuprojectCoreApiMsproject;

  if (!mikuprojectCoreApiMsproject) {
    throw new Error("mikuproject core api msproject module is not loaded");
  }

  const mikuprojectCoreApiWorkbook = (globalThis as typeof globalThis & {
    __mikuprojectCoreApiWorkbook?: {
      xlsx: {
        encodeWorkbook: (workbook: XlsxWorkbookLike) => Uint8Array;
      };
    };
  }).__mikuprojectCoreApiWorkbook;

  if (!mikuprojectCoreApiWorkbook) {
    throw new Error("mikuproject core api workbook module is not loaded");
  }

  const mikuprojectMainUtil = (globalThis as typeof globalThis & {
    __mikuprojectMainUtil?: {
      packZipEntries: (entries: Array<{ name: string; data: Uint8Array }>) => Uint8Array;
    };
  }).__mikuprojectMainUtil;

  if (!mikuprojectMainUtil) {
    throw new Error("mikuproject main util module is not loaded");
  }

  const mikuprojectWbsXlsx = (globalThis as typeof globalThis & {
    __mikuprojectWbsXlsx?: {
      exportWbsWorkbook: (model: ProjectModel, options?: ReportProgressOptions) => XlsxWorkbookLike;
    };
  }).__mikuprojectWbsXlsx;

  if (!mikuprojectWbsXlsx) {
    throw new Error("mikuproject WBS XLSX module is not loaded");
  }

  const mikuprojectNativeSvg = (globalThis as typeof globalThis & {
    __mikuprojectNativeSvg?: {
      exportNativeSvg: (
        model: ProjectModel,
        options?: ReportProgressOptions & { labelMode?: "near" | "list" }
      ) => string;
      exportWeeklyNativeSvg: (model: ProjectModel, options?: ReportDisplayRangeOptions) => string;
      exportMonthlyWbsCalendarSvgArchive: (model: ProjectModel) => {
        entries: Array<{ fileName: string; svg: string }>;
        zipBytes: Uint8Array;
      };
    };
  }).__mikuprojectNativeSvg;

  if (!mikuprojectNativeSvg) {
    throw new Error("mikuproject native SVG module is not loaded");
  }

  const mikuprojectWbsMarkdown = (globalThis as typeof globalThis & {
    __mikuprojectWbsMarkdown?: {
      exportWbsMarkdown: (model: ProjectModel, options?: ReportProgressOptions) => string;
    };
  }).__mikuprojectWbsMarkdown;

  if (!mikuprojectWbsMarkdown) {
    throw new Error("mikuproject WBS Markdown module is not loaded");
  }

  (globalThis as typeof globalThis & {
    __mikuprojectCoreApiReportAdapters?: {
      report: unknown;
    };
  }).__mikuprojectCoreApiReportAdapters = {
    report: {
      all: {
        export: (model: ProjectModel, options: ReportProgressOptions = {}) => {
          const entries = mikuprojectCoreApiReport.exportAllReportEntries(model, options);
          return {
            entries,
            zipBytes: mikuprojectMainUtil.packZipEntries(entries)
          };
        }
      },
      wbsXlsx: {
        exportWorkbook: mikuprojectWbsXlsx.exportWbsWorkbook,
        exportBytes: (model: ProjectModel, options: ReportProgressOptions = {}) =>
          mikuprojectCoreApiWorkbook.xlsx.encodeWorkbook(mikuprojectWbsXlsx.exportWbsWorkbook(model, options))
      },
      svg: {
        exportDaily: mikuprojectNativeSvg.exportNativeSvg,
        exportWeekly: mikuprojectNativeSvg.exportWeeklyNativeSvg,
        exportMonthlyCalendar: mikuprojectNativeSvg.exportMonthlyWbsCalendarSvgArchive
      },
      wbsMarkdown: {
        export: mikuprojectWbsMarkdown.exportWbsMarkdown
      },
      mermaid: {
        exportGantt: mikuprojectCoreApiMsproject.mermaid.exportGantt
      }
    }
  };
})();
