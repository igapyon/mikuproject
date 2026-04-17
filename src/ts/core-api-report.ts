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

  const mikuprojectXml = (globalThis as typeof globalThis & {
    __mikuprojectXml?: {
      exportMermaidGantt: (model: ProjectModel) => string;
    };
  }).__mikuprojectXml;
  if (!mikuprojectXml) {
    throw new Error("mikuproject XML module is not loaded");
  }

  const mikuprojectMainUtil = (globalThis as typeof globalThis & {
    __mikuprojectMainUtil?: {
      encodeUtf8: (value: string) => Uint8Array;
      packZipEntries: (entries: Array<{ name: string; data: Uint8Array }>) => Uint8Array;
    };
  }).__mikuprojectMainUtil;
  if (!mikuprojectMainUtil) {
    throw new Error("mikuproject main util module is not loaded");
  }

  const mikuprojectExcelIo = (globalThis as typeof globalThis & {
    __mikuprojectExcelIo?: {
      XlsxWorkbookCodec: new () => {
        exportWorkbook: (workbook: XlsxWorkbookLike) => Uint8Array;
      };
    };
  }).__mikuprojectExcelIo;
  if (!mikuprojectExcelIo) {
    throw new Error("mikuproject Excel IO module is not loaded");
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

  globalThis.__mikuprojectCoreApiReport = {
    exportAllReportEntries(
      model: ProjectModel,
      options: ReportProgressOptions = {}
    ): Array<{ name: string; data: Uint8Array }> {
      const codec = new mikuprojectExcelIo.XlsxWorkbookCodec();
      const monthlyArchive = mikuprojectNativeSvg.exportMonthlyWbsCalendarSvgArchive(model);
      const entries = [
        {
          name: "wbs.xlsx",
          data: codec.exportWorkbook(mikuprojectWbsXlsx.exportWbsWorkbook(model, options))
        },
        {
          name: "wbs.md",
          data: mikuprojectMainUtil.encodeUtf8(`${mikuprojectWbsMarkdown.exportWbsMarkdown(model, options)}\n`)
        },
        {
          name: "mermaid.mmd",
          data: mikuprojectMainUtil.encodeUtf8(`${mikuprojectXml.exportMermaidGantt(model)}\n`)
        },
        {
          name: "daily.svg",
          data: mikuprojectMainUtil.encodeUtf8(`${mikuprojectNativeSvg.exportNativeSvg(model, options)}\n`)
        },
        {
          name: "weekly.svg",
          data: mikuprojectMainUtil.encodeUtf8(`${mikuprojectNativeSvg.exportWeeklyNativeSvg(model, options)}\n`)
        }
      ];

      for (const entry of monthlyArchive.entries) {
        entries.push({
          name: `monthly-calendar/${entry.fileName}`,
          data: mikuprojectMainUtil.encodeUtf8(entry.svg)
        });
      }

      return entries;
    }
  };
})();
