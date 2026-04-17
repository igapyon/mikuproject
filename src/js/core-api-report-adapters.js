/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const mikuprojectCoreApiReport = globalThis.__mikuprojectCoreApiReport;
    if (!mikuprojectCoreApiReport) {
        throw new Error("mikuproject core api report module is not loaded");
    }
    const mikuprojectCoreApiMsproject = globalThis.__mikuprojectCoreApiMsproject;
    if (!mikuprojectCoreApiMsproject) {
        throw new Error("mikuproject core api msproject module is not loaded");
    }
    const mikuprojectCoreApiWorkbook = globalThis.__mikuprojectCoreApiWorkbook;
    if (!mikuprojectCoreApiWorkbook) {
        throw new Error("mikuproject core api workbook module is not loaded");
    }
    const mikuprojectMainUtil = globalThis.__mikuprojectMainUtil;
    if (!mikuprojectMainUtil) {
        throw new Error("mikuproject main util module is not loaded");
    }
    const mikuprojectWbsXlsx = globalThis.__mikuprojectWbsXlsx;
    if (!mikuprojectWbsXlsx) {
        throw new Error("mikuproject WBS XLSX module is not loaded");
    }
    const mikuprojectNativeSvg = globalThis.__mikuprojectNativeSvg;
    if (!mikuprojectNativeSvg) {
        throw new Error("mikuproject native SVG module is not loaded");
    }
    const mikuprojectWbsMarkdown = globalThis.__mikuprojectWbsMarkdown;
    if (!mikuprojectWbsMarkdown) {
        throw new Error("mikuproject WBS Markdown module is not loaded");
    }
    globalThis.__mikuprojectCoreApiReportAdapters = {
        report: {
            all: {
                export: (model, options = {}) => {
                    const entries = mikuprojectCoreApiReport.exportAllReportEntries(model, options);
                    return {
                        entries,
                        zipBytes: mikuprojectMainUtil.packZipEntries(entries)
                    };
                }
            },
            wbsXlsx: {
                exportWorkbook: mikuprojectWbsXlsx.exportWbsWorkbook,
                exportBytes: (model, options = {}) => mikuprojectCoreApiWorkbook.xlsx.encodeWorkbook(mikuprojectWbsXlsx.exportWbsWorkbook(model, options))
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
