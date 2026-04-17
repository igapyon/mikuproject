/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const mikuprojectXml = globalThis.__mikuprojectXml;
    if (!mikuprojectXml) {
        throw new Error("mikuproject XML module is not loaded");
    }
    const mikuprojectMainUtil = globalThis.__mikuprojectMainUtil;
    if (!mikuprojectMainUtil) {
        throw new Error("mikuproject main util module is not loaded");
    }
    const mikuprojectExcelIo = globalThis.__mikuprojectExcelIo;
    if (!mikuprojectExcelIo) {
        throw new Error("mikuproject Excel IO module is not loaded");
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
    globalThis.__mikuprojectCoreApiReport = {
        exportAllReportEntries(model, options = {}) {
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
