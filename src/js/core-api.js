/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const mikuprojectXml = globalThis.__mikuprojectXml;
    if (!mikuprojectXml) {
        throw new Error("mikuproject XML module is not loaded");
    }
    const mikuprojectAiJsonUtil = globalThis.__mikuprojectAiJsonUtil;
    if (!mikuprojectAiJsonUtil) {
        throw new Error("mikuproject AI JSON util module is not loaded");
    }
    const mikuprojectAiJsonSpec = globalThis.__mikuprojectAiJsonSpec;
    if (!mikuprojectAiJsonSpec) {
        throw new Error("mikuproject AI JSON spec module is not loaded");
    }
    const mikuprojectMainUtil = globalThis.__mikuprojectMainUtil;
    if (!mikuprojectMainUtil) {
        throw new Error("mikuproject main util module is not loaded");
    }
    const mikuprojectProjectWorkbookJson = globalThis.__mikuprojectProjectWorkbookJson;
    if (!mikuprojectProjectWorkbookJson) {
        throw new Error("mikuproject Project Workbook JSON module is not loaded");
    }
    const mikuprojectProjectXlsx = globalThis.__mikuprojectProjectXlsx;
    if (!mikuprojectProjectXlsx) {
        throw new Error("mikuproject Project XLSX module is not loaded");
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
    const mikuprojectProjectPatchJson = globalThis.__mikuprojectProjectPatchJson;
    if (!mikuprojectProjectPatchJson) {
        throw new Error("mikuproject Project Patch JSON module is not loaded");
    }
    function getAiJsonSpecText() {
        return mikuprojectAiJsonSpec.getAiJsonSpecText();
    }
    function getAiJsonSpec() {
        return mikuprojectAiJsonSpec.getAiJsonSpec();
    }
    function detectAiJsonDocumentKind(documentLike) {
        return mikuprojectAiJsonUtil.detectJsonDocumentKind(documentLike);
    }
    function parseAiJsonText(sourceText) {
        const jsonText = mikuprojectAiJsonUtil.extractLastJsonBlock(sourceText);
        const document = JSON.parse(jsonText);
        return {
            sourceText,
            jsonText,
            document,
            kind: detectAiJsonDocumentKind(document)
        };
    }
    function importAiJsonDocument(documentLike, options = {}) {
        const kind = detectAiJsonDocumentKind(documentLike);
        if (kind === "project_draft_view") {
            return {
                kind,
                mode: "replace",
                model: mikuprojectXml.importProjectDraftView(documentLike),
                warnings: []
            };
        }
        if (kind === "workbook_json") {
            if (options.baseModel) {
                return {
                    kind,
                    mode: "merge",
                    ...mikuprojectProjectWorkbookJson.importProjectWorkbookJson(documentLike, options.baseModel)
                };
            }
            return {
                kind,
                mode: "replace",
                ...mikuprojectProjectWorkbookJson.importProjectWorkbookJsonAsProjectModel(documentLike)
            };
        }
        if (kind === "patch_json") {
            if (!options.baseModel) {
                throw new Error("Patch JSON の適用には baseModel が必要です");
            }
            return {
                kind,
                mode: "patch",
                ...mikuprojectProjectPatchJson.importProjectPatchJson(documentLike, options.baseModel)
            };
        }
        throw new Error("AI JSON の format / view_type を判別できません");
    }
    function importAiJsonText(sourceText, options = {}) {
        const parsed = parseAiJsonText(sourceText);
        return {
            ...parsed,
            result: importAiJsonDocument(parsed.document, options)
        };
    }
    function importExternal(input) {
        const { source, mode, baseModel } = input;
        if (source.format === "ms_project_xml") {
            if (mode !== "replace") {
                throw new Error("MS Project XML は replace import のみ対応です");
            }
            return {
                kind: "ms_project_xml",
                mode,
                model: mikuprojectXml.importMsProjectXml(source.text),
                warnings: []
            };
        }
        if (source.format === "xlsx") {
            if (mode === "patch") {
                throw new Error("XLSX は replace または merge import のみ対応です");
            }
            const workbook = new mikuprojectExcelIo.XlsxWorkbookCodec().importWorkbook(source.bytes);
            if (mode === "replace") {
                return {
                    kind: "xlsx",
                    mode,
                    model: mikuprojectProjectXlsx.importProjectWorkbookAsProjectModel(workbook),
                    warnings: []
                };
            }
            if (mode === "merge") {
                if (!baseModel) {
                    throw new Error("XLSX merge import には baseModel が必要です");
                }
                if (typeof mikuprojectProjectXlsx.importProjectWorkbookDetailed === "function") {
                    return {
                        kind: "xlsx",
                        mode,
                        ...mikuprojectProjectXlsx.importProjectWorkbookDetailed(workbook, baseModel),
                        warnings: []
                    };
                }
                return {
                    kind: "xlsx",
                    mode,
                    model: mikuprojectProjectXlsx.importProjectWorkbook(workbook, baseModel),
                    changes: [],
                    warnings: []
                };
            }
        }
        if (source.format === "workbook_json") {
            if (mode === "patch") {
                throw new Error("workbook JSON は patch import に対応していません");
            }
            if (mode === "merge" && !baseModel) {
                throw new Error("workbook JSON merge import には baseModel が必要です");
            }
            return importAiJsonDocument(source.document, mode === "merge" ? { baseModel } : {});
        }
        if (source.format === "project_draft_view") {
            if (mode !== "replace") {
                throw new Error("project_draft_view は replace import のみ対応です");
            }
            return importAiJsonDocument(source.document);
        }
        if (source.format === "patch_json") {
            if (mode !== "patch") {
                throw new Error("patch JSON は patch import のみ対応です");
            }
            return importAiJsonDocument(source.document, { baseModel });
        }
        throw new Error(`未対応の import format です: ${source.format || "unknown"}`);
    }
    function exportAllReportEntries(model, options = {}) {
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
    globalThis.__mikuprojectCoreApi = {
        version: 1,
        getAiJsonSpecText,
        getAiJsonSpec,
        detectAiJsonDocumentKind,
        parseAiJsonText,
        importAiJsonDocument,
        importAiJsonText,
        importExternal,
        samples: {
            getSampleXml: () => mikuprojectXml.SAMPLE_XML,
            getSampleProjectDraftView: () => mikuprojectXml.SAMPLE_PROJECT_DRAFT_VIEW
        },
        projectModel: {
            normalize: mikuprojectXml.normalizeProjectModel,
            validate: mikuprojectXml.validateProjectModel
        },
        msProject: {
            parseXmlDocument: mikuprojectXml.parseXmlDocument,
            importFromXml: mikuprojectXml.importMsProjectXml,
            exportToXml: mikuprojectXml.exportMsProjectXml,
            importFromCsvParentId: mikuprojectXml.importCsvParentId,
            exportToCsvParentId: mikuprojectXml.exportCsvParentId
        },
        aiViews: {
            buildProjectDraftRequest: mikuprojectXml.buildProjectDraftRequest,
            importProjectDraftView: mikuprojectXml.importProjectDraftView,
            exportProjectOverviewView: mikuprojectXml.exportProjectOverviewView,
            exportTaskEditView: mikuprojectXml.exportTaskEditView,
            exportPhaseDetailView: mikuprojectXml.exportPhaseDetailView
        },
        workbookJson: {
            exportDocument: mikuprojectProjectWorkbookJson.exportProjectWorkbookJson,
            validateDocument: mikuprojectProjectWorkbookJson.validateWorkbookJsonDocument,
            importAsProjectModel: mikuprojectProjectWorkbookJson.importProjectWorkbookJsonAsProjectModel,
            importIntoProjectModel: mikuprojectProjectWorkbookJson.importProjectWorkbookJson
        },
        xlsx: {
            decodeWorkbook: (bytes) => new mikuprojectExcelIo.XlsxWorkbookCodec().importWorkbook(bytes),
            encodeWorkbook: (workbook) => new mikuprojectExcelIo.XlsxWorkbookCodec().exportWorkbook(workbook),
            exportWorkbook: mikuprojectProjectXlsx.exportProjectWorkbook,
            importAsProjectModel: mikuprojectProjectXlsx.importProjectWorkbookAsProjectModel,
            importIntoProjectModel: mikuprojectProjectXlsx.importProjectWorkbook
        },
        patchJson: {
            validateDocument: mikuprojectProjectPatchJson.validatePatchDocument,
            applyToProjectModel: mikuprojectProjectPatchJson.importProjectPatchJson
        },
        report: {
            all: {
                export: (model, options = {}) => {
                    const entries = exportAllReportEntries(model, options);
                    return {
                        entries,
                        zipBytes: mikuprojectMainUtil.packZipEntries(entries)
                    };
                }
            },
            wbsXlsx: {
                exportWorkbook: mikuprojectWbsXlsx.exportWbsWorkbook,
                exportBytes: (model, options = {}) => new mikuprojectExcelIo.XlsxWorkbookCodec().exportWorkbook(mikuprojectWbsXlsx.exportWbsWorkbook(model, options))
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
                exportGantt: mikuprojectXml.exportMermaidGantt
            }
        }
    };
})();
