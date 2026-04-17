/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const mikuprojectCoreApiImport = globalThis.__mikuprojectCoreApiImport;
    if (!mikuprojectCoreApiImport) {
        throw new Error("mikuproject core api import module is not loaded");
    }
    const mikuprojectCoreApiMsproject = globalThis.__mikuprojectCoreApiMsproject;
    if (!mikuprojectCoreApiMsproject) {
        throw new Error("mikuproject core api msproject module is not loaded");
    }
    const mikuprojectCoreApiWorkbook = globalThis.__mikuprojectCoreApiWorkbook;
    if (!mikuprojectCoreApiWorkbook) {
        throw new Error("mikuproject core api workbook module is not loaded");
    }
    const mikuprojectCoreApiReportPublic = globalThis.__mikuprojectCoreApiReportPublic;
    if (!mikuprojectCoreApiReportPublic) {
        throw new Error("mikuproject core api report public module is not loaded");
    }
    globalThis.__mikuprojectCoreApiRegistry = {
        getAiJsonSpecText: mikuprojectCoreApiImport.getAiJsonSpecText,
        getAiJsonSpec: mikuprojectCoreApiImport.getAiJsonSpec,
        detectAiJsonDocumentKind: mikuprojectCoreApiImport.detectAiJsonDocumentKind,
        parseAiJsonText: mikuprojectCoreApiImport.parseAiJsonText,
        importAiJsonDocument: mikuprojectCoreApiImport.importAiJsonDocument,
        importAiJsonText: mikuprojectCoreApiImport.importAiJsonText,
        importExternal: mikuprojectCoreApiImport.importExternal,
        samples: mikuprojectCoreApiMsproject.samples,
        projectModel: mikuprojectCoreApiMsproject.projectModel,
        msProject: mikuprojectCoreApiMsproject.msProject,
        aiViews: mikuprojectCoreApiMsproject.aiViews,
        workbookJson: mikuprojectCoreApiWorkbook.workbookJson,
        xlsx: mikuprojectCoreApiWorkbook.xlsx,
        patchJson: mikuprojectCoreApiWorkbook.patchJson,
        report: mikuprojectCoreApiReportPublic.report
    };
})();
