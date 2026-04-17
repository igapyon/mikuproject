/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const projectWorkbookJsonExport = globalThis.__mikuprojectProjectWorkbookJsonExport;
    if (!projectWorkbookJsonExport) {
        throw new Error("mikuproject Project Workbook JSON export module is not loaded");
    }
    const projectWorkbookJsonImport = globalThis.__mikuprojectProjectWorkbookJsonImport;
    if (!projectWorkbookJsonImport) {
        throw new Error("mikuproject Project Workbook JSON import module is not loaded");
    }
    globalThis.__mikuprojectProjectWorkbookJson = {
        exportProjectWorkbookJson: (model) => projectWorkbookJsonExport.exportProjectWorkbookJson(model),
        importProjectWorkbookJsonAsProjectModel: (documentLike) => projectWorkbookJsonImport.importProjectWorkbookJsonAsProjectModel(documentLike),
        importProjectWorkbookJson: (documentLike, baseModel) => projectWorkbookJsonImport.importProjectWorkbookJson(documentLike, baseModel),
        validateWorkbookJsonDocument: (documentLike) => projectWorkbookJsonImport.validateWorkbookJsonDocument(documentLike)
    };
})();
