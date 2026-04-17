/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const projectXlsxImport = globalThis.__mikuprojectProjectXlsxImport;
    if (!projectXlsxImport) {
        throw new Error("mikuproject Project XLSX import module is not loaded");
    }
    const projectXlsxExport = globalThis.__mikuprojectProjectXlsxExport;
    if (!projectXlsxExport) {
        throw new Error("mikuproject Project XLSX export module is not loaded");
    }
    globalThis.__mikuprojectProjectXlsx = {
        exportProjectWorkbook: (model) => projectXlsxExport.exportProjectWorkbook(model),
        importProjectWorkbook: (workbook, baseModel) => projectXlsxImport.importProjectWorkbook(workbook, baseModel),
        importProjectWorkbookAsProjectModel: (workbook) => projectXlsxImport.importProjectWorkbookAsProjectModel(workbook),
        importProjectWorkbookDetailed: (workbook, baseModel) => projectXlsxImport.importProjectWorkbookDetailed(workbook, baseModel)
    };
})();
