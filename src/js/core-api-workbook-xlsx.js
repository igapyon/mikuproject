/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const mikuprojectProjectXlsx = globalThis.__mikuprojectProjectXlsx;
    if (!mikuprojectProjectXlsx) {
        throw new Error("mikuproject Project XLSX module is not loaded");
    }
    const mikuprojectExcelIo = globalThis.__mikuprojectExcelIo;
    if (!mikuprojectExcelIo) {
        throw new Error("mikuproject Excel IO module is not loaded");
    }
    globalThis.__mikuprojectCoreApiWorkbookXlsx = {
        decodeWorkbook: (bytes) => new mikuprojectExcelIo.XlsxWorkbookCodec().importWorkbook(bytes),
        encodeWorkbook: (workbook) => new mikuprojectExcelIo.XlsxWorkbookCodec().exportWorkbook(workbook),
        exportWorkbook: mikuprojectProjectXlsx.exportProjectWorkbook,
        importAsProjectModel: mikuprojectProjectXlsx.importProjectWorkbookAsProjectModel,
        importIntoProjectModel: mikuprojectProjectXlsx.importProjectWorkbook,
        importIntoProjectModelDetailed: (workbook, baseModel) => {
            if (typeof mikuprojectProjectXlsx.importProjectWorkbookDetailed !== "function") {
                return undefined;
            }
            return mikuprojectProjectXlsx.importProjectWorkbookDetailed(workbook, baseModel);
        }
    };
})();
