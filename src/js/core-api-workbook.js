/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const mikuprojectProjectWorkbookJson = globalThis.__mikuprojectProjectWorkbookJson;
    if (!mikuprojectProjectWorkbookJson) {
        throw new Error("mikuproject Project Workbook JSON module is not loaded");
    }
    const mikuprojectCoreApiWorkbookXlsx = globalThis.__mikuprojectCoreApiWorkbookXlsx;
    if (!mikuprojectCoreApiWorkbookXlsx) {
        throw new Error("mikuproject core api workbook xlsx module is not loaded");
    }
    const mikuprojectProjectPatchJson = globalThis.__mikuprojectProjectPatchJson;
    if (!mikuprojectProjectPatchJson) {
        throw new Error("mikuproject Project Patch JSON module is not loaded");
    }
    globalThis.__mikuprojectCoreApiWorkbook = {
        workbookJson: {
            exportDocument: mikuprojectProjectWorkbookJson.exportProjectWorkbookJson,
            validateDocument: mikuprojectProjectWorkbookJson.validateWorkbookJsonDocument,
            importAsProjectModel: mikuprojectProjectWorkbookJson.importProjectWorkbookJsonAsProjectModel,
            importIntoProjectModel: mikuprojectProjectWorkbookJson.importProjectWorkbookJson
        },
        xlsx: mikuprojectCoreApiWorkbookXlsx,
        patchJson: {
            validateDocument: mikuprojectProjectPatchJson.validatePatchDocument,
            applyToProjectModel: mikuprojectProjectPatchJson.importProjectPatchJson
        }
    };
})();
