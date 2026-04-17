/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const mikuprojectCoreApiAiJson = globalThis.__mikuprojectCoreApiAiJson;
    if (!mikuprojectCoreApiAiJson) {
        throw new Error("mikuproject core api ai json module is not loaded");
    }
    const mikuprojectCoreApiExternalImport = globalThis.__mikuprojectCoreApiExternalImport;
    if (!mikuprojectCoreApiExternalImport) {
        throw new Error("mikuproject core api external import module is not loaded");
    }
    globalThis.__mikuprojectCoreApiImport = {
        getAiJsonSpecText: mikuprojectCoreApiAiJson.getAiJsonSpecText,
        getAiJsonSpec: mikuprojectCoreApiAiJson.getAiJsonSpec,
        detectAiJsonDocumentKind: mikuprojectCoreApiAiJson.detectAiJsonDocumentKind,
        parseAiJsonText: mikuprojectCoreApiAiJson.parseAiJsonText,
        importAiJsonDocument: mikuprojectCoreApiAiJson.importAiJsonDocument,
        importAiJsonText: mikuprojectCoreApiAiJson.importAiJsonText,
        importExternal: mikuprojectCoreApiExternalImport.importExternal
    };
})();
