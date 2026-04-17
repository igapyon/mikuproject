/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const mikuprojectAiJsonUtil = globalThis.__mikuprojectAiJsonUtil;
    if (!mikuprojectAiJsonUtil) {
        throw new Error("mikuproject AI JSON util module is not loaded");
    }
    const mikuprojectAiJsonSpec = globalThis.__mikuprojectAiJsonSpec;
    if (!mikuprojectAiJsonSpec) {
        throw new Error("mikuproject AI JSON spec module is not loaded");
    }
    const mikuprojectCoreApiAiJsonImport = globalThis.__mikuprojectCoreApiAiJsonImport;
    if (!mikuprojectCoreApiAiJsonImport) {
        throw new Error("mikuproject core api ai json import module is not loaded");
    }
    function getAiJsonSpecText() {
        return mikuprojectAiJsonSpec.getAiJsonSpecText();
    }
    function getAiJsonSpec() {
        return mikuprojectAiJsonSpec.getAiJsonSpec();
    }
    function parseAiJsonText(sourceText) {
        const jsonText = mikuprojectAiJsonUtil.extractLastJsonBlock(sourceText);
        const document = JSON.parse(jsonText);
        return {
            sourceText,
            jsonText,
            document,
            kind: mikuprojectCoreApiAiJsonImport.detectAiJsonDocumentKind(document)
        };
    }
    function importAiJsonText(sourceText, options = {}) {
        const parsed = parseAiJsonText(sourceText);
        return {
            ...parsed,
            result: mikuprojectCoreApiAiJsonImport.importAiJsonDocument(parsed.document, options)
        };
    }
    globalThis.__mikuprojectCoreApiAiJson = {
        getAiJsonSpecText,
        getAiJsonSpec,
        detectAiJsonDocumentKind: mikuprojectCoreApiAiJsonImport.detectAiJsonDocumentKind,
        parseAiJsonText,
        importAiJsonDocument: mikuprojectCoreApiAiJsonImport.importAiJsonDocument,
        importAiJsonText
    };
})();
