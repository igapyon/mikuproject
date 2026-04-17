/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const mikuprojectCoreApiExternalBinary = globalThis.__mikuprojectCoreApiExternalBinary;
    if (!mikuprojectCoreApiExternalBinary) {
        throw new Error("mikuproject core api external binary module is not loaded");
    }
    const mikuprojectCoreApiExternalDocument = globalThis.__mikuprojectCoreApiExternalDocument;
    if (!mikuprojectCoreApiExternalDocument) {
        throw new Error("mikuproject core api external document module is not loaded");
    }
    function importExternal(input) {
        const { source, mode, baseModel } = input;
        if (source.format === "ms_project_xml") {
            return mikuprojectCoreApiExternalBinary.importMsProjectXml(source.text, mode);
        }
        if (source.format === "xlsx") {
            if (mode === "patch") {
                throw new Error("XLSX は replace または merge import のみ対応です");
            }
            return mikuprojectCoreApiExternalBinary.importXlsx(source.bytes, mode, baseModel);
        }
        if (source.format === "workbook_json" ||
            source.format === "project_draft_view" ||
            source.format === "patch_json") {
            return mikuprojectCoreApiExternalDocument.importDocument(source.format, source.document, mode, baseModel);
        }
        throw new Error(`未対応の import format です: ${source.format || "unknown"}`);
    }
    globalThis.__mikuprojectCoreApiExternalImport = {
        importExternal
    };
})();
