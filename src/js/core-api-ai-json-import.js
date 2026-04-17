/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const mikuprojectCoreApiMsproject = globalThis.__mikuprojectCoreApiMsproject;
    if (!mikuprojectCoreApiMsproject) {
        throw new Error("mikuproject core api msproject module is not loaded");
    }
    const mikuprojectCoreApiWorkbook = globalThis.__mikuprojectCoreApiWorkbook;
    if (!mikuprojectCoreApiWorkbook) {
        throw new Error("mikuproject core api workbook module is not loaded");
    }
    const mikuprojectAiJsonUtil = globalThis.__mikuprojectAiJsonUtil;
    if (!mikuprojectAiJsonUtil) {
        throw new Error("mikuproject AI JSON util module is not loaded");
    }
    function detectAiJsonDocumentKind(documentLike) {
        return mikuprojectAiJsonUtil.detectJsonDocumentKind(documentLike);
    }
    function importAiJsonDocument(documentLike, options = {}) {
        const kind = detectAiJsonDocumentKind(documentLike);
        if (kind === "project_draft_view") {
            return {
                kind,
                mode: "replace",
                model: mikuprojectCoreApiMsproject.aiViews.importProjectDraftView(documentLike),
                warnings: []
            };
        }
        if (kind === "workbook_json") {
            if (options.baseModel) {
                return {
                    kind,
                    mode: "merge",
                    ...mikuprojectCoreApiWorkbook.workbookJson.importIntoProjectModel(documentLike, options.baseModel)
                };
            }
            return {
                kind,
                mode: "replace",
                ...mikuprojectCoreApiWorkbook.workbookJson.importAsProjectModel(documentLike)
            };
        }
        if (kind === "patch_json") {
            if (!options.baseModel) {
                throw new Error("Patch JSON の適用には baseModel が必要です");
            }
            return {
                kind,
                mode: "patch",
                ...mikuprojectCoreApiWorkbook.patchJson.applyToProjectModel(documentLike, options.baseModel)
            };
        }
        throw new Error("AI JSON の format / view_type を判別できません");
    }
    globalThis.__mikuprojectCoreApiAiJsonImport = {
        detectAiJsonDocumentKind,
        importAiJsonDocument
    };
})();
