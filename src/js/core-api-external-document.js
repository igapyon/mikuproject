/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const mikuprojectCoreApiAiJson = globalThis.__mikuprojectCoreApiAiJson;
    if (!mikuprojectCoreApiAiJson) {
        throw new Error("mikuproject core api ai json module is not loaded");
    }
    function importDocument(format, document, mode, baseModel) {
        if (format === "workbook_json") {
            if (mode === "patch") {
                throw new Error("workbook JSON は patch import に対応していません");
            }
            if (mode === "merge" && !baseModel) {
                throw new Error("workbook JSON merge import には baseModel が必要です");
            }
            return mikuprojectCoreApiAiJson.importAiJsonDocument(document, mode === "merge" ? { baseModel } : {});
        }
        if (format === "project_draft_view") {
            if (mode !== "replace") {
                throw new Error("project_draft_view は replace import のみ対応です");
            }
            return mikuprojectCoreApiAiJson.importAiJsonDocument(document);
        }
        if (format === "patch_json") {
            if (mode !== "patch") {
                throw new Error("patch JSON は patch import のみ対応です");
            }
            return mikuprojectCoreApiAiJson.importAiJsonDocument(document, { baseModel });
        }
        throw new Error(`未対応の import format です: ${format}`);
    }
    globalThis.__mikuprojectCoreApiExternalDocument = {
        importDocument
    };
})();
