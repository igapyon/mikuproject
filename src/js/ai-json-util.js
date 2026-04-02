/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    function extractLastJsonBlock(value) {
        var _a, _b;
        const matches = Array.from(value.matchAll(/```json\s*([\s\S]*?)```/g));
        if (matches.length > 0) {
            return ((_b = (_a = matches.at(-1)) === null || _a === void 0 ? void 0 : _a[1]) === null || _b === void 0 ? void 0 : _b.trim()) || "";
        }
        return value.trim();
    }
    function detectJsonDocumentKind(documentLike) {
        if (!documentLike || typeof documentLike !== "object") {
            return undefined;
        }
        const candidate = documentLike;
        if (candidate.format === "mikuproject_workbook_json") {
            return "workbook_json";
        }
        if (candidate.view_type === "project_draft_view") {
            return "project_draft_view";
        }
        if (Array.isArray(candidate.operations)) {
            return "patch_json";
        }
        return undefined;
    }
    globalThis.__mikuprojectAiJsonUtil = {
        extractLastJsonBlock,
        detectJsonDocumentKind
    };
})();
