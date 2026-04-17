/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const mainImport = {
        importProjectDraftText(input) {
            const sourceText = input.sourceText.trim();
            if (!sourceText) {
                throw new Error("project_draft_view JSON を入力してください");
            }
            const draft = JSON.parse(input.extractLastJsonBlock(sourceText));
            const model = input.importProjectDraftView(draft);
            return {
                model,
                issues: input.validateProjectModel(model)
            };
        },
        importPatchJsonText(input) {
            const sourceText = input.sourceText.trim();
            if (!sourceText) {
                throw new Error("Patch JSON を入力してください");
            }
            const documentLike = JSON.parse(input.extractLastJsonBlock(sourceText));
            const result = input.importProjectPatchJson(documentLike, input.ensureCurrentModel());
            return {
                model: result.model,
                issues: input.validateProjectModel(result.model),
                changes: result.changes,
                warnings: result.warnings,
                regeneratedXmlText: result.changes.length > 0
                    ? input.exportMsProjectXml(result.model)
                    : undefined
            };
        },
        detectAiEditJsonKind(input) {
            const sourceText = input.sourceText.trim();
            if (!sourceText) {
                throw new Error("project_draft_view または Patch JSON を入力してください");
            }
            const documentLike = JSON.parse(input.extractLastJsonBlock(sourceText));
            const kind = input.detectJsonDocumentKind(documentLike);
            if (kind === "project_draft_view" || kind === "patch_json") {
                return kind;
            }
            throw new Error("project_draft_view または Patch JSON を入力してください");
        },
        importWorkbookJsonText(input) {
            const sourceText = input.sourceText.trim();
            if (!sourceText) {
                throw new Error("workbook JSON を入力してください");
            }
            const documentLike = JSON.parse(input.extractLastJsonBlock(sourceText));
            const replaceResult = input.importProjectWorkbookJsonAsProjectModel(documentLike);
            const diffResult = input.previousModel
                ? input.importProjectWorkbookJson(documentLike, input.previousModel)
                : { changes: [], warnings: replaceResult.warnings };
            return {
                model: replaceResult.model,
                issues: input.validateProjectModel(replaceResult.model),
                changes: diffResult.changes,
                warnings: replaceResult.warnings
            };
        },
        async importXlsxBytes(input) {
            const codec = input.createWorkbookCodec();
            const workbook = typeof codec.importWorkbookAsync === "function"
                ? await codec.importWorkbookAsync(input.bytes)
                : codec.importWorkbook(input.bytes);
            const model = input.importProjectWorkbookAsProjectModel(workbook);
            const diffResult = input.previousModel
                ? input.importProjectWorkbookDetailed(workbook, input.previousModel)
                : { changes: [] };
            return {
                model,
                issues: input.validateProjectModel(model),
                changes: diffResult.changes
            };
        },
        importCsvText(input) {
            const csvText = input.csvText.trim();
            if (!csvText) {
                throw new Error("CSV が空です");
            }
            const model = input.importCsvParentId(csvText);
            return {
                model,
                issues: input.validateProjectModel(model)
            };
        }
    };
    globalThis.__mikuprojectMainImport = mainImport;
})();
