/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const mainFlow = {
        applyModelState(input) {
            input.updateSummary(input.model);
            input.renderValidationIssues(input.issues);
            input.renderImportWarnings(input.warnings || [], input.warningSourceLabel ? { sourceLabel: input.warningSourceLabel } : {});
            input.renderXlsxImportSummary(input.changes || [], input.changeSourceLabel || input.warnings
                ? {
                    sourceLabel: input.changeSourceLabel,
                    warnings: input.warnings
                }
                : {});
        },
        completeOutput(input) {
            input.setStatus(input.statusMessage);
            input.showToast(input.toastMessage);
            input.setActiveTab("output");
        },
        completeTransform(input) {
            input.setStatus(input.statusMessage);
            input.showToast(input.toastMessage);
            input.setActiveTab("transform", { skipTransformRefresh: true });
        }
    };
    globalThis.__mikuprojectMainFlow = mainFlow;
})();
