/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const mikuprojectXml = globalThis.__mikuprojectXml;
    if (!mikuprojectXml) {
        throw new Error("mikuproject XML module is not loaded");
    }
    globalThis.__mikuprojectCoreApiMsprojectAi = {
        aiViews: {
            buildProjectDraftRequest: mikuprojectXml.buildProjectDraftRequest,
            importProjectDraftView: mikuprojectXml.importProjectDraftView,
            exportProjectOverviewView: mikuprojectXml.exportProjectOverviewView,
            exportTaskEditView: mikuprojectXml.exportTaskEditView,
            exportPhaseDetailView: mikuprojectXml.exportPhaseDetailView
        },
        mermaid: {
            exportGantt: mikuprojectXml.exportMermaidGantt
        }
    };
})();
