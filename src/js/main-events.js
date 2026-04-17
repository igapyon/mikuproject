/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const mikuprojectMainButtonEvents = globalThis.__mikuprojectMainButtonEvents;
    if (!mikuprojectMainButtonEvents) {
        throw new Error("mikuproject main button events module is not loaded");
    }
    const mikuprojectMainInputEvents = globalThis.__mikuprojectMainInputEvents;
    if (!mikuprojectMainInputEvents) {
        throw new Error("mikuproject main input events module is not loaded");
    }
    const mainEvents = {
        bind(deps) {
            mikuprojectMainButtonEvents.bind({
                document: deps.document,
                setStatus: deps.setStatus,
                getImportFileInput: deps.getImportFileInput,
                loadSample: deps.loadSample,
                downloadAllOutputs: deps.downloadAllOutputs,
                setSvgPreviewMode: deps.setSvgPreviewMode,
                downloadCurrentSvg: deps.downloadCurrentSvg,
                downloadCurrentWeeklySvg: deps.downloadCurrentWeeklySvg,
                downloadCurrentMonthlyWbsSvgZip: deps.downloadCurrentMonthlyWbsSvgZip,
                downloadCurrentMermaidMmd: deps.downloadCurrentMermaidMmd,
                exportCurrentCsv: deps.exportCurrentCsv,
                exportCurrentProjectOverviewView: deps.exportCurrentProjectOverviewView,
                exportCurrentTaskEditView: deps.exportCurrentTaskEditView,
                exportCurrentAiProjectionBundle: deps.exportCurrentAiProjectionBundle,
                loadProjectDraftSample: deps.loadProjectDraftSample,
                copyAiPrompt: deps.copyAiPrompt,
                importAiEditJsonFromText: deps.importAiEditJsonFromText,
                exportCurrentPhaseDetailView: deps.exportCurrentPhaseDetailView,
                exportCurrentXlsx: deps.exportCurrentXlsx,
                exportCurrentWorkbookJson: deps.exportCurrentWorkbookJson,
                exportCurrentWbsXlsx: deps.exportCurrentWbsXlsx,
                downloadCurrentWbsMarkdown: deps.downloadCurrentWbsMarkdown,
                downloadCurrentXml: deps.downloadCurrentXml,
                runRoundTripCheck: deps.runRoundTripCheck
            });
            mikuprojectMainInputEvents.bind({
                setStatus: deps.setStatus,
                refreshXmlSaveState: deps.refreshXmlSaveState,
                markXmlSourceDirty: deps.markXmlSourceDirty,
                getImportFileInput: deps.getImportFileInput,
                getXmlInput: deps.getXmlInput,
                importFromFile: deps.importFromFile
            });
        }
    };
    globalThis.__mikuprojectMainEvents = mainEvents;
})();
