/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    function getButton(doc, id) {
        const element = doc.getElementById(id);
        if (!(element instanceof HTMLButtonElement)) {
            throw new Error(`button #${id} が見つかりません`);
        }
        return element;
    }
    const mainButtonEvents = {
        bind(deps) {
            const bindClick = (id, action, fallbackMessage, errorLogLabel) => {
                getButton(deps.document, id).addEventListener("click", () => {
                    try {
                        action();
                    }
                    catch (error) {
                        if (errorLogLabel) {
                            console.error(errorLogLabel, error);
                        }
                        deps.setStatus(error instanceof Error ? error.message : fallbackMessage);
                    }
                });
            };
            const bindAsyncClick = (id, action, fallbackMessage, errorLogLabel) => {
                getButton(deps.document, id).addEventListener("click", () => {
                    void action().catch((error) => {
                        if (errorLogLabel) {
                            console.error(errorLogLabel, error);
                        }
                        deps.setStatus(error instanceof Error ? error.message : fallbackMessage);
                    });
                });
            };
            getButton(deps.document, "loadSampleBtn").addEventListener("click", deps.loadSample);
            getButton(deps.document, "importFileBtn").addEventListener("click", () => {
                const input = deps.getImportFileInput();
                input.value = "";
                input.click();
            });
            bindClick("downloadAllOutputsBtn", deps.downloadAllOutputs, "All 出力保存に失敗しました", "[mikuproject] all outputs download failed");
            bindClick("previewDailySvgBtn", () => deps.setSvgPreviewMode("daily"), "SVG preview 切替に失敗しました");
            bindClick("previewWeeklySvgBtn", () => deps.setSvgPreviewMode("weekly"), "SVG preview 切替に失敗しました");
            bindClick("previewMonthlySvgBtn", () => deps.setSvgPreviewMode("monthly"), "SVG preview 切替に失敗しました");
            bindAsyncClick("downloadSvgBtn", deps.downloadCurrentSvg, "SVG 保存に失敗しました", "[mikuproject] native SVG download failed");
            bindClick("downloadWeeklySvgBtn", deps.downloadCurrentWeeklySvg, "Weekly SVG 保存に失敗しました", "[mikuproject] weekly SVG download failed");
            bindClick("downloadMonthlyCalendarSvgBtn", deps.downloadCurrentMonthlyWbsSvgZip, "月別 WBS カレンダー SVG 保存に失敗しました", "[mikuproject] monthly WBS SVG download failed");
            bindClick("exportMermaidMdBtn", deps.downloadCurrentMermaidMmd, "Mermaid 保存に失敗しました");
            bindClick("exportCsvBtn", deps.exportCurrentCsv, "CSV 生成に失敗しました");
            bindClick("exportProjectOverviewBtn", deps.exportCurrentProjectOverviewView, "project_overview_view 生成に失敗しました");
            bindClick("exportTaskEditBtn", deps.exportCurrentTaskEditView, "task_edit_view 生成に失敗しました");
            bindClick("exportAiBundleBtn", deps.exportCurrentAiProjectionBundle, "AI 連携用まとめ JSON 生成に失敗しました");
            getButton(deps.document, "loadProjectDraftSampleBtn").addEventListener("click", deps.loadProjectDraftSample);
            bindAsyncClick("copyAiPromptBtn", deps.copyAiPrompt, "生成AIプロンプトのコピーに失敗しました");
            bindAsyncClick("importProjectDraftBtn", deps.importAiEditJsonFromText, "編集用 JSON 取り込みに失敗しました");
            bindClick("exportPhaseDetailBtn", () => deps.exportCurrentPhaseDetailView("scoped"), "phase_detail_view 生成に失敗しました");
            bindClick("exportPhaseDetailFullBtn", () => deps.exportCurrentPhaseDetailView("full"), "phase_detail_view 生成に失敗しました");
            bindClick("exportXlsxBtn", deps.exportCurrentXlsx, "XLSX エクスポートに失敗しました");
            bindClick("exportWorkbookJsonBtn", deps.exportCurrentWorkbookJson, "JSON エクスポートに失敗しました");
            bindClick("exportWbsXlsxBtn", deps.exportCurrentWbsXlsx, "WBS XLSX エクスポートに失敗しました");
            bindClick("exportWbsMdBtn", deps.downloadCurrentWbsMarkdown, "WBS Markdown 保存に失敗しました");
            bindClick("downloadXmlBtn", deps.downloadCurrentXml, "XML 保存に失敗しました");
            bindClick("roundTripBtn", deps.runRoundTripCheck, "再読込テストに失敗しました");
        }
    };
    globalThis.__mikuprojectMainButtonEvents = mainButtonEvents;
})();
