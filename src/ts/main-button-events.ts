/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  type MainButtonEventsDeps = {
    document: Document;
    setStatus: (message: string) => void;
    getImportFileInput: () => HTMLInputElement;
    loadSample: () => void;
    downloadAllOutputs: () => void;
    setSvgPreviewMode: (mode: "daily" | "weekly" | "monthly") => void;
    downloadCurrentSvg: () => Promise<void>;
    downloadCurrentWeeklySvg: () => void;
    downloadCurrentMonthlyWbsSvgZip: () => void;
    downloadCurrentMermaidMmd: () => void;
    exportCurrentCsv: () => void;
    exportCurrentProjectOverviewView: () => void;
    exportCurrentTaskEditView: () => void;
    exportCurrentAiProjectionBundle: () => void;
    loadProjectDraftSample: () => void;
    copyAiPrompt: () => Promise<void>;
    importAiEditJsonFromText: () => Promise<void>;
    exportCurrentPhaseDetailView: (mode: "full" | "scoped") => void;
    exportCurrentXlsx: () => void;
    exportCurrentWorkbookJson: () => void;
    exportCurrentWbsXlsx: () => void;
    downloadCurrentWbsMarkdown: () => void;
    downloadCurrentXml: () => void;
    runRoundTripCheck: () => void;
  };

  function getButton(doc: Document, id: string): HTMLButtonElement {
    const element = doc.getElementById(id);
    if (!(element instanceof HTMLButtonElement)) {
      throw new Error(`button #${id} が見つかりません`);
    }
    return element;
  }

  const mainButtonEvents = {
    bind(deps: MainButtonEventsDeps): void {
      const bindClick = (id: string, action: () => void, fallbackMessage: string, errorLogLabel?: string): void => {
        getButton(deps.document, id).addEventListener("click", () => {
          try {
            action();
          } catch (error) {
            if (errorLogLabel) {
              console.error(errorLogLabel, error);
            }
            deps.setStatus(error instanceof Error ? error.message : fallbackMessage);
          }
        });
      };

      const bindAsyncClick = (id: string, action: () => Promise<void>, fallbackMessage: string, errorLogLabel?: string): void => {
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

  (globalThis as typeof globalThis & {
    __mikuprojectMainButtonEvents?: typeof mainButtonEvents;
  }).__mikuprojectMainButtonEvents = mainButtonEvents;
})();
