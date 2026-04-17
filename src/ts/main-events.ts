/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  type MainEventsDeps = {
    document: Document;
    setStatus: (message: string) => void;
    refreshXmlSaveState: () => void;
    markXmlSourceDirty: () => void;
    getImportFileInput: () => HTMLInputElement;
    getXmlInput: () => HTMLTextAreaElement;
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
    importFromFile: (file: File | null | undefined) => Promise<void>;
  };

  const mikuprojectMainButtonEvents = (globalThis as typeof globalThis & {
    __mikuprojectMainButtonEvents?: {
      bind: (deps: Omit<MainEventsDeps, "refreshXmlSaveState" | "markXmlSourceDirty" | "getXmlInput" | "importFromFile">) => void;
    };
  }).__mikuprojectMainButtonEvents;

  if (!mikuprojectMainButtonEvents) {
    throw new Error("mikuproject main button events module is not loaded");
  }

  const mikuprojectMainInputEvents = (globalThis as typeof globalThis & {
    __mikuprojectMainInputEvents?: {
      bind: (deps: {
        setStatus: (message: string) => void;
        refreshXmlSaveState: () => void;
        markXmlSourceDirty: () => void;
        getImportFileInput: () => HTMLInputElement;
        getXmlInput: () => HTMLTextAreaElement;
        importFromFile: (file: File | null | undefined) => Promise<void>;
      }) => void;
    };
  }).__mikuprojectMainInputEvents;

  if (!mikuprojectMainInputEvents) {
    throw new Error("mikuproject main input events module is not loaded");
  }

  const mainEvents = {
    bind(deps: MainEventsDeps): void {
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

  (globalThis as typeof globalThis & {
    __mikuprojectMainEvents?: typeof mainEvents;
  }).__mikuprojectMainEvents = mainEvents;
})();
