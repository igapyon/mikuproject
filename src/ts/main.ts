/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  const mikuprojectXml = (globalThis as typeof globalThis & {
    __mikuprojectXml?: {
      SAMPLE_XML: string;
      SAMPLE_PROJECT_DRAFT_VIEW: unknown;
      importMsProjectXml: (xmlText: string) => ProjectModel;
      importCsvParentId: (csvText: string) => ProjectModel;
      exportMsProjectXml: (model: ProjectModel) => string;
      exportMermaidGantt: (model: ProjectModel) => string;
      buildProjectDraftRequest: (input: {
        name: string;
        plannedStart?: string;
        goal?: string;
        teamCount?: number;
        mustHavePhases?: string[];
        mustHaveMilestones?: string[];
      }) => unknown;
      importProjectDraftView: (draft: unknown) => ProjectModel;
      exportProjectOverviewView: (model: ProjectModel) => unknown;
      exportTaskEditView: (model: ProjectModel, taskUid?: string) => unknown;
      exportPhaseDetailView: (
        model: ProjectModel,
        phaseUid?: string,
        options?: {
          mode?: "full" | "scoped";
          rootUid?: string;
          maxDepth?: number;
        }
      ) => unknown;
      exportCsvParentId: (model: ProjectModel) => string;
      normalizeProjectModel: (model: ProjectModel) => ProjectModel;
      validateProjectModel: (model: ProjectModel) => ValidationIssue[];
    };
  }).__mikuprojectXml;

  if (!mikuprojectXml) {
    throw new Error("mikuproject XML module is not loaded");
  }

  const mikuprojectAiJsonUtil = (globalThis as typeof globalThis & {
    __mikuprojectAiJsonUtil?: {
      extractLastJsonBlock: (value: string) => string;
      detectJsonDocumentKind: (documentLike: unknown) => "workbook_json" | "project_draft_view" | "patch_json" | undefined;
    };
  }).__mikuprojectAiJsonUtil;

  if (!mikuprojectAiJsonUtil) {
    throw new Error("mikuproject AI JSON util module is not loaded");
  }

  const mikuprojectAiJsonSpec = (globalThis as typeof globalThis & {
    __mikuprojectAiJsonSpec?: {
      getAiJsonSpecText: () => string;
    };
  }).__mikuprojectAiJsonSpec;

  const mikuprojectMainUtil = (globalThis as typeof globalThis & {
    __mikuprojectMainUtil?: {
      parseOptionalNonNegativeInteger: (raw: string) => number | undefined;
      formatTimestampCompact: (date: Date) => string;
      formatSaveStamp: (date: Date) => string;
      encodeUtf8: (value: string) => Uint8Array;
      packZipEntries: (entries: Array<{ name: string; data: Uint8Array }>) => Uint8Array;
    };
  }).__mikuprojectMainUtil;

  if (!mikuprojectMainUtil) {
    throw new Error("mikuproject main util module is not loaded");
  }

  const mikuprojectMainRender = (globalThis as typeof globalThis & {
    __mikuprojectMainRender?: {
      renderValidationIssues: (doc: Document, issues: ValidationIssue[]) => void;
      renderImportWarnings: (doc: Document, warnings: Array<{ message: string; scope?: "project" | "tasks" | "resources" | "assignments" | "calendars"; uid?: string; label?: string }>, options?: { sourceLabel?: string }) => void;
      renderXlsxImportSummary: (doc: Document, changes: Array<{
        scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
        uid: string;
        label: string;
        field: string;
        before: string | number | boolean | undefined;
        after: string | number | boolean;
      }>, options?: {
        sourceLabel?: string;
        warnings?: Array<{ message: string; scope?: "project" | "tasks" | "resources" | "assignments" | "calendars"; uid?: string; label?: string }>;
      }) => void;
      updateSummary: (doc: Document, model: ProjectModel | null, updateSvgButton: () => void) => void;
    };
  }).__mikuprojectMainRender;

  if (!mikuprojectMainRender) {
    throw new Error("mikuproject main render module is not loaded");
  }

  const mikuprojectMainIo = (globalThis as typeof globalThis & {
    __mikuprojectMainIo?: {
      buildOutputArchiveEntries: (deps: {
        model: ProjectModel;
        syncXmlTextFromModel: (model: ProjectModel) => string;
        formatTimestampCompact: (date: Date) => string;
        encodeUtf8: (value: string) => Uint8Array;
        createWorkbookCodec: () => {
          exportWorkbook: (workbook: unknown) => Uint8Array;
        };
        exportProjectWorkbook: (model: ProjectModel) => unknown;
        exportProjectWorkbookJson: (model: ProjectModel) => unknown;
        exportCsvParentId: (model: ProjectModel) => string;
        buildWbsOptions: (model: ProjectModel) => {
          holidayDates: string[];
          displayDaysBeforeBaseDate?: number;
          displayDaysAfterBaseDate?: number;
          useBusinessDaysForDisplayRange?: boolean;
          useBusinessDaysForProgressBand?: boolean;
        };
        exportWbsWorkbook: (
          model: ProjectModel,
          options: {
            holidayDates: string[];
            displayDaysBeforeBaseDate?: number;
            displayDaysAfterBaseDate?: number;
            useBusinessDaysForDisplayRange?: boolean;
            useBusinessDaysForProgressBand?: boolean;
          }
        ) => unknown;
        exportWbsMarkdown: (
          model: ProjectModel,
          options: {
            holidayDates: string[];
            displayDaysBeforeBaseDate?: number;
            displayDaysAfterBaseDate?: number;
            useBusinessDaysForDisplayRange?: boolean;
            useBusinessDaysForProgressBand?: boolean;
          }
        ) => string;
        exportNativeSvg: (
          model: ProjectModel,
          options: {
            holidayDates: string[];
            displayDaysBeforeBaseDate?: number;
            displayDaysAfterBaseDate?: number;
            useBusinessDaysForDisplayRange?: boolean;
            useBusinessDaysForProgressBand?: boolean;
          }
        ) => string;
        exportWeeklyNativeSvg: (
          model: ProjectModel,
          options: {
            holidayDates: string[];
            displayDaysBeforeBaseDate?: number;
            displayDaysAfterBaseDate?: number;
            useBusinessDaysForDisplayRange?: boolean;
            useBusinessDaysForProgressBand?: boolean;
          }
        ) => string;
        exportMonthlyWbsCalendarSvgArchive: (model: ProjectModel) => {
          entries: Array<{ fileName: string; svg: string }>;
        };
        exportMermaidGantt: (model: ProjectModel) => string;
        exportProjectOverviewView: (model: ProjectModel) => unknown;
        exportPhaseDetailView: (
          model: ProjectModel,
          phaseUid?: string,
          options?: {
            mode?: "full" | "scoped";
            rootUid?: string;
            maxDepth?: number;
          }
        ) => unknown;
        exportTaskEditView: (model: ProjectModel, taskUid?: string) => unknown;
      }) => Array<{ name: string; data: Uint8Array }>;
      detectImportKind: (input: {
        fileName: string;
        readJsonText?: string;
        extractLastJsonBlock: (value: string) => string;
        detectJsonDocumentKind: (documentLike: unknown) => "xml" | "xlsx" | "csv" | "editjson" | "workbook_json" | "project_draft_view" | "patch_json" | undefined;
      }) => "xml" | "xlsx" | "csv" | "editjson" | "workbook_json" | "project_draft_view" | "patch_json";
      assertRoundTripStable: (input: {
        currentModel: ProjectModel;
        exportMsProjectXml: (model: ProjectModel) => string;
        importMsProjectXml: (xmlText: string) => ProjectModel;
        validateProjectModel: (model: ProjectModel) => ValidationIssue[];
        normalizeProjectModel: (model: ProjectModel) => ProjectModel;
      }) => ValidationIssue[];
    };
  }).__mikuprojectMainIo;

  if (!mikuprojectMainIo) {
    throw new Error("mikuproject main IO module is not loaded");
  }

  const mikuprojectMainImport = (globalThis as typeof globalThis & {
    __mikuprojectMainImport?: {
      importProjectDraftText: (input: {
        sourceText: string;
        extractLastJsonBlock: (value: string) => string;
        importProjectDraftView: (draft: unknown) => ProjectModel;
        validateProjectModel: (model: ProjectModel) => ValidationIssue[];
      }) => { model: ProjectModel; issues: ValidationIssue[] };
      importPatchJsonText: (input: {
        sourceText: string;
        ensureCurrentModel: () => ProjectModel;
        extractLastJsonBlock: (value: string) => string;
        importProjectPatchJson: (documentLike: unknown, baseModel: ProjectModel) => {
          model: ProjectModel;
          changes: Array<{
            scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
            uid: string;
            label: string;
            field: string;
            before: string | number | boolean | undefined;
            after: string | number | boolean;
          }>;
          warnings: Array<{ message: string; scope?: "project" | "tasks" | "resources" | "assignments" | "calendars"; uid?: string; label?: string }>;
        };
        exportMsProjectXml: (model: ProjectModel) => string;
        validateProjectModel: (model: ProjectModel) => ValidationIssue[];
      }) => {
        model: ProjectModel;
        issues: ValidationIssue[];
        changes: Array<{
          scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
          uid: string;
          label: string;
          field: string;
          before: string | number | boolean | undefined;
          after: string | number | boolean;
        }>;
        warnings: Array<{ message: string; scope?: "project" | "tasks" | "resources" | "assignments" | "calendars"; uid?: string; label?: string }>;
        regeneratedXmlText?: string;
      };
      detectAiEditJsonKind: (input: {
        sourceText: string;
        extractLastJsonBlock: (value: string) => string;
        detectJsonDocumentKind: (documentLike: unknown) => "workbook_json" | "project_draft_view" | "patch_json" | undefined;
      }) => "project_draft_view" | "patch_json";
      importWorkbookJsonText: (input: {
        sourceText: string;
        previousModel: ProjectModel | null;
        extractLastJsonBlock: (value: string) => string;
        importProjectWorkbookJsonAsProjectModel: (documentLike: unknown) => {
          model: ProjectModel;
          warnings: Array<{ message: string; scope?: "project" | "tasks" | "resources" | "assignments" | "calendars"; uid?: string; label?: string }>;
        };
        importProjectWorkbookJson: (documentLike: unknown, baseModel: ProjectModel) => {
          model: ProjectModel;
          changes: Array<{
            scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
            uid: string;
            label: string;
            field: string;
            before: string | number | boolean | undefined;
            after: string | number | boolean;
          }>;
          warnings: Array<{ message: string; scope?: "project" | "tasks" | "resources" | "assignments" | "calendars"; uid?: string; label?: string }>;
        };
        validateProjectModel: (model: ProjectModel) => ValidationIssue[];
      }) => {
        model: ProjectModel;
        issues: ValidationIssue[];
        changes: Array<{
          scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
          uid: string;
          label: string;
          field: string;
          before: string | number | boolean | undefined;
          after: string | number | boolean;
        }>;
        warnings: Array<{ message: string; scope?: "project" | "tasks" | "resources" | "assignments" | "calendars"; uid?: string; label?: string }>;
      };
      importXlsxBytes: (input: {
        bytes: Uint8Array;
        previousModel: ProjectModel | null;
        createWorkbookCodec: () => {
          importWorkbook: (bytes: Uint8Array) => unknown;
          importWorkbookAsync?: (bytes: Uint8Array) => Promise<unknown>;
        };
        importProjectWorkbookAsProjectModel: (workbook: unknown) => ProjectModel;
        importProjectWorkbookDetailed: (workbook: unknown, baseModel: ProjectModel) => {
          model: ProjectModel;
          changes: Array<{
            scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
            uid: string;
            label: string;
            field: string;
            before: string | number | boolean | undefined;
            after: string | number | boolean;
          }>;
        };
        validateProjectModel: (model: ProjectModel) => ValidationIssue[];
      }) => Promise<{
        model: ProjectModel;
        issues: ValidationIssue[];
        changes: Array<{
          scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
          uid: string;
          label: string;
          field: string;
          before: string | number | boolean | undefined;
          after: string | number | boolean;
        }>;
      }>;
      importCsvText: (input: {
        csvText: string;
        importCsvParentId: (csvText: string) => ProjectModel;
        validateProjectModel: (model: ProjectModel) => ValidationIssue[];
      }) => { model: ProjectModel; issues: ValidationIssue[] };
    };
  }).__mikuprojectMainImport;

  if (!mikuprojectMainImport) {
    throw new Error("mikuproject main import module is not loaded");
  }

  const mikuprojectMainExport = (globalThis as typeof globalThis & {
    __mikuprojectMainExport?: {
      buildCsvExport: (input: {
        model: ProjectModel;
        exportCsvParentId: (model: ProjectModel) => string;
      }) => { fileName: string; text: string };
      buildProjectOverviewExport: (input: {
        model: ProjectModel;
        exportProjectOverviewView: (model: ProjectModel) => unknown;
      }) => { fileName: string; text: string };
      buildTaskEditExport: (input: {
        model: ProjectModel;
        requestedTaskUid?: string;
        exportTaskEditView: (model: ProjectModel, taskUid?: string) => unknown;
      }) => { fileName: string; text: string; resolvedTaskUid?: string };
      buildAiProjectionBundleExport: (input: {
        model: ProjectModel;
        exportProjectOverviewView: (model: ProjectModel) => unknown;
        exportPhaseDetailView: (
          model: ProjectModel,
          phaseUid?: string,
          options?: { mode?: "full" | "scoped"; rootUid?: string; maxDepth?: number }
        ) => unknown;
        exportTaskEditView: (model: ProjectModel, taskUid?: string) => unknown;
      }) => { fileName: string; text: string; phaseCount: number; taskCount: number };
      buildPhaseDetailExport: (input: {
        model: ProjectModel;
        mode: "full" | "scoped";
        requestedPhaseUid?: string;
        requestedRootUid?: string;
        requestedMaxDepth?: number;
        exportPhaseDetailView: (
          model: ProjectModel,
          phaseUid?: string,
          options?: { mode?: "full" | "scoped"; rootUid?: string; maxDepth?: number }
        ) => unknown;
      }) => {
        fileName: string;
        text: string;
        resolvedPhaseUid?: string;
        resolvedRootUid?: string;
        resolvedMaxDepth?: number;
        resolvedMode: "full" | "scoped";
      };
      buildXlsxExport: (input: {
        model: ProjectModel;
        createWorkbookCodec: () => { exportWorkbook: (workbook: unknown) => Uint8Array };
        exportProjectWorkbook: (model: ProjectModel) => unknown;
      }) => { fileName: string; bytes: Uint8Array };
      buildWorkbookJsonExport: (input: {
        model: ProjectModel;
        exportProjectWorkbookJson: (model: ProjectModel) => unknown;
      }) => { fileName: string; text: string };
      buildWbsXlsxExport: (input: {
        model: ProjectModel;
        options: {
          holidayDates: string[];
          displayDaysBeforeBaseDate?: number;
          displayDaysAfterBaseDate?: number;
          useBusinessDaysForDisplayRange?: boolean;
          useBusinessDaysForProgressBand?: boolean;
        };
        createWorkbookCodec: () => { exportWorkbook: (workbook: unknown) => Uint8Array };
        exportWbsWorkbook: (
          model: ProjectModel,
          options: {
            holidayDates: string[];
            displayDaysBeforeBaseDate?: number;
            displayDaysAfterBaseDate?: number;
            useBusinessDaysForDisplayRange?: boolean;
            useBusinessDaysForProgressBand?: boolean;
          }
        ) => unknown;
      }) => { fileName: string; bytes: Uint8Array };
      buildXmlExport: (input: { xmlText: string }) => { fileName: string; text: string };
      buildDailySvgExport: (input: { svg: string }) => { fileName: string; text: string };
      buildWeeklySvgExport: (input: { svg: string }) => { fileName: string; text: string };
      buildMonthlySvgZipExport: (input: { zipBytes: Uint8Array }) => { fileName: string; bytes: Uint8Array };
      buildMermaidExport: (input: { mermaidText: string }) => { fileName: string; text: string };
      buildWbsMarkdownExport: (input: { markdownText: string }) => { fileName: string; text: string };
    };
  }).__mikuprojectMainExport;

  if (!mikuprojectMainExport) {
    throw new Error("mikuproject main export module is not loaded");
  }

  const mikuprojectMainEvents = (globalThis as typeof globalThis & {
    __mikuprojectMainEvents?: {
      bind: (deps: {
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
      }) => void;
    };
  }).__mikuprojectMainEvents;

  if (!mikuprojectMainEvents) {
    throw new Error("mikuproject main events module is not loaded");
  }

  const mikuprojectMainPreview = (globalThis as typeof globalThis & {
    __mikuprojectMainPreview?: {
      createEmptyState: () => {
        currentNativeSvg: string;
        currentWeeklyPreviewSvg: string;
        currentMonthlyPreviewSvg: string;
        currentSvgPreviewMode: "daily" | "weekly" | "monthly";
      };
      renderPreviewMarkup: (state: {
        currentNativeSvg: string;
        currentWeeklyPreviewSvg: string;
        currentMonthlyPreviewSvg: string;
        currentSvgPreviewMode: "daily" | "weekly" | "monthly";
      }) => string;
      applyPreviewModeButtonClasses: (doc: Document, mode: "daily" | "weekly" | "monthly") => void;
      updateDownloadButtons: (doc: Document, hasModel: boolean) => void;
      buildRenderedState: (input: {
        model: ProjectModel | null;
        previousState: {
          currentNativeSvg: string;
          currentWeeklyPreviewSvg: string;
          currentMonthlyPreviewSvg: string;
          currentSvgPreviewMode: "daily" | "weekly" | "monthly";
        };
        buildWbsOptions: (model: ProjectModel) => {
          holidayDates: string[];
          displayDaysBeforeBaseDate?: number;
          displayDaysAfterBaseDate?: number;
          useBusinessDaysForDisplayRange?: boolean;
          useBusinessDaysForProgressBand?: boolean;
        };
        exportNativeSvg: (
          model: ProjectModel,
          options: {
            holidayDates: string[];
            displayDaysBeforeBaseDate?: number;
            displayDaysAfterBaseDate?: number;
            useBusinessDaysForDisplayRange?: boolean;
            useBusinessDaysForProgressBand?: boolean;
          }
        ) => string;
        exportWeeklyNativeSvg: (
          model: ProjectModel,
          options: {
            holidayDates: string[];
            displayDaysBeforeBaseDate?: number;
            displayDaysAfterBaseDate?: number;
            useBusinessDaysForDisplayRange?: boolean;
            useBusinessDaysForProgressBand?: boolean;
          }
        ) => string;
        exportMonthlyWbsCalendarSvgArchive: (model: ProjectModel) => {
          entries: Array<{ fileName: string; svg: string }>;
        };
      }) => {
        currentNativeSvg: string;
        currentWeeklyPreviewSvg: string;
        currentMonthlyPreviewSvg: string;
        currentSvgPreviewMode: "daily" | "weekly" | "monthly";
      };
      setMode: (
        state: {
          currentNativeSvg: string;
          currentWeeklyPreviewSvg: string;
          currentMonthlyPreviewSvg: string;
          currentSvgPreviewMode: "daily" | "weekly" | "monthly";
        },
        mode: "daily" | "weekly" | "monthly"
      ) => {
        currentNativeSvg: string;
        currentWeeklyPreviewSvg: string;
        currentMonthlyPreviewSvg: string;
        currentSvgPreviewMode: "daily" | "weekly" | "monthly";
      };
    };
  }).__mikuprojectMainPreview;

  if (!mikuprojectMainPreview) {
    throw new Error("mikuproject main preview module is not loaded");
  }

  const mikuprojectMainUi = (globalThis as typeof globalThis & {
    __mikuprojectMainUi?: {
      getTabButtons: (doc: Document) => HTMLButtonElement[];
      getTabPanels: (doc: Document) => HTMLElement[];
      setStatus: (doc: Document, message: string) => void;
      updateXmlSaveState: (doc: Document, input: { isDirty: boolean; lastSavedXmlStamp: string }) => void;
      setActiveTab: (doc: Document, tabId: "input" | "transform" | "output") => void;
      moveTabFocus: (doc: Document, currentButton: HTMLButtonElement, direction: -1 | 1) => HTMLButtonElement | null;
    };
  }).__mikuprojectMainUi;

  if (!mikuprojectMainUi) {
    throw new Error("mikuproject main UI module is not loaded");
  }

  const mikuprojectMainSupport = (globalThis as typeof globalThis & {
    __mikuprojectMainSupport?: {
      showToast: (doc: Document, message: string) => void;
      getAiPromptText: (doc: Document, getSpecText?: () => string) => string;
      copyTextToClipboard: (doc: Document, text: string) => Promise<void>;
      downloadBlob: (doc: Document, blob: Blob, filename: string) => void;
    };
  }).__mikuprojectMainSupport;

  if (!mikuprojectMainSupport) {
    throw new Error("mikuproject main support module is not loaded");
  }

  const mikuprojectMainDownloads = (globalThis as typeof globalThis & {
    __mikuprojectMainDownloads?: {
      downloadXml: (input: {
        xmlText: string;
        buildXmlExport: (input: { xmlText: string }) => { fileName: string; text: string };
        downloadBlob: (blob: Blob, filename: string) => void;
        markXmlSavedCurrent: () => void;
        completeOutput: (statusMessage: string, toastMessage: string) => void;
      }) => void;
      downloadDailySvg: (input: {
        svg: string;
        buildDailySvgExport: (input: { svg: string }) => { fileName: string; text: string };
        downloadBlob: (blob: Blob, filename: string) => void;
        completeOutput: (statusMessage: string, toastMessage: string) => void;
      }) => void;
      downloadWeeklySvg: (input: {
        svg: string;
        buildWeeklySvgExport: (input: { svg: string }) => { fileName: string; text: string };
        downloadBlob: (blob: Blob, filename: string) => void;
        completeOutput: (statusMessage: string, toastMessage: string) => void;
      }) => void;
      downloadMonthlySvgZip: (input: {
        zipBytes: Uint8Array;
        buildMonthlySvgZipExport: (input: { zipBytes: Uint8Array }) => { fileName: string; bytes: Uint8Array };
        downloadBlob: (blob: Blob, filename: string) => void;
        completeOutput: (statusMessage: string, toastMessage: string) => void;
      }) => void;
      downloadMermaid: (input: {
        mermaidText: string;
        buildMermaidExport: (input: { mermaidText: string }) => { fileName: string; text: string };
        downloadBlob: (blob: Blob, filename: string) => void;
        completeOutput: (statusMessage: string, toastMessage: string) => void;
      }) => void;
      downloadWbsMarkdown: (input: {
        markdownText: string;
        buildWbsMarkdownExport: (input: { markdownText: string }) => { fileName: string; text: string };
        downloadBlob: (blob: Blob, filename: string) => void;
        completeOutput: (statusMessage: string, toastMessage: string) => void;
      }) => void;
    };
  }).__mikuprojectMainDownloads;

  if (!mikuprojectMainDownloads) {
    throw new Error("mikuproject main downloads module is not loaded");
  }

  const mikuprojectMainOutputActions = (globalThis as typeof globalThis & {
    __mikuprojectMainOutputActions?: typeof globalThis.__mikuprojectMainOutputActions;
  }).__mikuprojectMainOutputActions;

  if (!mikuprojectMainOutputActions) {
    throw new Error("mikuproject main output actions module is not loaded");
  }

  const mikuprojectMainImportActions = (globalThis as typeof globalThis & {
    __mikuprojectMainImportActions?: {
      importProjectDraftText: (input: {
        sourceText: string;
        importProjectDraftText: typeof mikuprojectMainImport.importProjectDraftText;
        extractLastJsonBlock: (value: string) => string;
        importProjectDraftView: (draft: unknown) => ProjectModel;
        validateProjectModel: (model: ProjectModel) => ValidationIssue[];
        setCurrentModel: (model: ProjectModel) => void;
        syncXmlTextFromModel: (model: ProjectModel) => string;
        applyModelState: (input: { model: ProjectModel; issues: ValidationIssue[] }) => void;
        exportCurrentMermaid: (options?: { silent?: boolean }) => Promise<void>;
        completeTransform: (statusMessage: string, toastMessage: string) => void;
      }) => Promise<void>;
      importPatchJsonText: (input: {
        sourceText: string;
        importPatchJsonText: typeof mikuprojectMainImport.importPatchJsonText;
        ensureCurrentModel: () => ProjectModel;
        extractLastJsonBlock: (value: string) => string;
        importProjectPatchJson: typeof mikuprojectProjectPatchJson.importProjectPatchJson;
        exportMsProjectXml: (model: ProjectModel) => string;
        validateProjectModel: (model: ProjectModel) => ValidationIssue[];
        setCurrentModel: (model: ProjectModel) => void;
        applyModelState: (input: {
          model: ProjectModel;
          issues: ValidationIssue[];
          warnings: Array<{ message: string; scope?: "project" | "tasks" | "resources" | "assignments" | "calendars"; uid?: string; label?: string }>;
          changes: Array<{
            scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
            uid: string;
            label: string;
            field: string;
            before: string | number | boolean | undefined;
            after: string | number | boolean;
          }>;
        }) => void;
        writeXmlText: (xmlText: string) => void;
        markXmlDirty: () => void;
        setXmlSourceDirty: (dirty: boolean) => void;
        exportCurrentMermaid: (options?: { silent?: boolean }) => Promise<void>;
        completeTransform: (statusMessage: string, toastMessage: string) => void;
      }) => Promise<void>;
      importWorkbookJsonText: (input: {
        sourceText: string;
        importWorkbookJsonText: typeof mikuprojectMainImport.importWorkbookJsonText;
        previousModel: ProjectModel | null;
        extractLastJsonBlock: (value: string) => string;
        importProjectWorkbookJsonAsProjectModel: typeof mikuprojectProjectWorkbookJson.importProjectWorkbookJsonAsProjectModel;
        importProjectWorkbookJson: typeof mikuprojectProjectWorkbookJson.importProjectWorkbookJson;
        validateProjectModel: (model: ProjectModel) => ValidationIssue[];
        setCurrentModel: (model: ProjectModel) => void;
        syncXmlTextFromModel: (model: ProjectModel) => string;
        applyModelState: (input: {
          model: ProjectModel;
          issues: ValidationIssue[];
          warnings: Array<{ message: string; scope?: "project" | "tasks" | "resources" | "assignments" | "calendars"; uid?: string; label?: string }>;
          changes: Array<{
            scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
            uid: string;
            label: string;
            field: string;
            before: string | number | boolean | undefined;
            after: string | number | boolean;
          }>;
        }) => void;
        exportCurrentMermaid: (options?: { silent?: boolean }) => Promise<void>;
        completeTransform: (statusMessage: string, toastMessage: string) => void;
      }) => Promise<void>;
      importXlsxFile: (input: {
        file: File;
        importXlsxBytes: typeof mikuprojectMainImport.importXlsxBytes;
        previousModel: ProjectModel | null;
        createWorkbookCodec: () => { importWorkbook: (bytes: Uint8Array) => unknown; importWorkbookAsync?: (bytes: Uint8Array) => Promise<unknown> };
        importProjectWorkbookAsProjectModel: typeof mikuprojectProjectXlsx.importProjectWorkbookAsProjectModel;
        importProjectWorkbookDetailed: typeof mikuprojectProjectXlsx.importProjectWorkbookDetailed;
        validateProjectModel: (model: ProjectModel) => ValidationIssue[];
        setCurrentModel: (model: ProjectModel) => void;
        syncXmlTextFromModel: (model: ProjectModel) => string;
        updateSummary: (model: ProjectModel | null) => void;
        renderValidationIssues: (issues: ValidationIssue[]) => void;
        renderImportWarnings: (warnings: Array<{ message: string; scope?: "project" | "tasks" | "resources" | "assignments" | "calendars"; uid?: string; label?: string }>) => void;
        renderXlsxImportSummary: (changes: Array<{
          scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
          uid: string;
          label: string;
          field: string;
          before: string | number | boolean | undefined;
          after: string | number | boolean;
        }>, options?: { sourceLabel?: string }) => void;
        setStatus: (message: string) => void;
        showToast: (message: string) => void;
        setActiveTab: () => void;
        exportCurrentMermaid: (options?: { silent?: boolean }) => Promise<void>;
      }) => Promise<void>;
      importCsvFile: (input: {
        file: File;
        importCsvText: typeof mikuprojectMainImport.importCsvText;
        importCsvParentId: (csvText: string) => ProjectModel;
        validateProjectModel: (model: ProjectModel) => ValidationIssue[];
        setCurrentModel: (model: ProjectModel) => void;
        setXmlSourceDirty: (dirty: boolean) => void;
        updateSummary: (model: ProjectModel | null) => void;
        renderValidationIssues: (issues: ValidationIssue[]) => void;
        renderImportWarnings: (warnings: Array<{ message: string; scope?: "project" | "tasks" | "resources" | "assignments" | "calendars"; uid?: string; label?: string }>) => void;
        renderXlsxImportSummary: (changes: Array<never>) => void;
        setStatus: (message: string) => void;
        showToast: (message: string) => void;
        setActiveTab: () => void;
        exportCurrentMermaid: (options?: { silent?: boolean }) => Promise<void>;
      }) => Promise<void>;
    };
  }).__mikuprojectMainImportActions;

  if (!mikuprojectMainImportActions) {
    throw new Error("mikuproject main import actions module is not loaded");
  }

  const mikuprojectMainXmlActions = (globalThis as typeof globalThis & {
    __mikuprojectMainXmlActions?: {
      exportCurrentMermaid: (input: {
        currentModel: ProjectModel | null;
        exportMermaidGantt: (model: ProjectModel) => string;
        setMermaidText: (text: string) => void;
        renderSvgPreview: () => Promise<void>;
        setStatus: (message: string) => void;
        showToast: (message: string) => void;
        setActiveTab: () => void;
        silent?: boolean;
      }) => Promise<void>;
      loadSampleXml: (input: {
        loadSampleXml: typeof mikuprojectMainSamples.loadSampleXml;
        document: Document;
        readSampleXml: () => string;
        readSampleProjectDraftView: () => unknown;
        writeXmlText: (xmlText: string) => void;
        writeProjectDraftText: (text: string) => void;
        importSampleXml: (xmlText: string) => ProjectModel;
        validateProjectModel: (model: ProjectModel) => ValidationIssue[];
        markXmlDirty: () => void;
        applyModelState: (input: { model: ProjectModel; issues: ValidationIssue[] }) => void;
        updateSvgButton: () => void;
        setStatus: (message: string) => void;
        setActiveTab: (tabId: "input" | "transform" | "output") => void;
        setCurrentModel: (model: ProjectModel) => void;
        setXmlSourceDirty: (dirty: boolean) => void;
      }) => void;
      importXmlFromFile: (input: {
        file: File;
        writeXmlText: (xmlText: string) => void;
        markXmlDirty: () => void;
        importMsProjectXml: (xmlText: string) => ProjectModel;
        validateProjectModel: (model: ProjectModel) => ValidationIssue[];
        setCurrentModel: (model: ProjectModel) => void;
        setXmlSourceDirty: (dirty: boolean) => void;
        applyModelState: (input: { model: ProjectModel; issues: ValidationIssue[] }) => void;
        completeTransform: (statusMessage: string, toastMessage: string) => void;
        exportCurrentMermaid: (options?: { silent?: boolean }) => Promise<void>;
      }) => Promise<void>;
      runRoundTripCheck: (input: {
        currentModel: ProjectModel | null;
        parseCurrentXml: () => void;
        getCurrentModel: () => ProjectModel | null;
        assertRoundTripStable: typeof mikuprojectMainIo.assertRoundTripStable;
        exportMsProjectXml: (model: ProjectModel) => string;
        importMsProjectXml: (xmlText: string) => ProjectModel;
        validateProjectModel: (model: ProjectModel) => ValidationIssue[];
        normalizeProjectModel: (model: ProjectModel) => ProjectModel;
        renderValidationIssues: (issues: ValidationIssue[]) => void;
        setStatus: (message: string) => void;
        showToast: (message: string) => void;
        setActiveTab: () => void;
      }) => void;
    };
  }).__mikuprojectMainXmlActions;

  if (!mikuprojectMainXmlActions) {
    throw new Error("mikuproject main XML actions module is not loaded");
  }

  const mikuprojectMainArchiveActions = (globalThis as typeof globalThis & {
    __mikuprojectMainArchiveActions?: {
      buildOutputArchive: (input: {
        buildOutputArchiveEntries: () => Array<{ name: string; data: Uint8Array }>;
        formatTimestampCompact: (date: Date) => string;
        packZipEntries: (entries: Array<{ name: string; data: Uint8Array }>) => Uint8Array;
      }) => { fileName: string; zipBytes: Uint8Array; entryCount: number };
      downloadAllOutputs: (input: {
        buildOutputArchive: () => { fileName: string; zipBytes: Uint8Array; entryCount: number };
        downloadBlob: (blob: Blob, filename: string) => void;
        setStatus: (message: string) => void;
        showToast: (message: string) => void;
        setActiveTab: () => void;
      }) => void;
    };
  }).__mikuprojectMainArchiveActions;

  if (!mikuprojectMainArchiveActions) {
    throw new Error("mikuproject main archive actions module is not loaded");
  }

  const mikuprojectMainSaveState = (globalThis as typeof globalThis & {
    __mikuprojectMainSaveState?: {
      updateXmlSaveState: (input: {
        document: Document;
        isDirty: boolean;
        lastSavedXmlStamp: string;
        updateXmlSaveStateView: (doc: Document, input: { isDirty: boolean; lastSavedXmlStamp: string }) => void;
      }) => void;
      markXmlDirty: (input: {
        document: Document;
        lastSavedXmlStamp: string;
        updateXmlSaveStateView: (doc: Document, input: { isDirty: boolean; lastSavedXmlStamp: string }) => void;
      }) => void;
      markXmlSavedCurrent: (input: {
        readXmlText: () => string;
        formatSaveStamp: (date: Date) => string;
        writeSavedState: (state: { lastSavedXmlText: string; lastSavedXmlStamp: string }) => void;
        document: Document;
        updateXmlSaveStateView: (doc: Document, input: { isDirty: boolean; lastSavedXmlStamp: string }) => void;
      }) => void;
      refreshXmlSaveState: (input: {
        readXmlText: () => string;
        lastSavedXmlText: string;
        lastSavedXmlStamp: string;
        document: Document;
        updateXmlSaveStateView: (doc: Document, input: { isDirty: boolean; lastSavedXmlStamp: string }) => void;
      }) => void;
    };
  }).__mikuprojectMainSaveState;

  if (!mikuprojectMainSaveState) {
    throw new Error("mikuproject main save state module is not loaded");
  }

  const mikuprojectMainTabActions = (globalThis as typeof globalThis & {
    __mikuprojectMainTabActions?: {
      setActiveTab: (input: {
        document: Document;
        tabId: "input" | "transform" | "output";
        skipTransformRefresh?: boolean;
        isRefreshingTransformTab: boolean;
        setCurrentTabId: (tabId: "input" | "transform" | "output") => void;
        setActiveTabView: (doc: Document, tabId: "input" | "transform" | "output") => void;
        refreshTransformTab: () => Promise<void>;
        setStatus: (message: string) => void;
      }) => void;
      refreshTransformTab: (input: {
        isRefreshingTransformTab: boolean;
        setRefreshingTransformTab: (value: boolean) => void;
        refreshTransformTabImpl: typeof mikuprojectMainTransform.refreshTransformTab;
        currentModel: ProjectModel | null;
        isXmlSourceDirty: boolean;
        readXmlText: () => string;
        parseCurrentXml: (options?: { silent?: boolean }) => void;
        exportCurrentMermaid: (options?: { silent?: boolean }) => Promise<void>;
        setStatus: (message: string) => void;
      }) => Promise<void>;
      moveTabFocus: (input: {
        document: Document;
        currentButton: HTMLButtonElement;
        direction: -1 | 1;
        moveTabFocusView: typeof mikuprojectMainUi.moveTabFocus;
        setActiveTab: (tabId: "input" | "transform" | "output") => void;
      }) => void;
      bindTabs: (input: {
        document: Document;
        currentTabId: "input" | "transform" | "output";
        getTabButtons: () => HTMLButtonElement[];
        bindTabsView: typeof mikuprojectMainTransform.bindTabs;
        setActiveTab: (tabId: "input" | "transform" | "output") => void;
        moveTabFocus: (currentButton: HTMLButtonElement, direction: -1 | 1) => void;
      }) => void;
    };
  }).__mikuprojectMainTabActions;

  if (!mikuprojectMainTabActions) {
    throw new Error("mikuproject main tab actions module is not loaded");
  }

  const mikuprojectMainPreviewActions = (globalThis as typeof globalThis & {
    __mikuprojectMainPreviewActions?: {
      updateSvgPreviewModeButtons: (input: {
        document: Document;
        mode: "daily" | "weekly" | "monthly";
        applyPreviewModeButtonClasses: typeof mikuprojectMainPreview.applyPreviewModeButtonClasses;
      }) => void;
      renderCurrentSvgPreviewMarkup: (input: {
        state: typeof currentSvgPreviewState.state;
        renderPreviewMarkup: typeof mikuprojectMainPreview.renderPreviewMarkup;
        setSvgPreviewMarkup: (markup: string) => void;
      }) => void;
      updateSvgButton: (input: {
        document: Document;
        hasModel: boolean;
        updateDownloadButtons: typeof mikuprojectMainPreview.updateDownloadButtons;
      }) => void;
      renderSvgPreview: (input: {
        currentModel: ProjectModel | null;
        currentState: typeof currentSvgPreviewState.state;
        setState: (state: typeof currentSvgPreviewState.state) => void;
        buildRenderedState: typeof mikuprojectMainPreview.buildRenderedState;
        buildWbsOptions: (model: ProjectModel) => {
          holidayDates: string[];
          displayDaysBeforeBaseDate?: number;
          displayDaysAfterBaseDate?: number;
          useBusinessDaysForDisplayRange?: boolean;
          useBusinessDaysForProgressBand?: boolean;
        };
        exportNativeSvg: typeof mikuprojectNativeSvg.exportNativeSvg;
        exportWeeklyNativeSvg: typeof mikuprojectNativeSvg.exportWeeklyNativeSvg;
        exportMonthlyWbsCalendarSvgArchive: typeof mikuprojectNativeSvg.exportMonthlyWbsCalendarSvgArchive;
        renderCurrentSvgPreviewMarkup: () => void;
        updateSvgButton: () => void;
      }) => Promise<void>;
      setSvgPreviewMode: (input: {
        currentState: typeof currentSvgPreviewState.state;
        mode: "daily" | "weekly" | "monthly";
        setMode: typeof mikuprojectMainPreview.setMode;
        setState: (state: typeof currentSvgPreviewState.state) => void;
        updateSvgPreviewModeButtons: () => void;
        renderCurrentSvgPreviewMarkup: () => void;
      }) => void;
    };
  }).__mikuprojectMainPreviewActions;

  if (!mikuprojectMainPreviewActions) {
    throw new Error("mikuproject main preview actions module is not loaded");
  }

  const mikuprojectMainTransform = (globalThis as typeof globalThis & {
    __mikuprojectMainTransform?: {
      bindTabs: (input: {
        doc: Document;
        currentTabId: "input" | "transform" | "output";
        getTabButtons: () => HTMLButtonElement[];
        setActiveTab: (tabId: "input" | "transform" | "output") => void;
        moveTabFocus: (currentButton: HTMLButtonElement, direction: -1 | 1) => void;
      }) => void;
      buildWbsOptions: (input: {
        doc: Document;
        model: ProjectModel;
        parseOptionalNonNegativeInteger: (raw: string) => number | undefined;
        collectWbsHolidayDates: (model: ProjectModel) => string[];
      }) => {
        holidayDates: string[];
        displayDaysBeforeBaseDate?: number;
        displayDaysAfterBaseDate?: number;
        useBusinessDaysForDisplayRange?: boolean;
        useBusinessDaysForProgressBand?: boolean;
      };
      refreshTransformTab: (input: {
        currentModel: ProjectModel | null;
        isXmlSourceDirty: boolean;
        readXmlText: () => string;
        parseCurrentXml: (options?: { silent?: boolean }) => void;
        exportCurrentMermaid: (options?: { silent?: boolean }) => Promise<void>;
        setStatus: (message: string) => void;
      }) => Promise<void>;
    };
  }).__mikuprojectMainTransform;

  if (!mikuprojectMainTransform) {
    throw new Error("mikuproject main transform module is not loaded");
  }

  const mikuprojectMainFlow = (globalThis as typeof globalThis & {
    __mikuprojectMainFlow?: {
      applyModelState: (input: {
        model: ProjectModel;
        issues: ValidationIssue[];
        updateSummary: (model: ProjectModel | null) => void;
        renderValidationIssues: (issues: ValidationIssue[]) => void;
        renderImportWarnings: (warnings: Array<{ message: string; scope?: "project" | "tasks" | "resources" | "assignments" | "calendars"; uid?: string; label?: string }>, options?: { sourceLabel?: string }) => void;
        renderXlsxImportSummary: (changes: Array<{
          scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
          uid: string;
          label: string;
          field: string;
          before: string | number | boolean | undefined;
          after: string | number | boolean;
        }>, options?: {
          sourceLabel?: string;
          warnings?: Array<{ message: string; scope?: "project" | "tasks" | "resources" | "assignments" | "calendars"; uid?: string; label?: string }>;
        }) => void;
        warnings?: Array<{ message: string; scope?: "project" | "tasks" | "resources" | "assignments" | "calendars"; uid?: string; label?: string }>;
        warningSourceLabel?: string;
        changes?: Array<{
          scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
          uid: string;
          label: string;
          field: string;
          before: string | number | boolean | undefined;
          after: string | number | boolean;
        }>;
        changeSourceLabel?: string;
      }) => void;
      completeOutput: (input: {
        setStatus: (message: string) => void;
        showToast: (message: string) => void;
        setActiveTab: (tabId: "input" | "transform" | "output", options?: { skipTransformRefresh?: boolean }) => void;
        statusMessage: string;
        toastMessage: string;
      }) => void;
      completeTransform: (input: {
        setStatus: (message: string) => void;
        showToast: (message: string) => void;
        setActiveTab: (tabId: "input" | "transform" | "output", options?: { skipTransformRefresh?: boolean }) => void;
        statusMessage: string;
        toastMessage: string;
      }) => void;
    };
  }).__mikuprojectMainFlow;

  if (!mikuprojectMainFlow) {
    throw new Error("mikuproject main flow module is not loaded");
  }

  const mikuprojectMainModel = (globalThis as typeof globalThis & {
    __mikuprojectMainModel?: {
      syncXmlTextFromModel: (input: {
        model: ProjectModel;
        exportMsProjectXml: (model: ProjectModel) => string;
        writeXmlText: (xmlText: string) => void;
        clearXmlSourceDirty: () => void;
        refreshXmlSaveState: () => void;
      }) => string;
      ensureCurrentModel: (input: {
        currentModel: ProjectModel | null;
        readXmlText: () => string;
        importMsProjectXml: (xmlText: string) => ProjectModel;
        clearXmlSourceDirty: () => void;
      }) => ProjectModel;
      parseCurrentXml: (input: {
        readXmlText: () => string;
        importMsProjectXml: (xmlText: string) => ProjectModel;
        validateProjectModel: (model: ProjectModel) => ValidationIssue[];
        clearXmlSourceDirty: () => void;
      }) => { model: ProjectModel | null; issues: ValidationIssue[]; empty: boolean };
    };
  }).__mikuprojectMainModel;

  if (!mikuprojectMainModel) {
    throw new Error("mikuproject main model module is not loaded");
  }

  const mikuprojectExcelIo = (globalThis as typeof globalThis & {
    __mikuprojectExcelIo?: {
      XlsxWorkbookCodec: new () => {
        exportWorkbook: (workbook: unknown) => Uint8Array;
        importWorkbook: (bytes: Uint8Array) => unknown;
        importWorkbookAsync?: (bytes: Uint8Array) => Promise<unknown>;
      };
    };
  }).__mikuprojectExcelIo;

  if (!mikuprojectExcelIo) {
    throw new Error("mikuproject Excel IO module is not loaded");
  }

  const mikuprojectProjectXlsx = (globalThis as typeof globalThis & {
    __mikuprojectProjectXlsx?: {
      exportProjectWorkbook: (model: ProjectModel) => unknown;
      importProjectWorkbook: (workbook: unknown, baseModel: ProjectModel) => ProjectModel;
      importProjectWorkbookAsProjectModel: (workbook: unknown) => ProjectModel;
      importProjectWorkbookDetailed: (workbook: unknown, baseModel: ProjectModel) => {
        model: ProjectModel;
        changes: Array<{
          scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
          uid: string;
          label: string;
          field: string;
          before: string | number | boolean | undefined;
          after: string | number | boolean;
        }>;
      };
    };
  }).__mikuprojectProjectXlsx;

  if (!mikuprojectProjectXlsx) {
    throw new Error("mikuproject Project XLSX module is not loaded");
  }

  const mikuprojectProjectWorkbookJson = (globalThis as typeof globalThis & {
    __mikuprojectProjectWorkbookJson?: {
      exportProjectWorkbookJson: (model: ProjectModel) => unknown;
      importProjectWorkbookJsonAsProjectModel: (documentLike: unknown) => {
        model: ProjectModel;
        warnings: Array<{ message: string }>;
      };
      importProjectWorkbookJson: (documentLike: unknown, baseModel: ProjectModel) => {
        model: ProjectModel;
        changes: Array<{
          scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
          uid: string;
          label: string;
          field: string;
          before: string | number | boolean | undefined;
          after: string | number | boolean;
        }>;
        warnings: Array<{
          message: string;
        }>;
      };
      validateWorkbookJsonDocument: (documentLike: unknown) => {
        document: unknown;
        warnings: Array<{
          message: string;
        }>;
      };
    };
  }).__mikuprojectProjectWorkbookJson;

  if (!mikuprojectProjectWorkbookJson) {
    throw new Error("mikuproject Project Workbook JSON module is not loaded");
  }

  const mikuprojectProjectPatchJson = (globalThis as typeof globalThis & {
    __mikuprojectProjectPatchJson?: {
      importProjectPatchJson: (documentLike: unknown, baseModel: ProjectModel) => {
        model: ProjectModel;
        changes: Array<{
          scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
          uid: string;
          label: string;
          field: string;
          before: string | number | boolean | undefined;
          after: string | number | boolean;
        }>;
        warnings: Array<{
          message: string;
        }>;
      };
      validatePatchDocument: (documentLike: unknown) => {
        document: unknown;
        warnings: Array<{
          message: string;
        }>;
      };
    };
  }).__mikuprojectProjectPatchJson;

  if (!mikuprojectProjectPatchJson) {
    throw new Error("mikuproject Project Patch JSON module is not loaded");
  }

  const mikuprojectWbsXlsx = (globalThis as typeof globalThis & {
    __mikuprojectWbsXlsx?: {
      collectWbsHolidayDates: (model: ProjectModel) => string[];
      exportWbsWorkbook: (
        model: ProjectModel,
        options?: {
          holidayDates?: string[];
          displayDaysBeforeBaseDate?: number;
          displayDaysAfterBaseDate?: number;
          useBusinessDaysForDisplayRange?: boolean;
          useBusinessDaysForProgressBand?: boolean;
        }
      ) => unknown;
    };
  }).__mikuprojectWbsXlsx;

  if (!mikuprojectWbsXlsx) {
    throw new Error("mikuproject WBS XLSX module is not loaded");
  }

  const mikuprojectWbsMarkdown = (globalThis as typeof globalThis & {
    __mikuprojectWbsMarkdown?: {
      exportWbsMarkdown: (
        model: ProjectModel,
        options?: {
          holidayDates?: string[];
          displayDaysBeforeBaseDate?: number;
          displayDaysAfterBaseDate?: number;
          useBusinessDaysForDisplayRange?: boolean;
          useBusinessDaysForProgressBand?: boolean;
        }
      ) => string;
    };
  }).__mikuprojectWbsMarkdown;

  if (!mikuprojectWbsMarkdown) {
    throw new Error("mikuproject WBS Markdown module is not loaded");
  }

  const mikuprojectNativeSvg = (globalThis as typeof globalThis & {
    __mikuprojectNativeSvg?: {
      exportNativeSvg: (
        model: ProjectModel,
        options?: {
          holidayDates?: string[];
          displayDaysBeforeBaseDate?: number;
          displayDaysAfterBaseDate?: number;
          useBusinessDaysForDisplayRange?: boolean;
          useBusinessDaysForProgressBand?: boolean;
        }
      ) => string;
      exportWeeklyNativeSvg: (
        model: ProjectModel,
        options?: {
          holidayDates?: string[];
          displayDaysBeforeBaseDate?: number;
          displayDaysAfterBaseDate?: number;
          useBusinessDaysForDisplayRange?: boolean;
          useBusinessDaysForProgressBand?: boolean;
        }
      ) => string;
      exportMonthlyWbsCalendarSvgArchive: (
        model: ProjectModel
      ) => {
        entries: Array<{
          fileName: string;
          svg: string;
        }>;
        zipBytes: Uint8Array;
      };
    };
  }).__mikuprojectNativeSvg;

  if (!mikuprojectNativeSvg) {
    throw new Error("mikuproject native SVG module is not loaded");
  }

  let currentModel: ProjectModel | null = null;
  const currentSvgPreviewState = {
    state: mikuprojectMainPreview.createEmptyState()
  };
  let lastSavedXmlText = "";
  let lastSavedXmlStamp = "";
  let currentTabId: "input" | "transform" | "output" = "input";
  let isXmlSourceDirty = true;
  let isRefreshingTransformTab = false;

  function getElement<T extends HTMLElement>(id: string): T {
    const element = document.getElementById(id);
    if (!element) {
      throw new Error(`Element not found: ${id}`);
    }
    return element as T;
  }

  function getTextArea(id: string): HTMLTextAreaElement {
    return getElement<HTMLTextAreaElement>(id);
  }

  function getInput(id: string): HTMLInputElement {
    return getElement<HTMLInputElement>(id);
  }

  function getTabButtons(): HTMLButtonElement[] {
    return mikuprojectMainUi.getTabButtons(document);
  }

  function getTabPanels(): HTMLElement[] {
    return mikuprojectMainUi.getTabPanels(document);
  }

  function setActiveTab(
    tabId: "input" | "transform" | "output",
    options: { skipTransformRefresh?: boolean } = {}
  ): void {
    mikuprojectMainTabActions.setActiveTab({
      document,
      tabId,
      skipTransformRefresh: options.skipTransformRefresh,
      isRefreshingTransformTab,
      setCurrentTabId: (nextTabId) => {
        currentTabId = nextTabId;
      },
      setActiveTabView: mikuprojectMainUi.setActiveTab,
      refreshTransformTab,
      setStatus
    });
  }

  async function refreshTransformTab(): Promise<void> {
    await mikuprojectMainTabActions.refreshTransformTab({
      isRefreshingTransformTab,
      setRefreshingTransformTab: (value) => {
        isRefreshingTransformTab = value;
      },
      refreshTransformTabImpl: mikuprojectMainTransform.refreshTransformTab,
      currentModel,
      isXmlSourceDirty,
      readXmlText: () => getTextArea("xmlInput").value,
      parseCurrentXml,
      exportCurrentMermaid,
      setStatus
    });
  }

  function moveTabFocus(currentButton: HTMLButtonElement, direction: -1 | 1): void {
    mikuprojectMainTabActions.moveTabFocus({
      document,
      currentButton,
      direction,
      moveTabFocusView: mikuprojectMainUi.moveTabFocus,
      setActiveTab
    });
  }

  function bindTabs(): void {
    mikuprojectMainTabActions.bindTabs({
      document,
      currentTabId,
      getTabButtons,
      bindTabsView: mikuprojectMainTransform.bindTabs,
      setActiveTab,
      moveTabFocus
    });
  }

  function parseOptionalNonNegativeInteger(raw: string): number | undefined {
    return mikuprojectMainUtil.parseOptionalNonNegativeInteger(raw);
  }

  function showToast(message: string): void {
    mikuprojectMainSupport.showToast(document, message);
  }

  function getAiPromptText(): string {
    return mikuprojectMainSupport.getAiPromptText(
      document,
      mikuprojectAiJsonSpec && typeof mikuprojectAiJsonSpec.getAiJsonSpecText === "function"
        ? mikuprojectAiJsonSpec.getAiJsonSpecText
        : undefined
    );
  }

  async function copyTextToClipboard(text: string): Promise<void> {
    await mikuprojectMainSupport.copyTextToClipboard(document, text);
  }

  const mikuprojectMainSamples = (globalThis as typeof globalThis & {
    __mikuprojectMainSamples?: {
      copyAiPrompt: (deps: {
        document: Document;
        getAiPromptText: () => string;
        copyTextToClipboard: (text: string) => Promise<void>;
        showToast: (message: string) => void;
        setStatus: (message: string) => void;
      }) => Promise<void>;
      loadSampleXml: (deps: {
        document: Document;
        readSampleXml: () => string;
        readSampleProjectDraftView: () => unknown;
        writeXmlText: (xmlText: string) => void;
        writeProjectDraftText: (text: string) => void;
        importSampleXml: (xmlText: string) => ProjectModel;
        validateProjectModel: (model: ProjectModel) => ValidationIssue[];
        markXmlDirty: () => void;
        applyModelState: (input: { model: ProjectModel; issues: ValidationIssue[] }) => void;
        updateSvgButton: () => void;
        setStatus: (message: string) => void;
        setActiveTab: (tabId: "input" | "transform" | "output") => void;
      }) => ProjectModel;
      loadProjectDraftSample: (deps: {
        document: Document;
        readSampleXml: () => string;
        readSampleProjectDraftView: () => unknown;
        writeXmlText: (xmlText: string) => void;
        writeProjectDraftText: (text: string) => void;
        importSampleXml: (xmlText: string) => ProjectModel;
        validateProjectModel: (model: ProjectModel) => ValidationIssue[];
        markXmlDirty: () => void;
        applyModelState: (input: { model: ProjectModel; issues: ValidationIssue[] }) => void;
        updateSvgButton: () => void;
        setStatus: (message: string) => void;
        setActiveTab: (tabId: "input" | "transform" | "output") => void;
      }) => void;
    };
  }).__mikuprojectMainSamples;

  if (!mikuprojectMainSamples) {
    throw new Error("mikuproject main samples module is not loaded");
  }

  async function copyAiPrompt(): Promise<void> {
    await mikuprojectMainSamples.copyAiPrompt({
      document,
      getAiPromptText,
      copyTextToClipboard,
      showToast,
      setStatus
    });
  }

  function setSvgPreviewMarkup(markup: string): void {
    getElement<HTMLElement>("nativeSvgPreview").innerHTML = markup;
  }

  function updateSvgPreviewModeButtons(): void {
    mikuprojectMainPreviewActions.updateSvgPreviewModeButtons({
      document,
      mode: currentSvgPreviewState.state.currentSvgPreviewMode,
      applyPreviewModeButtonClasses: mikuprojectMainPreview.applyPreviewModeButtonClasses
    });
  }

  function renderCurrentSvgPreviewMarkup(): void {
    mikuprojectMainPreviewActions.renderCurrentSvgPreviewMarkup({
      state: currentSvgPreviewState.state,
      renderPreviewMarkup: mikuprojectMainPreview.renderPreviewMarkup,
      setSvgPreviewMarkup
    });
  }

  function updateSvgButton(): void {
    mikuprojectMainPreviewActions.updateSvgButton({
      document,
      hasModel: Boolean(currentModel),
      updateDownloadButtons: mikuprojectMainPreview.updateDownloadButtons
    });
  }

  function buildCurrentWbsOptions(model: ProjectModel): {
    holidayDates: string[];
    displayDaysBeforeBaseDate?: number;
    displayDaysAfterBaseDate?: number;
    useBusinessDaysForDisplayRange?: boolean;
    useBusinessDaysForProgressBand?: boolean;
  } {
    return mikuprojectMainTransform.buildWbsOptions({
      doc: document,
      model,
      parseOptionalNonNegativeInteger: parseOptionalNonNegativeInteger,
      collectWbsHolidayDates: mikuprojectWbsXlsx.collectWbsHolidayDates
    });
  }

  function downloadBlob(blob: Blob, filename: string): void {
    mikuprojectMainSupport.downloadBlob(document, blob, filename);
  }

  function formatTimestampCompact(date: Date): string {
    return mikuprojectMainUtil.formatTimestampCompact(date);
  }

  async function renderSvgPreview(): Promise<void> {
    await mikuprojectMainPreviewActions.renderSvgPreview({
      currentModel,
      currentState: currentSvgPreviewState.state,
      setState: (state) => {
        currentSvgPreviewState.state = state;
      },
      buildRenderedState: mikuprojectMainPreview.buildRenderedState,
      buildWbsOptions: buildCurrentWbsOptions,
      exportNativeSvg: mikuprojectNativeSvg.exportNativeSvg,
      exportWeeklyNativeSvg: mikuprojectNativeSvg.exportWeeklyNativeSvg,
      exportMonthlyWbsCalendarSvgArchive: mikuprojectNativeSvg.exportMonthlyWbsCalendarSvgArchive,
      renderCurrentSvgPreviewMarkup,
      updateSvgButton
    });
  }

  function setSvgPreviewMode(mode: "daily" | "weekly" | "monthly"): void {
    mikuprojectMainPreviewActions.setSvgPreviewMode({
      currentState: currentSvgPreviewState.state,
      mode,
      setMode: mikuprojectMainPreview.setMode,
      setState: (state) => {
        currentSvgPreviewState.state = state;
      },
      updateSvgPreviewModeButtons,
      renderCurrentSvgPreviewMarkup
    });
  }

  function buildCurrentOutputArchiveEntries(): Array<{ name: string; data: Uint8Array }> {
    return mikuprojectMainIo.buildOutputArchiveEntries({
      model: ensureCurrentModel(),
      syncXmlTextFromModel,
      formatTimestampCompact,
      encodeUtf8: mikuprojectMainUtil.encodeUtf8,
      createWorkbookCodec: () => new mikuprojectExcelIo.XlsxWorkbookCodec(),
      exportProjectWorkbook: mikuprojectProjectXlsx.exportProjectWorkbook,
      exportProjectWorkbookJson: mikuprojectProjectWorkbookJson.exportProjectWorkbookJson,
      exportCsvParentId: mikuprojectXml.exportCsvParentId,
      buildWbsOptions: buildCurrentWbsOptions,
      exportWbsWorkbook: mikuprojectWbsXlsx.exportWbsWorkbook,
      exportWbsMarkdown: mikuprojectWbsMarkdown.exportWbsMarkdown,
      exportNativeSvg: mikuprojectNativeSvg.exportNativeSvg,
      exportWeeklyNativeSvg: mikuprojectNativeSvg.exportWeeklyNativeSvg,
      exportMonthlyWbsCalendarSvgArchive: mikuprojectNativeSvg.exportMonthlyWbsCalendarSvgArchive,
      exportMermaidGantt: mikuprojectXml.exportMermaidGantt,
      exportProjectOverviewView: mikuprojectXml.exportProjectOverviewView,
      exportPhaseDetailView: mikuprojectXml.exportPhaseDetailView,
      exportTaskEditView: mikuprojectXml.exportTaskEditView
    });
  }

  function buildCurrentOutputArchive(): { fileName: string; zipBytes: Uint8Array; entryCount: number } {
    return mikuprojectMainArchiveActions.buildOutputArchive({
      buildOutputArchiveEntries: buildCurrentOutputArchiveEntries,
      formatTimestampCompact,
      packZipEntries: mikuprojectMainUtil.packZipEntries
    });
  }

  function downloadAllOutputs(): void {
    mikuprojectMainArchiveActions.downloadAllOutputs({
      buildOutputArchive: buildCurrentOutputArchive,
      downloadBlob,
      setStatus,
      showToast,
      setActiveTab: () => {
        setActiveTab("output");
      }
    });
  }

  function setStatus(message: string): void {
    mikuprojectMainUi.setStatus(document, message);
  }

  function formatSaveStamp(date: Date): string {
    return mikuprojectMainUtil.formatSaveStamp(date);
  }

  function updateXmlSaveState(isDirty: boolean): void {
    mikuprojectMainSaveState.updateXmlSaveState({
      document,
      isDirty,
      lastSavedXmlStamp,
      updateXmlSaveStateView: mikuprojectMainUi.updateXmlSaveState
    });
  }

  function markXmlDirty(): void {
    mikuprojectMainSaveState.markXmlDirty({
      document,
      lastSavedXmlStamp,
      updateXmlSaveStateView: mikuprojectMainUi.updateXmlSaveState
    });
  }

  function markXmlSavedCurrent(): void {
    mikuprojectMainSaveState.markXmlSavedCurrent({
      readXmlText: () => getTextArea("xmlInput").value,
      formatSaveStamp,
      writeSavedState: ({ lastSavedXmlText: nextSavedXmlText, lastSavedXmlStamp: nextSavedXmlStamp }) => {
        lastSavedXmlText = nextSavedXmlText;
        lastSavedXmlStamp = nextSavedXmlStamp;
      },
      document,
      updateXmlSaveStateView: mikuprojectMainUi.updateXmlSaveState
    });
  }

  function refreshXmlSaveState(): void {
    mikuprojectMainSaveState.refreshXmlSaveState({
      readXmlText: () => getTextArea("xmlInput").value,
      lastSavedXmlText,
      lastSavedXmlStamp,
      document,
      updateXmlSaveStateView: mikuprojectMainUi.updateXmlSaveState
    });
  }

  function syncXmlTextFromModel(model: ProjectModel): string {
    return mikuprojectMainModel.syncXmlTextFromModel({
      model,
      exportMsProjectXml: mikuprojectXml.exportMsProjectXml,
      writeXmlText: (xmlText) => {
        getTextArea("xmlInput").value = xmlText;
      },
      clearXmlSourceDirty: () => {
        isXmlSourceDirty = false;
      },
      refreshXmlSaveState
    });
  }

  function renderValidationIssues(issues: ValidationIssue[]): void {
    mikuprojectMainRender.renderValidationIssues(document, issues);
  }

  function renderImportWarnings(
    warnings: Array<{ message: string; scope?: "project" | "tasks" | "resources" | "assignments" | "calendars"; uid?: string; label?: string }>,
    options: { sourceLabel?: string } = {}
  ): void {
    mikuprojectMainRender.renderImportWarnings(document, warnings, options);
  }

  function renderXlsxImportSummary(changes: Array<{
    scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
    uid: string;
    label: string;
    field: string;
    before: string | number | boolean | undefined;
    after: string | number | boolean;
  }>, options: {
    sourceLabel?: string;
    warnings?: Array<{ message: string; scope?: "project" | "tasks" | "resources" | "assignments" | "calendars"; uid?: string; label?: string }>;
  } = {}): void {
    mikuprojectMainRender.renderXlsxImportSummary(document, changes, options);
  }

  function updateSummary(model: ProjectModel | null): void {
    mikuprojectMainRender.updateSummary(document, model, updateSvgButton);
  }

  function loadSample(): void {
    mikuprojectMainXmlActions.loadSampleXml({
      loadSampleXml: mikuprojectMainSamples.loadSampleXml,
      document,
      readSampleXml: () => mikuprojectXml.SAMPLE_XML,
      readSampleProjectDraftView: () => mikuprojectXml.SAMPLE_PROJECT_DRAFT_VIEW,
      writeXmlText: (xmlText) => {
        getTextArea("xmlInput").value = xmlText;
      },
      writeProjectDraftText: (text) => {
        getTextArea("projectDraftImportInput").value = text;
      },
      importSampleXml: mikuprojectXml.importMsProjectXml,
      validateProjectModel: mikuprojectXml.validateProjectModel,
      markXmlDirty,
      applyModelState: ({ model, issues }) => {
        mikuprojectMainFlow.applyModelState({
          model,
          issues,
          updateSummary,
          renderValidationIssues,
          renderImportWarnings,
          renderXlsxImportSummary
        });
      },
      updateSvgButton,
      setStatus,
      setActiveTab,
      setCurrentModel: (model) => {
        currentModel = model;
      },
      setXmlSourceDirty: (dirty) => {
        isXmlSourceDirty = dirty;
      }
    });
  }

  async function importXmlFromFile(file: File | null | undefined): Promise<void> {
    if (!file) {
      return;
    }
    await mikuprojectMainXmlActions.importXmlFromFile({
      file,
      writeXmlText: (xmlText) => {
        getTextArea("xmlInput").value = xmlText;
      },
      markXmlDirty,
      importMsProjectXml: mikuprojectXml.importMsProjectXml,
      validateProjectModel: mikuprojectXml.validateProjectModel,
      setCurrentModel: (model) => {
        currentModel = model;
      },
      setXmlSourceDirty: (dirty) => {
        isXmlSourceDirty = dirty;
      },
      applyModelState: ({ model, issues }) => {
        mikuprojectMainFlow.applyModelState({
          model,
          issues,
          updateSummary,
          renderValidationIssues,
          renderImportWarnings,
          renderXlsxImportSummary
        });
      },
      completeTransform: (statusMessage, toastMessage) => {
        mikuprojectMainFlow.completeTransform({ setStatus, showToast, setActiveTab, statusMessage, toastMessage });
      },
      exportCurrentMermaid
    });
  }

  function ensureCurrentModel(): ProjectModel {
    currentModel = mikuprojectMainModel.ensureCurrentModel({
      currentModel,
      readXmlText: () => getTextArea("xmlInput").value,
      importMsProjectXml: mikuprojectXml.importMsProjectXml,
      clearXmlSourceDirty: () => {
        isXmlSourceDirty = false;
      }
    });
    return currentModel;
  }

  function parseCurrentXml(options: { silent?: boolean } = {}): void {
    const result = mikuprojectMainModel.parseCurrentXml({
      readXmlText: () => getTextArea("xmlInput").value,
      importMsProjectXml: mikuprojectXml.importMsProjectXml,
      validateProjectModel: mikuprojectXml.validateProjectModel,
      clearXmlSourceDirty: () => {
        isXmlSourceDirty = false;
      }
    });
    if (result.empty || !result.model) {
      setStatus("XML が空です");
      return;
    }
    currentModel = result.model;
    mikuprojectMainFlow.applyModelState({
      model: currentModel,
      issues: result.issues,
      updateSummary,
      renderValidationIssues,
      renderImportWarnings,
      renderXlsxImportSummary
    });
    if (!options.silent) {
      mikuprojectMainFlow.completeTransform({
        setStatus,
        showToast,
        setActiveTab,
        statusMessage: result.issues.length > 0 ? `XML を解析しました。検証で ${result.issues.length} 件の問題があります` : "XML を内部モデルへ変換しました",
        toastMessage: "XML を解析しました"
      });
      return;
    }
    setActiveTab("transform", { skipTransformRefresh: true });
  }

  async function exportCurrentMermaid(options: { silent?: boolean } = {}): Promise<void> {
    await mikuprojectMainXmlActions.exportCurrentMermaid({
      currentModel,
      exportMermaidGantt: mikuprojectXml.exportMermaidGantt,
      setMermaidText: (text) => {
        getTextArea("mermaidOutput").value = text;
      },
      renderSvgPreview,
      setStatus,
      showToast,
      setActiveTab: () => {
        setActiveTab("transform", { skipTransformRefresh: true });
      },
      silent: options.silent
    });
  }

  function exportCurrentCsv(): void {
    const model = ensureCurrentModel();
    mikuprojectMainOutputActions.exportCsv({
      model,
      syncXmlTextFromModel,
      buildCsvExport: mikuprojectMainExport.buildCsvExport,
      exportCsvParentId: mikuprojectXml.exportCsvParentId,
      downloadBlob,
      completeOutput: (statusMessage, toastMessage) => {
        mikuprojectMainFlow.completeOutput({ setStatus, showToast, setActiveTab, statusMessage, toastMessage });
      }
    });
  }

  function exportCurrentProjectOverviewView(): void {
    const model = ensureCurrentModel();
    mikuprojectMainOutputActions.exportProjectOverview({
      model,
      syncXmlTextFromModel,
      buildProjectOverviewExport: mikuprojectMainExport.buildProjectOverviewExport,
      exportProjectOverviewView: mikuprojectXml.exportProjectOverviewView,
      setOutputText: (text) => {
        getTextArea("projectOverviewOutput").value = text;
      },
      downloadBlob,
      completeOutput: (statusMessage, toastMessage) => {
        mikuprojectMainFlow.completeOutput({ setStatus, showToast, setActiveTab, statusMessage, toastMessage });
      }
    });
  }

  function exportCurrentTaskEditView(): void {
    const model = ensureCurrentModel();
    mikuprojectMainOutputActions.exportTaskEdit({
      model,
      syncXmlTextFromModel,
      requestedTaskUid: getInput("taskEditUidInput").value.trim() || undefined,
      buildTaskEditExport: mikuprojectMainExport.buildTaskEditExport,
      exportTaskEditView: mikuprojectXml.exportTaskEditView,
      setResolvedTaskUid: (taskUid) => {
        getInput("taskEditUidInput").value = taskUid;
      },
      setOutputText: (text) => {
        getTextArea("taskEditOutput").value = text;
      },
      downloadBlob,
      completeOutput: (statusMessage, toastMessage) => {
        mikuprojectMainFlow.completeOutput({ setStatus, showToast, setActiveTab, statusMessage, toastMessage });
      }
    });
  }

  function exportCurrentAiProjectionBundle(): void {
    const model = ensureCurrentModel();
    mikuprojectMainOutputActions.exportAiProjectionBundle({
      model,
      syncXmlTextFromModel,
      buildAiProjectionBundleExport: mikuprojectMainExport.buildAiProjectionBundleExport,
      exportProjectOverviewView: mikuprojectXml.exportProjectOverviewView,
      exportPhaseDetailView: mikuprojectXml.exportPhaseDetailView,
      exportTaskEditView: mikuprojectXml.exportTaskEditView,
      setOutputText: (text) => {
        getTextArea("aiBundleOutput").value = text;
      },
      downloadBlob,
      completeOutput: (statusMessage, toastMessage) => {
        mikuprojectMainFlow.completeOutput({ setStatus, showToast, setActiveTab, statusMessage, toastMessage });
      }
    });
  }

  async function importProjectDraftFromText(): Promise<void> {
    await mikuprojectMainImportActions.importProjectDraftText({
      sourceText: getTextArea("projectDraftImportInput").value,
      importProjectDraftText: mikuprojectMainImport.importProjectDraftText,
      extractLastJsonBlock: mikuprojectAiJsonUtil.extractLastJsonBlock,
      importProjectDraftView: mikuprojectXml.importProjectDraftView,
      validateProjectModel: mikuprojectXml.validateProjectModel,
      setCurrentModel: (model) => {
        currentModel = model;
      },
      syncXmlTextFromModel,
      applyModelState: ({ model, issues }) => {
        mikuprojectMainFlow.applyModelState({
          model,
          issues,
          updateSummary,
          renderValidationIssues,
          renderImportWarnings,
          renderXlsxImportSummary
        });
      },
      exportCurrentMermaid,
      completeTransform: (statusMessage, toastMessage) => {
        mikuprojectMainFlow.completeTransform({ setStatus, showToast, setActiveTab, statusMessage, toastMessage });
      }
    });
  }

  async function importPatchJsonFromSourceText(sourceText: string): Promise<void> {
    await mikuprojectMainImportActions.importPatchJsonText({
      sourceText,
      importPatchJsonText: mikuprojectMainImport.importPatchJsonText,
      ensureCurrentModel,
      extractLastJsonBlock: mikuprojectAiJsonUtil.extractLastJsonBlock,
      importProjectPatchJson: mikuprojectProjectPatchJson.importProjectPatchJson,
      exportMsProjectXml: mikuprojectXml.exportMsProjectXml,
      validateProjectModel: mikuprojectXml.validateProjectModel,
      setCurrentModel: (model) => {
        currentModel = model;
      },
      applyModelState: ({ model, issues, warnings, changes }) => {
        mikuprojectMainFlow.applyModelState({
          model,
          issues,
          updateSummary,
          renderValidationIssues,
          renderImportWarnings,
          renderXlsxImportSummary,
          warnings,
          warningSourceLabel: "Patch JSON",
          changes,
          changeSourceLabel: "Patch JSON"
        });
      },
      writeXmlText: (xmlText) => {
        getTextArea("xmlInput").value = xmlText;
      },
      markXmlDirty,
      setXmlSourceDirty: (dirty) => {
        isXmlSourceDirty = dirty;
      },
      exportCurrentMermaid,
      completeTransform: (statusMessage, toastMessage) => {
        mikuprojectMainFlow.completeTransform({ setStatus, showToast, setActiveTab, statusMessage, toastMessage });
      }
    });
  }

  async function importAiEditJsonFromText(): Promise<void> {
    const sourceText = getTextArea("projectDraftImportInput").value;
    const kind = mikuprojectMainImport.detectAiEditJsonKind({
      sourceText,
      extractLastJsonBlock: mikuprojectAiJsonUtil.extractLastJsonBlock,
      detectJsonDocumentKind: mikuprojectAiJsonUtil.detectJsonDocumentKind
    });
    if (kind === "project_draft_view") {
      await importProjectDraftFromText();
      return;
    }
    await importPatchJsonFromSourceText(sourceText);
  }

  function loadProjectDraftSample(): void {
    mikuprojectMainSamples.loadProjectDraftSample({
      document,
      readSampleXml: () => mikuprojectXml.SAMPLE_XML,
      readSampleProjectDraftView: () => mikuprojectXml.SAMPLE_PROJECT_DRAFT_VIEW,
      writeXmlText: (xmlText) => {
        getTextArea("xmlInput").value = xmlText;
      },
      writeProjectDraftText: (text) => {
        getTextArea("projectDraftImportInput").value = text;
      },
      importSampleXml: mikuprojectXml.importMsProjectXml,
      validateProjectModel: mikuprojectXml.validateProjectModel,
      markXmlDirty,
      applyModelState: ({ model, issues }) => {
        mikuprojectMainFlow.applyModelState({
          model,
          issues,
          updateSummary,
          renderValidationIssues,
          renderImportWarnings,
          renderXlsxImportSummary
        });
      },
      updateSvgButton,
      setStatus,
      setActiveTab
    });
  }

  async function importProjectDraftFromFile(file?: File | null): Promise<void> {
    if (!file) {
      throw new Error("project_draft_view JSON ファイルを選択してください");
    }
    const sourceText = await file.text();
    getTextArea("projectDraftImportInput").value = sourceText;
    await importProjectDraftFromText();
  }

  async function importWorkbookJsonFromSourceText(sourceText: string): Promise<void> {
    await mikuprojectMainImportActions.importWorkbookJsonText({
      sourceText,
      importWorkbookJsonText: mikuprojectMainImport.importWorkbookJsonText,
      previousModel: currentModel,
      extractLastJsonBlock: mikuprojectAiJsonUtil.extractLastJsonBlock,
      importProjectWorkbookJsonAsProjectModel: mikuprojectProjectWorkbookJson.importProjectWorkbookJsonAsProjectModel,
      importProjectWorkbookJson: mikuprojectProjectWorkbookJson.importProjectWorkbookJson,
      validateProjectModel: mikuprojectXml.validateProjectModel,
      setCurrentModel: (model) => {
        currentModel = model;
      },
      syncXmlTextFromModel,
      applyModelState: ({ model, issues, warnings, changes }) => {
        mikuprojectMainFlow.applyModelState({
          model,
          issues,
          updateSummary,
          renderValidationIssues,
          renderImportWarnings,
          renderXlsxImportSummary,
          warnings,
          warningSourceLabel: "JSON Replace",
          changes,
          changeSourceLabel: "JSON Replace"
        });
      },
      exportCurrentMermaid,
      completeTransform: (statusMessage, toastMessage) => {
        mikuprojectMainFlow.completeTransform({ setStatus, showToast, setActiveTab, statusMessage, toastMessage });
      }
    });
  }

  async function importWorkbookJsonFromFile(file?: File | null): Promise<void> {
    if (!file) {
      throw new Error("workbook JSON ファイルを選択してください");
    }
    const sourceText = await file.text();
    await importWorkbookJsonFromSourceText(sourceText);
  }

  async function importFromFile(file: File | null | undefined): Promise<void> {
    if (!file) {
      return;
    }
    const sourceText = file.name.toLowerCase().endsWith(".json") || file.name.toLowerCase().endsWith(".editjson")
      ? await file.text()
      : undefined;
    const kind = mikuprojectMainIo.detectImportKind({
      fileName: file.name,
      readJsonText: sourceText,
      extractLastJsonBlock: mikuprojectAiJsonUtil.extractLastJsonBlock,
      detectJsonDocumentKind: mikuprojectAiJsonUtil.detectJsonDocumentKind
    });
    if (kind === "xml") {
      await importXmlFromFile(file);
      return;
    }
    if (kind === "xlsx") {
      await importXlsxFromFile(file);
      return;
    }
    if (kind === "csv") {
      await importCsvFromFile(file);
      return;
    }
    if (kind === "editjson") {
      getTextArea("projectDraftImportInput").value = sourceText;
      await importAiEditJsonFromText();
      return;
    }
    if (kind === "workbook_json") {
      await importWorkbookJsonFromSourceText(sourceText);
      return;
    }
    if (kind === "project_draft_view") {
      getTextArea("projectDraftImportInput").value = sourceText;
      await importProjectDraftFromText();
      return;
    }
    getTextArea("projectDraftImportInput").value = sourceText;
    await importPatchJsonFromSourceText(sourceText);
  }

  function exportCurrentPhaseDetailView(mode: "full" | "scoped" = "scoped"): void {
    const model = ensureCurrentModel();
    const requestedPhaseUid = getInput("phaseDetailUidInput").value.trim() || undefined;
    const requestedRootUid = mode === "scoped" ? getInput("phaseDetailRootUidInput").value.trim() || undefined : undefined;
    const maxDepthText = getInput("phaseDetailMaxDepthInput").value.trim();
    const requestedMaxDepth = mode === "scoped" && maxDepthText !== "" ? Number.parseInt(maxDepthText, 10) : undefined;
    if (typeof requestedMaxDepth === "number" && (!Number.isFinite(requestedMaxDepth) || requestedMaxDepth < 0)) {
      throw new Error("max depth は 0 以上の整数で指定してください");
    }
    mikuprojectMainOutputActions.exportPhaseDetail({
      model,
      syncXmlTextFromModel,
      mode,
      requestedPhaseUid,
      requestedRootUid,
      requestedMaxDepth,
      buildPhaseDetailExport: mikuprojectMainExport.buildPhaseDetailExport,
      exportPhaseDetailView: mikuprojectXml.exportPhaseDetailView,
      setResolvedPhaseUid: (uid) => {
        getInput("phaseDetailUidInput").value = uid;
      },
      setResolvedRootUid: (uid) => {
        getInput("phaseDetailRootUidInput").value = uid;
      },
      setResolvedMaxDepth: (maxDepth) => {
        getInput("phaseDetailMaxDepthInput").value = typeof maxDepth === "number" ? String(maxDepth) : "";
      },
      setOutputText: (text) => {
        getTextArea("phaseDetailOutput").value = text;
      },
      downloadBlob,
      completeOutput: (statusMessage, toastMessage) => {
        mikuprojectMainFlow.completeOutput({ setStatus, showToast, setActiveTab, statusMessage, toastMessage });
      }
    });
  }

  function exportCurrentXlsx(): void {
    const model = ensureCurrentModel();
    mikuprojectMainOutputActions.exportXlsx({
      model,
      syncXmlTextFromModel,
      buildXlsxExport: mikuprojectMainExport.buildXlsxExport,
      createWorkbookCodec: () => new mikuprojectExcelIo.XlsxWorkbookCodec(),
      exportProjectWorkbook: mikuprojectProjectXlsx.exportProjectWorkbook,
      downloadBlob,
      completeOutput: (statusMessage, toastMessage) => {
        mikuprojectMainFlow.completeOutput({ setStatus, showToast, setActiveTab, statusMessage, toastMessage });
      }
    });
  }

  function exportCurrentWorkbookJson(): void {
    const model = ensureCurrentModel();
    mikuprojectMainOutputActions.exportWorkbookJson({
      model,
      syncXmlTextFromModel,
      buildWorkbookJsonExport: mikuprojectMainExport.buildWorkbookJsonExport,
      exportProjectWorkbookJson: mikuprojectProjectWorkbookJson.exportProjectWorkbookJson,
      setOutputText: (text) => {
        getTextArea("workbookJsonOutput").value = text;
      },
      downloadBlob,
      completeOutput: (statusMessage, toastMessage) => {
        mikuprojectMainFlow.completeOutput({ setStatus, showToast, setActiveTab, statusMessage, toastMessage });
      }
    });
  }

  function exportCurrentWbsXlsx(): void {
    const model = ensureCurrentModel();
    const options = buildCurrentWbsOptions(model);
    mikuprojectMainOutputActions.exportWbsXlsx({
      model,
      syncXmlTextFromModel,
      options,
      buildWbsXlsxExport: mikuprojectMainExport.buildWbsXlsxExport,
      createWorkbookCodec: () => new mikuprojectExcelIo.XlsxWorkbookCodec(),
      exportWbsWorkbook: mikuprojectWbsXlsx.exportWbsWorkbook,
      downloadBlob,
      completeOutput: (statusMessage, toastMessage) => {
        mikuprojectMainFlow.completeOutput({ setStatus, showToast, setActiveTab, statusMessage, toastMessage });
      }
    });
  }

  async function importXlsxFromFile(file: File | null | undefined): Promise<void> {
    if (!file) {
      return;
    }
    await mikuprojectMainImportActions.importXlsxFile({
      file,
      importXlsxBytes: mikuprojectMainImport.importXlsxBytes,
      previousModel: currentModel,
      createWorkbookCodec: () => new mikuprojectExcelIo.XlsxWorkbookCodec(),
      importProjectWorkbookAsProjectModel: mikuprojectProjectXlsx.importProjectWorkbookAsProjectModel,
      importProjectWorkbookDetailed: mikuprojectProjectXlsx.importProjectWorkbookDetailed,
      validateProjectModel: mikuprojectXml.validateProjectModel,
      setCurrentModel: (model) => {
        currentModel = model;
      },
      syncXmlTextFromModel,
      updateSummary,
      renderValidationIssues,
      renderImportWarnings,
      renderXlsxImportSummary,
      setStatus,
      showToast,
      setActiveTab: () => {
        setActiveTab("transform", { skipTransformRefresh: true });
      },
      exportCurrentMermaid
    });
  }

  async function importCsvFromFile(file: File | null | undefined): Promise<void> {
    if (!file) {
      return;
    }
    await mikuprojectMainImportActions.importCsvFile({
      file,
      importCsvText: mikuprojectMainImport.importCsvText,
      importCsvParentId: mikuprojectXml.importCsvParentId,
      validateProjectModel: mikuprojectXml.validateProjectModel,
      setCurrentModel: (model) => {
        currentModel = model;
      },
      setXmlSourceDirty: (dirty) => {
        isXmlSourceDirty = dirty;
      },
      updateSummary,
      renderValidationIssues,
      renderImportWarnings,
      renderXlsxImportSummary: () => {
        renderXlsxImportSummary([]);
      },
      setStatus,
      showToast,
      setActiveTab: () => {
        setActiveTab("transform", { skipTransformRefresh: true });
      },
      exportCurrentMermaid
    });
  }

  function downloadCurrentXml(): void {
    const model = ensureCurrentModel();
    mikuprojectMainDownloads.downloadXml({
      xmlText: syncXmlTextFromModel(model),
      buildXmlExport: mikuprojectMainExport.buildXmlExport,
      downloadBlob,
      markXmlSavedCurrent,
      completeOutput: (statusMessage, toastMessage) => {
        mikuprojectMainFlow.completeOutput({
          setStatus,
          showToast,
          setActiveTab,
          statusMessage,
          toastMessage
        });
      }
    });
  }

  async function downloadCurrentSvg(): Promise<void> {
    const model = ensureCurrentModel();
    syncXmlTextFromModel(model);
    const mermaidText = mikuprojectXml.exportMermaidGantt(model);
    getTextArea("mermaidOutput").value = mermaidText;
    await renderSvgPreview();
    if (!currentSvgPreviewState.state.currentNativeSvg) {
      setStatus("出力する SVG がありません");
      return;
    }
    mikuprojectMainDownloads.downloadDailySvg({
      svg: currentSvgPreviewState.state.currentNativeSvg,
      buildDailySvgExport: mikuprojectMainExport.buildDailySvgExport,
      downloadBlob,
      completeOutput: (statusMessage, toastMessage) => {
        mikuprojectMainFlow.completeOutput({
          setStatus,
          showToast,
          setActiveTab,
          statusMessage,
          toastMessage
        });
      }
    });
  }

  function downloadCurrentWeeklySvg(): void {
    const model = ensureCurrentModel();
    syncXmlTextFromModel(model);
    const weeklySvg = mikuprojectNativeSvg.exportWeeklyNativeSvg(model, buildCurrentWbsOptions(model));
    if (!weeklySvg) {
      setStatus("出力する Weekly SVG がありません");
      return;
    }
    mikuprojectMainDownloads.downloadWeeklySvg({
      svg: weeklySvg,
      buildWeeklySvgExport: mikuprojectMainExport.buildWeeklySvgExport,
      downloadBlob,
      completeOutput: (statusMessage, toastMessage) => {
        mikuprojectMainFlow.completeOutput({
          setStatus,
          showToast,
          setActiveTab,
          statusMessage,
          toastMessage
        });
      }
    });
  }

  function downloadCurrentMonthlyWbsSvgZip(): void {
    const model = ensureCurrentModel();
    syncXmlTextFromModel(model);
    const archive = mikuprojectNativeSvg.exportMonthlyWbsCalendarSvgArchive(model);
    if (!archive.entries.length || archive.zipBytes.byteLength === 0) {
      throw new Error("出力する月別 WBS カレンダー SVG がありません");
    }
    mikuprojectMainDownloads.downloadMonthlySvgZip({
      zipBytes: archive.zipBytes,
      buildMonthlySvgZipExport: mikuprojectMainExport.buildMonthlySvgZipExport,
      downloadBlob,
      completeOutput: (statusMessage, toastMessage) => {
        const suffix = ` (${archive.entries.length} か月分, ZIP)`;
        mikuprojectMainFlow.completeOutput({
          setStatus,
          showToast,
          setActiveTab,
          statusMessage: `${statusMessage}${suffix}`,
          toastMessage
        });
      }
    });
  }

  function downloadCurrentMermaidMmd(): void {
    const model = ensureCurrentModel();
    syncXmlTextFromModel(model);
    const mermaidText = mikuprojectXml.exportMermaidGantt(model);
    getTextArea("mermaidOutput").value = mermaidText;
    mikuprojectMainDownloads.downloadMermaid({
      mermaidText,
      buildMermaidExport: mikuprojectMainExport.buildMermaidExport,
      downloadBlob,
      completeOutput: (statusMessage, toastMessage) => {
        mikuprojectMainFlow.completeOutput({
          setStatus,
          showToast,
          setActiveTab,
          statusMessage,
          toastMessage
        });
      }
    });
  }

  function downloadCurrentWbsMarkdown(): void {
    const model = ensureCurrentModel();
    syncXmlTextFromModel(model);
    const options = buildCurrentWbsOptions(model);
    const markdownText = mikuprojectWbsMarkdown.exportWbsMarkdown(model, options);
    mikuprojectMainDownloads.downloadWbsMarkdown({
      markdownText,
      buildWbsMarkdownExport: mikuprojectMainExport.buildWbsMarkdownExport,
      downloadBlob,
      completeOutput: (statusMessage, toastMessage) => {
        mikuprojectMainFlow.completeOutput({
          setStatus,
          showToast,
          setActiveTab,
          statusMessage,
          toastMessage
        });
      }
    });
  }

  function runRoundTripCheck(): void {
    mikuprojectMainXmlActions.runRoundTripCheck({
      currentModel,
      parseCurrentXml,
      getCurrentModel: () => currentModel,
      assertRoundTripStable: mikuprojectMainIo.assertRoundTripStable,
      exportMsProjectXml: mikuprojectXml.exportMsProjectXml,
      importMsProjectXml: mikuprojectXml.importMsProjectXml,
      validateProjectModel: mikuprojectXml.validateProjectModel,
      normalizeProjectModel: mikuprojectXml.normalizeProjectModel,
      renderValidationIssues,
      setStatus,
      showToast,
      setActiveTab: () => {
        setActiveTab("transform");
      }
    });
  }

  function bindEvents(): void {
    mikuprojectMainEvents.bind({
      document,
      setStatus,
      refreshXmlSaveState,
      markXmlSourceDirty: () => {
        isXmlSourceDirty = true;
      },
      getImportFileInput: () => getElement<HTMLInputElement>("importFileInput"),
      getXmlInput: () => getTextArea("xmlInput"),
      loadSample,
      downloadAllOutputs,
      setSvgPreviewMode,
      downloadCurrentSvg,
      downloadCurrentWeeklySvg,
      downloadCurrentMonthlyWbsSvgZip,
      downloadCurrentMermaidMmd,
      exportCurrentCsv,
      exportCurrentProjectOverviewView,
      exportCurrentTaskEditView,
      exportCurrentAiProjectionBundle,
      loadProjectDraftSample,
      copyAiPrompt,
      importAiEditJsonFromText,
      exportCurrentPhaseDetailView,
      exportCurrentXlsx,
      exportCurrentWorkbookJson,
      exportCurrentWbsXlsx,
      downloadCurrentWbsMarkdown,
      downloadCurrentXml,
      runRoundTripCheck,
      importFromFile
    });
  }

  function initialize(): void {
    bindTabs();
    bindEvents();
    updateSummary(null);
    renderValidationIssues([]);
    renderImportWarnings([]);
    renderXlsxImportSummary([]);
    updateSvgPreviewModeButtons();
    updateSvgButton();
    loadSample();
  }

  (globalThis as typeof globalThis & {
    __mikuprojectMainTestHooks?: {
      parseCurrentXml: () => void;
      buildCurrentOutputArchiveEntries: () => Array<{ name: string; data: Uint8Array }>;
      importFromFile: (file: File | null | undefined) => Promise<void>;
      runRoundTripCheck: () => void;
      exportCurrentMermaid: () => Promise<void>;
      renderValidationIssues: (issues: ValidationIssue[]) => void;
      renderXlsxImportSummary: (changes: Array<{
        scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
        uid: string;
        label: string;
        field: string;
        before: string | number | boolean | undefined;
        after: string | number | boolean;
      }>, options?: {
        sourceLabel?: string;
        warnings?: Array<{ message: string; scope?: "project" | "tasks" | "resources" | "assignments" | "calendars"; uid?: string; label?: string }>;
      }) => void;
    };
  }).__mikuprojectMainTestHooks = {
    parseCurrentXml,
    buildCurrentOutputArchiveEntries,
    importFromFile,
    runRoundTripCheck,
    exportCurrentMermaid,
    renderValidationIssues,
    renderXlsxImportSummary
  };

  document.addEventListener("DOMContentLoaded", initialize);
})();
