/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const mikuprojectXml = globalThis.__mikuprojectXml;
    if (!mikuprojectXml) {
        throw new Error("mikuproject XML module is not loaded");
    }
    const mikuprojectAiJsonUtil = globalThis.__mikuprojectAiJsonUtil;
    if (!mikuprojectAiJsonUtil) {
        throw new Error("mikuproject AI JSON util module is not loaded");
    }
    const mikuprojectAiJsonSpec = globalThis.__mikuprojectAiJsonSpec;
    const mikuprojectMainUtil = globalThis.__mikuprojectMainUtil;
    if (!mikuprojectMainUtil) {
        throw new Error("mikuproject main util module is not loaded");
    }
    const mikuprojectMainRender = globalThis.__mikuprojectMainRender;
    if (!mikuprojectMainRender) {
        throw new Error("mikuproject main render module is not loaded");
    }
    const mikuprojectMainIo = globalThis.__mikuprojectMainIo;
    if (!mikuprojectMainIo) {
        throw new Error("mikuproject main IO module is not loaded");
    }
    const mikuprojectMainImport = globalThis.__mikuprojectMainImport;
    if (!mikuprojectMainImport) {
        throw new Error("mikuproject main import module is not loaded");
    }
    const mikuprojectMainExport = globalThis.__mikuprojectMainExport;
    if (!mikuprojectMainExport) {
        throw new Error("mikuproject main export module is not loaded");
    }
    const mikuprojectMainEvents = globalThis.__mikuprojectMainEvents;
    if (!mikuprojectMainEvents) {
        throw new Error("mikuproject main events module is not loaded");
    }
    const mikuprojectMainPreview = globalThis.__mikuprojectMainPreview;
    if (!mikuprojectMainPreview) {
        throw new Error("mikuproject main preview module is not loaded");
    }
    const mikuprojectMainUi = globalThis.__mikuprojectMainUi;
    if (!mikuprojectMainUi) {
        throw new Error("mikuproject main UI module is not loaded");
    }
    const mikuprojectMainSupport = globalThis.__mikuprojectMainSupport;
    if (!mikuprojectMainSupport) {
        throw new Error("mikuproject main support module is not loaded");
    }
    const mikuprojectMainDownloads = globalThis.__mikuprojectMainDownloads;
    if (!mikuprojectMainDownloads) {
        throw new Error("mikuproject main downloads module is not loaded");
    }
    const mikuprojectMainOutputActions = globalThis.__mikuprojectMainOutputActions;
    if (!mikuprojectMainOutputActions) {
        throw new Error("mikuproject main output actions module is not loaded");
    }
    const mikuprojectMainImportActions = globalThis.__mikuprojectMainImportActions;
    if (!mikuprojectMainImportActions) {
        throw new Error("mikuproject main import actions module is not loaded");
    }
    const mikuprojectMainXmlActions = globalThis.__mikuprojectMainXmlActions;
    if (!mikuprojectMainXmlActions) {
        throw new Error("mikuproject main XML actions module is not loaded");
    }
    const mikuprojectMainArchiveActions = globalThis.__mikuprojectMainArchiveActions;
    if (!mikuprojectMainArchiveActions) {
        throw new Error("mikuproject main archive actions module is not loaded");
    }
    const mikuprojectMainSaveState = globalThis.__mikuprojectMainSaveState;
    if (!mikuprojectMainSaveState) {
        throw new Error("mikuproject main save state module is not loaded");
    }
    const mikuprojectMainTabActions = globalThis.__mikuprojectMainTabActions;
    if (!mikuprojectMainTabActions) {
        throw new Error("mikuproject main tab actions module is not loaded");
    }
    const mikuprojectMainPreviewActions = globalThis.__mikuprojectMainPreviewActions;
    if (!mikuprojectMainPreviewActions) {
        throw new Error("mikuproject main preview actions module is not loaded");
    }
    const mikuprojectMainTransform = globalThis.__mikuprojectMainTransform;
    if (!mikuprojectMainTransform) {
        throw new Error("mikuproject main transform module is not loaded");
    }
    const mikuprojectMainFlow = globalThis.__mikuprojectMainFlow;
    if (!mikuprojectMainFlow) {
        throw new Error("mikuproject main flow module is not loaded");
    }
    const mikuprojectMainModel = globalThis.__mikuprojectMainModel;
    if (!mikuprojectMainModel) {
        throw new Error("mikuproject main model module is not loaded");
    }
    const mikuprojectExcelIo = globalThis.__mikuprojectExcelIo;
    if (!mikuprojectExcelIo) {
        throw new Error("mikuproject Excel IO module is not loaded");
    }
    const mikuprojectProjectXlsx = globalThis.__mikuprojectProjectXlsx;
    if (!mikuprojectProjectXlsx) {
        throw new Error("mikuproject Project XLSX module is not loaded");
    }
    const mikuprojectProjectWorkbookJson = globalThis.__mikuprojectProjectWorkbookJson;
    if (!mikuprojectProjectWorkbookJson) {
        throw new Error("mikuproject Project Workbook JSON module is not loaded");
    }
    const mikuprojectProjectPatchJson = globalThis.__mikuprojectProjectPatchJson;
    if (!mikuprojectProjectPatchJson) {
        throw new Error("mikuproject Project Patch JSON module is not loaded");
    }
    const mikuprojectWbsXlsx = globalThis.__mikuprojectWbsXlsx;
    if (!mikuprojectWbsXlsx) {
        throw new Error("mikuproject WBS XLSX module is not loaded");
    }
    const mikuprojectWbsMarkdown = globalThis.__mikuprojectWbsMarkdown;
    if (!mikuprojectWbsMarkdown) {
        throw new Error("mikuproject WBS Markdown module is not loaded");
    }
    const mikuprojectNativeSvg = globalThis.__mikuprojectNativeSvg;
    if (!mikuprojectNativeSvg) {
        throw new Error("mikuproject native SVG module is not loaded");
    }
    let currentModel = null;
    const currentSvgPreviewState = {
        state: mikuprojectMainPreview.createEmptyState()
    };
    let lastSavedXmlText = "";
    let lastSavedXmlStamp = "";
    let currentTabId = "input";
    let isXmlSourceDirty = true;
    let isRefreshingTransformTab = false;
    function getElement(id) {
        const element = document.getElementById(id);
        if (!element) {
            throw new Error(`Element not found: ${id}`);
        }
        return element;
    }
    function getTextArea(id) {
        return getElement(id);
    }
    function getInput(id) {
        return getElement(id);
    }
    function getTabButtons() {
        return mikuprojectMainUi.getTabButtons(document);
    }
    function getTabPanels() {
        return mikuprojectMainUi.getTabPanels(document);
    }
    function setActiveTab(tabId, options = {}) {
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
    async function refreshTransformTab() {
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
    function moveTabFocus(currentButton, direction) {
        mikuprojectMainTabActions.moveTabFocus({
            document,
            currentButton,
            direction,
            moveTabFocusView: mikuprojectMainUi.moveTabFocus,
            setActiveTab
        });
    }
    function bindTabs() {
        mikuprojectMainTabActions.bindTabs({
            document,
            currentTabId,
            getTabButtons,
            bindTabsView: mikuprojectMainTransform.bindTabs,
            setActiveTab,
            moveTabFocus
        });
    }
    function parseOptionalNonNegativeInteger(raw) {
        return mikuprojectMainUtil.parseOptionalNonNegativeInteger(raw);
    }
    function showToast(message) {
        mikuprojectMainSupport.showToast(document, message);
    }
    function getAiPromptText() {
        return mikuprojectMainSupport.getAiPromptText(document, mikuprojectAiJsonSpec && typeof mikuprojectAiJsonSpec.getAiJsonSpecText === "function"
            ? mikuprojectAiJsonSpec.getAiJsonSpecText
            : undefined);
    }
    async function copyTextToClipboard(text) {
        await mikuprojectMainSupport.copyTextToClipboard(document, text);
    }
    const mikuprojectMainSamples = globalThis.__mikuprojectMainSamples;
    if (!mikuprojectMainSamples) {
        throw new Error("mikuproject main samples module is not loaded");
    }
    async function copyAiPrompt() {
        await mikuprojectMainSamples.copyAiPrompt({
            document,
            getAiPromptText,
            copyTextToClipboard,
            showToast,
            setStatus
        });
    }
    function setSvgPreviewMarkup(markup) {
        getElement("nativeSvgPreview").innerHTML = markup;
    }
    function updateSvgPreviewModeButtons() {
        mikuprojectMainPreviewActions.updateSvgPreviewModeButtons({
            document,
            mode: currentSvgPreviewState.state.currentSvgPreviewMode,
            applyPreviewModeButtonClasses: mikuprojectMainPreview.applyPreviewModeButtonClasses
        });
    }
    function renderCurrentSvgPreviewMarkup() {
        mikuprojectMainPreviewActions.renderCurrentSvgPreviewMarkup({
            state: currentSvgPreviewState.state,
            renderPreviewMarkup: mikuprojectMainPreview.renderPreviewMarkup,
            setSvgPreviewMarkup
        });
    }
    function updateSvgButton() {
        mikuprojectMainPreviewActions.updateSvgButton({
            document,
            hasModel: Boolean(currentModel),
            updateDownloadButtons: mikuprojectMainPreview.updateDownloadButtons
        });
    }
    function buildCurrentWbsOptions(model) {
        return mikuprojectMainTransform.buildWbsOptions({
            doc: document,
            model,
            parseOptionalNonNegativeInteger: parseOptionalNonNegativeInteger,
            collectWbsHolidayDates: mikuprojectWbsXlsx.collectWbsHolidayDates
        });
    }
    function downloadBlob(blob, filename) {
        mikuprojectMainSupport.downloadBlob(document, blob, filename);
    }
    function formatTimestampCompact(date) {
        return mikuprojectMainUtil.formatTimestampCompact(date);
    }
    async function renderSvgPreview() {
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
    function setSvgPreviewMode(mode) {
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
    function buildCurrentOutputArchiveEntries() {
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
    function buildCurrentOutputArchive() {
        return mikuprojectMainArchiveActions.buildOutputArchive({
            buildOutputArchiveEntries: buildCurrentOutputArchiveEntries,
            formatTimestampCompact,
            packZipEntries: mikuprojectMainUtil.packZipEntries
        });
    }
    function downloadAllOutputs() {
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
    function setStatus(message) {
        mikuprojectMainUi.setStatus(document, message);
    }
    function formatSaveStamp(date) {
        return mikuprojectMainUtil.formatSaveStamp(date);
    }
    function updateXmlSaveState(isDirty) {
        mikuprojectMainSaveState.updateXmlSaveState({
            document,
            isDirty,
            lastSavedXmlStamp,
            updateXmlSaveStateView: mikuprojectMainUi.updateXmlSaveState
        });
    }
    function markXmlDirty() {
        mikuprojectMainSaveState.markXmlDirty({
            document,
            lastSavedXmlStamp,
            updateXmlSaveStateView: mikuprojectMainUi.updateXmlSaveState
        });
    }
    function markXmlSavedCurrent() {
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
    function refreshXmlSaveState() {
        mikuprojectMainSaveState.refreshXmlSaveState({
            readXmlText: () => getTextArea("xmlInput").value,
            lastSavedXmlText,
            lastSavedXmlStamp,
            document,
            updateXmlSaveStateView: mikuprojectMainUi.updateXmlSaveState
        });
    }
    function syncXmlTextFromModel(model) {
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
    function renderValidationIssues(issues) {
        mikuprojectMainRender.renderValidationIssues(document, issues);
    }
    function renderImportWarnings(warnings, options = {}) {
        mikuprojectMainRender.renderImportWarnings(document, warnings, options);
    }
    function renderXlsxImportSummary(changes, options = {}) {
        mikuprojectMainRender.renderXlsxImportSummary(document, changes, options);
    }
    function updateSummary(model) {
        mikuprojectMainRender.updateSummary(document, model, updateSvgButton);
    }
    function loadSample() {
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
    async function importXmlFromFile(file) {
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
    function ensureCurrentModel() {
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
    function parseCurrentXml(options = {}) {
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
    async function exportCurrentMermaid(options = {}) {
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
    function exportCurrentCsv() {
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
    function exportCurrentProjectOverviewView() {
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
    function exportCurrentTaskEditView() {
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
    function exportCurrentAiProjectionBundle() {
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
    async function importProjectDraftFromText() {
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
    async function importPatchJsonFromSourceText(sourceText) {
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
    async function importAiEditJsonFromText() {
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
    function loadProjectDraftSample() {
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
    async function importProjectDraftFromFile(file) {
        if (!file) {
            throw new Error("project_draft_view JSON ファイルを選択してください");
        }
        const sourceText = await file.text();
        getTextArea("projectDraftImportInput").value = sourceText;
        await importProjectDraftFromText();
    }
    async function importWorkbookJsonFromSourceText(sourceText) {
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
    async function importWorkbookJsonFromFile(file) {
        if (!file) {
            throw new Error("workbook JSON ファイルを選択してください");
        }
        const sourceText = await file.text();
        await importWorkbookJsonFromSourceText(sourceText);
    }
    async function importFromFile(file) {
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
    function exportCurrentPhaseDetailView(mode = "scoped") {
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
    function exportCurrentXlsx() {
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
    function exportCurrentWorkbookJson() {
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
    function exportCurrentWbsXlsx() {
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
    async function importXlsxFromFile(file) {
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
    async function importCsvFromFile(file) {
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
    function downloadCurrentXml() {
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
    async function downloadCurrentSvg() {
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
    function downloadCurrentWeeklySvg() {
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
    function downloadCurrentMonthlyWbsSvgZip() {
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
    function downloadCurrentMermaidMmd() {
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
    function downloadCurrentWbsMarkdown() {
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
    function runRoundTripCheck() {
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
    function bindEvents() {
        mikuprojectMainEvents.bind({
            document,
            setStatus,
            refreshXmlSaveState,
            markXmlSourceDirty: () => {
                isXmlSourceDirty = true;
            },
            getImportFileInput: () => getElement("importFileInput"),
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
    function initialize() {
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
    globalThis.__mikuprojectMainTestHooks = {
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
