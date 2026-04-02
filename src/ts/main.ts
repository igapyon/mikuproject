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
      renderImportWarnings: (doc: Document, warnings: Array<{ message: string }>) => void;
      renderXlsxImportSummary: (doc: Document, changes: Array<{
        scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
        uid: string;
        label: string;
        field: string;
        before: string | number | boolean | undefined;
        after: string | number | boolean;
      }>) => void;
      updateSummary: (doc: Document, model: ProjectModel | null, updateSvgButton: () => void) => void;
    };
  }).__mikuprojectMainRender;

  if (!mikuprojectMainRender) {
    throw new Error("mikuproject main render module is not loaded");
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
  let currentNativeSvg = "";
  let currentWeeklyPreviewSvg = "";
  let currentMonthlyPreviewSvg = "";
  let lastSavedXmlText = "";
  let lastSavedXmlStamp = "";
  let currentTabId: "input" | "transform" | "output" = "input";
  let currentSvgPreviewMode: "daily" | "weekly" | "monthly" = "daily";
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
    return Array.from(document.querySelectorAll<HTMLButtonElement>(".md-top-tab[data-tab]"));
  }

  function getTabPanels(): HTMLElement[] {
    return Array.from(document.querySelectorAll<HTMLElement>(".md-tab-panel[data-tab-panel]"));
  }

  function setActiveTab(
    tabId: "input" | "transform" | "output",
    options: { skipTransformRefresh?: boolean } = {}
  ): void {
    currentTabId = tabId;
    for (const button of getTabButtons()) {
      const isActive = button.dataset.tab === tabId;
      button.classList.toggle("is-active", isActive);
      button.setAttribute("aria-selected", isActive ? "true" : "false");
      button.tabIndex = isActive ? 0 : -1;
    }
    for (const panel of getTabPanels()) {
      panel.hidden = panel.dataset.tabPanel !== tabId;
    }
    if (tabId === "transform" && !options.skipTransformRefresh && !isRefreshingTransformTab) {
      void refreshTransformTab().catch((error) => {
        setStatus(error instanceof Error ? error.message : "Transform の更新に失敗しました");
      });
    }
  }

  async function refreshTransformTab(): Promise<void> {
    if (isRefreshingTransformTab) {
      return;
    }
    isRefreshingTransformTab = true;
    try {
      if (!currentModel || isXmlSourceDirty) {
        const xmlText = getTextArea("xmlInput").value.trim();
        if (!xmlText) {
          setStatus("XML が空です");
          return;
        }
        parseCurrentXml({ silent: true });
      }
      await exportCurrentMermaid({ silent: true });
    } finally {
      isRefreshingTransformTab = false;
    }
  }

  function moveTabFocus(currentButton: HTMLButtonElement, direction: -1 | 1): void {
    const buttons = getTabButtons();
    const currentIndex = buttons.indexOf(currentButton);
    if (currentIndex < 0) {
      return;
    }
    const nextIndex = (currentIndex + direction + buttons.length) % buttons.length;
    const nextButton = buttons[nextIndex];
    nextButton.focus();
    const nextTab = nextButton.dataset.tab;
    if (nextTab === "input" || nextTab === "transform" || nextTab === "output") {
      setActiveTab(nextTab);
    }
  }

  function bindTabs(): void {
    const buttons = getTabButtons();
    if (buttons.length === 0) {
      return;
    }
    for (const button of buttons) {
      button.addEventListener("click", () => {
        const tabId = button.dataset.tab;
        if (tabId === "input" || tabId === "transform" || tabId === "output") {
          setActiveTab(tabId);
        }
      });
      button.addEventListener("keydown", (event) => {
        if (event.key === "ArrowRight" || event.key === "ArrowDown") {
          event.preventDefault();
          moveTabFocus(button, 1);
          return;
        }
        if (event.key === "ArrowLeft" || event.key === "ArrowUp") {
          event.preventDefault();
          moveTabFocus(button, -1);
          return;
        }
        if (event.key === "Home") {
          event.preventDefault();
          buttons[0].focus();
          setActiveTab("input");
          return;
        }
        if (event.key === "End") {
          event.preventDefault();
          buttons[buttons.length - 1].focus();
          setActiveTab("output");
        }
      });
    }
    setActiveTab(currentTabId);
  }

  function parseOptionalNonNegativeInteger(raw: string): number | undefined {
    return mikuprojectMainUtil.parseOptionalNonNegativeInteger(raw);
  }

  function parseWbsDisplayDaysBeforeBaseDate(): number | undefined {
    return parseOptionalNonNegativeInteger(getInput("wbsDisplayDaysBeforeInput").value);
  }

  function parseWbsDisplayDaysAfterBaseDate(): number | undefined {
    return parseOptionalNonNegativeInteger(getInput("wbsDisplayDaysAfterInput").value);
  }

  function useBusinessDaysForWbsDisplayRange(): boolean {
    return true;
  }

  function useBusinessDaysForWbsProgressBand(): boolean {
    return true;
  }

  function showToast(message: string): void {
    const toast = document.getElementById("toast") as (HTMLElement & { show?: (text: string, duration?: number) => void }) | null;
    if (toast && typeof toast.show === "function") {
      toast.show(message, 2200);
    }
  }

  function getAiPromptText(): string {
    const template = document.getElementById("aiPromptTemplate") as HTMLTemplateElement | null;
    if (!template) {
      return "";
    }
    return (template.content?.textContent || template.textContent || "").trim();
  }

  async function copyTextToClipboard(text: string): Promise<void> {
    if (
      typeof navigator !== "undefined" &&
      navigator.clipboard &&
      typeof navigator.clipboard.writeText === "function"
    ) {
      await navigator.clipboard.writeText(text);
      return;
    }

    const textarea = document.createElement("textarea");
    textarea.value = text;
    textarea.setAttribute("readonly", "readonly");
    textarea.style.position = "fixed";
    textarea.style.opacity = "0";
    document.body.appendChild(textarea);
    textarea.select();
    document.execCommand("copy");
    document.body.removeChild(textarea);
  }

  async function copyAiPrompt(): Promise<void> {
    const promptText = getAiPromptText();
    if (!promptText) {
      throw new Error("生成AIプロンプトが見つかりません");
    }
    await copyTextToClipboard(promptText);
    showToast("生成AIプロンプトをクリップボードにコピーしました");
    setStatus("生成AIプロンプトをクリップボードにコピーしました");
  }

  function setSvgPreviewMarkup(markup: string): void {
    getElement<HTMLElement>("nativeSvgPreview").innerHTML = markup;
  }

  function updateSvgPreviewModeButtons(): void {
    const dailyButton = getElement<HTMLButtonElement>("previewDailySvgBtn");
    const weeklyButton = getElement<HTMLButtonElement>("previewWeeklySvgBtn");
    const monthlyButton = getElement<HTMLButtonElement>("previewMonthlySvgBtn");
    dailyButton.classList.toggle("md-button--primary", currentSvgPreviewMode === "daily");
    dailyButton.classList.toggle("md-button--tonal", currentSvgPreviewMode !== "daily");
    weeklyButton.classList.toggle("md-button--primary", currentSvgPreviewMode === "weekly");
    weeklyButton.classList.toggle("md-button--tonal", currentSvgPreviewMode !== "weekly");
    monthlyButton.classList.toggle("md-button--primary", currentSvgPreviewMode === "monthly");
    monthlyButton.classList.toggle("md-button--tonal", currentSvgPreviewMode !== "monthly");
  }

  function renderCurrentSvgPreviewMarkup(): void {
    if (currentSvgPreviewMode === "weekly") {
      if (currentWeeklyPreviewSvg) {
        setSvgPreviewMarkup(currentWeeklyPreviewSvg);
        return;
      }
      setSvgPreviewMarkup(`<div class="md-preview-empty">Weekly SVG を生成すると、ここにプレビューを表示します。</div>`);
      return;
    }
    if (currentSvgPreviewMode === "monthly") {
      if (currentMonthlyPreviewSvg) {
        setSvgPreviewMarkup(currentMonthlyPreviewSvg);
        return;
      }
      setSvgPreviewMarkup(`<div class="md-preview-empty">Monthly Calendar SVG を生成すると、ここにプレビューを表示します。</div>`);
      return;
    }
    if (currentNativeSvg) {
      setSvgPreviewMarkup(currentNativeSvg);
      return;
    }
    setSvgPreviewMarkup(`<div class="md-preview-empty">Daily SVG を生成すると、ここにプレビューを表示します。</div>`);
  }

  function updateSvgButton(): void {
    const nativeSvgButton = getElement<HTMLButtonElement>("downloadSvgBtn");
    const weeklySvgButton = getElement<HTMLButtonElement>("downloadWeeklySvgBtn");
    const monthlyWbsButton = getElement<HTMLButtonElement>("downloadMonthlyCalendarSvgBtn");
    const disabled = !currentModel;
    nativeSvgButton.disabled = disabled;
    weeklySvgButton.disabled = disabled;
    monthlyWbsButton.disabled = disabled;
  }

  function buildCurrentWbsOptions(model: ProjectModel): {
    holidayDates: string[];
    displayDaysBeforeBaseDate?: number;
    displayDaysAfterBaseDate?: number;
    useBusinessDaysForDisplayRange?: boolean;
    useBusinessDaysForProgressBand?: boolean;
  } {
    return {
      holidayDates: mikuprojectWbsXlsx.collectWbsHolidayDates(model),
      displayDaysBeforeBaseDate: parseWbsDisplayDaysBeforeBaseDate(),
      displayDaysAfterBaseDate: parseWbsDisplayDaysAfterBaseDate(),
      useBusinessDaysForDisplayRange: useBusinessDaysForWbsDisplayRange(),
      useBusinessDaysForProgressBand: useBusinessDaysForWbsProgressBand()
    };
  }

  function downloadBlob(blob: Blob, filename: string): void {
    const objectUrl = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = objectUrl;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    window.setTimeout(() => URL.revokeObjectURL(objectUrl), 0);
  }

  function formatTimestampCompact(date: Date): string {
    return mikuprojectMainUtil.formatTimestampCompact(date);
  }

  async function renderSvgPreview(): Promise<void> {
    if (!currentModel) {
      currentNativeSvg = "";
      currentWeeklyPreviewSvg = "";
      currentMonthlyPreviewSvg = "";
      updateSvgButton();
      renderCurrentSvgPreviewMarkup();
      return;
    }
    const wbsOptions = buildCurrentWbsOptions(currentModel);
    currentNativeSvg = mikuprojectNativeSvg.exportNativeSvg(currentModel, wbsOptions);
    currentWeeklyPreviewSvg = mikuprojectNativeSvg.exportWeeklyNativeSvg(currentModel, wbsOptions);
    const monthlyArchive = mikuprojectNativeSvg.exportMonthlyWbsCalendarSvgArchive(currentModel);
    currentMonthlyPreviewSvg = monthlyArchive.entries.length > 0
      ? monthlyArchive.entries.map((entry) => entry.svg).join("")
      : "";
    renderCurrentSvgPreviewMarkup();
    updateSvgButton();
  }

  function setSvgPreviewMode(mode: "daily" | "weekly" | "monthly"): void {
    currentSvgPreviewMode = mode;
    updateSvgPreviewModeButtons();
    renderCurrentSvgPreviewMarkup();
  }

  function buildCurrentOutputArchive(): { fileName: string; zipBytes: Uint8Array; entryCount: number } {
    const model = ensureCurrentModel();
    const xmlText = syncXmlTextFromModel(model);
    const now = new Date();
    const stamp = formatTimestampCompact(now);
    const dateOnlyStamp = [
      now.getFullYear(),
      String(now.getMonth() + 1).padStart(2, "0"),
      String(now.getDate()).padStart(2, "0")
    ].join("");
    const codec = new mikuprojectExcelIo.XlsxWorkbookCodec();
    const workbook = mikuprojectProjectXlsx.exportProjectWorkbook(model);
    const workbookJsonText = JSON.stringify(mikuprojectProjectWorkbookJson.exportProjectWorkbookJson(model), null, 2);
    const csvText = mikuprojectXml.exportCsvParentId(model);
    const wbsOptions = buildCurrentWbsOptions(model);
    const wbsWorkbook = mikuprojectWbsXlsx.exportWbsWorkbook(model, wbsOptions);
    const wbsMarkdown = mikuprojectWbsMarkdown.exportWbsMarkdown(model, wbsOptions);
    const dailySvg = mikuprojectNativeSvg.exportNativeSvg(model, wbsOptions);
    const weeklySvg = mikuprojectNativeSvg.exportWeeklyNativeSvg(model, wbsOptions);
    const monthlyArchive = mikuprojectNativeSvg.exportMonthlyWbsCalendarSvgArchive(model);
    const mermaidText = mikuprojectXml.exportMermaidGantt(model);
    const projectOverview = mikuprojectXml.exportProjectOverviewView(model) as { phases?: Array<{ uid?: string }> };
    const phaseDetailViewsFull = (projectOverview.phases || [])
      .map((phase) => phase?.uid)
      .filter((uid): uid is string => Boolean(uid))
      .map((phaseUid) => mikuprojectXml.exportPhaseDetailView(model, phaseUid, { mode: "full" }));
    const aiBundle = {
      view_type: "ai_projection_bundle",
      project_overview_view: projectOverview,
      phase_detail_views_full: phaseDetailViewsFull
    };
    const phaseDetailFull = mikuprojectXml.exportPhaseDetailView(model, undefined, { mode: "full" });
    const allReadmeText = [
      "mikuproject ALL ZIP",
      "",
      "GitHub: https://github.com/igapyon/mikuproject",
      "mikuproject is a local single-file web app that converts MS Project XML into XLSX, Markdown, SVG, Mermaid, and AI-facing JSON exports.",
      "",
      "This archive contains the main outputs generated from the current model.",
      "",
      "Files:",
      "- mikuproject-export-*.xml: regenerated MS Project XML",
      "- mikuproject-export-*.xlsx: workbook XLSX export",
      "- mikuproject-workbook-*.json: workbook JSON export",
      "- mikuproject-export-*.csv: CSV + ParentID export",
      "- mikuproject-wbs-*.xlsx: WBS workbook export",
      "- mikuproject-wbs-*.md: WBS Markdown export",
      "- mikuproject-wbs-daily-*.svg: daily WBS SVG export",
      "- mikuproject-wbs-weekly-*.svg: weekly WBS SVG export",
      "- mikuproject-wbs-mermaid-*.md: Mermaid gantt export",
      "- monthly-calendar/YYYY-MM.svg: month-by-month calendar SVG export",
      "- *.editjson: AI-facing projection exports"
    ].join("\n");

    const entries = [
      { name: "README.txt", data: mikuprojectMainUtil.encodeUtf8(`${allReadmeText}\n`) },
      { name: `mikuproject-export-${stamp}.xml`, data: mikuprojectMainUtil.encodeUtf8(`${xmlText}\n`) },
      { name: `mikuproject-export-${stamp}.xlsx`, data: codec.exportWorkbook(workbook) },
      { name: `mikuproject-workbook-${stamp}.json`, data: mikuprojectMainUtil.encodeUtf8(`${workbookJsonText}\n`) },
      { name: `mikuproject-export-${stamp}.csv`, data: mikuprojectMainUtil.encodeUtf8(`${csvText}\n`) },
      { name: `mikuproject-wbs-${stamp}.xlsx`, data: codec.exportWorkbook(wbsWorkbook) },
      { name: `mikuproject-wbs-${dateOnlyStamp}.md`, data: mikuprojectMainUtil.encodeUtf8(`${wbsMarkdown}\n`) },
      { name: `mikuproject-wbs-daily-${stamp}.svg`, data: mikuprojectMainUtil.encodeUtf8(dailySvg) },
      { name: `mikuproject-wbs-weekly-${stamp}.svg`, data: mikuprojectMainUtil.encodeUtf8(weeklySvg) },
      { name: `mikuproject-wbs-mermaid-${stamp}.md`, data: mikuprojectMainUtil.encodeUtf8(`\`\`\`mermaid\n${mermaidText}\n\`\`\`\n`) }
    ];
    for (const entry of monthlyArchive.entries) {
      entries.push({
        name: `monthly-calendar/${entry.fileName}`,
        data: mikuprojectMainUtil.encodeUtf8(entry.svg)
      });
    }
    entries.push(
      { name: "mikuproject-project-overview-view.editjson", data: mikuprojectMainUtil.encodeUtf8(`${JSON.stringify(projectOverview, null, 2)}\n`) },
      { name: "mikuproject-full-bundle.editjson", data: mikuprojectMainUtil.encodeUtf8(`${JSON.stringify(aiBundle, null, 2)}\n`) },
      { name: "mikuproject-phase-detail-view-full.editjson", data: mikuprojectMainUtil.encodeUtf8(`${JSON.stringify(phaseDetailFull, null, 2)}\n`) }
    );

    return {
      fileName: `mikuproject-all-${stamp}.zip`,
      zipBytes: mikuprojectMainUtil.packZipEntries(entries),
      entryCount: entries.length
    };
  }

  function downloadAllOutputs(): void {
    const archive = buildCurrentOutputArchive();
    downloadBlob(
      new Blob([archive.zipBytes], { type: "application/zip" }),
      archive.fileName
    );
    setStatus(`All 出力を保存しました (${archive.entryCount} 件, ZIP)`);
    showToast("All を保存しました");
    setActiveTab("output");
  }

  function setStatus(message: string): void {
    getElement<HTMLElement>("statusMessage").textContent = message;
  }

  function formatSaveStamp(date: Date): string {
    return mikuprojectMainUtil.formatSaveStamp(date);
  }

  function updateXmlSaveState(isDirty: boolean): void {
    const node = getElement<HTMLElement>("xmlSaveState");
    node.textContent = isDirty
      ? "XML 保存状態: 未保存"
      : `XML 保存状態: 保存済み (${lastSavedXmlStamp || "-"})`;
    node.classList.toggle("md-save-state--dirty", isDirty);
    node.classList.toggle("md-save-state--clean", !isDirty);
  }

  function markXmlDirty(): void {
    updateXmlSaveState(true);
  }

  function markXmlSavedCurrent(): void {
    lastSavedXmlText = getTextArea("xmlInput").value;
    lastSavedXmlStamp = formatSaveStamp(new Date());
    updateXmlSaveState(false);
  }

  function refreshXmlSaveState(): void {
    updateXmlSaveState(getTextArea("xmlInput").value !== lastSavedXmlText);
  }

  function syncXmlTextFromModel(model: ProjectModel): string {
    const xmlText = mikuprojectXml.exportMsProjectXml(model);
    getTextArea("xmlInput").value = xmlText;
    isXmlSourceDirty = false;
    refreshXmlSaveState();
    return xmlText;
  }

  function renderValidationIssues(issues: ValidationIssue[]): void {
    mikuprojectMainRender.renderValidationIssues(document, issues);
  }

  function renderImportWarnings(warnings: Array<{ message: string }>): void {
    mikuprojectMainRender.renderImportWarnings(document, warnings);
  }

  function renderXlsxImportSummary(changes: Array<{
    scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
    uid: string;
    label: string;
    field: string;
    before: string | number | boolean | undefined;
    after: string | number | boolean;
  }>): void {
    mikuprojectMainRender.renderXlsxImportSummary(document, changes);
  }

  function updateSummary(model: ProjectModel | null): void {
    mikuprojectMainRender.updateSummary(document, model, updateSvgButton);
  }

  function loadSample(): void {
    const sampleXml = mikuprojectXml.SAMPLE_XML;
    getTextArea("xmlInput").value = sampleXml;
    currentModel = mikuprojectXml.importMsProjectXml(sampleXml);
    isXmlSourceDirty = true;
    markXmlDirty();
    updateSummary(currentModel);
    renderValidationIssues(mikuprojectXml.validateProjectModel(currentModel));
    renderImportWarnings([]);
    renderXlsxImportSummary([]);
    updateSvgButton();
    setStatus("サンプル XML を読み込みました");
    setActiveTab("input");
  }

  async function importXmlFromFile(file: File | null | undefined): Promise<void> {
    if (!file) {
      return;
    }
    const xmlText = await file.text();
    getTextArea("xmlInput").value = xmlText;
    markXmlDirty();
    currentModel = mikuprojectXml.importMsProjectXml(xmlText);
    isXmlSourceDirty = false;
    const issues = mikuprojectXml.validateProjectModel(currentModel);
    updateSummary(currentModel);
    renderValidationIssues(issues);
    renderImportWarnings([]);
    renderXlsxImportSummary([]);
    setStatus(issues.length > 0 ? `XML ファイルを読み込んで解析しました。検証で ${issues.length} 件の問題があります` : "XML ファイルを読み込んで解析しました");
    showToast("XML を読み込んで解析しました");
    setActiveTab("transform", { skipTransformRefresh: true });
    await exportCurrentMermaid({ silent: true });
  }

  function ensureCurrentModel(): ProjectModel {
    if (currentModel) {
      return currentModel;
    }
    const xmlText = getTextArea("xmlInput").value.trim();
    if (!xmlText) {
      throw new Error("内部モデルがありません");
    }
    currentModel = mikuprojectXml.importMsProjectXml(xmlText);
    isXmlSourceDirty = false;
    return currentModel;
  }

  function parseCurrentXml(options: { silent?: boolean } = {}): void {
    const xmlText = getTextArea("xmlInput").value.trim();
    if (!xmlText) {
      setStatus("XML が空です");
      return;
    }
    currentModel = mikuprojectXml.importMsProjectXml(xmlText);
    isXmlSourceDirty = false;
    const issues = mikuprojectXml.validateProjectModel(currentModel);
    updateSummary(currentModel);
    renderValidationIssues(issues);
    renderImportWarnings([]);
    renderXlsxImportSummary([]);
    if (!options.silent) {
      setStatus(issues.length > 0 ? `XML を解析しました。検証で ${issues.length} 件の問題があります` : "XML を内部モデルへ変換しました");
      showToast("XML を解析しました");
    }
    setActiveTab("transform", { skipTransformRefresh: true });
  }

  async function exportCurrentMermaid(options: { silent?: boolean } = {}): Promise<void> {
    if (!currentModel) {
      setStatus("内部モデルがありません");
      return;
    }
    const mermaidText = mikuprojectXml.exportMermaidGantt(currentModel);
    getTextArea("mermaidOutput").value = mermaidText;
    await renderSvgPreview();
    if (!options.silent) {
      setStatus("内部モデルから Mermaid gantt を生成し、native SVG preview を更新しました");
      showToast("Mermaid を生成しました");
    }
    setActiveTab("transform", { skipTransformRefresh: true });
  }

  function exportCurrentCsv(): void {
    const model = ensureCurrentModel();
    syncXmlTextFromModel(model);
    const csvText = mikuprojectXml.exportCsvParentId(model);
    const now = new Date();
    const stamp = [
      now.getFullYear(),
      String(now.getMonth() + 1).padStart(2, "0"),
      String(now.getDate()).padStart(2, "0"),
      String(now.getHours()).padStart(2, "0"),
      String(now.getMinutes()).padStart(2, "0")
    ].join("");
    downloadBlob(
      new Blob([`${csvText}\n`], { type: "text/csv;charset=utf-8" }),
      `mikuproject-export-${stamp}.csv`
    );
    setStatus("内部モデルから CSV + ParentID を生成して保存しました");
    showToast("CSV を保存しました");
    setActiveTab("output");
  }

  function exportCurrentProjectOverviewView(): void {
    const model = ensureCurrentModel();
    syncXmlTextFromModel(model);
    const viewText = JSON.stringify(mikuprojectXml.exportProjectOverviewView(model), null, 2);
    getTextArea("projectOverviewOutput").value = viewText;
    downloadBlob(
      new Blob([`${viewText}\n`], { type: "application/json;charset=utf-8" }),
      "mikuproject-project-overview-view.editjson"
    );
    setStatus("project_overview_view を生成して保存しました");
    showToast("project_overview_view を保存しました");
    setActiveTab("output");
  }

  function exportCurrentAiProjectionBundle(): void {
    const model = ensureCurrentModel();
    syncXmlTextFromModel(model);
    const projectOverview = mikuprojectXml.exportProjectOverviewView(model) as {
      phases?: Array<{ uid?: string }>;
    };
    const phaseDetailViewsFull = (projectOverview.phases || [])
      .map((phase) => phase?.uid)
      .filter((uid): uid is string => Boolean(uid))
      .map((phaseUid) => mikuprojectXml.exportPhaseDetailView(model, phaseUid, { mode: "full" }));
    const bundle = {
      view_type: "ai_projection_bundle",
      project_overview_view: projectOverview,
      phase_detail_views_full: phaseDetailViewsFull
    };
    const bundleText = JSON.stringify(bundle, null, 2);
    getTextArea("aiBundleOutput").value = bundleText;
    downloadBlob(
      new Blob([`${bundleText}\n`], { type: "application/json;charset=utf-8" }),
      "mikuproject-full-bundle.editjson"
    );
    setStatus(`AI 連携用まとめ JSON を生成して保存しました (phase_detail_view full ${phaseDetailViewsFull.length} 件)`);
    showToast("AI 連携用まとめ JSON を保存しました");
    setActiveTab("output");
  }

  async function importProjectDraftFromText(): Promise<void> {
    const sourceText = getTextArea("projectDraftImportInput").value.trim();
    if (!sourceText) {
      throw new Error("project_draft_view JSON を入力してください");
    }
    const jsonText = mikuprojectAiJsonUtil.extractLastJsonBlock(sourceText);
    const draft = JSON.parse(jsonText);
    currentModel = mikuprojectXml.importProjectDraftView(draft);
    syncXmlTextFromModel(currentModel);
    const issues = mikuprojectXml.validateProjectModel(currentModel);
    updateSummary(currentModel);
    renderValidationIssues(issues);
    renderImportWarnings([]);
    renderXlsxImportSummary([]);
    await exportCurrentMermaid({ silent: true });
    setStatus(issues.length > 0 ? `project_draft_view を取り込みました。検証で ${issues.length} 件の問題があります` : "project_draft_view を取り込みました");
    showToast("project_draft_view を取り込みました");
    setActiveTab("transform", { skipTransformRefresh: true });
  }

  async function importPatchJsonFromSourceText(sourceText: string): Promise<void> {
    const trimmedSourceText = sourceText.trim();
    if (!trimmedSourceText) {
      throw new Error("Patch JSON を入力してください");
    }
    const documentLike = JSON.parse(mikuprojectAiJsonUtil.extractLastJsonBlock(trimmedSourceText));
    const baseModel = ensureCurrentModel();
    const result = mikuprojectProjectPatchJson.importProjectPatchJson(documentLike, baseModel);
    currentModel = result.model;
    const issues = mikuprojectXml.validateProjectModel(currentModel);
    updateSummary(currentModel);
    renderValidationIssues(issues);
    renderImportWarnings(result.warnings);
    renderXlsxImportSummary(result.changes);
    if (result.changes.length > 0) {
      getTextArea("xmlInput").value = mikuprojectXml.exportMsProjectXml(currentModel);
      markXmlDirty();
    }
    isXmlSourceDirty = false;
    const summaryText = result.changes.length > 0
      ? `Patch JSON を読み込んで ${result.changes.length} 件の変更を反映しました。XML は再生成済みで、必要なら XML Export で保存できます`
      : "Patch JSON に反映対象の変更はありませんでした。XML は未変更です";
    const warningText = result.warnings.length > 0 ? `。Patch JSON 取込で ${result.warnings.length} 件の warning を無視しました` : "";
    setStatus(issues.length > 0 ? `${summaryText}${warningText}。検証で ${issues.length} 件の問題があります` : `${summaryText}${warningText}`);
    showToast("Patch JSON を反映しました");
    setActiveTab("transform", { skipTransformRefresh: true });
    await exportCurrentMermaid({ silent: true });
  }

  async function importAiEditJsonFromText(): Promise<void> {
    const sourceText = getTextArea("projectDraftImportInput").value.trim();
    if (!sourceText) {
      throw new Error("project_draft_view または Patch JSON を入力してください");
    }
    const jsonText = mikuprojectAiJsonUtil.extractLastJsonBlock(sourceText);
    const documentLike = JSON.parse(jsonText);
    const kind = mikuprojectAiJsonUtil.detectJsonDocumentKind(documentLike);
    if (kind === "project_draft_view") {
      await importProjectDraftFromText();
      return;
    }
    if (kind === "patch_json") {
      await importPatchJsonFromSourceText(sourceText);
      return;
    }
    throw new Error("project_draft_view または Patch JSON を入力してください");
  }

  function loadProjectDraftSample(): void {
    const sampleDraftText = JSON.stringify(mikuprojectXml.SAMPLE_PROJECT_DRAFT_VIEW, null, 2);
    getTextArea("projectDraftImportInput").value = sampleDraftText;
    setStatus("サンプル project_draft_view を読み込みました");
    setActiveTab("input");
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
    const trimmedSourceText = sourceText.trim();
    if (!trimmedSourceText) {
      throw new Error("workbook JSON を入力してください");
    }
    const documentLike = JSON.parse(mikuprojectAiJsonUtil.extractLastJsonBlock(trimmedSourceText));
    const baseModel = ensureCurrentModel();
    const result = mikuprojectProjectWorkbookJson.importProjectWorkbookJson(documentLike, baseModel);
    currentModel = result.model;
    const issues = mikuprojectXml.validateProjectModel(currentModel);
    updateSummary(currentModel);
    renderValidationIssues(issues);
    renderImportWarnings(result.warnings);
    renderXlsxImportSummary(result.changes);
    if (result.changes.length > 0) {
      getTextArea("xmlInput").value = mikuprojectXml.exportMsProjectXml(currentModel);
      markXmlDirty();
    }
    isXmlSourceDirty = false;
    const summaryText = result.changes.length > 0
      ? `JSON を読み込んで ${result.changes.length} 件の変更を反映しました。XML は再生成済みで、必要なら XML Export で保存できます`
      : "JSON に反映対象の変更はありませんでした。XML は未変更です";
    const warningText = result.warnings.length > 0 ? `。JSON 取込で ${result.warnings.length} 件の warning を無視しました` : "";
    setStatus(issues.length > 0 ? `${summaryText}${warningText}。検証で ${issues.length} 件の問題があります` : `${summaryText}${warningText}`);
    showToast("JSON を反映しました");
    setActiveTab("transform", { skipTransformRefresh: true });
    await exportCurrentMermaid({ silent: true });
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
    const normalizedName = file.name.trim().toLowerCase();
    if (normalizedName.endsWith(".xml")) {
      await importXmlFromFile(file);
      return;
    }
    if (normalizedName.endsWith(".xlsx")) {
      await importXlsxFromFile(file);
      return;
    }
    if (normalizedName.endsWith(".csv")) {
      await importCsvFromFile(file);
      return;
    }
    if (normalizedName.endsWith(".editjson")) {
      const sourceText = await file.text();
      getTextArea("projectDraftImportInput").value = sourceText;
      await importAiEditJsonFromText();
      return;
    }
    if (normalizedName.endsWith(".json")) {
      const sourceText = await file.text();
      const documentLike = JSON.parse(mikuprojectAiJsonUtil.extractLastJsonBlock(sourceText));
      const kind = mikuprojectAiJsonUtil.detectJsonDocumentKind(documentLike);
      if (kind === "workbook_json") {
        await importWorkbookJsonFromSourceText(sourceText);
        return;
      }
      if (kind === "project_draft_view") {
        getTextArea("projectDraftImportInput").value = sourceText;
        await importProjectDraftFromText();
        return;
      }
      if (kind === "patch_json") {
        getTextArea("projectDraftImportInput").value = sourceText;
        await importPatchJsonFromSourceText(sourceText);
        return;
      }
      throw new Error("JSON の format / view_type を判別できません。workbook JSON、project_draft_view、または Patch JSON を指定してください");
    }
    throw new Error("対応していないファイル形式です。.xml / .xlsx / .json / .editjson / .csv を指定してください");
  }

  function exportCurrentPhaseDetailView(mode: "full" | "scoped" = "scoped"): void {
    const model = ensureCurrentModel();
    syncXmlTextFromModel(model);
    const requestedPhaseUid = getInput("phaseDetailUidInput").value.trim() || undefined;
    const requestedRootUid = mode === "scoped" ? getInput("phaseDetailRootUidInput").value.trim() || undefined : undefined;
    const maxDepthText = getInput("phaseDetailMaxDepthInput").value.trim();
    const requestedMaxDepth = mode === "scoped" && maxDepthText !== "" ? Number.parseInt(maxDepthText, 10) : undefined;
    if (typeof requestedMaxDepth === "number" && (!Number.isFinite(requestedMaxDepth) || requestedMaxDepth < 0)) {
      throw new Error("max depth は 0 以上の整数で指定してください");
    }
    const view = mikuprojectXml.exportPhaseDetailView(model, requestedPhaseUid, {
      mode,
      rootUid: requestedRootUid,
      maxDepth: requestedMaxDepth
    }) as {
      phase?: { uid?: string };
      scope?: { mode?: "full" | "scoped"; root_uid?: string | null; max_depth?: number | null };
    };
    if (view.phase?.uid) {
      getInput("phaseDetailUidInput").value = view.phase.uid;
    }
    getInput("phaseDetailRootUidInput").value = view.scope?.root_uid || "";
    getInput("phaseDetailMaxDepthInput").value = typeof view.scope?.max_depth === "number" ? String(view.scope.max_depth) : "";
    const viewText = JSON.stringify(view, null, 2);
    getTextArea("phaseDetailOutput").value = viewText;
    const phaseSuffix = view.phase?.uid ? `-${view.phase.uid}` : "";
    const modeSuffix = view.scope?.mode === "scoped" ? "-scoped" : "-full";
    const rootSuffix = view.scope?.root_uid ? `-root-${view.scope.root_uid}` : "";
    const depthSuffix = typeof view.scope?.max_depth === "number" ? `-depth-${view.scope.max_depth}` : "";
    downloadBlob(
      new Blob([`${viewText}\n`], { type: "application/json;charset=utf-8" }),
      `mikuproject-phase-detail-view${phaseSuffix}${modeSuffix}${rootSuffix}${depthSuffix}.editjson`
    );
    setStatus(`phase_detail_view (${view.scope?.mode === "scoped" ? "scoped" : "full"}) を生成して保存しました`);
    showToast(`phase_detail_view (${view.scope?.mode === "scoped" ? "scoped" : "full"}) を保存しました`);
    setActiveTab("output");
  }

  function exportCurrentXlsx(): void {
    const model = ensureCurrentModel();
    syncXmlTextFromModel(model);
    const workbook = mikuprojectProjectXlsx.exportProjectWorkbook(model);
    const codec = new mikuprojectExcelIo.XlsxWorkbookCodec();
    const bytes = codec.exportWorkbook(workbook);
    const now = new Date();
    const stamp = [
      now.getFullYear(),
      String(now.getMonth() + 1).padStart(2, "0"),
      String(now.getDate()).padStart(2, "0"),
      String(now.getHours()).padStart(2, "0"),
      String(now.getMinutes()).padStart(2, "0")
    ].join("");
    downloadBlob(
      new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
      `mikuproject-export-${stamp}.xlsx`
    );
    setStatus("XLSX ファイルをエクスポートしました");
    showToast("XLSX を保存しました");
    setActiveTab("output");
  }

  function exportCurrentWorkbookJson(): void {
    const model = ensureCurrentModel();
    syncXmlTextFromModel(model);
    const jsonText = JSON.stringify(mikuprojectProjectWorkbookJson.exportProjectWorkbookJson(model), null, 2);
    const now = new Date();
    const stamp = [
      now.getFullYear(),
      String(now.getMonth() + 1).padStart(2, "0"),
      String(now.getDate()).padStart(2, "0"),
      String(now.getHours()).padStart(2, "0"),
      String(now.getMinutes()).padStart(2, "0")
    ].join("");
    getTextArea("workbookJsonOutput").value = jsonText;
    downloadBlob(
      new Blob([`${jsonText}\n`], { type: "application/json;charset=utf-8" }),
      `mikuproject-workbook-${stamp}.json`
    );
    setStatus("XLSX 相当の workbook JSON を生成して保存しました");
    showToast("JSON を保存しました");
    setActiveTab("output");
  }

  function exportCurrentWbsXlsx(): void {
    const model = ensureCurrentModel();
    syncXmlTextFromModel(model);
    const options = buildCurrentWbsOptions(model);
    const workbook = mikuprojectWbsXlsx.exportWbsWorkbook(model, options);
    const codec = new mikuprojectExcelIo.XlsxWorkbookCodec();
    const bytes = codec.exportWorkbook(workbook);
    const now = new Date();
    const stamp = [
      now.getFullYear(),
      String(now.getMonth() + 1).padStart(2, "0"),
      String(now.getDate()).padStart(2, "0"),
      String(now.getHours()).padStart(2, "0"),
      String(now.getMinutes()).padStart(2, "0")
    ].join("");
    downloadBlob(
      new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
      `mikuproject-wbs-${stamp}.xlsx`
    );
    const displayRangeText = options.displayDaysBeforeBaseDate !== undefined || options.displayDaysAfterBaseDate !== undefined
      ? ` / 表示期間 営業日 基準日前 ${options.displayDaysBeforeBaseDate || 0} 日, 基準日後 ${options.displayDaysAfterBaseDate || 0} 日`
      : "";
    const progressBandText = " / 進捗帯 営業日";
    setStatus(`WBS XLSX ファイルをエクスポートしました${options.holidayDates.length > 0 ? ` (祝日 ${options.holidayDates.length} 件)` : ""}${displayRangeText}${progressBandText}`);
    showToast("WBS XLSX を保存しました");
    setActiveTab("output");
  }

  async function importXlsxFromFile(file: File | null | undefined): Promise<void> {
    if (!file) {
      return;
    }
    const baseModel = ensureCurrentModel();
    const bytes = new Uint8Array(await file.arrayBuffer());
    const codec = new mikuprojectExcelIo.XlsxWorkbookCodec();
    const workbook = typeof codec.importWorkbookAsync === "function"
      ? await codec.importWorkbookAsync(bytes)
      : codec.importWorkbook(bytes);
    const result = mikuprojectProjectXlsx.importProjectWorkbookDetailed(workbook, baseModel);
    currentModel = result.model;
    const issues = mikuprojectXml.validateProjectModel(currentModel);
    updateSummary(currentModel);
    renderValidationIssues(issues);
    renderImportWarnings([]);
    renderXlsxImportSummary(result.changes);
    if (result.changes.length > 0) {
      getTextArea("xmlInput").value = mikuprojectXml.exportMsProjectXml(currentModel);
      markXmlDirty();
    }
    const summaryText = result.changes.length > 0
      ? `XLSX を読み込んで ${result.changes.length} 件の変更を反映しました。XML は再生成済みで、必要なら XML Export で保存できます`
      : "XLSX に反映対象の変更はありませんでした。XML は未変更です";
    isXmlSourceDirty = false;
    setStatus(issues.length > 0 ? `${summaryText}。検証で ${issues.length} 件の問題があります` : summaryText);
    showToast("XLSX を反映しました");
    setActiveTab("transform", { skipTransformRefresh: true });
    await exportCurrentMermaid({ silent: true });
  }

  async function importCsvFromFile(file: File | null | undefined): Promise<void> {
    if (!file) {
      return;
    }
    const csvText = (await file.text()).trim();
    if (!csvText) {
      setStatus("CSV が空です");
      return;
    }
    currentModel = mikuprojectXml.importCsvParentId(csvText);
    isXmlSourceDirty = false;
    const issues = mikuprojectXml.validateProjectModel(currentModel);
    updateSummary(currentModel);
    renderValidationIssues(issues);
    renderImportWarnings([]);
    renderXlsxImportSummary([]);
    setStatus(issues.length > 0 ? `CSV ファイルを読み込んで解析しました。検証で ${issues.length} 件の問題があります` : "CSV + ParentID を内部モデルへ変換しました");
    showToast("CSV を読み込みました");
    setActiveTab("transform", { skipTransformRefresh: true });
    await exportCurrentMermaid({ silent: true });
  }

  function downloadCurrentXml(): void {
    const model = ensureCurrentModel();
    const xmlText = syncXmlTextFromModel(model);
    const blob = new Blob([`${xmlText}\n`], { type: "application/xml;charset=utf-8" });
    const objectUrl = URL.createObjectURL(blob);
    const link = document.createElement("a");
    const now = new Date();
    const stamp = [
      now.getFullYear(),
      String(now.getMonth() + 1).padStart(2, "0"),
      String(now.getDate()).padStart(2, "0"),
      String(now.getHours()).padStart(2, "0"),
      String(now.getMinutes()).padStart(2, "0")
    ].join("");
    link.href = objectUrl;
    link.download = `mikuproject-export-${stamp}.xml`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    window.setTimeout(() => URL.revokeObjectURL(objectUrl), 0);
    markXmlSavedCurrent();
    setStatus("XML ファイルをエクスポートしました");
    showToast("XML を保存しました");
    setActiveTab("output");
  }

  async function downloadCurrentSvg(): Promise<void> {
    const model = ensureCurrentModel();
    syncXmlTextFromModel(model);
    const mermaidText = mikuprojectXml.exportMermaidGantt(model);
    getTextArea("mermaidOutput").value = mermaidText;
    await renderSvgPreview();
    if (!currentNativeSvg) {
      setStatus("出力する SVG がありません");
      return;
    }
    const now = new Date();
    const stamp = [
      now.getFullYear(),
      String(now.getMonth() + 1).padStart(2, "0"),
      String(now.getDate()).padStart(2, "0"),
      String(now.getHours()).padStart(2, "0"),
      String(now.getMinutes()).padStart(2, "0")
    ].join("");
    downloadBlob(
      new Blob([currentNativeSvg], { type: "image/svg+xml;charset=utf-8" }),
      `mikuproject-wbs-daily-${stamp}.svg`
    );
    setStatus("Daily SVG を保存しました");
    showToast("Daily SVG を保存しました");
    setActiveTab("output");
  }

  function downloadCurrentWeeklySvg(): void {
    const model = ensureCurrentModel();
    syncXmlTextFromModel(model);
    const weeklySvg = mikuprojectNativeSvg.exportWeeklyNativeSvg(model, buildCurrentWbsOptions(model));
    if (!weeklySvg) {
      setStatus("出力する Weekly SVG がありません");
      return;
    }
    const now = new Date();
    const stamp = [
      now.getFullYear(),
      String(now.getMonth() + 1).padStart(2, "0"),
      String(now.getDate()).padStart(2, "0"),
      String(now.getHours()).padStart(2, "0"),
      String(now.getMinutes()).padStart(2, "0")
    ].join("");
    downloadBlob(
      new Blob([weeklySvg], { type: "image/svg+xml;charset=utf-8" }),
      `mikuproject-wbs-weekly-${stamp}.svg`
    );
    setStatus("Weekly SVG を保存しました");
    showToast("Weekly SVG を保存しました");
    setActiveTab("output");
  }

  function downloadCurrentMonthlyWbsSvgZip(): void {
    const model = ensureCurrentModel();
    syncXmlTextFromModel(model);
    const archive = mikuprojectNativeSvg.exportMonthlyWbsCalendarSvgArchive(model);
    if (!archive.entries.length || archive.zipBytes.byteLength === 0) {
      throw new Error("出力する月別 WBS カレンダー SVG がありません");
    }
    const now = new Date();
    const stamp = [
      now.getFullYear(),
      String(now.getMonth() + 1).padStart(2, "0"),
      String(now.getDate()).padStart(2, "0"),
      String(now.getHours()).padStart(2, "0"),
      String(now.getMinutes()).padStart(2, "0")
    ].join("");
    downloadBlob(
      new Blob([archive.zipBytes], { type: "application/zip" }),
      `mikuproject-monthly-wbs-calendar-${stamp}.zip`
    );
    setStatus(`Monthly Calendar SVG を保存しました (${archive.entries.length} か月分, ZIP)`);
    showToast("Monthly Calendar SVG を保存しました");
    setActiveTab("output");
  }

  function downloadCurrentMermaidMarkdown(): void {
    const model = ensureCurrentModel();
    syncXmlTextFromModel(model);
    const mermaidText = mikuprojectXml.exportMermaidGantt(model);
    getTextArea("mermaidOutput").value = mermaidText;
    const now = new Date();
    const stamp = [
      now.getFullYear(),
      String(now.getMonth() + 1).padStart(2, "0"),
      String(now.getDate()).padStart(2, "0"),
      String(now.getHours()).padStart(2, "0"),
      String(now.getMinutes()).padStart(2, "0")
    ].join("");
    const markdownText = `\`\`\`mermaid\n${mermaidText}\n\`\`\`\n`;
    downloadBlob(
      new Blob([markdownText], { type: "text/markdown;charset=utf-8" }),
      `mikuproject-wbs-mermaid-${stamp}.md`
    );
    setStatus("Mermaid Markdown を保存しました");
    showToast("Mermaid Markdown を保存しました");
    setActiveTab("output");
  }

  function downloadCurrentWbsMarkdown(): void {
    const model = ensureCurrentModel();
    syncXmlTextFromModel(model);
    const options = buildCurrentWbsOptions(model);
    const markdownText = mikuprojectWbsMarkdown.exportWbsMarkdown(model, options);
    const now = new Date();
    const stamp = [
      now.getFullYear(),
      String(now.getMonth() + 1).padStart(2, "0"),
      String(now.getDate()).padStart(2, "0")
    ].join("");
    downloadBlob(
      new Blob([markdownText], { type: "text/markdown;charset=utf-8" }),
      `mikuproject-wbs-${stamp}.md`
    );
    setStatus("WBS Markdown を保存しました");
    showToast("WBS Markdown を保存しました");
    setActiveTab("output");
  }

  function runRoundTripCheck(): void {
    if (!currentModel) {
      parseCurrentXml();
      if (!currentModel) {
        return;
      }
    }
    const exportedXml = mikuprojectXml.exportMsProjectXml(currentModel);
    const reparsedModel = mikuprojectXml.importMsProjectXml(exportedXml);
    const validationIssues = mikuprojectXml.validateProjectModel(reparsedModel);
    renderValidationIssues(validationIssues);
    if (validationIssues.some((issue) => issue.level === "error")) {
      throw new Error(validationIssues.map((issue) => issue.message).join("\n"));
    }
    const normalizedLeft = JSON.stringify(mikuprojectXml.normalizeProjectModel(currentModel));
    const normalizedRight = JSON.stringify(mikuprojectXml.normalizeProjectModel(reparsedModel));
    if (normalizedLeft !== normalizedRight) {
      throw new Error("再読込後の内部モデルが一致しません");
    }
    setStatus("再読込テストに成功しました");
    showToast("再読込テスト成功");
    setActiveTab("transform");
  }

  function bindEvents(): void {
    getElement<HTMLButtonElement>("loadSampleBtn").addEventListener("click", loadSample);
    getElement<HTMLInputElement>("importFileInput").addEventListener("click", (event) => {
      const input = event.target as HTMLInputElement | null;
      if (input) {
        input.value = "";
      }
    });
    getElement<HTMLButtonElement>("importFileBtn").addEventListener("click", () => {
      const input = getElement<HTMLInputElement>("importFileInput");
      input.value = "";
      input.click();
    });
    getElement<HTMLButtonElement>("downloadAllOutputsBtn").addEventListener("click", () => {
      try {
        downloadAllOutputs();
      } catch (error) {
        console.error("[mikuproject] all outputs download failed", error);
        setStatus(error instanceof Error ? error.message : "All 出力保存に失敗しました");
      }
    });
    getElement<HTMLButtonElement>("previewDailySvgBtn").addEventListener("click", () => {
      setSvgPreviewMode("daily");
    });
    getElement<HTMLButtonElement>("previewWeeklySvgBtn").addEventListener("click", () => {
      setSvgPreviewMode("weekly");
    });
    getElement<HTMLButtonElement>("previewMonthlySvgBtn").addEventListener("click", () => {
      setSvgPreviewMode("monthly");
    });
    getElement<HTMLButtonElement>("downloadSvgBtn").addEventListener("click", () => {
      void downloadCurrentSvg().catch((error) => {
        console.error("[mikuproject] native SVG download failed", error);
        setStatus(error instanceof Error ? error.message : "SVG 保存に失敗しました");
      });
    });
    getElement<HTMLButtonElement>("downloadWeeklySvgBtn").addEventListener("click", () => {
      try {
        downloadCurrentWeeklySvg();
      } catch (error) {
        console.error("[mikuproject] weekly SVG download failed", error);
        setStatus(error instanceof Error ? error.message : "Weekly SVG 保存に失敗しました");
      }
    });
    getElement<HTMLButtonElement>("downloadMonthlyCalendarSvgBtn").addEventListener("click", () => {
      try {
        downloadCurrentMonthlyWbsSvgZip();
      } catch (error) {
        console.error("[mikuproject] monthly WBS SVG download failed", error);
        setStatus(error instanceof Error ? error.message : "月別 WBS カレンダー SVG 保存に失敗しました");
      }
    });
    getElement<HTMLButtonElement>("exportMermaidMdBtn").addEventListener("click", () => {
      try {
        downloadCurrentMermaidMarkdown();
      } catch (error) {
        setStatus(error instanceof Error ? error.message : "Mermaid Markdown 保存に失敗しました");
      }
    });
    getElement<HTMLButtonElement>("exportCsvBtn").addEventListener("click", () => {
      try {
        exportCurrentCsv();
      } catch (error) {
        setStatus(error instanceof Error ? error.message : "CSV 生成に失敗しました");
      }
    });
    getElement<HTMLButtonElement>("exportProjectOverviewBtn").addEventListener("click", () => {
      try {
        exportCurrentProjectOverviewView();
      } catch (error) {
        setStatus(error instanceof Error ? error.message : "project_overview_view 生成に失敗しました");
      }
    });
    getElement<HTMLButtonElement>("exportAiBundleBtn").addEventListener("click", () => {
      try {
        exportCurrentAiProjectionBundle();
      } catch (error) {
        setStatus(error instanceof Error ? error.message : "AI 連携用まとめ JSON 生成に失敗しました");
      }
    });
    getElement<HTMLButtonElement>("loadProjectDraftSampleBtn").addEventListener("click", loadProjectDraftSample);
    getElement<HTMLButtonElement>("copyAiPromptBtn").addEventListener("click", async () => {
      try {
        await copyAiPrompt();
      } catch (error) {
        setStatus(error instanceof Error ? error.message : "生成AIプロンプトのコピーに失敗しました");
      }
    });
    getElement<HTMLButtonElement>("importProjectDraftBtn").addEventListener("click", async () => {
      try {
        await importAiEditJsonFromText();
      } catch (error) {
        setStatus(error instanceof Error ? error.message : "編集用 JSON 取り込みに失敗しました");
      }
    });
    getElement<HTMLButtonElement>("exportPhaseDetailBtn").addEventListener("click", () => {
      try {
        exportCurrentPhaseDetailView("scoped");
      } catch (error) {
        setStatus(error instanceof Error ? error.message : "phase_detail_view 生成に失敗しました");
      }
    });
    getElement<HTMLButtonElement>("exportPhaseDetailFullBtn").addEventListener("click", () => {
      try {
        exportCurrentPhaseDetailView("full");
      } catch (error) {
        setStatus(error instanceof Error ? error.message : "phase_detail_view 生成に失敗しました");
      }
    });
    getElement<HTMLButtonElement>("exportXlsxBtn").addEventListener("click", () => {
      try {
        exportCurrentXlsx();
      } catch (error) {
        setStatus(error instanceof Error ? error.message : "XLSX エクスポートに失敗しました");
      }
    });
    getElement<HTMLButtonElement>("exportWorkbookJsonBtn").addEventListener("click", () => {
      try {
        exportCurrentWorkbookJson();
      } catch (error) {
        setStatus(error instanceof Error ? error.message : "JSON エクスポートに失敗しました");
      }
    });
    getElement<HTMLButtonElement>("exportWbsXlsxBtn").addEventListener("click", () => {
      try {
        exportCurrentWbsXlsx();
      } catch (error) {
        setStatus(error instanceof Error ? error.message : "WBS XLSX エクスポートに失敗しました");
      }
    });
    getElement<HTMLButtonElement>("exportWbsMdBtn").addEventListener("click", () => {
      try {
        downloadCurrentWbsMarkdown();
      } catch (error) {
        setStatus(error instanceof Error ? error.message : "WBS Markdown 保存に失敗しました");
      }
    });
    getElement<HTMLButtonElement>("downloadXmlBtn").addEventListener("click", () => {
      try {
        downloadCurrentXml();
      } catch (error) {
        setStatus(error instanceof Error ? error.message : "XML 保存に失敗しました");
      }
    });
    getElement<HTMLButtonElement>("roundTripBtn").addEventListener("click", () => {
      try {
        runRoundTripCheck();
      } catch (error) {
        setStatus(error instanceof Error ? error.message : "再読込テストに失敗しました");
      }
    });
    getElement<HTMLInputElement>("importFileInput").addEventListener("change", async (event) => {
      const input = event.target as HTMLInputElement | null;
      const file = input?.files && input.files[0];
      if (file) {
        setStatus(`${file.name} を読み込んでいます...`);
      }
      try {
        await importFromFile(file);
      } catch (error) {
        console.error("[mikuproject] file import failed", error);
        setStatus(error instanceof Error ? error.message : "ファイル読込に失敗しました");
      } finally {
        if (input) {
          input.value = "";
        }
      }
    });
    getTextArea("xmlInput").addEventListener("input", () => {
      isXmlSourceDirty = true;
      refreshXmlSaveState();
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
      exportCurrentMermaid: () => Promise<void>;
      renderValidationIssues: (issues: ValidationIssue[]) => void;
      renderXlsxImportSummary: (changes: Array<{
        scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
        uid: string;
        label: string;
        field: string;
        before: string | number | boolean | undefined;
        after: string | number | boolean;
      }>) => void;
    };
  }).__mikuprojectMainTestHooks = {
    parseCurrentXml,
    exportCurrentMermaid,
    renderValidationIssues,
    renderXlsxImportSummary
  };

  document.addEventListener("DOMContentLoaded", initialize);
})();
