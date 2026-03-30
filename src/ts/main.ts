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
    };
  }).__mikuprojectNativeSvg;

  if (!mikuprojectNativeSvg) {
    throw new Error("mikuproject native SVG module is not loaded");
  }

  const mermaidApi = (globalThis as typeof globalThis & {
    mermaid?: {
      initialize: (config: Record<string, unknown>) => void;
      render: (id: string, source: string) => Promise<{ svg: string }>;
    };
  }).mermaid;

  let currentModel: ProjectModel | null = null;
  let currentNativeSvg = "";
  let mermaidRenderCount = 0;
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

  function parseHolidayDateList(raw: string): string[] {
    if (!raw) {
      return [];
    }
    const seen = new Set<string>();
    const holidays: string[] = [];
    for (const token of raw.split(/[\s,、;]+/)) {
      const value = token.trim();
      if (!value) {
        continue;
      }
      const match = value.match(/^(\d{4}-\d{2}-\d{2})/);
      if (!match) {
        continue;
      }
      const dateText = match[1];
      if (seen.has(dateText)) {
        continue;
      }
      seen.add(dateText);
      holidays.push(dateText);
    }
    return holidays;
  }

  function parseWbsDefaultHolidayDates(): string[] {
    return parseHolidayDateList(getTextArea("wbsHolidayDatesInput").value.trim());
  }

  function parseOptionalNonNegativeInteger(raw: string): number | undefined {
    const value = raw.trim();
    if (!value) {
      return undefined;
    }
    const parsed = Number(value);
    if (!Number.isFinite(parsed)) {
      return undefined;
    }
    return Math.max(0, Math.floor(parsed));
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

  function updateWbsHolidaySummary(holidayDates: string[]): void {
    const summary = getElement<HTMLElement>("wbsHolidaySummary");
    if (holidayDates.length === 0) {
      summary.textContent = "既定祝日: 0 件";
      return;
    }
    summary.textContent = `既定祝日: ${holidayDates.length} 件 (${holidayDates.join(", ")})`;
  }

  function syncWbsHolidayDatesInput(model: ProjectModel | null): void {
    const input = getTextArea("wbsHolidayDatesInput");
    if (!model) {
      input.value = "";
      updateWbsHolidaySummary([]);
      return;
    }
    const holidayDates = mikuprojectWbsXlsx.collectWbsHolidayDates(model);
    input.value = holidayDates.join("\n");
    updateWbsHolidaySummary(holidayDates);
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

  function setMermaidError(message: string): void {
    const errorNode = getElement<HTMLElement>("mermaidSvgError");
    errorNode.textContent = message;
    errorNode.classList.remove("md-hidden");
  }

  function clearMermaidError(): void {
    const errorNode = getElement<HTMLElement>("mermaidSvgError");
    errorNode.textContent = "";
    errorNode.classList.add("md-hidden");
  }

  function setMermaidPreviewMarkup(markup: string): void {
    getElement<HTMLElement>("mermaidSvgPreview").innerHTML = markup;
  }

  function updateMermaidSvgButton(): void {
    getElement<HTMLButtonElement>("downloadMermaidSvgBtn").disabled = !currentModel;
  }

  function buildCurrentWbsOptions(model: ProjectModel): {
    holidayDates: string[];
    displayDaysBeforeBaseDate?: number;
    displayDaysAfterBaseDate?: number;
    useBusinessDaysForDisplayRange?: boolean;
    useBusinessDaysForProgressBand?: boolean;
  } {
    syncWbsHolidayDatesInput(model);
    return {
      holidayDates: parseWbsDefaultHolidayDates(),
      displayDaysBeforeBaseDate: parseWbsDisplayDaysBeforeBaseDate(),
      displayDaysAfterBaseDate: parseWbsDisplayDaysAfterBaseDate(),
      useBusinessDaysForDisplayRange: useBusinessDaysForWbsDisplayRange(),
      useBusinessDaysForProgressBand: useBusinessDaysForWbsProgressBand()
    };
  }

  function normalizeSvgForXml(svgText: string): string {
    if (!svgText) {
      return "";
    }

    const candidate = svgText
      .replace(/<br\s*>/gi, "<br/>")
      .replace(/<br([^/>]*)><\/br>/gi, "<br$1/>");

    try {
      const parsed = new DOMParser().parseFromString(candidate, "image/svg+xml");
      if (parsed.querySelector("parsererror")) {
        return candidate;
      }
      return new XMLSerializer().serializeToString(parsed.documentElement);
    } catch (_error) {
      return candidate;
    }
  }

  function applyMermaidSvgTheme(svgText: string): string {
    if (!svgText) {
      return svgText;
    }

    const styleBlock = [
      "<style>",
      ".section0, .section2, .section4, .section6, .section8, rect.section0, rect.section2, rect.section4, rect.section6, rect.section8, .section0 rect, .section2 rect, .section4 rect, .section6 rect, .section8 rect, g.section0 rect, g.section2 rect, g.section4 rect, g.section6 rect, g.section8 rect { fill: #d8efe8 !important; }",
      ".section1, .section3, .section5, .section7, .section9, rect.section1, rect.section3, rect.section5, rect.section7, rect.section9, .section1 rect, .section3 rect, .section5 rect, .section7 rect, .section9 rect, g.section1 rect, g.section3 rect, g.section5 rect, g.section7 rect, g.section9 rect { fill: #ffe8c7 !important; }",
      ".task, .task0, .task1, .task2, .task3, .task4 { fill: #8f95e8 !important; stroke: #5d63cf !important; }",
      ".active, .active0, .active1, .active2, .active3, .active4 { fill: #7fc8a9 !important; stroke: #3f8f72 !important; }",
      ".done, .done0, .done1, .done2, .done3, .done4 { fill: #5ba98f !important; stroke: #2e6e5b !important; }",
      ".milestone, .milestone0, .milestone1, .milestone2, .milestone3 { fill: #666666 !important; stroke: #444444 !important; }",
      ".grid .tick line, .tick line { stroke: #707b94 !important; }",
      ".today { stroke: #ff3b30 !important; }",
      ".taskText, .taskTextOutsideRight, .taskTextOutsideLeft, .sectionTitle, .titleText, text { fill: #1d2740; }",
      "</style>"
    ].join("");

    if (svgText.includes(".task0") || svgText.includes(".section0")) {
      return svgText.replace(/(<svg\b[^>]*>)/i, `$1${styleBlock}`);
    }
    return svgText;
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

  function getMermaidRenderConfig(): Record<string, unknown> {
    return {
      startOnLoad: false,
      securityLevel: "strict",
      theme: "default",
      themeVariables: {
        sectionBkgColor: "#eaf6f3",
        altSectionBkgColor: "#fff3df",
        sectionBkgColor2: "#eaf6f3",
        sectionBkgColor3: "#fff3df",
        sectionBkgColor4: "#eaf6f3",
        taskBkgColor: "#8f95e8",
        taskBorderColor: "#5d63cf",
        taskTextColor: "#1d2740",
        activeTaskBkgColor: "#7fc8a9",
        activeTaskBorderColor: "#3f8f72",
        doneTaskBkgColor: "#5ba98f",
        doneTaskBorderColor: "#2e6e5b",
        milestoneBkgColor: "#666666",
        milestoneBorderColor: "#444444",
        gridColor: "#707b94",
        lineColor: "#707b94",
        todayLineColor: "#ff3b30"
      }
    };
  }

  async function renderMermaidPreview(_source: string): Promise<void> {
    if (!currentModel) {
      currentNativeSvg = "";
      updateMermaidSvgButton();
      setMermaidPreviewMarkup(`<div class="md-preview-empty">SVG を生成すると、ここにプレビューを表示します。</div>`);
      return;
    }
    clearMermaidError();
    currentNativeSvg = mikuprojectNativeSvg.exportNativeSvg(currentModel, buildCurrentWbsOptions(currentModel));
    setMermaidPreviewMarkup(currentNativeSvg);
    updateMermaidSvgButton();

    if (mermaidApi) {
      const renderId = `mikuprojectMermaidRender${++mermaidRenderCount}`;
      mermaidApi.initialize(getMermaidRenderConfig());
      try {
        await mermaidApi.render(renderId, _source);
      } catch (_error) {
        // Mermaid text export remains useful even if native SVG preview is used.
      }
    }
  }

  function setStatus(message: string): void {
    getElement<HTMLElement>("statusMessage").textContent = message;
  }

  function formatSaveStamp(date: Date): string {
    return [
      date.getFullYear(),
      String(date.getMonth() + 1).padStart(2, "0"),
      String(date.getDate()).padStart(2, "0")
    ].join("-") + " " + [
      String(date.getHours()).padStart(2, "0"),
      String(date.getMinutes()).padStart(2, "0")
    ].join(":");
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

  function renderPreviewList(containerId: string, items: string[]): void {
    const container = getElement<HTMLElement>(containerId);
    if (items.length === 0) {
      container.innerHTML = `<div class="md-preview-empty">まだ表示できる項目がありません。</div>`;
      return;
    }
    container.innerHTML = items.join("");
  }

  function formatFirstBaselineSummary<T extends { baselines: Array<{ number?: number; start?: string; finish?: string; work?: string; cost?: number }> }>(item: T): string {
    const baseline = item.baselines[0];
    if (!baseline) {
      return "-";
    }
    return `#${baseline.number ?? "-"} ${baseline.start || "-"} -> ${baseline.finish || "-"} / Work=${baseline.work || "-"} / Cost=${baseline.cost ?? "-"}`;
  }

  function formatFirstTimephasedSummary<T extends { timephasedData: Array<{ type?: number; start?: string; finish?: string; unit?: number; value?: string }> }>(item: T): string {
    const timephasedData = item.timephasedData[0];
    if (!timephasedData) {
      return "-";
    }
    return `Type=${timephasedData.type ?? "-"} ${timephasedData.start || "-"} -> ${timephasedData.finish || "-"} / Unit=${timephasedData.unit ?? "-"} / Value=${timephasedData.value || "-"}`;
  }

  function formatFirstExtendedAttributeSummary<T extends { extendedAttributes: Array<{ fieldID?: string; value?: string }> }>(item: T): string {
    const attribute = item.extendedAttributes[0];
    if (!attribute) {
      return "-";
    }
    return `FieldID=${attribute.fieldID || "-"} / Value=${attribute.value || "-"}`;
  }

  function formatFirstProjectExtendedAttributeSummary(project: ProjectInfo): string {
    const attribute = project.extendedAttributes[0];
    if (!attribute) {
      return "-";
    }
    return `FieldID=${attribute.fieldID || "-"} / FieldName=${attribute.fieldName || "-"} / Alias=${attribute.alias || "-"}`;
  }

  function formatFirstOutlineCodeSummary(project: ProjectInfo): string {
    const outlineCode = project.outlineCodes[0];
    if (!outlineCode) {
      return "-";
    }
    return `FieldID=${outlineCode.fieldID || "-"} / FieldName=${outlineCode.fieldName || "-"} / Alias=${outlineCode.alias || "-"}`;
  }

  function formatFirstWbsMaskSummary(project: ProjectInfo): string {
    const wbsMask = project.wbsMasks[0];
    if (!wbsMask) {
      return "-";
    }
    return `Level=${wbsMask.level} / Mask=${wbsMask.mask || "-"} / Length=${wbsMask.length ?? "-"} / Sequence=${wbsMask.sequence ?? "-"}`;
  }

  function formatCalendarWeekDaySummary(calendar: CalendarModel): string {
    const weekDay = calendar.weekDays[0];
    if (!weekDay) {
      return "-";
    }
    const workingTimes = weekDay.workingTimes.length > 0
      ? weekDay.workingTimes.map((item) => `${item.fromTime}-${item.toTime}`).join(", ")
      : "-";
    return `DayType=${weekDay.dayType} / Working=${weekDay.dayWorking ? 1 : 0} / Times=${workingTimes}`;
  }

  function formatCalendarExceptionSummary(calendar: CalendarModel): string {
    const exception = calendar.exceptions[0];
    if (!exception) {
      return "-";
    }
    return `${exception.name || "(no name)"} ${exception.fromDate || "-"} -> ${exception.toDate || "-"} / Working=${exception.dayWorking ? 1 : 0}`;
  }

  function formatCalendarWorkWeekSummary(calendar: CalendarModel): string {
    const workWeek = calendar.workWeeks[0];
    if (!workWeek) {
      return "-";
    }
    return `${workWeek.name || "(no name)"} ${workWeek.fromDate || "-"} -> ${workWeek.toDate || "-"} / WeekDays=${workWeek.weekDays.length}`;
  }

  function formatCalendarReferenceSummary(model: ProjectModel, calendar: CalendarModel): string {
    const projectRefs = model.project.calendarUID === calendar.uid ? 1 : 0;
    const taskRefs = model.tasks.filter((task) => task.calendarUID === calendar.uid).length;
    const resourceRefs = model.resources.filter((resource) => resource.calendarUID === calendar.uid).length;
    const baseRefs = model.calendars.filter((item) => item.baseCalendarUID === calendar.uid).length;
    return `Project=${projectRefs} / Tasks=${taskRefs} / Resources=${resourceRefs} / BaseOf=${baseRefs}`;
  }

  function formatCalendarLink(model: ProjectModel, calendarUID?: string): string {
    if (!calendarUID) {
      return "-";
    }
    const calendar = model.calendars.find((item) => item.uid === calendarUID);
    return calendar ? `${calendarUID} (${calendar.name || "(no name)"})` : `${calendarUID} (missing)`;
  }

  function formatTaskLink(model: ProjectModel, taskUID?: string): string {
    if (!taskUID) {
      return "-";
    }
    const task = model.tasks.find((item) => item.uid === taskUID);
    return task ? `${taskUID} (${task.name || "(no name)"})` : `${taskUID} (missing)`;
  }

  function formatResourceLink(model: ProjectModel, resourceUID?: string): string {
    if (!resourceUID) {
      return "-";
    }
    const resource = model.resources.find((item) => item.uid === resourceUID);
    return resource ? `${resourceUID} (${resource.name || "(no name)"})` : `${resourceUID} (missing)`;
  }

  function renderValidationIssues(issues: ValidationIssue[]): void {
    const container = getElement<HTMLElement>("validationIssues");
    const label = container.previousElementSibling as HTMLElement | null;
    if (issues.length === 0) {
      container.classList.add("md-hidden");
      container.innerHTML = "";
      label?.classList.add("md-hidden");
      updateFeedbackVisibility();
      return;
    }
    const sections: ValidationIssue["scope"][] = ["project", "tasks", "resources", "assignments", "calendars"];
    const sectionLabels: Record<ValidationIssue["scope"], string> = {
      project: "Project",
      tasks: "Tasks",
      resources: "Resources",
      assignments: "Assignments",
      calendars: "Calendars"
    };
    container.classList.remove("md-hidden");
    label?.classList.remove("md-hidden");
    container.innerHTML = `
      <div class="md-issues__title">検証メッセージ</div>
      ${sections
        .map((scope) => {
          const scopedIssues = issues.filter((issue) => issue.scope === scope);
          if (scopedIssues.length === 0) {
            return "";
          }
          return `
            <div class="md-issues__section">
              <div class="md-issues__section-title">${sectionLabels[scope]}</div>
              <ul class="md-issues__list">
                ${scopedIssues.map((issue) => `<li class="md-issues__item">[${issue.level}] ${issue.message}</li>`).join("")}
              </ul>
            </div>
          `;
        })
        .join("")}
    `;
    updateFeedbackVisibility();
  }

  function renderImportWarnings(warnings: Array<{ message: string }>): void {
    const container = getElement<HTMLElement>("importWarnings");
    const label = container.previousElementSibling as HTMLElement | null;
    if (warnings.length === 0) {
      container.classList.add("md-hidden");
      container.innerHTML = "";
      label?.classList.add("md-hidden");
      updateFeedbackVisibility();
      return;
    }
    container.classList.remove("md-hidden");
    label?.classList.remove("md-hidden");
    container.innerHTML = `
      <div class="md-issues__title">取込 warning</div>
      <ul class="md-issues__list">
        ${warnings.map((warning) => `<li class="md-issues__item">${escapeHtml(warning.message)}</li>`).join("")}
      </ul>
    `;
    updateFeedbackVisibility();
  }

  function renderXlsxImportSummary(changes: Array<{
    scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
    uid: string;
    label: string;
    field: string;
    before: string | number | boolean | undefined;
    after: string | number | boolean;
  }>): void {
    const container = getElement<HTMLElement>("xlsxImportSummary");
    const label = container.previousElementSibling as HTMLElement | null;
    if (changes.length === 0) {
      container.classList.add("md-hidden");
      container.innerHTML = "";
      label?.classList.add("md-hidden");
      updateFeedbackVisibility();
      return;
    }
    const scopeLabel: Record<"project" | "tasks" | "resources" | "assignments" | "calendars", string> = {
      project: "Project",
      tasks: "Tasks",
      resources: "Resources",
      assignments: "Assignments",
      calendars: "Calendars"
    };
    const scopeCounts: Record<"project" | "tasks" | "resources" | "assignments" | "calendars", number> = {
      project: 0,
      tasks: 0,
      resources: 0,
      assignments: 0,
      calendars: 0
    };
    const groupedByScope = new Map<"project" | "tasks" | "resources" | "assignments" | "calendars", Array<{
      uid: string;
      label: string;
      items: Array<{
        field: string;
        before: string | number | boolean | undefined;
        after: string | number | boolean;
      }>;
    }>>();
    const groupedChanges = new Map<string, {
      scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
      uid: string;
      label: string;
      items: Array<{
        field: string;
        before: string | number | boolean | undefined;
        after: string | number | boolean;
      }>;
    }>();
    for (const change of changes) {
      const groupKey = `${change.scope}:${change.uid}:${change.label}`;
      const currentGroup = groupedChanges.get(groupKey);
      if (currentGroup) {
        currentGroup.items.push({
          field: change.field,
          before: change.before,
          after: change.after
        });
        continue;
      }
      groupedChanges.set(groupKey, {
        scope: change.scope,
        uid: change.uid,
        label: change.label,
        items: [{
          field: change.field,
          before: change.before,
          after: change.after
        }]
      });
      scopeCounts[change.scope] += 1;
    }
    for (const group of groupedChanges.values()) {
      const scopedGroups = groupedByScope.get(group.scope) || [];
      scopedGroups.push({
        uid: group.uid,
        label: group.label,
        items: group.items
      });
      groupedByScope.set(group.scope, scopedGroups);
    }
    const changedScopes = (["project", "tasks", "resources", "assignments", "calendars"] as const).filter((scope) => scopeCounts[scope] > 0);
    const unchangedScopes = (["project", "tasks", "resources", "assignments", "calendars"] as const).filter((scope) => scopeCounts[scope] === 0);
    container.classList.remove("md-hidden");
    label?.classList.remove("md-hidden");
    container.innerHTML = `
      <div class="md-xlsx-summary__title">XLSX Import 反映結果</div>
      <div class="md-xlsx-summary__counts">
        ${changedScopes.map((scope) => `<span class="md-xlsx-summary__count">${scopeLabel[scope]} ${scopeCounts[scope]}</span>`).join("")}
      </div>
      ${unchangedScopes.length > 0 ? `<div class="md-xlsx-summary__unchanged">変更なし: ${unchangedScopes.map((scope) => scopeLabel[scope]).join(", ")}</div>` : ""}
      ${changedScopes.map((scope) => `
        <div class="md-xlsx-summary__section">
          <div class="md-xlsx-summary__section-title">${scopeLabel[scope]}</div>
          <ul class="md-xlsx-summary__list">
            ${(groupedByScope.get(scope) || []).map((group) => `
              <li class="md-xlsx-summary__item">
                <div class="md-xlsx-summary__item-title">UID=${group.uid} ${escapeHtml(group.label)}</div>
                <div class="md-xlsx-summary__item-body">
                  ${group.items.map((item) => `${escapeHtml(item.field)}: ${escapeHtml(formatChangeValue(item.before))} -> ${escapeHtml(formatChangeValue(item.after))}`).join(" / ")}
                </div>
              </li>
            `).join("")}
          </ul>
        </div>
      `).join("")}
      <div class="md-xlsx-summary__hint">反映後の XML は更新済みです。必要なら XML Export で保存できます。</div>
    `;
    updateFeedbackVisibility();
  }

  function updateFeedbackVisibility(): void {
    const stack = document.querySelector<HTMLElement>(".md-feedback-stack");
    const validationIssues = getElement<HTMLElement>("validationIssues");
    const importWarnings = getElement<HTMLElement>("importWarnings");
    const xlsxImportSummary = getElement<HTMLElement>("xlsxImportSummary");
    const shouldShow = !validationIssues.classList.contains("md-hidden")
      || !importWarnings.classList.contains("md-hidden")
      || !xlsxImportSummary.classList.contains("md-hidden");
    stack?.classList.toggle("md-hidden", !shouldShow);
  }

  function formatChangeValue(value: string | number | boolean | undefined): string {
    if (value === undefined) {
      return "(empty)";
    }
    return String(value);
  }

  function escapeHtml(value: string): string {
    return value
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function updateSummary(model: ProjectModel | null): void {
    updateMermaidSvgButton();
    syncWbsHolidayDatesInput(model);
    getElement<HTMLElement>("summaryProjectName").textContent = model?.project.name || "-";
    getElement<HTMLElement>("summaryTaskCount").textContent = String(model?.tasks.length || 0);
    getElement<HTMLElement>("summaryResourceCount").textContent = String(model?.resources.length || 0);
    getElement<HTMLElement>("summaryAssignmentCount").textContent = String(model?.assignments.length || 0);
    getElement<HTMLElement>("summaryCalendarCount").textContent = String(model?.calendars.length || 0);
    getTextArea("modelOutput").value = model ? JSON.stringify(model, null, 2) : "";
    renderPreviewList("projectPreview", model ? [`
      <div class="md-preview-item">
        <div class="md-preview-item__title">${model.project.name || "(no name)"}</div>
        <div class="md-preview-item__meta">Title=${model.project.title || "-"}
Author=${model.project.author || "-"} / Company=${model.project.company || "-"}
Start=${model.project.startDate || "-"} / Finish=${model.project.finishDate || "-"}
Calendar=${formatCalendarLink(model, model.project.calendarUID)}
OutlineCodes=${model.project.outlineCodes.length} / WBSMasks=${model.project.wbsMasks.length} / Ext=${model.project.extendedAttributes.length}
OutlineCode1=${formatFirstOutlineCodeSummary(model.project)}
WBSMask1=${formatFirstWbsMaskSummary(model.project)}
Ext1=${formatFirstProjectExtendedAttributeSummary(model.project)}</div>
      </div>
    `] : []);
    renderPreviewList("taskPreview", model ? model.tasks.map((task) => `
      <div class="md-preview-item">
        <div class="md-preview-item__title">${task.name || "(no name)"}</div>
        <div class="md-preview-item__meta">UID=${task.uid} / ID=${task.id} / Outline=${task.outlineNumber || task.outlineLevel}
Calendar=${formatCalendarLink(model, task.calendarUID)}
Start=${task.start || "-"}
Finish=${task.finish || "-"}
Predecessors=${task.predecessors.map((item) => item.predecessorUid).join(", ") || "-"}
Ext=${task.extendedAttributes.length} / Baselines=${task.baselines.length} / Timephased=${task.timephasedData.length}
Ext1=${formatFirstExtendedAttributeSummary(task)}
Baseline1=${formatFirstBaselineSummary(task)}
Timephased1=${formatFirstTimephasedSummary(task)}</div>
      </div>
    `) : []);
    renderPreviewList("resourcePreview", model ? model.resources.map((resource) => `
      <div class="md-preview-item">
        <div class="md-preview-item__title">${resource.name || "(no name)"}</div>
        <div class="md-preview-item__meta">UID=${resource.uid} / ID=${resource.id}
Initials=${resource.initials || "-"}
Group=${resource.group || "-"}
Calendar=${formatCalendarLink(model, resource.calendarUID)}
Ext=${resource.extendedAttributes.length} / Baselines=${resource.baselines.length} / Timephased=${resource.timephasedData.length}
Ext1=${formatFirstExtendedAttributeSummary(resource)}
Baseline1=${formatFirstBaselineSummary(resource)}
Timephased1=${formatFirstTimephasedSummary(resource)}</div>
      </div>
    `) : []);
    renderPreviewList("assignmentPreview", model ? model.assignments.map((assignment) => `
      <div class="md-preview-item">
        <div class="md-preview-item__title">Assignment ${assignment.uid || "-"}</div>
        <div class="md-preview-item__meta">Task=${formatTaskLink(model, assignment.taskUid)}
Resource=${formatResourceLink(model, assignment.resourceUid)}
Start=${assignment.start || "-"}
Finish=${assignment.finish || "-"}
Ext=${assignment.extendedAttributes.length} / Baselines=${assignment.baselines.length} / Timephased=${assignment.timephasedData.length}
Ext1=${formatFirstExtendedAttributeSummary(assignment)}
Baseline1=${formatFirstBaselineSummary(assignment)}
Timephased1=${formatFirstTimephasedSummary(assignment)}</div>
      </div>
    `) : []);
    renderPreviewList("calendarPreview", model ? model.calendars.map((calendar) => `
      <div class="md-preview-item">
        <div class="md-preview-item__title">${calendar.name || "(no name)"}</div>
        <div class="md-preview-item__meta">UID=${calendar.uid}
Base=${calendar.isBaseCalendar ? 1 : 0} / Baseline=${calendar.isBaselineCalendar ? 1 : 0} / BaseCalendarUID=${calendar.baseCalendarUID || "-"}
WeekDays=${calendar.weekDays.length} / Exceptions=${calendar.exceptions.length} / WorkWeeks=${calendar.workWeeks.length}
Refs=${formatCalendarReferenceSummary(model, calendar)}
WeekDay1=${formatCalendarWeekDaySummary(calendar)}
Exception1=${formatCalendarExceptionSummary(calendar)}
WorkWeek1=${formatCalendarWorkWeekSummary(calendar)}</div>
      </div>
    `) : []);
  }

  function loadSample(): void {
    currentModel = null;
    getTextArea("xmlInput").value = mikuprojectXml.SAMPLE_XML;
    isXmlSourceDirty = true;
    markXmlDirty();
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
    await renderMermaidPreview(mermaidText);
    if (!options.silent) {
      setStatus("内部モデルから Mermaid gantt を生成し、SVG プレビューを更新しました");
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

  function extractLastJsonBlock(value: string): string {
    const matches = Array.from(value.matchAll(/```json\s*([\s\S]*?)```/g));
    if (matches.length > 0) {
      return matches.at(-1)?.[1]?.trim() || "";
    }
    return value.trim();
  }

  function detectJsonDocumentKind(documentLike: unknown): "workbook_json" | "project_draft_view" | undefined {
    if (!documentLike || typeof documentLike !== "object") {
      return undefined;
    }
    const candidate = documentLike as {
      format?: string;
      view_type?: string;
    };
    if (candidate.format === "mikuproject_workbook_json") {
      return "workbook_json";
    }
    if (candidate.view_type === "project_draft_view") {
      return "project_draft_view";
    }
    return undefined;
  }

  async function importProjectDraftFromText(): Promise<void> {
    const sourceText = getTextArea("projectDraftImportInput").value.trim();
    if (!sourceText) {
      throw new Error("project_draft_view JSON を入力してください");
    }
    const jsonText = extractLastJsonBlock(sourceText);
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
    const documentLike = JSON.parse(extractLastJsonBlock(trimmedSourceText));
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
      await importProjectDraftFromFile(file);
      return;
    }
    if (normalizedName.endsWith(".json")) {
      const sourceText = await file.text();
      const documentLike = JSON.parse(extractLastJsonBlock(sourceText));
      const kind = detectJsonDocumentKind(documentLike);
      if (kind === "workbook_json") {
        await importWorkbookJsonFromSourceText(sourceText);
        return;
      }
      if (kind === "project_draft_view") {
        getTextArea("projectDraftImportInput").value = sourceText;
        await importProjectDraftFromText();
        return;
      }
      throw new Error("JSON の format / view_type を判別できません。workbook JSON か project_draft_view を指定してください");
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
    const defaultHolidayDates = parseWbsDefaultHolidayDates();
    const displayDaysBeforeBaseDate = parseWbsDisplayDaysBeforeBaseDate();
    const displayDaysAfterBaseDate = parseWbsDisplayDaysAfterBaseDate();
    const useBusinessDaysForDisplayRange = useBusinessDaysForWbsDisplayRange();
    const useBusinessDaysForProgressBand = useBusinessDaysForWbsProgressBand();
    const workbook = mikuprojectWbsXlsx.exportWbsWorkbook(model, {
      holidayDates: defaultHolidayDates,
      displayDaysBeforeBaseDate,
      displayDaysAfterBaseDate,
      useBusinessDaysForDisplayRange,
      useBusinessDaysForProgressBand
    });
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
    const displayRangeText = displayDaysBeforeBaseDate !== undefined || displayDaysAfterBaseDate !== undefined
      ? ` / 表示期間 営業日 基準日前 ${displayDaysBeforeBaseDate || 0} 日, 基準日後 ${displayDaysAfterBaseDate || 0} 日`
      : "";
    const progressBandText = " / 進捗帯 営業日";
    setStatus(`WBS XLSX ファイルをエクスポートしました${defaultHolidayDates.length > 0 ? ` (祝日 ${defaultHolidayDates.length} 件)` : ""}${displayRangeText}${progressBandText}`);
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

  async function downloadCurrentMermaidSvg(): Promise<void> {
    const model = ensureCurrentModel();
    syncXmlTextFromModel(model);
    const mermaidText = mikuprojectXml.exportMermaidGantt(model);
    getTextArea("mermaidOutput").value = mermaidText;
    await renderMermaidPreview(mermaidText);
    if (!currentNativeSvg) {
      setStatus("出力する SVG がありません");
      return;
    }
    downloadBlob(new Blob([currentNativeSvg], { type: "image/svg+xml;charset=utf-8" }), "mikuproject-native.svg");
    setStatus("SVG を保存しました");
    showToast("SVG を保存しました");
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
      String(now.getDate()).padStart(2, "0")
    ].join("");
    const markdownText = `\`\`\`mermaid\n${mermaidText}\n\`\`\`\n`;
    downloadBlob(
      new Blob([markdownText], { type: "text/markdown;charset=utf-8" }),
      `mermaid-${stamp}.md`
    );
    setStatus("Mermaid Markdown を保存しました");
    showToast("Mermaid Markdown を保存しました");
    setActiveTab("output");
  }

  function downloadCurrentWbsMarkdown(): void {
    const model = ensureCurrentModel();
    syncXmlTextFromModel(model);
    const defaultHolidayDates = mikuprojectWbsXlsx.collectWbsHolidayDates(model);
    syncWbsHolidayDatesInput(model);
    const displayDaysBeforeBaseDate = parseOptionalNonNegativeInteger(getInput("wbsDisplayDaysBeforeInput").value);
    const displayDaysAfterBaseDate = parseOptionalNonNegativeInteger(getInput("wbsDisplayDaysAfterInput").value);
    const useBusinessDaysForDisplayRange = useBusinessDaysForWbsDisplayRange();
    const useBusinessDaysForProgressBand = useBusinessDaysForWbsProgressBand();
    const markdownText = mikuprojectWbsMarkdown.exportWbsMarkdown(model, {
      holidayDates: defaultHolidayDates,
      displayDaysBeforeBaseDate,
      displayDaysAfterBaseDate,
      useBusinessDaysForDisplayRange,
      useBusinessDaysForProgressBand
    });
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
    getElement<HTMLButtonElement>("downloadMermaidSvgBtn").addEventListener("click", () => {
      void downloadCurrentMermaidSvg().catch((error) => {
        setStatus(error instanceof Error ? error.message : "SVG 保存に失敗しました");
      });
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
        await importProjectDraftFromText();
      } catch (error) {
        setStatus(error instanceof Error ? error.message : "project_draft_view 取り込みに失敗しました");
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
    updateMermaidSvgButton();
    clearMermaidError();
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
      updateFeedbackVisibility: () => void;
    };
  }).__mikuprojectMainTestHooks = {
    parseCurrentXml,
    exportCurrentMermaid,
    renderValidationIssues,
    renderXlsxImportSummary,
    updateFeedbackVisibility
  };

  document.addEventListener("DOMContentLoaded", initialize);
})();
