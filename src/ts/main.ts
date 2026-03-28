(() => {
  const mikuprojectXml = (globalThis as typeof globalThis & {
    __mikuprojectXml?: {
      SAMPLE_XML: string;
      importMsProjectXml: (xmlText: string) => ProjectModel;
      importCsvParentId: (csvText: string) => ProjectModel;
      exportMsProjectXml: (model: ProjectModel) => string;
      exportMermaidGantt: (model: ProjectModel) => string;
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

  const mermaidApi = (globalThis as typeof globalThis & {
    mermaid?: {
      initialize: (config: Record<string, unknown>) => void;
      render: (id: string, source: string) => Promise<{ svg: string }>;
    };
  }).mermaid;

  let currentModel: ProjectModel | null = null;
  let currentMermaidSvg = "";
  let mermaidRenderCount = 0;
  let lastSavedXmlText = "";
  let lastSavedXmlStamp = "";

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

  function parseWbsAdditionalHolidayDates(): string[] {
    return parseHolidayDateList(getTextArea("wbsExtraHolidayDatesInput").value.trim());
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
    return getInput("wbsBusinessDayRangeInput").checked;
  }

  function useBusinessDaysForWbsProgressBand(): boolean {
    return getInput("wbsBusinessDayProgressInput").checked;
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
      getTextArea("wbsExtraHolidayDatesInput").value = "";
      updateWbsHolidaySummary([]);
      return;
    }
    const holidayDates = mikuprojectWbsXlsx.collectWbsHolidayDates(model);
    input.value = holidayDates.join("\n");
    updateWbsHolidaySummary(holidayDates);
  }

  function resetWbsHolidayDatesInput(): void {
    const model = ensureCurrentModel();
    const holidayDates = mikuprojectWbsXlsx.collectWbsHolidayDates(model);
    getTextArea("wbsHolidayDatesInput").value = holidayDates.join("\n");
    getTextArea("wbsExtraHolidayDatesInput").value = "";
    updateWbsHolidaySummary(holidayDates);
    setStatus(`WBS 祝日入力を既定値へ戻しました${holidayDates.length > 0 ? ` (${holidayDates.length} 件)` : ""}`);
    showToast("WBS 祝日を既定値へ戻しました");
  }

  function showToast(message: string): void {
    const toast = document.getElementById("toast") as (HTMLElement & { show?: (text: string, duration?: number) => void }) | null;
    if (toast && typeof toast.show === "function") {
      toast.show(message, 2200);
    }
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
    getElement<HTMLButtonElement>("downloadMermaidSvgBtn").disabled = !currentMermaidSvg;
  }

  function activateTopTab(tabId: string): void {
    const buttons = Array.from(document.querySelectorAll<HTMLElement>(".md-top-tab[data-tab]"));
    const panels = Array.from(document.querySelectorAll<HTMLElement>(".md-tab-panel[data-tab-panel]"));
    buttons.forEach((button) => {
      const isActive = button.dataset.tab === tabId;
      button.classList.toggle("is-active", isActive);
      button.setAttribute("aria-selected", isActive ? "true" : "false");
    });
    panels.forEach((panel) => {
      panel.hidden = panel.dataset.tabPanel !== tabId;
    });
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

  async function renderMermaidPreview(source: string): Promise<void> {
    if (!mermaidApi) {
      currentMermaidSvg = "";
      updateMermaidSvgButton();
      setMermaidPreviewMarkup(`<div class="md-preview-empty">Mermaid ライブラリを読み込めなかったため、プレビューできません。</div>`);
      setMermaidError("Mermaid ライブラリが利用できません。");
      return;
    }

    clearMermaidError();
    const renderId = `mikuprojectMermaidRender${++mermaidRenderCount}`;
    mermaidApi.initialize({
      startOnLoad: false,
      securityLevel: "strict",
      theme: "default"
    });

    try {
      const result = await mermaidApi.render(renderId, source);
      currentMermaidSvg = normalizeSvgForXml(result.svg);
      setMermaidPreviewMarkup(currentMermaidSvg);
      updateMermaidSvgButton();
    } catch (error) {
      currentMermaidSvg = "";
      updateMermaidSvgButton();
      setMermaidPreviewMarkup(`<div class="md-preview-empty">Mermaid のプレビューを表示できませんでした。</div>`);
      const message = error instanceof Error ? error.message : String(error);
      setMermaidError(`SVG プレビュー生成に失敗しました: ${message}`);
      throw error;
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
    const xlsxImportSummary = getElement<HTMLElement>("xlsxImportSummary");
    const shouldShow = !validationIssues.classList.contains("md-hidden") || !xlsxImportSummary.classList.contains("md-hidden");
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
    getTextArea("xmlInput").value = mikuprojectXml.SAMPLE_XML;
    markXmlDirty();
    setStatus("サンプル XML を読み込みました");
  }

  async function importXmlFromFile(file: File | null | undefined): Promise<void> {
    if (!file) {
      return;
    }
    const xmlText = await file.text();
    getTextArea("xmlInput").value = xmlText;
    markXmlDirty();
    currentModel = mikuprojectXml.importMsProjectXml(xmlText);
    const issues = mikuprojectXml.validateProjectModel(currentModel);
    updateSummary(currentModel);
    renderValidationIssues(issues);
    renderXlsxImportSummary([]);
    setStatus(issues.length > 0 ? `XML ファイルを読み込んで解析しました。検証で ${issues.length} 件の問題があります` : "XML ファイルを読み込んで解析しました");
    showToast("XML を読み込んで解析しました");
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
    return currentModel;
  }

  function parseCurrentXml(): void {
    const xmlText = getTextArea("xmlInput").value.trim();
    if (!xmlText) {
      setStatus("XML が空です");
      return;
    }
    currentModel = mikuprojectXml.importMsProjectXml(xmlText);
    const issues = mikuprojectXml.validateProjectModel(currentModel);
    updateSummary(currentModel);
    renderValidationIssues(issues);
    renderXlsxImportSummary([]);
    setStatus(issues.length > 0 ? `XML を解析しました。検証で ${issues.length} 件の問題があります` : "XML を内部モデルへ変換しました");
    showToast("XML を解析しました");
  }

  function exportCurrentModel(): void {
    if (!currentModel) {
      setStatus("内部モデルがありません");
      return;
    }
    getTextArea("xmlInput").value = mikuprojectXml.exportMsProjectXml(currentModel);
    markXmlDirty();
    renderValidationIssues([]);
    renderXlsxImportSummary([]);
    setStatus("内部モデルから XML を再生成しました");
    showToast("XML を再生成しました");
  }

  async function exportCurrentMermaid(): Promise<void> {
    if (!currentModel) {
      setStatus("内部モデルがありません");
      return;
    }
    const mermaidText = mikuprojectXml.exportMermaidGantt(currentModel);
    getTextArea("mermaidOutput").value = mermaidText;
    await renderMermaidPreview(mermaidText);
    setStatus("内部モデルから Mermaid gantt を生成し、SVG プレビューを更新しました");
    showToast("Mermaid を生成しました");
  }

  function exportCurrentCsv(): void {
    if (!currentModel) {
      setStatus("内部モデルがありません");
      return;
    }
    getTextArea("csvOutput").value = mikuprojectXml.exportCsvParentId(currentModel);
    setStatus("内部モデルから CSV + ParentID を生成しました");
    showToast("CSV を生成しました");
  }

  function exportCurrentXlsx(): void {
    const model = ensureCurrentModel();
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
  }

  function exportCurrentWbsXlsx(): void {
    const model = ensureCurrentModel();
    const defaultHolidayDates = parseWbsDefaultHolidayDates();
    const additionalHolidayDates = parseWbsAdditionalHolidayDates();
    const displayDaysBeforeBaseDate = parseWbsDisplayDaysBeforeBaseDate();
    const displayDaysAfterBaseDate = parseWbsDisplayDaysAfterBaseDate();
    const useBusinessDaysForDisplayRange = useBusinessDaysForWbsDisplayRange();
    const useBusinessDaysForProgressBand = useBusinessDaysForWbsProgressBand();
    const effectiveHolidayDates = Array.from(new Set([...defaultHolidayDates, ...additionalHolidayDates]));
    const workbook = mikuprojectWbsXlsx.exportWbsWorkbook(model, {
      holidayDates: effectiveHolidayDates,
      displayDaysBeforeBaseDate,
      displayDaysAfterBaseDate,
      useBusinessDaysForDisplayRange,
      useBusinessDaysForProgressBand
    });
    if (defaultHolidayDates.length === 0 && effectiveHolidayDates.length > 0) {
      getTextArea("wbsHolidayDatesInput").value = effectiveHolidayDates.join("\n");
    }
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
      ? ` / 表示期間 ${useBusinessDaysForDisplayRange ? "営業日" : "暦日"} 基準日前 ${displayDaysBeforeBaseDate || 0} 日, 基準日後 ${displayDaysAfterBaseDate || 0} 日`
      : "";
    const progressBandText = useBusinessDaysForProgressBand ? " / 進捗帯 営業日" : "";
    setStatus(`WBS XLSX ファイルをエクスポートしました${effectiveHolidayDates.length > 0 ? ` (祝日 ${effectiveHolidayDates.length} 件)` : ""}${displayRangeText}${progressBandText}`);
    showToast("WBS XLSX を保存しました");
  }

  async function importXlsxFromFile(file: File | null | undefined): Promise<void> {
    if (!file) {
      return;
    }
    const baseModel = ensureCurrentModel();
    const bytes = new Uint8Array(await file.arrayBuffer());
    const codec = new mikuprojectExcelIo.XlsxWorkbookCodec();
    const workbook = codec.importWorkbook(bytes);
    const result = mikuprojectProjectXlsx.importProjectWorkbookDetailed(workbook, baseModel);
    currentModel = result.model;
    const issues = mikuprojectXml.validateProjectModel(currentModel);
    updateSummary(currentModel);
    renderValidationIssues(issues);
    renderXlsxImportSummary(result.changes);
    if (result.changes.length > 0) {
      getTextArea("xmlInput").value = mikuprojectXml.exportMsProjectXml(currentModel);
      markXmlDirty();
    }
    const summaryText = result.changes.length > 0
      ? `XLSX を読み込んで ${result.changes.length} 件の変更を反映しました。XML は再生成済みで、必要なら XML Export で保存できます`
      : "XLSX に反映対象の変更はありませんでした。XML は未変更です";
    setStatus(issues.length > 0 ? `${summaryText}。検証で ${issues.length} 件の問題があります` : summaryText);
    showToast("XLSX を反映しました");
  }

  function parseCurrentCsv(): void {
    const csvText = getTextArea("csvInput").value.trim();
    if (!csvText) {
      setStatus("CSV が空です");
      return;
    }
    currentModel = mikuprojectXml.importCsvParentId(csvText);
    const issues = mikuprojectXml.validateProjectModel(currentModel);
    updateSummary(currentModel);
    renderValidationIssues(issues);
    renderXlsxImportSummary([]);
    setStatus(issues.length > 0 ? `CSV を解析しました。検証で ${issues.length} 件の問題があります` : "CSV + ParentID を内部モデルへ変換しました");
    showToast("CSV を解析しました");
  }

  function downloadCurrentXml(): void {
    const xmlText = getTextArea("xmlInput").value.trim();
    if (!xmlText) {
      setStatus("出力する XML がありません");
      return;
    }
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
  }

  function downloadCurrentMermaidSvg(): void {
    if (!currentMermaidSvg) {
      setStatus("出力する SVG がありません");
      return;
    }
    downloadBlob(new Blob([currentMermaidSvg], { type: "image/svg+xml;charset=utf-8" }), "mikuproject-mermaid.svg");
    setStatus("Mermaid SVG を保存しました");
    showToast("SVG を保存しました");
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
  }

  function bindEvents(): void {
    Array.from(document.querySelectorAll<HTMLButtonElement>(".md-top-tab[data-tab]")).forEach((button) => {
      button.addEventListener("click", () => {
        activateTopTab(button.dataset.tab || "input");
      });
    });
    getElement<HTMLButtonElement>("loadSampleBtn").addEventListener("click", loadSample);
    getElement<HTMLButtonElement>("importXmlBtn").addEventListener("click", () => {
      getElement<HTMLInputElement>("importXmlInput").click();
    });
    getElement<HTMLButtonElement>("parseXmlBtn").addEventListener("click", () => {
      try {
        parseCurrentXml();
      } catch (error) {
        setStatus(error instanceof Error ? error.message : "XML 解析に失敗しました");
      }
    });
    getElement<HTMLButtonElement>("exportXmlBtn").addEventListener("click", () => {
      try {
        exportCurrentModel();
      } catch (error) {
        setStatus(error instanceof Error ? error.message : "XML 再生成に失敗しました");
      }
    });
    getElement<HTMLButtonElement>("exportMermaidBtn").addEventListener("click", () => {
      void exportCurrentMermaid().catch((error) => {
        setStatus(error instanceof Error ? error.message : "Mermaid 生成に失敗しました");
      });
    });
    getElement<HTMLButtonElement>("downloadMermaidSvgBtn").addEventListener("click", () => {
      try {
        downloadCurrentMermaidSvg();
      } catch (error) {
        setStatus(error instanceof Error ? error.message : "SVG 保存に失敗しました");
      }
    });
    getElement<HTMLButtonElement>("exportCsvBtn").addEventListener("click", () => {
      try {
        exportCurrentCsv();
      } catch (error) {
        setStatus(error instanceof Error ? error.message : "CSV 生成に失敗しました");
      }
    });
    getElement<HTMLButtonElement>("exportXlsxBtn").addEventListener("click", () => {
      try {
        exportCurrentXlsx();
      } catch (error) {
        setStatus(error instanceof Error ? error.message : "XLSX エクスポートに失敗しました");
      }
    });
    getElement<HTMLButtonElement>("exportWbsXlsxBtn").addEventListener("click", () => {
      try {
        exportCurrentWbsXlsx();
      } catch (error) {
        setStatus(error instanceof Error ? error.message : "WBS XLSX エクスポートに失敗しました");
      }
    });
    getElement<HTMLButtonElement>("resetWbsHolidayDatesBtn").addEventListener("click", () => {
      try {
        resetWbsHolidayDatesInput();
      } catch (error) {
        setStatus(error instanceof Error ? error.message : "WBS 祝日入力のリセットに失敗しました");
      }
    });
    getElement<HTMLButtonElement>("parseCsvBtn").addEventListener("click", () => {
      try {
        parseCurrentCsv();
      } catch (error) {
        setStatus(error instanceof Error ? error.message : "CSV 解析に失敗しました");
      }
    });
    getElement<HTMLButtonElement>("importXlsxBtn").addEventListener("click", () => {
      getElement<HTMLInputElement>("importXlsxInput").click();
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
    getElement<HTMLInputElement>("importXmlInput").addEventListener("change", async (event) => {
      const input = event.target as HTMLInputElement | null;
      const file = input?.files && input.files[0];
      try {
        await importXmlFromFile(file);
      } catch (error) {
        setStatus(error instanceof Error ? error.message : "XML 読み込みに失敗しました");
      } finally {
        if (input) {
          input.value = "";
        }
      }
    });
    getElement<HTMLInputElement>("importXlsxInput").addEventListener("change", async (event) => {
      const input = event.target as HTMLInputElement | null;
      const file = input?.files && input.files[0];
      try {
        await importXlsxFromFile(file);
      } catch (error) {
        setStatus(error instanceof Error ? error.message : "XLSX 読み込みに失敗しました");
      } finally {
        if (input) {
          input.value = "";
        }
      }
    });
    getTextArea("xmlInput").addEventListener("input", () => {
      refreshXmlSaveState();
    });
  }

  function initialize(): void {
    bindEvents();
    activateTopTab("input");
    updateSummary(null);
    renderValidationIssues([]);
    renderXlsxImportSummary([]);
    updateMermaidSvgButton();
    clearMermaidError();
    loadSample();
  }

  document.addEventListener("DOMContentLoaded", initialize);
})();
