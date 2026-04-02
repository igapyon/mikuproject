/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  type ImportChange = {
    scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
    uid: string;
    label: string;
    field: string;
    before: string | number | boolean | undefined;
    after: string | number | boolean;
  };

  function getElement<T extends HTMLElement>(doc: Document, id: string): T {
    const element = doc.getElementById(id);
    if (!element) {
      throw new Error(`Element not found: ${id}`);
    }
    return element as T;
  }

  function getTextArea(doc: Document, id: string): HTMLTextAreaElement {
    return getElement<HTMLTextAreaElement>(doc, id);
  }

  function renderPreviewList(doc: Document, containerId: string, items: string[]): void {
    const container = getElement<HTMLElement>(doc, containerId);
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

  function formatChangeValue(value: string | number | boolean | undefined): string {
    if (value === undefined) {
      return "(empty)";
    }
    return String(value);
  }

  function formatImportFieldLabel(field: string): string {
    const labelMap: Record<string, string> = {
      name: "Name",
      title: "Title",
      author: "Author",
      company: "Company",
      initials: "Initials",
      group: "Group",
      notes: "Notes",
      start: "Start",
      finish: "Finish",
      units: "Units",
      work: "Work",
      planned_start: "Start",
      planned_finish: "Finish",
      planned_duration: "Duration",
      planned_duration_hours: "DurationHours",
      is_milestone: "Milestone",
      parent_uid: "ParentUID",
      position: "Position",
      taskUid: "TaskUID",
      resourceUid: "ResourceUID",
      percentComplete: "PercentComplete",
      percentWorkComplete: "PercentWorkComplete",
      critical: "Critical",
      maxUnits: "MaxUnits",
      standardRate: "StandardRate",
      overtimeRate: "OvertimeRate",
      costPerUse: "CostPerUse",
      calendarUID: "CalendarUID",
      isBaseCalendar: "IsBaseCalendar",
      baseCalendarUID: "BaseCalendarUID",
      currentDate: "CurrentDate",
      startDate: "StartDate",
      finishDate: "FinishDate",
      statusDate: "StatusDate",
      minutesPerDay: "MinutesPerDay",
      minutesPerWeek: "MinutesPerWeek",
      daysPerMonth: "DaysPerMonth",
      scheduleFromStart: "ScheduleFromStart"
    };
    return labelMap[field] || field;
  }

  function formatImportSummaryHint(sourceLabel?: string): string {
    if (sourceLabel === "Patch JSON") {
      return "Patch JSON の部分適用結果です。反映後の XML は更新済みで、必要なら XML Export で保存できます。";
    }
    if (sourceLabel === "JSON Replace") {
      return "workbook JSON による全置換結果です。差分表示は取込前の project との差を示します。";
    }
    if (sourceLabel === "JSON Import") {
      return "workbook JSON の取込結果です。反映後の XML は更新済みで、必要なら XML Export で保存できます。";
    }
    if (sourceLabel === "XLSX Replace") {
      return "XLSX による全置換結果です。差分表示は取込前の project との差を示します。";
    }
    if (sourceLabel === "XLSX Import") {
      return "Excel 編集結果の取込内容です。反映後の XML は更新済みで、必要なら XML Export で保存できます。";
    }
    return "反映後の XML は更新済みです。必要なら XML Export で保存できます。";
  }

  function formatWarningDigest(
    warnings: Array<{ message: string; scope?: "project" | "tasks" | "resources" | "assignments" | "calendars"; uid?: string; label?: string }>,
    sourceLabel?: string
  ): string {
    if (warnings.length === 0) {
      return "";
    }
    const groupedTaskWarnings = Array.from(new Map(
      warnings
        .filter((warning) => warning.uid)
        .map((warning) => [`${warning.uid}:${warning.label || ""}`, `UID=${warning.uid} ${warning.label || ""}`])
    ).values());
    const warningPrefix = sourceLabel ? `${sourceLabel} warning` : "warning";
    if (groupedTaskWarnings.length === 0) {
      return `${warningPrefix} ${warnings.length} 件`;
    }
    return `${warningPrefix} ${warnings.length} 件 / ${groupedTaskWarnings.join(", ")}`;
  }

  function escapeHtml(value: string): string {
    return value
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function updateFeedbackVisibility(doc: Document): void {
    const stack = doc.querySelector<HTMLElement>(".md-feedback-stack");
    const validationIssues = getElement<HTMLElement>(doc, "validationIssues");
    const importWarnings = getElement<HTMLElement>(doc, "importWarnings");
    const xlsxImportSummary = getElement<HTMLElement>(doc, "xlsxImportSummary");
    const shouldShow = !validationIssues.classList.contains("md-hidden")
      || !importWarnings.classList.contains("md-hidden")
      || !xlsxImportSummary.classList.contains("md-hidden");
    stack?.classList.toggle("md-hidden", !shouldShow);
  }

  function renderValidationIssues(doc: Document, issues: ValidationIssue[]): void {
    const container = getElement<HTMLElement>(doc, "validationIssues");
    const label = container.previousElementSibling as HTMLElement | null;
    if (issues.length === 0) {
      container.classList.add("md-hidden");
      container.innerHTML = "";
      label?.classList.add("md-hidden");
      updateFeedbackVisibility(doc);
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
    updateFeedbackVisibility(doc);
  }

  function renderImportWarnings(
    doc: Document,
    warnings: Array<{ message: string; scope?: "project" | "tasks" | "resources" | "assignments" | "calendars"; uid?: string; label?: string }>,
    options: { sourceLabel?: string } = {}
  ): void {
    const container = getElement<HTMLElement>(doc, "importWarnings");
    const label = container.previousElementSibling as HTMLElement | null;
    if (warnings.length === 0) {
      container.classList.add("md-hidden");
      container.innerHTML = "";
      label?.classList.add("md-hidden");
      updateFeedbackVisibility(doc);
      return;
    }
    const genericWarnings = warnings.filter((warning) => !warning.uid);
    const groupedWarnings = new Map<string, { uid: string; label?: string; items: string[] }>();
    warnings
      .filter((warning) => warning.uid)
      .forEach((warning) => {
        const uid = warning.uid as string;
        const key = `${warning.scope || "generic"}:${uid}:${warning.label || ""}`;
        const current = groupedWarnings.get(key);
        if (current) {
          current.items.push(warning.message);
          return;
        }
        groupedWarnings.set(key, {
          uid,
          label: warning.label,
          items: [warning.message]
        });
      });
    const title = options.sourceLabel ? `${options.sourceLabel} warning` : "取込 warning";
    container.classList.remove("md-hidden");
    label?.classList.remove("md-hidden");
    container.innerHTML = `
      <div class="md-issues__title">${escapeHtml(title)}</div>
      ${genericWarnings.length > 0 ? `
        <ul class="md-issues__list">
          ${genericWarnings.map((warning) => `<li class="md-issues__item">${escapeHtml(warning.message)}</li>`).join("")}
        </ul>
      ` : ""}
      ${Array.from(groupedWarnings.values()).map((group) => `
        <div class="md-issues__section">
          <div class="md-issues__section-title">UID=${escapeHtml(group.uid)} ${escapeHtml(group.label || "")}</div>
          <ul class="md-issues__list">
            ${group.items.map((message) => `<li class="md-issues__item">${escapeHtml(message)}</li>`).join("")}
          </ul>
        </div>
      `).join("")}
    `;
    updateFeedbackVisibility(doc);
  }

  function renderXlsxImportSummary(
    doc: Document,
    changes: ImportChange[],
    options: {
      sourceLabel?: string;
      warnings?: Array<{ message: string; scope?: "project" | "tasks" | "resources" | "assignments" | "calendars"; uid?: string; label?: string }>;
    } = {}
  ): void {
    const container = getElement<HTMLElement>(doc, "xlsxImportSummary");
    const label = container.previousElementSibling as HTMLElement | null;
    if (changes.length === 0) {
      container.classList.add("md-hidden");
      container.innerHTML = "";
      label?.classList.add("md-hidden");
      updateFeedbackVisibility(doc);
      return;
    }
    const scopeLabel: Record<ImportChange["scope"], string> = {
      project: "Project",
      tasks: "Tasks",
      resources: "Resources",
      assignments: "Assignments",
      calendars: "Calendars"
    };
    const scopeCounts: Record<ImportChange["scope"], number> = {
      project: 0,
      tasks: 0,
      resources: 0,
      assignments: 0,
      calendars: 0
    };
    const scopeFieldCounts: Record<ImportChange["scope"], number> = {
      project: 0,
      tasks: 0,
      resources: 0,
      assignments: 0,
      calendars: 0
    };
    const groupedByScope = new Map<ImportChange["scope"], Array<{
      uid: string;
      label: string;
      items: Array<{
        field: string;
        before: string | number | boolean | undefined;
        after: string | number | boolean;
      }>;
    }>>();
    const groupedChanges = new Map<string, {
      scope: ImportChange["scope"];
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
        scopeFieldCounts[change.scope] += 1;
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
      scopeFieldCounts[change.scope] += 1;
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
    const allScopes = ["project", "tasks", "resources", "assignments", "calendars"] as const;
    const changedScopes = allScopes.filter((scope) => scopeCounts[scope] > 0);
    const unchangedScopes = allScopes.filter((scope) => scopeCounts[scope] === 0);
    container.classList.remove("md-hidden");
    label?.classList.remove("md-hidden");
    const title = options.sourceLabel ? `${options.sourceLabel} 反映結果` : "Import 反映結果";
    container.innerHTML = `
      <div class="md-xlsx-summary__title">${escapeHtml(title)}</div>
      <div class="md-xlsx-summary__counts">
        ${changedScopes.map((scope) => `<span class="md-xlsx-summary__count">${scopeLabel[scope]} ${scopeCounts[scope]} / ${scopeFieldCounts[scope]} fields</span>`).join("")}
        ${options.warnings && options.warnings.length > 0 ? `<span class="md-xlsx-summary__count md-xlsx-summary__count--warning">${escapeHtml(formatWarningDigest(options.warnings, options.sourceLabel))}</span>` : ""}
      </div>
      ${unchangedScopes.length > 0 ? `<div class="md-xlsx-summary__unchanged">変更なし: ${unchangedScopes.map((scope) => scopeLabel[scope]).join(", ")}</div>` : ""}
      ${changedScopes.map((scope) => `
        <div class="md-xlsx-summary__section">
          <div class="md-xlsx-summary__section-title">${scopeLabel[scope]}</div>
          <ul class="md-xlsx-summary__list">
            ${(groupedByScope.get(scope) || []).map((group) => `
              <li class="md-xlsx-summary__item">
                <div class="md-xlsx-summary__item-title-row">
                  <div class="md-xlsx-summary__item-title">UID=${group.uid} ${escapeHtml(group.label)}</div>
                  <div class="md-xlsx-summary__item-badge">${group.items.length} fields</div>
                </div>
                <div class="md-xlsx-summary__item-body">
                  <div class="md-xlsx-summary__change-head">
                    <span class="md-xlsx-summary__field">Field</span>
                    <span class="md-xlsx-summary__head-label">Before</span>
                    <span class="md-xlsx-summary__arrow"></span>
                    <span class="md-xlsx-summary__head-label">After</span>
                  </div>
                  ${group.items.map((item) => `
                    <div class="md-xlsx-summary__change-row">
                      <span class="md-xlsx-summary__field">${escapeHtml(formatImportFieldLabel(item.field))}</span>
                      <span class="md-xlsx-summary__value md-xlsx-summary__value--before">${escapeHtml(formatChangeValue(item.before))}</span>
                      <span class="md-xlsx-summary__arrow">→</span>
                      <span class="md-xlsx-summary__value md-xlsx-summary__value--after">${escapeHtml(formatChangeValue(item.after))}</span>
                    </div>
                  `).join("")}
                </div>
              </li>
            `).join("")}
          </ul>
        </div>
      `).join("")}
      <div class="md-xlsx-summary__hint">${escapeHtml(formatImportSummaryHint(options.sourceLabel))}</div>
    `;
    updateFeedbackVisibility(doc);
  }

  function updateSummary(doc: Document, model: ProjectModel | null, updateSvgButton: () => void): void {
    updateSvgButton();
    getElement<HTMLElement>(doc, "summaryProjectName").textContent = model?.project.name || "-";
    getElement<HTMLElement>(doc, "summaryTaskCount").textContent = String(model?.tasks.length || 0);
    getElement<HTMLElement>(doc, "summaryResourceCount").textContent = String(model?.resources.length || 0);
    getElement<HTMLElement>(doc, "summaryAssignmentCount").textContent = String(model?.assignments.length || 0);
    getElement<HTMLElement>(doc, "summaryCalendarCount").textContent = String(model?.calendars.length || 0);
    getTextArea(doc, "modelOutput").value = model ? JSON.stringify(model, null, 2) : "";
    renderPreviewList(doc, "projectPreview", model ? [`
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
    renderPreviewList(doc, "taskPreview", model ? model.tasks.map((task) => `
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
    renderPreviewList(doc, "resourcePreview", model ? model.resources.map((resource) => `
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
    renderPreviewList(doc, "assignmentPreview", model ? model.assignments.map((assignment) => `
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
    renderPreviewList(doc, "calendarPreview", model ? model.calendars.map((calendar) => `
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

  (globalThis as typeof globalThis & {
    __mikuprojectMainRender?: {
      renderValidationIssues: (doc: Document, issues: ValidationIssue[]) => void;
      renderImportWarnings: (doc: Document, warnings: Array<{ message: string; scope?: "project" | "tasks" | "resources" | "assignments" | "calendars"; uid?: string; label?: string }>, options?: { sourceLabel?: string }) => void;
      renderXlsxImportSummary: (doc: Document, changes: ImportChange[], options?: {
        sourceLabel?: string;
        warnings?: Array<{ message: string; scope?: "project" | "tasks" | "resources" | "assignments" | "calendars"; uid?: string; label?: string }>;
      }) => void;
      updateSummary: (doc: Document, model: ProjectModel | null, updateSvgButton: () => void) => void;
    };
  }).__mikuprojectMainRender = {
    renderValidationIssues,
    renderImportWarnings,
    renderXlsxImportSummary,
    updateSummary
  };
})();
