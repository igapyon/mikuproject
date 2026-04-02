/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    function getElement(doc, id) {
        const element = doc.getElementById(id);
        if (!element) {
            throw new Error(`Element not found: ${id}`);
        }
        return element;
    }
    function getTextArea(doc, id) {
        return getElement(doc, id);
    }
    function renderPreviewList(doc, containerId, items) {
        const container = getElement(doc, containerId);
        if (items.length === 0) {
            container.innerHTML = `<div class="md-preview-empty">まだ表示できる項目がありません。</div>`;
            return;
        }
        container.innerHTML = items.join("");
    }
    function formatFirstBaselineSummary(item) {
        var _a, _b;
        const baseline = item.baselines[0];
        if (!baseline) {
            return "-";
        }
        return `#${(_a = baseline.number) !== null && _a !== void 0 ? _a : "-"} ${baseline.start || "-"} -> ${baseline.finish || "-"} / Work=${baseline.work || "-"} / Cost=${(_b = baseline.cost) !== null && _b !== void 0 ? _b : "-"}`;
    }
    function formatFirstTimephasedSummary(item) {
        var _a, _b;
        const timephasedData = item.timephasedData[0];
        if (!timephasedData) {
            return "-";
        }
        return `Type=${(_a = timephasedData.type) !== null && _a !== void 0 ? _a : "-"} ${timephasedData.start || "-"} -> ${timephasedData.finish || "-"} / Unit=${(_b = timephasedData.unit) !== null && _b !== void 0 ? _b : "-"} / Value=${timephasedData.value || "-"}`;
    }
    function formatFirstExtendedAttributeSummary(item) {
        const attribute = item.extendedAttributes[0];
        if (!attribute) {
            return "-";
        }
        return `FieldID=${attribute.fieldID || "-"} / Value=${attribute.value || "-"}`;
    }
    function formatFirstProjectExtendedAttributeSummary(project) {
        const attribute = project.extendedAttributes[0];
        if (!attribute) {
            return "-";
        }
        return `FieldID=${attribute.fieldID || "-"} / FieldName=${attribute.fieldName || "-"} / Alias=${attribute.alias || "-"}`;
    }
    function formatFirstOutlineCodeSummary(project) {
        const outlineCode = project.outlineCodes[0];
        if (!outlineCode) {
            return "-";
        }
        return `FieldID=${outlineCode.fieldID || "-"} / FieldName=${outlineCode.fieldName || "-"} / Alias=${outlineCode.alias || "-"}`;
    }
    function formatFirstWbsMaskSummary(project) {
        var _a, _b;
        const wbsMask = project.wbsMasks[0];
        if (!wbsMask) {
            return "-";
        }
        return `Level=${wbsMask.level} / Mask=${wbsMask.mask || "-"} / Length=${(_a = wbsMask.length) !== null && _a !== void 0 ? _a : "-"} / Sequence=${(_b = wbsMask.sequence) !== null && _b !== void 0 ? _b : "-"}`;
    }
    function formatCalendarWeekDaySummary(calendar) {
        const weekDay = calendar.weekDays[0];
        if (!weekDay) {
            return "-";
        }
        const workingTimes = weekDay.workingTimes.length > 0
            ? weekDay.workingTimes.map((item) => `${item.fromTime}-${item.toTime}`).join(", ")
            : "-";
        return `DayType=${weekDay.dayType} / Working=${weekDay.dayWorking ? 1 : 0} / Times=${workingTimes}`;
    }
    function formatCalendarExceptionSummary(calendar) {
        const exception = calendar.exceptions[0];
        if (!exception) {
            return "-";
        }
        return `${exception.name || "(no name)"} ${exception.fromDate || "-"} -> ${exception.toDate || "-"} / Working=${exception.dayWorking ? 1 : 0}`;
    }
    function formatCalendarWorkWeekSummary(calendar) {
        const workWeek = calendar.workWeeks[0];
        if (!workWeek) {
            return "-";
        }
        return `${workWeek.name || "(no name)"} ${workWeek.fromDate || "-"} -> ${workWeek.toDate || "-"} / WeekDays=${workWeek.weekDays.length}`;
    }
    function formatCalendarReferenceSummary(model, calendar) {
        const projectRefs = model.project.calendarUID === calendar.uid ? 1 : 0;
        const taskRefs = model.tasks.filter((task) => task.calendarUID === calendar.uid).length;
        const resourceRefs = model.resources.filter((resource) => resource.calendarUID === calendar.uid).length;
        const baseRefs = model.calendars.filter((item) => item.baseCalendarUID === calendar.uid).length;
        return `Project=${projectRefs} / Tasks=${taskRefs} / Resources=${resourceRefs} / BaseOf=${baseRefs}`;
    }
    function formatCalendarLink(model, calendarUID) {
        if (!calendarUID) {
            return "-";
        }
        const calendar = model.calendars.find((item) => item.uid === calendarUID);
        return calendar ? `${calendarUID} (${calendar.name || "(no name)"})` : `${calendarUID} (missing)`;
    }
    function formatTaskLink(model, taskUID) {
        if (!taskUID) {
            return "-";
        }
        const task = model.tasks.find((item) => item.uid === taskUID);
        return task ? `${taskUID} (${task.name || "(no name)"})` : `${taskUID} (missing)`;
    }
    function formatResourceLink(model, resourceUID) {
        if (!resourceUID) {
            return "-";
        }
        const resource = model.resources.find((item) => item.uid === resourceUID);
        return resource ? `${resourceUID} (${resource.name || "(no name)"})` : `${resourceUID} (missing)`;
    }
    function formatChangeValue(value) {
        if (value === undefined) {
            return "(empty)";
        }
        return String(value);
    }
    function escapeHtml(value) {
        return value
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#39;");
    }
    function updateFeedbackVisibility(doc) {
        const stack = doc.querySelector(".md-feedback-stack");
        const validationIssues = getElement(doc, "validationIssues");
        const importWarnings = getElement(doc, "importWarnings");
        const xlsxImportSummary = getElement(doc, "xlsxImportSummary");
        const shouldShow = !validationIssues.classList.contains("md-hidden")
            || !importWarnings.classList.contains("md-hidden")
            || !xlsxImportSummary.classList.contains("md-hidden");
        stack === null || stack === void 0 ? void 0 : stack.classList.toggle("md-hidden", !shouldShow);
    }
    function renderValidationIssues(doc, issues) {
        const container = getElement(doc, "validationIssues");
        const label = container.previousElementSibling;
        if (issues.length === 0) {
            container.classList.add("md-hidden");
            container.innerHTML = "";
            label === null || label === void 0 ? void 0 : label.classList.add("md-hidden");
            updateFeedbackVisibility(doc);
            return;
        }
        const sections = ["project", "tasks", "resources", "assignments", "calendars"];
        const sectionLabels = {
            project: "Project",
            tasks: "Tasks",
            resources: "Resources",
            assignments: "Assignments",
            calendars: "Calendars"
        };
        container.classList.remove("md-hidden");
        label === null || label === void 0 ? void 0 : label.classList.remove("md-hidden");
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
    function renderImportWarnings(doc, warnings) {
        const container = getElement(doc, "importWarnings");
        const label = container.previousElementSibling;
        if (warnings.length === 0) {
            container.classList.add("md-hidden");
            container.innerHTML = "";
            label === null || label === void 0 ? void 0 : label.classList.add("md-hidden");
            updateFeedbackVisibility(doc);
            return;
        }
        container.classList.remove("md-hidden");
        label === null || label === void 0 ? void 0 : label.classList.remove("md-hidden");
        container.innerHTML = `
      <div class="md-issues__title">取込 warning</div>
      <ul class="md-issues__list">
        ${warnings.map((warning) => `<li class="md-issues__item">${escapeHtml(warning.message)}</li>`).join("")}
      </ul>
    `;
        updateFeedbackVisibility(doc);
    }
    function renderXlsxImportSummary(doc, changes) {
        const container = getElement(doc, "xlsxImportSummary");
        const label = container.previousElementSibling;
        if (changes.length === 0) {
            container.classList.add("md-hidden");
            container.innerHTML = "";
            label === null || label === void 0 ? void 0 : label.classList.add("md-hidden");
            updateFeedbackVisibility(doc);
            return;
        }
        const scopeLabel = {
            project: "Project",
            tasks: "Tasks",
            resources: "Resources",
            assignments: "Assignments",
            calendars: "Calendars"
        };
        const scopeCounts = {
            project: 0,
            tasks: 0,
            resources: 0,
            assignments: 0,
            calendars: 0
        };
        const groupedByScope = new Map();
        const groupedChanges = new Map();
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
        const allScopes = ["project", "tasks", "resources", "assignments", "calendars"];
        const changedScopes = allScopes.filter((scope) => scopeCounts[scope] > 0);
        const unchangedScopes = allScopes.filter((scope) => scopeCounts[scope] === 0);
        container.classList.remove("md-hidden");
        label === null || label === void 0 ? void 0 : label.classList.remove("md-hidden");
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
        updateFeedbackVisibility(doc);
    }
    function updateSummary(doc, model, updateSvgButton) {
        updateSvgButton();
        getElement(doc, "summaryProjectName").textContent = (model === null || model === void 0 ? void 0 : model.project.name) || "-";
        getElement(doc, "summaryTaskCount").textContent = String((model === null || model === void 0 ? void 0 : model.tasks.length) || 0);
        getElement(doc, "summaryResourceCount").textContent = String((model === null || model === void 0 ? void 0 : model.resources.length) || 0);
        getElement(doc, "summaryAssignmentCount").textContent = String((model === null || model === void 0 ? void 0 : model.assignments.length) || 0);
        getElement(doc, "summaryCalendarCount").textContent = String((model === null || model === void 0 ? void 0 : model.calendars.length) || 0);
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
    globalThis.__mikuprojectMainRender = {
        renderValidationIssues,
        renderImportWarnings,
        renderXlsxImportSummary,
        updateSummary
    };
})();
