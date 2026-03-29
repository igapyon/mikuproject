/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const mikuprojectXml = globalThis.__mikuprojectXml;
    if (!mikuprojectXml) {
        throw new Error("mikuproject XML module is not loaded");
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
    const mikuprojectWbsXlsx = globalThis.__mikuprojectWbsXlsx;
    if (!mikuprojectWbsXlsx) {
        throw new Error("mikuproject WBS XLSX module is not loaded");
    }
    const mermaidApi = globalThis.mermaid;
    let currentModel = null;
    let currentMermaidSvg = "";
    let mermaidRenderCount = 0;
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
        return Array.from(document.querySelectorAll(".md-top-tab[data-tab]"));
    }
    function getTabPanels() {
        return Array.from(document.querySelectorAll(".md-tab-panel[data-tab-panel]"));
    }
    function setActiveTab(tabId, options = {}) {
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
    async function refreshTransformTab() {
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
        }
        finally {
            isRefreshingTransformTab = false;
        }
    }
    function moveTabFocus(currentButton, direction) {
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
    function bindTabs() {
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
    function parseHolidayDateList(raw) {
        if (!raw) {
            return [];
        }
        const seen = new Set();
        const holidays = [];
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
    function parseWbsDefaultHolidayDates() {
        return parseHolidayDateList(getTextArea("wbsHolidayDatesInput").value.trim());
    }
    function parseOptionalNonNegativeInteger(raw) {
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
    function parseWbsDisplayDaysBeforeBaseDate() {
        return parseOptionalNonNegativeInteger(getInput("wbsDisplayDaysBeforeInput").value);
    }
    function parseWbsDisplayDaysAfterBaseDate() {
        return parseOptionalNonNegativeInteger(getInput("wbsDisplayDaysAfterInput").value);
    }
    function useBusinessDaysForWbsDisplayRange() {
        return true;
    }
    function useBusinessDaysForWbsProgressBand() {
        return true;
    }
    function updateWbsHolidaySummary(holidayDates) {
        const summary = getElement("wbsHolidaySummary");
        if (holidayDates.length === 0) {
            summary.textContent = "既定祝日: 0 件";
            return;
        }
        summary.textContent = `既定祝日: ${holidayDates.length} 件 (${holidayDates.join(", ")})`;
    }
    function syncWbsHolidayDatesInput(model) {
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
    function showToast(message) {
        const toast = document.getElementById("toast");
        if (toast && typeof toast.show === "function") {
            toast.show(message, 2200);
        }
    }
    function getAiPromptText() {
        var _a;
        const template = document.getElementById("aiPromptTemplate");
        if (!template) {
            return "";
        }
        return (((_a = template.content) === null || _a === void 0 ? void 0 : _a.textContent) || template.textContent || "").trim();
    }
    async function copyTextToClipboard(text) {
        if (typeof navigator !== "undefined" &&
            navigator.clipboard &&
            typeof navigator.clipboard.writeText === "function") {
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
    async function copyAiPrompt() {
        const promptText = getAiPromptText();
        if (!promptText) {
            throw new Error("生成AIプロンプトが見つかりません");
        }
        await copyTextToClipboard(promptText);
        showToast("生成AIプロンプトをクリップボードにコピーしました");
        setStatus("生成AIプロンプトをクリップボードにコピーしました");
    }
    function setMermaidError(message) {
        const errorNode = getElement("mermaidSvgError");
        errorNode.textContent = message;
        errorNode.classList.remove("md-hidden");
    }
    function clearMermaidError() {
        const errorNode = getElement("mermaidSvgError");
        errorNode.textContent = "";
        errorNode.classList.add("md-hidden");
    }
    function setMermaidPreviewMarkup(markup) {
        getElement("mermaidSvgPreview").innerHTML = markup;
    }
    function updateMermaidSvgButton() {
        getElement("downloadMermaidSvgBtn").disabled = !currentModel;
    }
    function normalizeSvgForXml(svgText) {
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
        }
        catch (_error) {
            return candidate;
        }
    }
    function applyMermaidSvgTheme(svgText) {
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
            ".milestone, .milestone0, .milestone1, .milestone2, .milestone3 { fill: #8f95e8 !important; stroke: #5d63cf !important; }",
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
    function downloadBlob(blob, filename) {
        const objectUrl = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = objectUrl;
        link.download = filename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        window.setTimeout(() => URL.revokeObjectURL(objectUrl), 0);
    }
    function getMermaidRenderConfig() {
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
                milestoneBkgColor: "#8f95e8",
                milestoneBorderColor: "#5d63cf",
                gridColor: "#707b94",
                lineColor: "#707b94",
                todayLineColor: "#ff3b30"
            }
        };
    }
    async function renderMermaidPreview(source) {
        if (!mermaidApi) {
            currentMermaidSvg = "";
            updateMermaidSvgButton();
            setMermaidPreviewMarkup(`<div class="md-preview-empty">Mermaid ライブラリを読み込めなかったため、プレビューできません。</div>`);
            setMermaidError("Mermaid ライブラリが利用できません。");
            return;
        }
        clearMermaidError();
        const renderId = `mikuprojectMermaidRender${++mermaidRenderCount}`;
        mermaidApi.initialize(getMermaidRenderConfig());
        try {
            const result = await mermaidApi.render(renderId, source);
            currentMermaidSvg = applyMermaidSvgTheme(normalizeSvgForXml(result.svg));
            setMermaidPreviewMarkup(currentMermaidSvg);
            updateMermaidSvgButton();
        }
        catch (error) {
            currentMermaidSvg = "";
            updateMermaidSvgButton();
            setMermaidPreviewMarkup(`<div class="md-preview-empty">Mermaid のプレビューを表示できませんでした。</div>`);
            const message = error instanceof Error ? error.message : String(error);
            setMermaidError(`SVG プレビュー生成に失敗しました: ${message}`);
            throw error;
        }
    }
    function setStatus(message) {
        getElement("statusMessage").textContent = message;
    }
    function formatSaveStamp(date) {
        return [
            date.getFullYear(),
            String(date.getMonth() + 1).padStart(2, "0"),
            String(date.getDate()).padStart(2, "0")
        ].join("-") + " " + [
            String(date.getHours()).padStart(2, "0"),
            String(date.getMinutes()).padStart(2, "0")
        ].join(":");
    }
    function updateXmlSaveState(isDirty) {
        const node = getElement("xmlSaveState");
        node.textContent = isDirty
            ? "XML 保存状態: 未保存"
            : `XML 保存状態: 保存済み (${lastSavedXmlStamp || "-"})`;
        node.classList.toggle("md-save-state--dirty", isDirty);
        node.classList.toggle("md-save-state--clean", !isDirty);
    }
    function markXmlDirty() {
        updateXmlSaveState(true);
    }
    function markXmlSavedCurrent() {
        lastSavedXmlText = getTextArea("xmlInput").value;
        lastSavedXmlStamp = formatSaveStamp(new Date());
        updateXmlSaveState(false);
    }
    function refreshXmlSaveState() {
        updateXmlSaveState(getTextArea("xmlInput").value !== lastSavedXmlText);
    }
    function syncXmlTextFromModel(model) {
        const xmlText = mikuprojectXml.exportMsProjectXml(model);
        getTextArea("xmlInput").value = xmlText;
        isXmlSourceDirty = false;
        refreshXmlSaveState();
        return xmlText;
    }
    function renderPreviewList(containerId, items) {
        const container = getElement(containerId);
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
    function renderValidationIssues(issues) {
        const container = getElement("validationIssues");
        const label = container.previousElementSibling;
        if (issues.length === 0) {
            container.classList.add("md-hidden");
            container.innerHTML = "";
            label === null || label === void 0 ? void 0 : label.classList.add("md-hidden");
            updateFeedbackVisibility();
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
        updateFeedbackVisibility();
    }
    function renderImportWarnings(warnings) {
        const container = getElement("importWarnings");
        const label = container.previousElementSibling;
        if (warnings.length === 0) {
            container.classList.add("md-hidden");
            container.innerHTML = "";
            label === null || label === void 0 ? void 0 : label.classList.add("md-hidden");
            updateFeedbackVisibility();
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
        updateFeedbackVisibility();
    }
    function renderXlsxImportSummary(changes) {
        const container = getElement("xlsxImportSummary");
        const label = container.previousElementSibling;
        if (changes.length === 0) {
            container.classList.add("md-hidden");
            container.innerHTML = "";
            label === null || label === void 0 ? void 0 : label.classList.add("md-hidden");
            updateFeedbackVisibility();
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
        const changedScopes = ["project", "tasks", "resources", "assignments", "calendars"].filter((scope) => scopeCounts[scope] > 0);
        const unchangedScopes = ["project", "tasks", "resources", "assignments", "calendars"].filter((scope) => scopeCounts[scope] === 0);
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
        updateFeedbackVisibility();
    }
    function updateFeedbackVisibility() {
        const stack = document.querySelector(".md-feedback-stack");
        const validationIssues = getElement("validationIssues");
        const importWarnings = getElement("importWarnings");
        const xlsxImportSummary = getElement("xlsxImportSummary");
        const shouldShow = !validationIssues.classList.contains("md-hidden")
            || !importWarnings.classList.contains("md-hidden")
            || !xlsxImportSummary.classList.contains("md-hidden");
        stack === null || stack === void 0 ? void 0 : stack.classList.toggle("md-hidden", !shouldShow);
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
    function updateSummary(model) {
        updateMermaidSvgButton();
        syncWbsHolidayDatesInput(model);
        getElement("summaryProjectName").textContent = (model === null || model === void 0 ? void 0 : model.project.name) || "-";
        getElement("summaryTaskCount").textContent = String((model === null || model === void 0 ? void 0 : model.tasks.length) || 0);
        getElement("summaryResourceCount").textContent = String((model === null || model === void 0 ? void 0 : model.resources.length) || 0);
        getElement("summaryAssignmentCount").textContent = String((model === null || model === void 0 ? void 0 : model.assignments.length) || 0);
        getElement("summaryCalendarCount").textContent = String((model === null || model === void 0 ? void 0 : model.calendars.length) || 0);
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
    function loadSample() {
        currentModel = null;
        getTextArea("xmlInput").value = mikuprojectXml.SAMPLE_XML;
        isXmlSourceDirty = true;
        markXmlDirty();
        setStatus("サンプル XML を読み込みました");
        setActiveTab("input");
    }
    async function importXmlFromFile(file) {
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
    function ensureCurrentModel() {
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
    function parseCurrentXml(options = {}) {
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
    async function exportCurrentMermaid(options = {}) {
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
    function exportCurrentCsv() {
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
        downloadBlob(new Blob([`${csvText}\n`], { type: "text/csv;charset=utf-8" }), `mikuproject-export-${stamp}.csv`);
        setStatus("内部モデルから CSV + ParentID を生成して保存しました");
        showToast("CSV を保存しました");
        setActiveTab("output");
    }
    function exportCurrentProjectOverviewView() {
        const model = ensureCurrentModel();
        syncXmlTextFromModel(model);
        const viewText = JSON.stringify(mikuprojectXml.exportProjectOverviewView(model), null, 2);
        getTextArea("projectOverviewOutput").value = viewText;
        downloadBlob(new Blob([`${viewText}\n`], { type: "application/json;charset=utf-8" }), "mikuproject-project-overview-view.editjson");
        setStatus("project_overview_view を生成して保存しました");
        showToast("project_overview_view を保存しました");
        setActiveTab("output");
    }
    function exportCurrentAiProjectionBundle() {
        const model = ensureCurrentModel();
        syncXmlTextFromModel(model);
        const projectOverview = mikuprojectXml.exportProjectOverviewView(model);
        const phaseDetailViewsFull = (projectOverview.phases || [])
            .map((phase) => phase === null || phase === void 0 ? void 0 : phase.uid)
            .filter((uid) => Boolean(uid))
            .map((phaseUid) => mikuprojectXml.exportPhaseDetailView(model, phaseUid, { mode: "full" }));
        const bundle = {
            view_type: "ai_projection_bundle",
            project_overview_view: projectOverview,
            phase_detail_views_full: phaseDetailViewsFull
        };
        const bundleText = JSON.stringify(bundle, null, 2);
        getTextArea("aiBundleOutput").value = bundleText;
        downloadBlob(new Blob([`${bundleText}\n`], { type: "application/json;charset=utf-8" }), "mikuproject-full-bundle.editjson");
        setStatus(`AI 連携用まとめ JSON を生成して保存しました (phase_detail_view full ${phaseDetailViewsFull.length} 件)`);
        showToast("AI 連携用まとめ JSON を保存しました");
        setActiveTab("output");
    }
    function extractLastJsonBlock(value) {
        var _a, _b;
        const matches = Array.from(value.matchAll(/```json\s*([\s\S]*?)```/g));
        if (matches.length > 0) {
            return ((_b = (_a = matches.at(-1)) === null || _a === void 0 ? void 0 : _a[1]) === null || _b === void 0 ? void 0 : _b.trim()) || "";
        }
        return value.trim();
    }
    function detectJsonDocumentKind(documentLike) {
        if (!documentLike || typeof documentLike !== "object") {
            return undefined;
        }
        const candidate = documentLike;
        if (candidate.format === "mikuproject_workbook_json") {
            return "workbook_json";
        }
        if (candidate.view_type === "project_draft_view") {
            return "project_draft_view";
        }
        return undefined;
    }
    async function importProjectDraftFromText() {
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
    function loadProjectDraftSample() {
        const sampleDraftText = JSON.stringify(mikuprojectXml.SAMPLE_PROJECT_DRAFT_VIEW, null, 2);
        getTextArea("projectDraftImportInput").value = sampleDraftText;
        setStatus("サンプル project_draft_view を読み込みました");
        setActiveTab("input");
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
    function exportCurrentPhaseDetailView(mode = "scoped") {
        var _a, _b, _c, _d, _e, _f, _g, _h, _j;
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
        });
        if ((_a = view.phase) === null || _a === void 0 ? void 0 : _a.uid) {
            getInput("phaseDetailUidInput").value = view.phase.uid;
        }
        getInput("phaseDetailRootUidInput").value = ((_b = view.scope) === null || _b === void 0 ? void 0 : _b.root_uid) || "";
        getInput("phaseDetailMaxDepthInput").value = typeof ((_c = view.scope) === null || _c === void 0 ? void 0 : _c.max_depth) === "number" ? String(view.scope.max_depth) : "";
        const viewText = JSON.stringify(view, null, 2);
        getTextArea("phaseDetailOutput").value = viewText;
        const phaseSuffix = ((_d = view.phase) === null || _d === void 0 ? void 0 : _d.uid) ? `-${view.phase.uid}` : "";
        const modeSuffix = ((_e = view.scope) === null || _e === void 0 ? void 0 : _e.mode) === "scoped" ? "-scoped" : "-full";
        const rootSuffix = ((_f = view.scope) === null || _f === void 0 ? void 0 : _f.root_uid) ? `-root-${view.scope.root_uid}` : "";
        const depthSuffix = typeof ((_g = view.scope) === null || _g === void 0 ? void 0 : _g.max_depth) === "number" ? `-depth-${view.scope.max_depth}` : "";
        downloadBlob(new Blob([`${viewText}\n`], { type: "application/json;charset=utf-8" }), `mikuproject-phase-detail-view${phaseSuffix}${modeSuffix}${rootSuffix}${depthSuffix}.editjson`);
        setStatus(`phase_detail_view (${((_h = view.scope) === null || _h === void 0 ? void 0 : _h.mode) === "scoped" ? "scoped" : "full"}) を生成して保存しました`);
        showToast(`phase_detail_view (${((_j = view.scope) === null || _j === void 0 ? void 0 : _j.mode) === "scoped" ? "scoped" : "full"}) を保存しました`);
        setActiveTab("output");
    }
    function exportCurrentXlsx() {
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
        downloadBlob(new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), `mikuproject-export-${stamp}.xlsx`);
        setStatus("XLSX ファイルをエクスポートしました");
        showToast("XLSX を保存しました");
        setActiveTab("output");
    }
    function exportCurrentWorkbookJson() {
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
        downloadBlob(new Blob([`${jsonText}\n`], { type: "application/json;charset=utf-8" }), `mikuproject-workbook-${stamp}.json`);
        setStatus("XLSX 相当の workbook JSON を生成して保存しました");
        showToast("JSON を保存しました");
        setActiveTab("output");
    }
    function exportCurrentWbsXlsx() {
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
        downloadBlob(new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), `mikuproject-wbs-${stamp}.xlsx`);
        const displayRangeText = displayDaysBeforeBaseDate !== undefined || displayDaysAfterBaseDate !== undefined
            ? ` / 表示期間 営業日 基準日前 ${displayDaysBeforeBaseDate || 0} 日, 基準日後 ${displayDaysAfterBaseDate || 0} 日`
            : "";
        const progressBandText = " / 進捗帯 営業日";
        setStatus(`WBS XLSX ファイルをエクスポートしました${defaultHolidayDates.length > 0 ? ` (祝日 ${defaultHolidayDates.length} 件)` : ""}${displayRangeText}${progressBandText}`);
        showToast("WBS XLSX を保存しました");
        setActiveTab("output");
    }
    async function importXlsxFromFile(file) {
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
    async function importCsvFromFile(file) {
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
    function downloadCurrentXml() {
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
    async function downloadCurrentMermaidSvg() {
        const model = ensureCurrentModel();
        syncXmlTextFromModel(model);
        const mermaidText = mikuprojectXml.exportMermaidGantt(model);
        getTextArea("mermaidOutput").value = mermaidText;
        await renderMermaidPreview(mermaidText);
        if (!currentMermaidSvg) {
            setStatus("出力する SVG がありません");
            return;
        }
        downloadBlob(new Blob([currentMermaidSvg], { type: "image/svg+xml;charset=utf-8" }), "mikuproject-mermaid.svg");
        setStatus("Mermaid SVG を保存しました");
        showToast("SVG を保存しました");
        setActiveTab("output");
    }
    function downloadCurrentMermaidMarkdown() {
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
        downloadBlob(new Blob([markdownText], { type: "text/markdown;charset=utf-8" }), `mermaid-${stamp}.md`);
        setStatus("Mermaid Markdown を保存しました");
        showToast("Mermaid Markdown を保存しました");
        setActiveTab("output");
    }
    function runRoundTripCheck() {
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
    function bindEvents() {
        getElement("loadSampleBtn").addEventListener("click", loadSample);
        getElement("importFileBtn").addEventListener("click", () => {
            getElement("importFileInput").click();
        });
        getElement("downloadMermaidSvgBtn").addEventListener("click", () => {
            void downloadCurrentMermaidSvg().catch((error) => {
                setStatus(error instanceof Error ? error.message : "SVG 保存に失敗しました");
            });
        });
        getElement("exportMermaidMdBtn").addEventListener("click", () => {
            try {
                downloadCurrentMermaidMarkdown();
            }
            catch (error) {
                setStatus(error instanceof Error ? error.message : "Mermaid Markdown 保存に失敗しました");
            }
        });
        getElement("exportCsvBtn").addEventListener("click", () => {
            try {
                exportCurrentCsv();
            }
            catch (error) {
                setStatus(error instanceof Error ? error.message : "CSV 生成に失敗しました");
            }
        });
        getElement("exportProjectOverviewBtn").addEventListener("click", () => {
            try {
                exportCurrentProjectOverviewView();
            }
            catch (error) {
                setStatus(error instanceof Error ? error.message : "project_overview_view 生成に失敗しました");
            }
        });
        getElement("exportAiBundleBtn").addEventListener("click", () => {
            try {
                exportCurrentAiProjectionBundle();
            }
            catch (error) {
                setStatus(error instanceof Error ? error.message : "AI 連携用まとめ JSON 生成に失敗しました");
            }
        });
        getElement("loadProjectDraftSampleBtn").addEventListener("click", loadProjectDraftSample);
        getElement("copyAiPromptBtn").addEventListener("click", async () => {
            try {
                await copyAiPrompt();
            }
            catch (error) {
                setStatus(error instanceof Error ? error.message : "生成AIプロンプトのコピーに失敗しました");
            }
        });
        getElement("importProjectDraftBtn").addEventListener("click", async () => {
            try {
                await importProjectDraftFromText();
            }
            catch (error) {
                setStatus(error instanceof Error ? error.message : "project_draft_view 取り込みに失敗しました");
            }
        });
        getElement("exportPhaseDetailBtn").addEventListener("click", () => {
            try {
                exportCurrentPhaseDetailView("scoped");
            }
            catch (error) {
                setStatus(error instanceof Error ? error.message : "phase_detail_view 生成に失敗しました");
            }
        });
        getElement("exportPhaseDetailFullBtn").addEventListener("click", () => {
            try {
                exportCurrentPhaseDetailView("full");
            }
            catch (error) {
                setStatus(error instanceof Error ? error.message : "phase_detail_view 生成に失敗しました");
            }
        });
        getElement("exportXlsxBtn").addEventListener("click", () => {
            try {
                exportCurrentXlsx();
            }
            catch (error) {
                setStatus(error instanceof Error ? error.message : "XLSX エクスポートに失敗しました");
            }
        });
        getElement("exportWorkbookJsonBtn").addEventListener("click", () => {
            try {
                exportCurrentWorkbookJson();
            }
            catch (error) {
                setStatus(error instanceof Error ? error.message : "JSON エクスポートに失敗しました");
            }
        });
        getElement("exportWbsXlsxBtn").addEventListener("click", () => {
            try {
                exportCurrentWbsXlsx();
            }
            catch (error) {
                setStatus(error instanceof Error ? error.message : "WBS XLSX エクスポートに失敗しました");
            }
        });
        getElement("downloadXmlBtn").addEventListener("click", () => {
            try {
                downloadCurrentXml();
            }
            catch (error) {
                setStatus(error instanceof Error ? error.message : "XML 保存に失敗しました");
            }
        });
        getElement("roundTripBtn").addEventListener("click", () => {
            try {
                runRoundTripCheck();
            }
            catch (error) {
                setStatus(error instanceof Error ? error.message : "再読込テストに失敗しました");
            }
        });
        getElement("importFileInput").addEventListener("change", async (event) => {
            const input = event.target;
            const file = (input === null || input === void 0 ? void 0 : input.files) && input.files[0];
            try {
                await importFromFile(file);
            }
            catch (error) {
                setStatus(error instanceof Error ? error.message : "ファイル読込に失敗しました");
            }
            finally {
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
    function initialize() {
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
    globalThis.__mikuprojectMainTestHooks = {
        parseCurrentXml,
        exportCurrentMermaid,
        renderValidationIssues,
        renderXlsxImportSummary,
        updateFeedbackVisibility
    };
    document.addEventListener("DOMContentLoaded", initialize);
})();
