/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const markdownEscape = globalThis.__mikuprojectMarkdownEscape;
    if (!markdownEscape) {
        throw new Error("mikuproject markdown escape module is not loaded");
    }
    function exportWbsMarkdown(model, options = {}) {
        const nonWorkingDayTypes = collectProjectNonWorkingDayTypes(model);
        const holidaySet = new Set([
            ...collectWbsHolidayDates(model),
            ...(options.holidayDates || []).map((day) => String(day || "").slice(0, 10)).filter(Boolean)
        ]);
        const dateBand = buildDisplayDateBand(model.project.startDate, model.project.finishDate, model.project.currentDate, options.displayDaysBeforeBaseDate, options.displayDaysAfterBaseDate, holidaySet, nonWorkingDayTypes, options.useBusinessDaysForDisplayRange);
        const calendarNameByUid = new Map(model.calendars.map((calendar) => [calendar.uid, calendar.name]));
        const resourceNameByUid = new Map(model.resources.map((resource) => [resource.uid, resource.name]));
        const predecessorNameByUid = new Map(model.tasks.map((task) => [task.uid, task.name]));
        const resourceNamesByTaskUid = new Map();
        for (const assignment of model.assignments) {
            const resourceName = resourceNameByUid.get(assignment.resourceUid);
            if (!resourceName) {
                continue;
            }
            const resourceNames = resourceNamesByTaskUid.get(assignment.taskUid) || [];
            if (!resourceNames.includes(resourceName)) {
                resourceNames.push(resourceName);
            }
            resourceNamesByTaskUid.set(assignment.taskUid, resourceNames);
        }
        const sections = [
            "# プロジェクト情報",
            "",
            ...buildKeyValueTable([
                ["プロジェクト名", model.project.name || "-"],
                ["カレンダ", formatCalendarLabel(model.project.calendarUID, calendarNameByUid)],
                ["開始日", formatWbsDate(model.project.startDate)],
                ["終了日", formatWbsDate(model.project.finishDate)],
                ["現在日", formatWbsDate(model.project.currentDate)],
                ["祝日", String(holidaySet.size)]
            ]),
            "",
            "# WBS ツリー",
            "",
            ...wrapFenceBlock(buildTreeLines(model.tasks, holidaySet, nonWorkingDayTypes, options.useBusinessDaysForProgressBand), "text"),
            "",
            "---",
            "",
            "# WBS テーブル",
            "",
            ...buildWbsTable(model.tasks, holidaySet, nonWorkingDayTypes, resourceNamesByTaskUid, predecessorNameByUid, options.useBusinessDaysForProgressBand),
            "",
            "---",
            "",
            "# サマリ",
            "",
            ...buildKeyValueTable([
                ["表示日", String(dateBand.length)],
                ["表示週", String(dateBand.length > 0 ? Math.ceil(dateBand.length / 7) : 0)],
                ["営業日", String(countBusinessDays(dateBand, holidaySet, nonWorkingDayTypes))],
                ["前日数", formatOptionalCount(options.displayDaysBeforeBaseDate)],
                ["後日数", formatOptionalCount(options.displayDaysAfterBaseDate)],
                ["表示", options.useBusinessDaysForDisplayRange ? "営業日" : "暦日"],
                ["進捗", options.useBusinessDaysForProgressBand ? "営業日" : "暦日"],
                ["基準日", formatWbsDate(model.project.currentDate)],
                ["タスク", String(model.tasks.length)],
                ["リソース", String(model.resources.length)],
                ["割当", String(model.assignments.length)],
                ["カレンダ", String(model.calendars.length)]
            ]),
            ""
        ];
        return sections.join("\n");
    }
    function buildKeyValueTable(rows) {
        return [
            "| 項目 | 値 |",
            "| --- | --- |",
            ...rows.map(([label, value]) => `| ${escapeMarkdownCell(label)} | ${escapeMarkdownCell(value)} |`)
        ];
    }
    function buildWbsTable(tasks, holidaySet, nonWorkingDayTypes, resourceNamesByTaskUid, predecessorNameByUid, useBusinessDaysForProgressBand) {
        const header = [
            "WBS",
            "種別",
            "階層",
            "名称",
            "開始",
            "終了",
            "期間",
            "タスク詳細",
            "進捗",
            "担当",
            "リソース",
            "先行"
        ];
        const lines = [
            `| ${header.join(" | ")} |`,
            `| ${header.map(() => "---").join(" | ")} |`
        ];
        for (const task of tasks) {
            const resourceNames = resourceNamesByTaskUid.get(task.uid) || [];
            lines.push([
                task.wbs || task.outlineNumber || "-",
                classifyTaskKind(task),
                String(task.outlineLevel || "-"),
                formatTableTaskLabel(task),
                formatWbsDate(task.start),
                formatWbsDate(task.finish),
                formatDurationLabel(task, holidaySet, nonWorkingDayTypes, useBusinessDaysForProgressBand),
                formatNoteCell(task.notes),
                formatPercentCell(task.percentComplete),
                firstResourceName(resourceNames) || "-",
                resourceNames.join(", ") || "-",
                task.predecessors.map((item) => predecessorNameByUid.get(item.predecessorUid) || item.predecessorUid).join(", ") || "-"
            ].map((value) => escapeMarkdownCell(value)).join(" | ").replace(/^/, "| ").concat(" |"));
        }
        return lines;
    }
    function buildTreeLines(tasks, holidaySet, nonWorkingDayTypes, useBusinessDaysForProgressBand) {
        const lines = [];
        for (const task of tasks) {
            const indent = task.outlineLevel > 1 ? `${"　".repeat(Math.max(0, task.outlineLevel - 2))}┗　` : "";
            const taskLine = `${indent}${formatTreeInlineText(task.wbs || task.outlineNumber || "-")} ${formatTreeInlineText(task.name || "-")} (${formatTreeInlineText(formatTreeDateRange(task.start, task.finish))}): ${formatTreeInlineText(formatPercentCell(task.percentComplete))}`;
            lines.push(taskLine);
            if (task.notes && task.notes.trim()) {
                const noteIndent = `${"　".repeat(Math.max(0, task.outlineLevel - 1))}　`;
                const noteLines = normalizeFenceText(task.notes)
                    .split("\n")
                    .filter(Boolean);
                if (noteLines.length > 0) {
                    lines.push(`${noteIndent}詳細: ${noteLines[0]}`);
                    for (const line of noteLines.slice(1)) {
                        lines.push(`${noteIndent}      ${line}`);
                    }
                }
            }
        }
        return lines.length > 0 ? lines : ["(task なし)"];
    }
    function classifyTaskKind(task) {
        if (task.summary) {
            return "フェーズ";
        }
        if (task.milestone) {
            return "マイル";
        }
        return "タスク";
    }
    function formatTableTaskLabel(task) {
        return task.name || "-";
    }
    function formatNoteCell(value) {
        const normalized = normalizeTextBlock(value);
        return normalized ? normalized : "-";
    }
    function formatPercentCell(value) {
        if (value === undefined || value === null || !Number.isFinite(value)) {
            return "-";
        }
        return `${Math.max(0, Math.min(100, Math.round(value)))}%`;
    }
    function formatFlagCell(flag, label) {
        return flag ? label : "-";
    }
    function formatOptionalCount(value) {
        if (value === undefined || value === null || !Number.isFinite(value)) {
            return "-";
        }
        return String(Math.max(0, Math.floor(value)));
    }
    function escapeMarkdownCell(value) {
        return markdownEscape.escapeMarkdownTableCell(String(value || ""));
    }
    function formatTreeInlineText(value) {
        return normalizeFenceText(value).replace(/\n+/g, " / ");
    }
    // normalizeTextBlock only performs text normalization for export:
    // newline normalization, control-character removal, tab expansion,
    // trailing-space trimming, and blank-line compaction. It does not apply
    // Markdown-specific escaping.
    function normalizeTextBlock(value) {
        return String(value || "").trim()
            .replace(/\r\n?/g, "\n")
            .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, "")
            .replace(/\t/g, "  ")
            .split("\n")
            .map((line) => line.trimEnd())
            .filter((line, index, lines) => !(line === "" && lines[index - 1] === ""))
            .join("\n");
    }
    // Code-fence text does not need normal Markdown escaping, but it does need
    // normalized content so fence selection and indentation stay predictable.
    function normalizeFenceText(value) {
        return normalizeTextBlock(value);
    }
    function wrapFenceBlock(lines, infoString) {
        const content = lines.join("\n");
        const backtickRun = longestFenceRun(content, "`");
        const tildeRun = longestFenceRun(content, "~");
        const fenceChar = tildeRun <= backtickRun ? "~" : "`";
        const fenceLength = Math.max(3, (fenceChar === "~" ? tildeRun : backtickRun) + 1);
        const fence = fenceChar.repeat(fenceLength);
        return [`${fence}${infoString}`, ...lines, fence];
    }
    function longestFenceRun(text, char) {
        let maxRun = 0;
        let currentRun = 0;
        for (const currentChar of text) {
            if (currentChar === char) {
                currentRun += 1;
                if (currentRun > maxRun) {
                    maxRun = currentRun;
                }
            }
            else {
                currentRun = 0;
            }
        }
        return maxRun;
    }
    function collectWbsHolidayDates(model) {
        const holidaySet = new Set();
        for (const calendar of model.calendars) {
            for (const exception of calendar.exceptions || []) {
                if (exception.dayWorking !== false && (exception.workingTimes || []).length > 0) {
                    continue;
                }
                for (const day of expandExceptionDays(exception)) {
                    holidaySet.add(day);
                }
            }
        }
        return Array.from(holidaySet).sort();
    }
    function expandExceptionDays(exception) {
        const singleDay = exception.fromDate ? formatDateOnly(parseDateOnly(exception.fromDate)) : "";
        if (!exception.fromDate || !exception.toDate) {
            return singleDay ? [singleDay] : [];
        }
        return buildDateBand(exception.fromDate, exception.toDate);
    }
    function resolveProjectCalendar(model) {
        if (model.project.calendarUID) {
            const projectCalendar = model.calendars.find((calendar) => calendar.uid === model.project.calendarUID);
            if (projectCalendar) {
                return projectCalendar;
            }
        }
        return model.calendars.find((calendar) => calendar.isBaseCalendar) || model.calendars[0];
    }
    function resolveCalendarDayWorking(calendarByUid, calendar, dayType, visiting = new Set()) {
        if (!calendar) {
            return undefined;
        }
        if (visiting.has(calendar.uid)) {
            return undefined;
        }
        visiting.add(calendar.uid);
        const weekDay = calendar.weekDays.find((item) => item.dayType === dayType);
        if (weekDay) {
            return weekDay.dayWorking;
        }
        if (calendar.baseCalendarUID) {
            return resolveCalendarDayWorking(calendarByUid, calendarByUid.get(calendar.baseCalendarUID), dayType, visiting);
        }
        return undefined;
    }
    function collectProjectNonWorkingDayTypes(model) {
        const calendarByUid = new Map(model.calendars.map((calendar) => [calendar.uid, calendar]));
        const projectCalendar = resolveProjectCalendar(model);
        const nonWorkingDayTypes = new Set();
        for (let dayType = 1; dayType <= 7; dayType += 1) {
            const dayWorking = resolveCalendarDayWorking(calendarByUid, projectCalendar, dayType);
            if (dayWorking === false) {
                nonWorkingDayTypes.add(dayType);
                continue;
            }
            if (dayWorking === undefined && (dayType === 1 || dayType === 7)) {
                nonWorkingDayTypes.add(dayType);
            }
        }
        return nonWorkingDayTypes;
    }
    function buildDateBand(startDate, finishDate) {
        const start = parseDateOnly(startDate);
        const finish = parseDateOnly(finishDate);
        if (!start || !finish || start.getTime() > finish.getTime()) {
            return [];
        }
        const days = [];
        const cursor = new Date(start.getTime());
        while (cursor.getTime() <= finish.getTime()) {
            days.push(formatDateOnly(cursor));
            cursor.setDate(cursor.getDate() + 1);
        }
        return days;
    }
    function buildDisplayDateBand(startDate, finishDate, baseDate, displayDaysBeforeBaseDate, displayDaysAfterBaseDate, holidaySet, nonWorkingDayTypes, useBusinessDaysForDisplayRange) {
        const fullBand = buildDateBand(startDate, finishDate);
        const before = normalizeDisplayDayCount(displayDaysBeforeBaseDate);
        const after = normalizeDisplayDayCount(displayDaysAfterBaseDate);
        if (before === null && after === null) {
            return fullBand;
        }
        const base = parseDateOnly(baseDate);
        if (!base || fullBand.length === 0) {
            return fullBand;
        }
        const projectStart = parseDateOnly(startDate);
        const projectFinish = parseDateOnly(finishDate);
        if (!projectStart || !projectFinish) {
            return fullBand;
        }
        const from = useBusinessDaysForDisplayRange
            ? shiftBusinessDays(base, -(before || 0), holidaySet, nonWorkingDayTypes)
            : shiftCalendarDays(base, -(before || 0));
        const to = useBusinessDaysForDisplayRange
            ? shiftBusinessDays(base, after || 0, holidaySet, nonWorkingDayTypes)
            : shiftCalendarDays(base, after || 0);
        const clampedStart = from.getTime() < projectStart.getTime() ? projectStart : from;
        const clampedFinish = to.getTime() > projectFinish.getTime() ? projectFinish : to;
        if (clampedStart.getTime() > clampedFinish.getTime()) {
            return fullBand;
        }
        return buildDateBand(formatDateOnly(clampedStart), formatDateOnly(clampedFinish));
    }
    function normalizeDisplayDayCount(value) {
        if (value === undefined || value === null || !Number.isFinite(value)) {
            return null;
        }
        return Math.max(0, Math.floor(value));
    }
    function shiftCalendarDays(base, offset) {
        const result = new Date(base.getTime());
        result.setDate(result.getDate() + offset);
        return result;
    }
    function shiftBusinessDays(base, offset, holidaySet, nonWorkingDayTypes) {
        const result = new Date(base.getTime());
        const direction = offset < 0 ? -1 : 1;
        let remaining = Math.abs(offset);
        while (remaining > 0) {
            result.setDate(result.getDate() + direction);
            const day = formatDateOnly(result);
            if (isWeeklyNonWorkingDay(day, nonWorkingDayTypes) || holidaySet.has(day)) {
                continue;
            }
            remaining -= 1;
        }
        return result;
    }
    function countBusinessDays(dateBand, holidaySet, nonWorkingDayTypes) {
        return dateBand.filter((day) => !isWeeklyNonWorkingDay(day, nonWorkingDayTypes) && !holidaySet.has(day)).length;
    }
    function enumerateBusinessDays(startDate, finishDate, holidaySet, nonWorkingDayTypes) {
        return buildDateBand(startDate, finishDate).filter((day) => !isWeeklyNonWorkingDay(day, nonWorkingDayTypes) && !holidaySet.has(day));
    }
    function formatDurationLabel(task, holidaySet, nonWorkingDayTypes, useBusinessDaysForProgressBand) {
        if (useBusinessDaysForProgressBand) {
            const businessDays = enumerateBusinessDays(task.start, task.finish, holidaySet, nonWorkingDayTypes).length;
            return businessDays > 0 ? `${businessDays}営業日` : "-";
        }
        const calendarDays = buildDateBand(task.start, task.finish).length;
        return calendarDays > 0 ? `${calendarDays}日` : "-";
    }
    function formatWbsDate(value) {
        return value ? value.slice(0, 10) : "-";
    }
    function formatTreeDate(value) {
        if (!value) {
            return "-";
        }
        const match = String(value).match(/^(\d{4})-(\d{2})-(\d{2})/);
        if (!match) {
            return String(value).slice(0, 10) || "-";
        }
        const [, _year, month, day] = match;
        return `${Number(month)}/${Number(day)}`;
    }
    function formatTreeDateRange(start, finish) {
        const startLabel = formatTreeDate(start);
        const finishLabel = formatTreeDate(finish);
        if (startLabel === finishLabel) {
            return startLabel;
        }
        return `${startLabel} - ${finishLabel}`;
    }
    function parseDateOnly(value) {
        if (!value) {
            return null;
        }
        const match = String(value).match(/^(\d{4})-(\d{2})-(\d{2})/);
        if (!match) {
            return null;
        }
        const [, year, month, day] = match;
        return new Date(Number(year), Number(month) - 1, Number(day));
    }
    function formatDateOnly(value) {
        if (!value) {
            return "";
        }
        const year = value.getFullYear();
        const month = String(value.getMonth() + 1).padStart(2, "0");
        const day = String(value.getDate()).padStart(2, "0");
        return `${year}-${month}-${day}`;
    }
    function isWeeklyNonWorkingDay(day, nonWorkingDayTypes) {
        const target = parseDateOnly(day);
        if (!target) {
            return false;
        }
        const dayType = target.getDay() === 0 ? 1 : target.getDay() + 1;
        return nonWorkingDayTypes.has(dayType);
    }
    function firstResourceName(resourceNames) {
        if (!resourceNames || resourceNames.length === 0) {
            return "";
        }
        return resourceNames[0];
    }
    function formatCalendarLabel(calendarUID, calendarNameByUid) {
        if (!calendarUID) {
            return "-";
        }
        const calendarName = calendarNameByUid.get(calendarUID);
        return calendarName ? `${calendarUID} ${calendarName}` : calendarUID;
    }
    globalThis.__mikuprojectWbsMarkdown = {
        exportWbsMarkdown
    };
})();
