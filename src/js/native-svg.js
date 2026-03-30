/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const DAY_WIDTH = 56;
    const LIST_LABEL_WIDTH = 360;
    const NEAR_LEFT_LABEL_WIDTH = 220;
    const NEAR_RIGHT_LABEL_WIDTH = 300;
    const HEADER_HEIGHT = 82;
    const ROW_HEIGHT = 38;
    const LEFT_PADDING = 20;
    const TOP_PADDING = 22;
    const RIGHT_PADDING = 24;
    const BOTTOM_PADDING = 28;
    function exportNativeSvg(model, options = {}) {
        const labelMode = options.labelMode || "near";
        const holidaySet = new Set([
            ...collectWbsHolidayDates(model),
            ...(options.holidayDates || []).map((day) => String(day || "").slice(0, 10)).filter(Boolean)
        ]);
        const nonWorkingDayTypes = collectProjectNonWorkingDayTypes(model);
        const dateBand = buildDisplayDateBand(model.project.startDate, model.project.finishDate, model.project.currentDate, options.displayDaysBeforeBaseDate, options.displayDaysAfterBaseDate, holidaySet, nonWorkingDayTypes, options.useBusinessDaysForDisplayRange);
        const rows = buildTaskRows(model.tasks, dateBand);
        const chartWidth = dateBand.length * DAY_WIDTH;
        const leftLabelWidth = labelMode === "list" ? LIST_LABEL_WIDTH : NEAR_LEFT_LABEL_WIDTH;
        const rightLabelWidth = labelMode === "list" ? 0 : NEAR_RIGHT_LABEL_WIDTH;
        const svgWidth = LEFT_PADDING + leftLabelWidth + chartWidth + rightLabelWidth + RIGHT_PADDING;
        const svgHeight = TOP_PADDING + HEADER_HEIGHT + (rows.length * ROW_HEIGHT) + BOTTOM_PADDING;
        const todayIndex = indexOfDate(dateBand, model.project.currentDate);
        const parts = [
            `<svg xmlns="http://www.w3.org/2000/svg" width="${svgWidth}" height="${svgHeight}" viewBox="0 0 ${svgWidth} ${svgHeight}" role="img" aria-label="${escapeXml(model.project.name || "Project")}">`,
            "<style>",
            "text { font-family: 'Hiragino Sans', 'Yu Gothic', sans-serif; fill: #1d2740; }",
            ".title { font-size: 18px; font-weight: 700; }",
            ".axis { font-size: 12px; fill: #5b6370; }",
            ".label { font-size: 13px; }",
            ".phaseLabel { font-size: 13px; font-weight: 700; }",
            ".grid { stroke: #c9d3e1; stroke-width: 1; }",
            ".today { stroke: #ff6b5a; stroke-width: 2; }",
            "</style>",
            `<rect x="0" y="0" width="${svgWidth}" height="${svgHeight}" fill="#ffffff"/>`,
            `<text class="title" x="${LEFT_PADDING + leftLabelWidth + (chartWidth / 2)}" y="${TOP_PADDING + 18}" text-anchor="middle">${escapeXml(model.project.name || "-")}</text>`
        ];
        const chartOriginX = LEFT_PADDING + leftLabelWidth;
        const chartOriginY = TOP_PADDING + HEADER_HEIGHT;
        for (let index = 0; index < dateBand.length; index += 1) {
            const day = dateBand[index];
            const x = chartOriginX + (index * DAY_WIDTH);
            const isHoliday = holidaySet.has(day);
            const isWeekend = isWeeklyNonWorkingDay(day, nonWorkingDayTypes);
            const fill = isHoliday ? "#fce4ec" : (isWeekend ? "#eef3f8" : "#ffffff");
            parts.push(`<rect x="${x}" y="${TOP_PADDING + 30}" width="${DAY_WIDTH}" height="${svgHeight - TOP_PADDING - BOTTOM_PADDING - 8}" fill="${fill}"/>`);
            parts.push(`<line class="grid" x1="${x}" y1="${TOP_PADDING + 26}" x2="${x}" y2="${svgHeight - BOTTOM_PADDING}" />`);
            parts.push(`<text class="axis" x="${x + (DAY_WIDTH / 2)}" y="${TOP_PADDING + 54}" text-anchor="middle">${escapeXml(formatSvgAxisDate(day))}</text>`);
        }
        parts.push(`<line class="grid" x1="${chartOriginX + chartWidth}" y1="${TOP_PADDING + 26}" x2="${chartOriginX + chartWidth}" y2="${svgHeight - BOTTOM_PADDING}" />`);
        if (todayIndex >= 0) {
            const todayX = chartOriginX + (todayIndex * DAY_WIDTH) + (DAY_WIDTH / 2);
            parts.push(`<line class="today" x1="${todayX}" y1="${TOP_PADDING + 26}" x2="${todayX}" y2="${svgHeight - BOTTOM_PADDING}" />`);
        }
        for (const row of rows) {
            const rowY = chartOriginY + row.y;
            if (row.startIndex !== null && row.endIndex !== null) {
                const barX = chartOriginX + (row.startIndex * DAY_WIDTH) + 6;
                const barWidth = Math.max(12, ((row.endIndex - row.startIndex + 1) * DAY_WIDTH) - 12);
                const barY = rowY + 8;
                if (row.kind === "milestone") {
                    const centerX = chartOriginX + (row.startIndex * DAY_WIDTH) + (DAY_WIDTH / 2);
                    const centerY = rowY + (ROW_HEIGHT / 2);
                    const isCompleted = (row.task.percentComplete || 0) >= 100;
                    const fill = isCompleted ? "#d9efff" : "#ffffff";
                    parts.push(`<polygon points="${centerX},${centerY - 13} ${centerX + 13},${centerY} ${centerX},${centerY + 13} ${centerX - 13},${centerY}" fill="${fill}" stroke="#4f95d6" stroke-width="3"/>`);
                }
                else if (row.kind === "phase") {
                    const lineY = rowY + (ROW_HEIGHT / 2);
                    const startX = barX;
                    const endX = barX + barWidth;
                    const trackStroke = "#8eb9ea";
                    const progressStroke = "#2f79d0";
                    const phaseStrokeWidth = 3;
                    const progressEndX = startX + Math.max(0, Math.min(barWidth, Math.round(barWidth * (Math.max(0, Math.min(100, row.task.percentComplete || 0)) / 100))));
                    parts.push(`<line x1="${startX}" y1="${lineY}" x2="${endX}" y2="${lineY}" stroke="${trackStroke}" stroke-width="${phaseStrokeWidth}" stroke-linecap="round"/>`);
                    if (progressEndX > startX) {
                        parts.push(`<line x1="${startX}" y1="${lineY}" x2="${progressEndX}" y2="${lineY}" stroke="${progressStroke}" stroke-width="${phaseStrokeWidth}" stroke-linecap="round"/>`);
                    }
                }
                else {
                    const trackFill = "#d9efff";
                    const trackStroke = "#4f95d6";
                    const progressFill = "#3f86d8";
                    parts.push(`<rect x="${barX}" y="${barY}" width="${barWidth}" height="22" rx="4" fill="${trackFill}" stroke="${trackStroke}" stroke-width="3"/>`);
                    const progressWidth = Math.max(0, Math.min(barWidth, Math.round(barWidth * (Math.max(0, Math.min(100, row.task.percentComplete || 0)) / 100))));
                    if (progressWidth > 0) {
                        parts.push(`<rect x="${barX}" y="${barY}" width="${progressWidth}" height="22" rx="4" fill="${progressFill}" stroke="none"/>`);
                    }
                }
            }
            const labelPlacement = resolveLabelPlacement(row, chartOriginX, chartWidth, svgWidth, labelMode);
            parts.push(`<text class="${row.kind === "phase" ? "phaseLabel" : "label"}" x="${labelPlacement.x}" y="${rowY + 24}" text-anchor="${labelPlacement.anchor}">${escapeXml(formatTaskLabel(row.task, labelMode))}</text>`);
        }
        parts.push("</svg>");
        return parts.join("");
    }
    function buildTaskRows(tasks, dateBand) {
        return tasks.map((task, index) => ({
            task,
            label: task.name || "-",
            kind: task.summary ? "phase" : (task.milestone ? "milestone" : "task"),
            startIndex: indexOfDate(dateBand, task.start),
            endIndex: indexOfDate(dateBand, task.finish),
            y: index * ROW_HEIGHT
        }));
    }
    function formatTaskLabel(task, labelMode) {
        if (labelMode === "list") {
            return `${"　".repeat(Math.max(0, task.outlineLevel - 1))}${task.name || "-"}`;
        }
        return task.name || "-";
    }
    function resolveLabelPlacement(row, chartOriginX, chartWidth, svgWidth, labelMode) {
        if (labelMode === "list" || row.startIndex === null || row.endIndex === null) {
            return { x: LEFT_PADDING + 10, anchor: "start" };
        }
        const textWidth = estimateLabelWidth(row.label, row.kind === "phase");
        const gap = 12;
        const chartEndX = chartOriginX + chartWidth;
        const shapeStartX = row.kind === "milestone"
            ? chartOriginX + (row.startIndex * DAY_WIDTH) + (DAY_WIDTH / 2) - 13
            : chartOriginX + (row.startIndex * DAY_WIDTH) + 6;
        const shapeEndX = row.kind === "milestone"
            ? chartOriginX + (row.startIndex * DAY_WIDTH) + (DAY_WIDTH / 2) + 13
            : chartOriginX + (row.endIndex * DAY_WIDTH) + DAY_WIDTH - 6;
        const preferredRightX = shapeEndX + gap;
        if ((preferredRightX + textWidth) <= (svgWidth - RIGHT_PADDING)) {
            return { x: preferredRightX, anchor: "start" };
        }
        const preferredLeftX = shapeStartX - gap;
        if ((preferredLeftX - textWidth) >= LEFT_PADDING) {
            return { x: preferredLeftX, anchor: "end" };
        }
        const fallbackX = Math.max(LEFT_PADDING + 10, chartOriginX - gap);
        if ((fallbackX + textWidth) <= chartEndX) {
            return { x: fallbackX, anchor: "start" };
        }
        return { x: preferredLeftX, anchor: "end" };
    }
    function estimateLabelWidth(label, isPhase) {
        const basePerChar = isPhase ? 14 : 13;
        return Math.max(48, Math.ceil(String(label || "").length * basePerChar));
    }
    function formatSvgAxisDate(day) {
        const match = day.match(/^(\d{4})-(\d{2})-(\d{2})$/);
        if (!match) {
            return day;
        }
        return `${Number(match[2])}/${Number(match[3])}`;
    }
    function indexOfDate(dateBand, value) {
        const key = String(value || "").slice(0, 10);
        if (!key) {
            return null;
        }
        const index = dateBand.indexOf(key);
        return index >= 0 ? index : null;
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
    function isWeeklyNonWorkingDay(day, nonWorkingDayTypes) {
        const date = parseDateOnly(day);
        if (!date) {
            return false;
        }
        const dayType = date.getDay() === 0 ? 1 : date.getDay() + 1;
        return nonWorkingDayTypes.has(dayType);
    }
    function parseDateOnly(value) {
        const text = String(value || "").trim().slice(0, 10);
        if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) {
            return null;
        }
        const parsed = new Date(`${text}T00:00:00`);
        return Number.isNaN(parsed.getTime()) ? null : parsed;
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
    function escapeXml(value) {
        return String(value || "")
            .replaceAll("&", "&amp;")
            .replaceAll("<", "&lt;")
            .replaceAll(">", "&gt;")
            .replaceAll("\"", "&quot;");
    }
    globalThis.__mikuprojectNativeSvg = {
        exportNativeSvg
    };
})();
