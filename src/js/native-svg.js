/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const DAY_WIDTH = 56;
    const LIST_LABEL_WIDTH = 360;
    const NEAR_LEFT_LABEL_WIDTH = 0;
    const NEAR_RIGHT_LABEL_WIDTH = 0;
    const HEADER_HEIGHT = 82;
    const ROW_HEIGHT = 38;
    const LEFT_PADDING = 0;
    const TOP_PADDING = 22;
    const RIGHT_PADDING = 0;
    const BOTTOM_PADDING = 28;
    const ZIP_TEXT_ENCODER = new TextEncoder();
    const ZIP_CRC32_TABLE = buildCrc32Table();
    const MONTHLY_CELL_WIDTH = 164;
    const MONTHLY_CELL_HEIGHT = 126;
    const MONTHLY_HEADER_HEIGHT = 108;
    const MONTHLY_WEEKDAY_HEIGHT = 34;
    const MONTHLY_LEFT_PADDING = 24;
    const MONTHLY_TOP_PADDING = 24;
    const MONTHLY_RIGHT_PADDING = 24;
    const MONTHLY_BOTTOM_PADDING = 24;
    const MONTHLY_MAX_ITEMS_PER_DAY = 3;
    const MONTHLY_MAX_LABEL_CHARS = 15;
    const ZIP_FIXED_MOD_TIME = 0;
    const ZIP_FIXED_MOD_DATE = ((2025 - 1980) << 9) | (1 << 5) | 1;
    const WEEK_WIDTH = 38;
    const WEEKLY_HEADER_HEIGHT = 96;
    const WEEKLY_TOP_PADDING = 22;
    const WEEKLY_BOTTOM_PADDING = 28;
    const WEEKLY_LEFT_PADDING = 0;
    const WEEKLY_RIGHT_PADDING = 0;
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
        const chartOriginXBase = LEFT_PADDING + leftLabelWidth;
        const labelPlacements = rows.map((row) => resolveLabelPlacement(row, chartOriginXBase, chartWidth, labelMode));
        const labelMinX = labelPlacements.reduce((min, placement) => {
            const placementMinX = placement.anchor === "start" ? placement.x : placement.x - placement.width;
            return Math.min(min, placementMinX);
        }, chartOriginXBase);
        const labelMaxX = labelPlacements.reduce((max, placement) => {
            const placementMaxX = placement.anchor === "start" ? placement.x + placement.width : placement.x;
            return Math.max(max, placementMaxX);
        }, chartOriginXBase + chartWidth + rightLabelWidth);
        const shiftX = labelMode === "near" ? Math.max(0, -labelMinX) : 0;
        const chartOriginX = chartOriginXBase + shiftX;
        const svgWidth = (labelMaxX + shiftX) + RIGHT_PADDING;
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
            ".onBarLabelBg { fill: rgba(255,255,255,0.9); }",
            ".grid { stroke: #c9d3e1; stroke-width: 1; }",
            ".today { stroke: #ff6b5a; stroke-width: 2; }",
            "</style>",
            `<rect x="0" y="0" width="${svgWidth}" height="${svgHeight}" fill="#ffffff"/>`,
            `<text class="title" x="${chartOriginX + (chartWidth / 2)}" y="${TOP_PADDING + 18}" text-anchor="middle">${escapeXml(model.project.name || "-")}</text>`
        ];
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
            const useOnBarLabel = shouldUseDailyOnBarLabel(row, dateBand.length, labelMode);
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
                if (useOnBarLabel) {
                    const labelText = escapeXml(formatTaskLabel(row.task, labelMode));
                    const labelWidth = estimateLabelWidth(row.label, row.kind === "phase");
                    const labelCenterX = barX + (barWidth / 2);
                    const labelX = labelCenterX - (labelWidth / 2) - 8;
                    const labelY = row.kind === "phase" ? rowY + 15 : rowY + 10;
                    const labelHeight = row.kind === "phase" ? 18 : 20;
                    parts.push(`<rect class="onBarLabelBg" x="${labelX}" y="${labelY}" width="${labelWidth + 16}" height="${labelHeight}" rx="4"/>`);
                    parts.push(`<text class="${row.kind === "phase" ? "phaseLabel" : "label"}" x="${labelCenterX}" y="${rowY + 24}" text-anchor="middle">${labelText}</text>`);
                    continue;
                }
            }
            const labelPlacement = labelPlacements[rows.indexOf(row)];
            parts.push(`<text class="${row.kind === "phase" ? "phaseLabel" : "label"}" x="${labelPlacement.x + shiftX}" y="${rowY + 24}" text-anchor="${labelPlacement.anchor}">${escapeXml(formatTaskLabel(row.task, labelMode))}</text>`);
        }
        parts.push("</svg>");
        return parts.join("");
    }
    function exportWeeklyNativeSvg(model, options = {}) {
        const holidaySet = new Set([
            ...collectWbsHolidayDates(model),
            ...(options.holidayDates || []).map((day) => String(day || "").slice(0, 10)).filter(Boolean)
        ]);
        const nonWorkingDayTypes = collectProjectNonWorkingDayTypes(model);
        const dateBand = buildDisplayDateBand(model.project.startDate, model.project.finishDate, model.project.currentDate, options.displayDaysBeforeBaseDate, options.displayDaysAfterBaseDate, holidaySet, nonWorkingDayTypes, options.useBusinessDaysForDisplayRange);
        if (dateBand.length === 0) {
            return exportNativeSvg(model, { ...options, labelMode: "list" });
        }
        const weeklyBand = buildWeeklyBand(dateBand[0], dateBand[dateBand.length - 1]);
        const rows = buildWeeklyTaskRows(model.tasks, weeklyBand);
        const chartWidth = weeklyBand.length * WEEK_WIDTH;
        const chartOriginXBase = WEEKLY_LEFT_PADDING;
        const labelPlacements = rows.map((row) => resolveWeeklyLabelPlacement(row, chartOriginXBase, chartWidth));
        const labelMaxX = labelPlacements.reduce((max, placement) => {
            const placementMaxX = placement.anchor === "start" ? placement.x + placement.width : placement.x;
            return Math.max(max, placementMaxX);
        }, chartOriginXBase + chartWidth);
        const chartOriginX = chartOriginXBase;
        const svgWidth = labelMaxX + WEEKLY_RIGHT_PADDING;
        const svgHeight = WEEKLY_TOP_PADDING + WEEKLY_HEADER_HEIGHT + (rows.length * ROW_HEIGHT) + WEEKLY_BOTTOM_PADDING;
        const todayWeekIndex = indexOfWeekForDate(weeklyBand, model.project.currentDate);
        const parts = [
            `<svg xmlns="http://www.w3.org/2000/svg" width="${svgWidth}" height="${svgHeight}" viewBox="0 0 ${svgWidth} ${svgHeight}" role="img" aria-label="${escapeXml(model.project.name || "Project")} weekly overview">`,
            "<style>",
            "text { font-family: 'Hiragino Sans', 'Yu Gothic', sans-serif; fill: #1d2740; }",
            ".title { font-size: 18px; font-weight: 700; }",
            ".meta { font-size: 12px; fill: #5b6370; }",
            ".monthAxis { font-size: 13px; font-weight: 700; fill: #475467; }",
            ".weekAxis { font-size: 10px; fill: #667085; }",
            ".label { font-size: 13px; }",
            ".phaseLabel { font-size: 13px; font-weight: 700; }",
            ".grid { stroke: #c9d3e1; stroke-width: 1; }",
            ".today { stroke: #ff6b5a; stroke-width: 2; }",
            ".monthBoundary { stroke: #98a2b3; stroke-width: 1.5; }",
            "</style>",
            `<rect x="0" y="0" width="${svgWidth}" height="${svgHeight}" fill="#ffffff"/>`,
            `<text class="title" x="${chartOriginX + (chartWidth / 2)}" y="${WEEKLY_TOP_PADDING + 18}" text-anchor="middle">${escapeXml(model.project.name || "-")} weekly overview</text>`,
            `<text class="meta" x="${chartOriginX + (chartWidth / 2)}" y="${WEEKLY_TOP_PADDING + 40}" text-anchor="middle">${escapeXml(`project range ${String(model.project.startDate || "").slice(0, 10)} - ${String(model.project.finishDate || "").slice(0, 10)}`)}</text>`
        ];
        const chartOriginY = WEEKLY_TOP_PADDING + WEEKLY_HEADER_HEIGHT;
        const monthSpans = buildWeeklyMonthSpans(weeklyBand);
        for (const span of monthSpans) {
            const x = chartOriginX + (span.startIndex * WEEK_WIDTH);
            const width = (span.endIndex - span.startIndex + 1) * WEEK_WIDTH;
            parts.push(`<text class="monthAxis" x="${x + (width / 2)}" y="${WEEKLY_TOP_PADDING + 66}" text-anchor="middle">${escapeXml(span.monthKey)}</text>`);
            parts.push(`<line class="monthBoundary" x1="${x}" y1="${WEEKLY_TOP_PADDING + 72}" x2="${x}" y2="${svgHeight - WEEKLY_BOTTOM_PADDING}" />`);
        }
        parts.push(`<line class="monthBoundary" x1="${chartOriginX + chartWidth}" y1="${WEEKLY_TOP_PADDING + 72}" x2="${chartOriginX + chartWidth}" y2="${svgHeight - WEEKLY_BOTTOM_PADDING}" />`);
        for (let index = 0; index < weeklyBand.length; index += 1) {
            const week = weeklyBand[index];
            const x = chartOriginX + (index * WEEK_WIDTH);
            parts.push(`<line class="grid" x1="${x}" y1="${WEEKLY_TOP_PADDING + 72}" x2="${x}" y2="${svgHeight - WEEKLY_BOTTOM_PADDING}" />`);
            parts.push(`<text class="weekAxis" x="${x + (WEEK_WIDTH / 2)}" y="${WEEKLY_TOP_PADDING + 86}" text-anchor="middle">${escapeXml(formatWeeklyAxisLabel(week.startDay))}</text>`);
        }
        if (todayWeekIndex >= 0) {
            const todayX = chartOriginX + (todayWeekIndex * WEEK_WIDTH) + (WEEK_WIDTH / 2);
            parts.push(`<line class="today" x1="${todayX}" y1="${WEEKLY_TOP_PADDING + 72}" x2="${todayX}" y2="${svgHeight - WEEKLY_BOTTOM_PADDING}" />`);
        }
        for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
            const row = rows[rowIndex];
            const rowY = chartOriginY + row.y;
            if (row.startIndex === null || row.endIndex === null) {
                const fallbackLabelPlacement = labelPlacements[rowIndex];
                parts.push(`<text class="${row.kind === "phase" ? "phaseLabel" : "label"}" x="${fallbackLabelPlacement.x}" y="${rowY + 24}" text-anchor="${fallbackLabelPlacement.anchor}">${escapeXml(formatTaskLabel(row.task, "near"))}</text>`);
                continue;
            }
            const barX = chartOriginX + (row.startIndex * WEEK_WIDTH) + 4;
            const barWidth = Math.max(10, ((row.endIndex - row.startIndex + 1) * WEEK_WIDTH) - 8);
            const barY = rowY + 8;
            if (row.kind === "milestone") {
                const centerX = chartOriginX + (row.startIndex * WEEK_WIDTH) + (WEEK_WIDTH / 2);
                const centerY = rowY + (ROW_HEIGHT / 2);
                const isCompleted = (row.task.percentComplete || 0) >= 100;
                const fill = isCompleted ? "#d9efff" : "#ffffff";
                parts.push(`<polygon points="${centerX},${centerY - 11} ${centerX + 11},${centerY} ${centerX},${centerY + 11} ${centerX - 11},${centerY}" fill="${fill}" stroke="#4f95d6" stroke-width="3"/>`);
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
            const labelPlacement = labelPlacements[rowIndex];
            parts.push(`<text class="${row.kind === "phase" ? "phaseLabel" : "label"}" x="${labelPlacement.x}" y="${rowY + 24}" text-anchor="${labelPlacement.anchor}">${escapeXml(formatTaskLabel(row.task, "near"))}</text>`);
        }
        parts.push("</svg>");
        return parts.join("");
    }
    function exportMonthlyWbsCalendarSvgArchive(model) {
        const holidaySet = new Set(collectWbsHolidayDates(model));
        const nonWorkingDayTypes = collectProjectNonWorkingDayTypes(model);
        const months = buildProjectMonthKeys(model.project.startDate, model.project.finishDate);
        const entries = months.map((monthKey) => ({
            fileName: `mikuproject-monthly-wbs-calendar-${monthKey}.svg`,
            svg: exportMonthlyWbsCalendarSvg(model, monthKey, holidaySet, nonWorkingDayTypes)
        }));
        const zipEntries = entries.map((entry) => ({
            name: entry.fileName,
            data: encodeUtf8(entry.svg)
        }));
        const zipBytes = packZip(zipEntries);
        return {
            entries,
            zipBytes
        };
    }
    function exportMonthlyWbsCalendarSvg(model, monthKey, holidaySet, nonWorkingDayTypes) {
        const monthStart = parseDateOnly(`${monthKey}-01`);
        if (!monthStart) {
            throw new Error(`Invalid month key: ${monthKey}`);
        }
        const monthEnd = new Date(monthStart.getFullYear(), monthStart.getMonth() + 1, 0);
        const gridStart = startOfWeekSunday(monthStart);
        const gridEnd = endOfWeekSaturday(monthEnd);
        const calendarDays = buildDateBand(formatDateOnly(gridStart), formatDateOnly(gridEnd));
        const weeks = [];
        for (let index = 0; index < calendarDays.length; index += 7) {
            weeks.push(calendarDays.slice(index, index + 7));
        }
        const dayItems = buildMonthlyDayItemMap(model.tasks, calendarDays, holidaySet, nonWorkingDayTypes);
        const svgWidth = MONTHLY_LEFT_PADDING + (7 * MONTHLY_CELL_WIDTH) + MONTHLY_RIGHT_PADDING;
        const svgHeight = MONTHLY_TOP_PADDING + MONTHLY_HEADER_HEIGHT + MONTHLY_WEEKDAY_HEIGHT + (weeks.length * MONTHLY_CELL_HEIGHT) + MONTHLY_BOTTOM_PADDING;
        const title = `${model.project.name || "-"} ${formatMonthLabel(monthStart)}`;
        const weekdayLabels = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
        const parts = [
            `<svg xmlns="http://www.w3.org/2000/svg" width="${svgWidth}" height="${svgHeight}" viewBox="0 0 ${svgWidth} ${svgHeight}" role="img" aria-label="${escapeXml(title)}">`,
            "<style>",
            "text { font-family: 'Hiragino Sans', 'Yu Gothic', sans-serif; fill: #1d2740; }",
            ".title { font-size: 28px; font-weight: 700; }",
            ".meta { font-size: 13px; fill: #5b6370; }",
            ".weekday { font-size: 13px; font-weight: 700; fill: #5b6370; }",
            ".dayNumber { font-size: 16px; font-weight: 700; }",
            ".dayNumberMuted { font-size: 16px; font-weight: 700; fill: #97a2b0; }",
            ".itemText { font-size: 12px; }",
            ".moreText { font-size: 11px; fill: #5b6370; }",
            ".summaryText { font-size: 12px; font-weight: 700; }",
            ".milestoneText { font-size: 12px; font-weight: 700; fill: #9a284d; }",
            ".cellBorder { stroke: #d4dbe5; stroke-width: 1; }",
            "</style>",
            `<rect x="0" y="0" width="${svgWidth}" height="${svgHeight}" fill="#ffffff"/>`,
            `<text class="title" x="${MONTHLY_LEFT_PADDING}" y="${MONTHLY_TOP_PADDING + 28}">${escapeXml(model.project.name || "-")}</text>`,
            `<text class="meta" x="${MONTHLY_LEFT_PADDING}" y="${MONTHLY_TOP_PADDING + 56}">${escapeXml(formatMonthLabel(monthStart))}</text>`,
            `<text class="meta" x="${MONTHLY_LEFT_PADDING}" y="${MONTHLY_TOP_PADDING + 78}">${escapeXml(`project range ${String(model.project.startDate || "").slice(0, 10)} - ${String(model.project.finishDate || "").slice(0, 10)}`)}</text>`
        ];
        const weekdayY = MONTHLY_TOP_PADDING + MONTHLY_HEADER_HEIGHT;
        for (let dayIndex = 0; dayIndex < weekdayLabels.length; dayIndex += 1) {
            const x = MONTHLY_LEFT_PADDING + (dayIndex * MONTHLY_CELL_WIDTH);
            parts.push(`<text class="weekday" x="${x + 10}" y="${weekdayY}">${weekdayLabels[dayIndex]}</text>`);
        }
        const gridOriginY = MONTHLY_TOP_PADDING + MONTHLY_HEADER_HEIGHT + MONTHLY_WEEKDAY_HEIGHT;
        for (let weekIndex = 0; weekIndex < weeks.length; weekIndex += 1) {
            const week = weeks[weekIndex];
            for (let dayIndex = 0; dayIndex < week.length; dayIndex += 1) {
                const day = week[dayIndex];
                const x = MONTHLY_LEFT_PADDING + (dayIndex * MONTHLY_CELL_WIDTH);
                const y = gridOriginY + (weekIndex * MONTHLY_CELL_HEIGHT);
                const date = parseDateOnly(day);
                const inMonth = !!date && date.getMonth() === monthStart.getMonth() && date.getFullYear() === monthStart.getFullYear();
                const isHoliday = holidaySet.has(day);
                const isWeekend = isWeeklyNonWorkingDay(day, nonWorkingDayTypes);
                const background = !inMonth
                    ? "#f8fafc"
                    : (isHoliday ? "#fce7ef" : (isWeekend ? "#eef4f8" : "#ffffff"));
                parts.push(`<rect x="${x}" y="${y}" width="${MONTHLY_CELL_WIDTH}" height="${MONTHLY_CELL_HEIGHT}" fill="${background}" class="cellBorder"/>`);
                parts.push(`<text class="${inMonth ? "dayNumber" : "dayNumberMuted"}" x="${x + 10}" y="${y + 22}">${escapeXml(String(date ? date.getDate() : ""))}</text>`);
                const items = dayItems.get(day) || [];
                const visibleItems = items.slice(0, MONTHLY_MAX_ITEMS_PER_DAY);
                for (let itemIndex = 0; itemIndex < visibleItems.length; itemIndex += 1) {
                    const item = visibleItems[itemIndex];
                    const itemY = y + 42 + (itemIndex * 24);
                    parts.push(`<text class="${item.className}" x="${x + 10}" y="${itemY}">${escapeXml(item.label)}</text>`);
                }
                if (items.length > MONTHLY_MAX_ITEMS_PER_DAY) {
                    parts.push(`<text class="moreText" x="${x + 10}" y="${y + MONTHLY_CELL_HEIGHT - 10}">${escapeXml(`+${items.length - MONTHLY_MAX_ITEMS_PER_DAY} more`)}</text>`);
                }
            }
        }
        parts.push("</svg>");
        return parts.join("");
    }
    function buildMonthlyDayItemMap(tasks, calendarDays, holidaySet, nonWorkingDayTypes) {
        var _a, _b, _c;
        const dayItems = new Map();
        const calendarDaySet = new Set(calendarDays);
        for (const day of calendarDays) {
            dayItems.set(day, []);
        }
        for (const task of tasks) {
            const startDay = String(task.start || "").slice(0, 10);
            const finishDay = String(task.finish || "").slice(0, 10);
            if (!startDay || !finishDay) {
                continue;
            }
            if (task.milestone) {
                if (calendarDaySet.has(startDay)) {
                    (_a = dayItems.get(startDay)) === null || _a === void 0 ? void 0 : _a.push({ label: truncateLabel(task.name || "-"), className: "milestoneText" });
                }
                continue;
            }
            if (task.summary) {
                for (const day of uniqueDayList([startDay, finishDay])) {
                    if (calendarDaySet.has(day)) {
                        (_b = dayItems.get(day)) === null || _b === void 0 ? void 0 : _b.push({ label: truncateLabel(task.name || "-"), className: "summaryText" });
                    }
                }
                continue;
            }
            for (const day of buildDateBand(startDay, finishDay)) {
                if (!calendarDaySet.has(day)) {
                    continue;
                }
                const isBoundaryDay = day === startDay || day === finishDay;
                const isHoliday = holidaySet.has(day);
                const isWeekend = isWeeklyNonWorkingDay(day, nonWorkingDayTypes);
                if ((isHoliday || isWeekend) && !isBoundaryDay) {
                    continue;
                }
                (_c = dayItems.get(day)) === null || _c === void 0 ? void 0 : _c.push({ label: truncateLabel(task.name || "-"), className: "itemText" });
            }
        }
        return dayItems;
    }
    function truncateLabel(value) {
        const text = String(value || "").trim() || "-";
        if (text.length <= MONTHLY_MAX_LABEL_CHARS) {
            return text;
        }
        return `${text.slice(0, MONTHLY_MAX_LABEL_CHARS - 3)}...`;
    }
    function uniqueDayList(days) {
        return Array.from(new Set(days.filter(Boolean)));
    }
    function buildProjectMonthKeys(startDate, finishDate) {
        const start = parseDateOnly(startDate);
        const finish = parseDateOnly(finishDate);
        if (!start || !finish) {
            return [];
        }
        const cursor = new Date(start.getFullYear(), start.getMonth(), 1);
        const limit = new Date(finish.getFullYear(), finish.getMonth(), 1);
        const months = [];
        while (cursor.getTime() <= limit.getTime()) {
            months.push(`${cursor.getFullYear()}-${String(cursor.getMonth() + 1).padStart(2, "0")}`);
            cursor.setMonth(cursor.getMonth() + 1);
        }
        return months;
    }
    function startOfWeekSunday(date) {
        const result = new Date(date.getTime());
        result.setDate(result.getDate() - result.getDay());
        return result;
    }
    function endOfWeekSaturday(date) {
        const result = new Date(date.getTime());
        result.setDate(result.getDate() + (6 - result.getDay()));
        return result;
    }
    function formatMonthLabel(date) {
        return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
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
    function buildWeeklyTaskRows(tasks, weeklyBand) {
        return tasks.map((task, index) => ({
            task,
            label: task.name || "-",
            kind: task.summary ? "phase" : (task.milestone ? "milestone" : "task"),
            startIndex: indexOfWeekForDate(weeklyBand, task.start),
            endIndex: task.milestone ? indexOfWeekForDate(weeklyBand, task.start) : indexOfWeekForDate(weeklyBand, task.finish),
            y: index * ROW_HEIGHT
        }));
    }
    function formatTaskLabel(task, labelMode) {
        if (labelMode === "list") {
            return `${"　".repeat(Math.max(0, task.outlineLevel - 1))}${task.name || "-"}`;
        }
        return task.name || "-";
    }
    function resolveLabelPlacement(row, chartOriginX, chartWidth, labelMode) {
        if (labelMode === "list" || row.startIndex === null || row.endIndex === null) {
            return { x: LEFT_PADDING + 10, anchor: "start", width: estimateLabelWidth(row.label, row.kind === "phase") };
        }
        const textWidth = estimateLabelWidth(row.label, row.kind === "phase");
        const gap = 12;
        const shapeStartX = row.kind === "milestone"
            ? chartOriginX + (row.startIndex * DAY_WIDTH) + (DAY_WIDTH / 2) - 13
            : chartOriginX + (row.startIndex * DAY_WIDTH) + 6;
        const shapeEndX = row.kind === "milestone"
            ? chartOriginX + (row.startIndex * DAY_WIDTH) + (DAY_WIDTH / 2) + 13
            : chartOriginX + (row.endIndex * DAY_WIDTH) + DAY_WIDTH - 6;
        const chartMidX = chartOriginX + (chartWidth / 2);
        if (shapeEndX <= chartMidX) {
            return { x: shapeEndX + gap, anchor: "start", width: textWidth };
        }
        return { x: shapeStartX - gap, anchor: "end", width: textWidth };
    }
    function estimateLabelWidth(label, isPhase) {
        const basePerChar = isPhase ? 14 : 13;
        return Math.max(48, Math.ceil(String(label || "").length * basePerChar));
    }
    function shouldUseDailyOnBarLabel(row, bandLength, labelMode) {
        if (labelMode !== "near") {
            return false;
        }
        if (row.kind === "milestone") {
            return false;
        }
        if (row.startIndex === null || row.endIndex === null || bandLength <= 0) {
            return false;
        }
        return row.startIndex === 0 && row.endIndex === bandLength - 1;
    }
    function resolveWeeklyLabelPlacement(row, chartOriginX, chartWidth) {
        if (row.startIndex === null || row.endIndex === null) {
            return { x: chartOriginX + 10, anchor: "start", width: estimateLabelWidth(row.label, row.kind === "phase") };
        }
        const textWidth = estimateLabelWidth(row.label, row.kind === "phase");
        const gap = 10;
        const shapeStartX = row.kind === "milestone"
            ? chartOriginX + (row.startIndex * WEEK_WIDTH) + (WEEK_WIDTH / 2) - 11
            : chartOriginX + (row.startIndex * WEEK_WIDTH) + 4;
        const shapeEndX = row.kind === "milestone"
            ? chartOriginX + (row.startIndex * WEEK_WIDTH) + (WEEK_WIDTH / 2) + 11
            : chartOriginX + (row.endIndex * WEEK_WIDTH) + WEEK_WIDTH - 4;
        const chartMidX = chartOriginX + (chartWidth / 2);
        if (shapeEndX <= chartMidX) {
            return { x: shapeEndX + gap, anchor: "start", width: textWidth };
        }
        return { x: Math.max(textWidth, shapeStartX - gap), anchor: "end", width: textWidth };
    }
    function formatSvgAxisDate(day) {
        const match = day.match(/^(\d{4})-(\d{2})-(\d{2})$/);
        if (!match) {
            return day;
        }
        return `${Number(match[2])}/${Number(match[3])}`;
    }
    function formatWeeklyAxisLabel(day) {
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
    function indexOfWeekForDate(weeklyBand, value) {
        const key = String(value || "").slice(0, 10);
        if (!key) {
            return null;
        }
        for (let index = 0; index < weeklyBand.length; index += 1) {
            const week = weeklyBand[index];
            if (week.startDay <= key && key <= week.endDay) {
                return index;
            }
        }
        return null;
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
    function buildWeeklyBand(startDate, finishDate) {
        const start = parseDateOnly(startDate);
        const finish = parseDateOnly(finishDate);
        if (!start || !finish || start.getTime() > finish.getTime()) {
            return [];
        }
        const cursor = startOfWeekSunday(start);
        const limit = endOfWeekSaturday(finish);
        const weeks = [];
        while (cursor.getTime() <= limit.getTime()) {
            const weekStart = new Date(cursor.getTime());
            const weekEnd = new Date(cursor.getTime());
            weekEnd.setDate(weekEnd.getDate() + 6);
            weeks.push({
                startDay: formatDateOnly(weekStart),
                endDay: formatDateOnly(weekEnd),
                monthKey: formatMonthLabel(new Date(weekStart.getFullYear(), weekStart.getMonth(), 1))
            });
            cursor.setDate(cursor.getDate() + 7);
        }
        return weeks;
    }
    function buildWeeklyMonthSpans(weeklyBand) {
        const spans = [];
        for (let index = 0; index < weeklyBand.length; index += 1) {
            const current = weeklyBand[index];
            const last = spans[spans.length - 1];
            if (last && last.monthKey === current.monthKey) {
                last.endIndex = index;
            }
            else {
                spans.push({
                    monthKey: current.monthKey,
                    startIndex: index,
                    endIndex: index
                });
            }
        }
        return spans;
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
    function encodeUtf8(value) {
        return ZIP_TEXT_ENCODER.encode(value);
    }
    function packZip(entries) {
        const localParts = [];
        const centralParts = [];
        let offset = 0;
        for (const entry of entries) {
            const nameBytes = encodeUtf8(entry.name);
            const crc32 = computeCrc32(entry.data);
            const localHeader = new Uint8Array(30 + nameBytes.length);
            const localView = new DataView(localHeader.buffer);
            localView.setUint32(0, 0x04034b50, true);
            localView.setUint16(4, 20, true);
            localView.setUint16(6, 0, true);
            localView.setUint16(8, 0, true);
            localView.setUint16(10, ZIP_FIXED_MOD_TIME, true);
            localView.setUint16(12, ZIP_FIXED_MOD_DATE, true);
            localView.setUint32(14, crc32, true);
            localView.setUint32(18, entry.data.byteLength, true);
            localView.setUint32(22, entry.data.byteLength, true);
            localView.setUint16(26, nameBytes.length, true);
            localView.setUint16(28, 0, true);
            localHeader.set(nameBytes, 30);
            const centralHeader = new Uint8Array(46 + nameBytes.length);
            const centralView = new DataView(centralHeader.buffer);
            centralView.setUint32(0, 0x02014b50, true);
            centralView.setUint16(4, 20, true);
            centralView.setUint16(6, 20, true);
            centralView.setUint16(8, 0, true);
            centralView.setUint16(10, 0, true);
            centralView.setUint16(12, ZIP_FIXED_MOD_TIME, true);
            centralView.setUint16(14, ZIP_FIXED_MOD_DATE, true);
            centralView.setUint32(16, crc32, true);
            centralView.setUint32(20, entry.data.byteLength, true);
            centralView.setUint32(24, entry.data.byteLength, true);
            centralView.setUint16(28, nameBytes.length, true);
            centralView.setUint16(30, 0, true);
            centralView.setUint16(32, 0, true);
            centralView.setUint16(34, 0, true);
            centralView.setUint16(36, 0, true);
            centralView.setUint32(38, 0, true);
            centralView.setUint32(42, offset, true);
            centralHeader.set(nameBytes, 46);
            localParts.push(localHeader, entry.data);
            centralParts.push(centralHeader);
            offset += localHeader.byteLength + entry.data.byteLength;
        }
        const centralDirectoryOffset = offset;
        const centralDirectorySize = centralParts.reduce((sum, part) => sum + part.byteLength, 0);
        const endOfCentralDirectory = new Uint8Array(22);
        const endView = new DataView(endOfCentralDirectory.buffer);
        endView.setUint32(0, 0x06054b50, true);
        endView.setUint16(4, 0, true);
        endView.setUint16(6, 0, true);
        endView.setUint16(8, entries.length, true);
        endView.setUint16(10, entries.length, true);
        endView.setUint32(12, centralDirectorySize, true);
        endView.setUint32(16, centralDirectoryOffset, true);
        endView.setUint16(20, 0, true);
        return concatUint8Arrays([...localParts, ...centralParts, endOfCentralDirectory]);
    }
    function concatUint8Arrays(parts) {
        const totalLength = parts.reduce((sum, part) => sum + part.byteLength, 0);
        const result = new Uint8Array(totalLength);
        let offset = 0;
        for (const part of parts) {
            result.set(part, offset);
            offset += part.byteLength;
        }
        return result;
    }
    function computeCrc32(bytes) {
        let crc = 0xffffffff;
        for (const byte of bytes) {
            crc = (crc >>> 8) ^ ZIP_CRC32_TABLE[(crc ^ byte) & 0xff];
        }
        return (crc ^ 0xffffffff) >>> 0;
    }
    function buildCrc32Table() {
        const table = new Uint32Array(256);
        for (let index = 0; index < 256; index += 1) {
            let value = index;
            for (let bit = 0; bit < 8; bit += 1) {
                value = (value & 1) !== 0 ? (0xedb88320 ^ (value >>> 1)) : (value >>> 1);
            }
            table[index] = value >>> 0;
        }
        return table;
    }
    globalThis.__mikuprojectNativeSvg = {
        exportNativeSvg,
        exportWeeklyNativeSvg,
        exportMonthlyWbsCalendarSvgArchive
    };
})();
