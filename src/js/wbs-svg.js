/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const WEEK_WIDTH = 38;
    const DAY_WIDTH = WEEK_WIDTH;
    const LIST_LABEL_WIDTH = 360;
    const NEAR_LEFT_LABEL_WIDTH = 0;
    const NEAR_RIGHT_LABEL_WIDTH = 0;
    const HEADER_HEIGHT = 82;
    const ROW_HEIGHT = 38;
    const LEFT_PADDING = 0;
    const TOP_PADDING = 22;
    const RIGHT_PADDING = 0;
    const BOTTOM_PADDING = 28;
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
    const WEEKLY_HEADER_HEIGHT = 96;
    const WEEKLY_TOP_PADDING = 22;
    const WEEKLY_BOTTOM_PADDING = 28;
    const WEEKLY_LEFT_PADDING = 0;
    const WEEKLY_RIGHT_PADDING = 0;
    const WEEKLY_TRIM_PADDING = 16;
    const wbsDateband = globalThis.__mikuprojectWbsDateband;
    const wbsSvgLabels = globalThis.__mikuprojectWbsSvgLabels;
    const wbsSvgCalendar = globalThis.__mikuprojectWbsSvgCalendar;
    const wbsSvgTimeline = globalThis.__mikuprojectWbsSvgTimeline;
    const wbsSvgZip = globalThis.__mikuprojectWbsSvgZip;
    const wbsSvgBars = globalThis.__mikuprojectWbsSvgBars;
    const wbsSvgViewport = globalThis.__mikuprojectWbsSvgViewport;
    const wbsSvgAxis = globalThis.__mikuprojectWbsSvgAxis;
    const wbsSvgScaffold = globalThis.__mikuprojectWbsSvgScaffold;
    const wbsSvgRender = globalThis.__mikuprojectWbsSvgRender;
    const wbsSvgPublic = globalThis.__mikuprojectWbsSvgPublic;
    if (!wbsDateband) {
        throw new Error("mikuproject WBS dateband module is not loaded");
    }
    if (!wbsSvgLabels) {
        throw new Error("mikuproject WBS svg labels module is not loaded");
    }
    if (!wbsSvgCalendar) {
        throw new Error("mikuproject WBS svg calendar module is not loaded");
    }
    if (!wbsSvgTimeline) {
        throw new Error("mikuproject WBS svg timeline module is not loaded");
    }
    if (!wbsSvgZip) {
        throw new Error("mikuproject WBS svg zip module is not loaded");
    }
    if (!wbsSvgBars) {
        throw new Error("mikuproject WBS svg bars module is not loaded");
    }
    if (!wbsSvgViewport) {
        throw new Error("mikuproject WBS svg viewport module is not loaded");
    }
    if (!wbsSvgAxis) {
        throw new Error("mikuproject WBS svg axis module is not loaded");
    }
    if (!wbsSvgScaffold) {
        throw new Error("mikuproject WBS svg scaffold module is not loaded");
    }
    if (!wbsSvgRender) {
        throw new Error("mikuproject WBS svg render module is not loaded");
    }
    if (!wbsSvgPublic) {
        throw new Error("mikuproject WBS svg public module is not loaded");
    }
    const WBS_DATEBAND = wbsDateband.createWbsDatebandHelper();
    const WBS_SVG_LABELS = wbsSvgLabels.createWbsSvgLabelsHelper({
        leftPadding: LEFT_PADDING,
        dayWidth: DAY_WIDTH,
        weekWidth: WEEK_WIDTH,
        rowHeight: ROW_HEIGHT
    });
    const WBS_SVG_CALENDAR = wbsSvgCalendar.createWbsSvgCalendarHelper({
        monthlyCellWidth: MONTHLY_CELL_WIDTH,
        monthlyCellHeight: MONTHLY_CELL_HEIGHT,
        monthlyHeaderHeight: MONTHLY_HEADER_HEIGHT,
        monthlyWeekdayHeight: MONTHLY_WEEKDAY_HEIGHT,
        monthlyLeftPadding: MONTHLY_LEFT_PADDING,
        monthlyTopPadding: MONTHLY_TOP_PADDING,
        monthlyRightPadding: MONTHLY_RIGHT_PADDING,
        monthlyBottomPadding: MONTHLY_BOTTOM_PADDING,
        monthlyMaxItemsPerDay: MONTHLY_MAX_ITEMS_PER_DAY,
        monthlyMaxLabelChars: MONTHLY_MAX_LABEL_CHARS,
        collectWbsHolidayDates: (model) => collectWbsHolidayDates(model),
        collectProjectNonWorkingDayTypes: (model) => collectProjectNonWorkingDayTypes(model),
        parseDateOnly: (value) => parseDateOnly(value),
        formatDateOnly: (value) => formatDateOnly(value),
        buildDateBand: (startDate, finishDate) => buildDateBand(startDate, finishDate),
        isWeeklyNonWorkingDay: (day, nonWorkingDayTypes) => isWeeklyNonWorkingDay(day, nonWorkingDayTypes),
        escapeXml: (value) => escapeXml(value),
        encodeUtf8: (value) => encodeUtf8(value),
        packZip: (entries) => packZip(entries)
    });
    const WBS_SVG_TIMELINE = wbsSvgTimeline.createWbsSvgTimelineHelper({
        rowHeight: ROW_HEIGHT,
        parseDateOnly: (value) => parseDateOnly(value),
        formatDateOnly: (value) => formatDateOnly(value)
    });
    const WBS_SVG_ZIP = wbsSvgZip.createWbsSvgZipHelper({
        fixedModTime: ZIP_FIXED_MOD_TIME,
        fixedModDate: ZIP_FIXED_MOD_DATE
    });
    const WBS_SVG_BARS = wbsSvgBars.createWbsSvgBarsHelper();
    const WBS_SVG_VIEWPORT = wbsSvgViewport.createWbsSvgViewportHelper();
    const WBS_SVG_AXIS = wbsSvgAxis.createWbsSvgAxisHelper();
    const WBS_SVG_SCAFFOLD = wbsSvgScaffold.createWbsSvgScaffoldHelper();
    const WBS_SVG_RENDER = wbsSvgRender.createWbsSvgRenderHelper({
        constants: {
            dayWidth: DAY_WIDTH,
            weekWidth: WEEK_WIDTH,
            listLabelWidth: LIST_LABEL_WIDTH,
            nearLeftLabelWidth: NEAR_LEFT_LABEL_WIDTH,
            nearRightLabelWidth: NEAR_RIGHT_LABEL_WIDTH,
            topPadding: TOP_PADDING,
            headerHeight: HEADER_HEIGHT,
            bottomPadding: BOTTOM_PADDING,
            leftPadding: LEFT_PADDING,
            rightPadding: RIGHT_PADDING,
            rowHeight: ROW_HEIGHT,
            weeklyTopPadding: WEEKLY_TOP_PADDING,
            weeklyHeaderHeight: WEEKLY_HEADER_HEIGHT,
            weeklyBottomPadding: WEEKLY_BOTTOM_PADDING,
            weeklyLeftPadding: WEEKLY_LEFT_PADDING,
            weeklyRightPadding: WEEKLY_RIGHT_PADDING,
            weeklyTrimPadding: WEEKLY_TRIM_PADDING
        },
        collectWbsHolidayDates: (model) => collectWbsHolidayDates(model),
        collectProjectNonWorkingDayTypes: (model) => collectProjectNonWorkingDayTypes(model),
        buildDisplayDateBand: (startDate, finishDate, baseDate, displayDaysBeforeBaseDate, displayDaysAfterBaseDate, holidaySet, nonWorkingDayTypes, useBusinessDaysForDisplayRange) => buildDisplayDateBand(startDate, finishDate, baseDate, displayDaysBeforeBaseDate, displayDaysAfterBaseDate, holidaySet, nonWorkingDayTypes, useBusinessDaysForDisplayRange),
        buildTaskRows: (tasks, dateBand) => buildTaskRows(tasks, dateBand),
        buildWeeklyTaskRows: (tasks, weeklyBand) => buildWeeklyTaskRows(tasks, weeklyBand),
        buildWeeklyBand: (startDate, finishDate) => buildWeeklyBand(startDate, finishDate),
        buildWeeklyMonthSpans: (weeklyBand) => buildWeeklyMonthSpans(weeklyBand),
        formatTaskLabel: (task, labelMode) => formatTaskLabel(task, labelMode),
        resolveLabelPlacement: (row, chartOriginX, chartWidth, labelMode) => resolveLabelPlacement(row, chartOriginX, chartWidth, labelMode),
        resolveWeeklyLabelPlacement: (row, chartOriginX, chartWidth) => resolveWeeklyLabelPlacement(row, chartOriginX, chartWidth),
        shouldUseDailyOnBarLabel: (row, chartOriginX, chartWidth, bandLength, labelMode) => shouldUseDailyOnBarLabel(row, chartOriginX, chartWidth, bandLength, labelMode),
        indexOfDate: (dateBand, value) => indexOfDate(dateBand, value),
        indexOfWeekForDate: (weeklyBand, value) => indexOfWeekForDate(weeklyBand, value),
        formatSvgAxisDate: (day) => formatSvgAxisDate(day),
        formatWeeklyAxisLabel: (day) => formatWeeklyAxisLabel(day),
        isWeeklyNonWorkingDay: (day, nonWorkingDayTypes) => isWeeklyNonWorkingDay(day, nonWorkingDayTypes),
        buildDependencyConnectorElements: (tasks, rows, chartOriginX, chartOriginY, config) => buildDependencyConnectorElements(tasks, rows, chartOriginX, chartOriginY, config),
        escapeXml: (value) => escapeXml(value),
        bars: WBS_SVG_BARS,
        viewport: WBS_SVG_VIEWPORT,
        axis: WBS_SVG_AXIS,
        scaffold: WBS_SVG_SCAFFOLD
    });
    const collectWbsHolidayDates = (model) => WBS_DATEBAND.collectWbsHolidayDates(model);
    const collectProjectNonWorkingDayTypes = (model) => WBS_DATEBAND.collectProjectNonWorkingDayTypes(model);
    const buildDateBand = (startDate, finishDate) => WBS_DATEBAND.buildDateBand(startDate, finishDate);
    const buildDisplayDateBand = (startDate, finishDate, baseDate, displayDaysBeforeBaseDate, displayDaysAfterBaseDate, holidaySet, nonWorkingDayTypes, useBusinessDaysForDisplayRange) => WBS_DATEBAND.buildDisplayDateBand(startDate, finishDate, baseDate, displayDaysBeforeBaseDate, displayDaysAfterBaseDate, holidaySet, nonWorkingDayTypes, useBusinessDaysForDisplayRange);
    const isWeeklyNonWorkingDay = (day, nonWorkingDayTypes) => WBS_DATEBAND.isWeeklyNonWorkingDay(day, nonWorkingDayTypes);
    const parseDateOnly = (value) => WBS_DATEBAND.parseDateOnly(value);
    const formatDateOnly = (value) => WBS_DATEBAND.formatDateOnly(value);
    const formatTaskLabel = (task, labelMode) => WBS_SVG_LABELS.formatTaskLabel(task, labelMode);
    const resolveLabelPlacement = (row, chartOriginX, chartWidth, labelMode) => WBS_SVG_LABELS.resolveLabelPlacement(row, chartOriginX, chartWidth, labelMode);
    const resolveWeeklyLabelPlacement = (row, chartOriginX, chartWidth) => WBS_SVG_LABELS.resolveWeeklyLabelPlacement(row, chartOriginX, chartWidth);
    const shouldUseDailyOnBarLabel = (row, chartOriginX, chartWidth, bandLength, labelMode) => WBS_SVG_LABELS.shouldUseDailyOnBarLabel(row, chartOriginX, chartWidth, bandLength, labelMode);
    const buildDependencyConnectorElements = (tasks, rows, chartOriginX, chartOriginY, config) => WBS_SVG_LABELS.buildDependencyConnectorElements(tasks, rows, chartOriginX, chartOriginY, config);
    const exportMonthlyWbsCalendarSvgArchive = (model) => WBS_SVG_CALENDAR.exportMonthlyWbsCalendarSvgArchive(model);
    const exportMonthlyWbsCalendarSvg = (model, monthKey, holidaySet, nonWorkingDayTypes) => WBS_SVG_CALENDAR.exportMonthlyWbsCalendarSvg(model, monthKey, holidaySet, nonWorkingDayTypes);
    const buildTaskRows = (tasks, dateBand) => WBS_SVG_TIMELINE.buildTaskRows(tasks, dateBand);
    const buildWeeklyTaskRows = (tasks, weeklyBand) => WBS_SVG_TIMELINE.buildWeeklyTaskRows(tasks, weeklyBand);
    const formatSvgAxisDate = (day) => WBS_SVG_TIMELINE.formatSvgAxisDate(day);
    const formatWeeklyAxisLabel = (day) => WBS_SVG_TIMELINE.formatWeeklyAxisLabel(day);
    const indexOfDate = (dateBand, value) => WBS_SVG_TIMELINE.indexOfDate(dateBand, value);
    const indexOfWeekForDate = (weeklyBand, value) => WBS_SVG_TIMELINE.indexOfWeekForDate(weeklyBand, value);
    const buildWeeklyBand = (startDate, finishDate) => WBS_SVG_TIMELINE.buildWeeklyBand(startDate, finishDate);
    const buildWeeklyMonthSpans = (weeklyBand) => WBS_SVG_TIMELINE.buildWeeklyMonthSpans(weeklyBand);
    const escapeXml = (value) => WBS_SVG_ZIP.escapeXml(value);
    const encodeUtf8 = (value) => WBS_SVG_ZIP.encodeUtf8(value);
    const packZip = (entries) => WBS_SVG_ZIP.packZip(entries);
    function exportNativeSvg(model, options = {}) {
        return WBS_SVG_RENDER.exportNativeSvg(model, options);
    }
    function exportWeeklyNativeSvg(model, options = {}) {
        return WBS_SVG_RENDER.exportWeeklyNativeSvg(model, options);
    }
    const NATIVE_SVG_FACADE = wbsSvgPublic.createNativeSvgFacade({
        collectWbsHolidayDates: (model) => collectWbsHolidayDates(model),
        exportNativeSvg: (model, options) => exportNativeSvg(model, options),
        exportWeeklyNativeSvg: (model, options) => exportWeeklyNativeSvg(model, options),
        exportMonthlyWbsCalendarSvgArchive: (model) => exportMonthlyWbsCalendarSvgArchive(model)
    });
    globalThis.__mikuprojectNativeSvg = NATIVE_SVG_FACADE;
})();
