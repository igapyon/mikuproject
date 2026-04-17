/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const HEADER_FILL = "#D9EAF7";
    const HEADER_ID_FILL = "#E1EDF8";
    const HEADER_STRUCTURE_FILL = "#E6F0DF";
    const HEADER_SCHEDULE_FILL = "#FDE7D3";
    const HEADER_STATUS_FILL = "#FBE4EC";
    const HEADER_ASSIGNMENT_FILL = "#E2F1EF";
    const SUMMARY_SCHEDULE_FILL = "#FDF1E4";
    const SUMMARY_ASSIGNMENT_FILL = "#E8F4F1";
    const PHASE_FILL = "#EEF7E8";
    const TASK_KIND_FILL = "#EEF2F6";
    const MILESTONE_FILL = "#FFF4E0";
    const IDENTIFIER_FILL = "#F7F9FC";
    const PLACEHOLDER_FILL = "#F5F7FA";
    const BAND_FILL = "#F4F7FB";
    const ACTIVE_BAND_FILL = "#D9EFFF";
    const PROGRESS_BAND_FILL = "#8EB9EA";
    const WEEKEND_BAND_FILL = "#EEF3F8";
    const WEEK_START_BAND_FILL = "#E3EEF9";
    const MONTH_BOUNDARY_WEEK_FILL = "#D6E7F8";
    const MONTH_START_HEADER_FILL = "#DCEAF7";
    const SATURDAY_HEADER_FILL = "#EEF3F8";
    const SUNDAY_HEADER_FILL = "#EEF3F8";
    const TODAY_BAND_FILL = "#FFE6A7";
    const TODAY_ACTIVE_BAND_FILL = "#C9DFF8";
    const TODAY_PROGRESS_BAND_FILL = "#6F9FD8";
    const HOLIDAY_BAND_FILL = "#FCE4EC";
    const DIVIDER_FILL = "#D9E2EA";
    const BASEDATE_GUIDE_TAIL_FILL = "#FFF8E1";
    const NAME_COLUMN_FILL = "#FBFCFE";
    const SCHEDULE_COLUMN_FILL = "#FCFAF7";
    const PROGRESS_COLUMN_FILL = "#FCF8FB";
    const REFERENCE_COLUMN_FILL = "#F8FBFB";
    const wbsXlsxLayout = globalThis.__mikuprojectWbsXlsxLayout;
    const wbsDateband = globalThis.__mikuprojectWbsDateband;
    const wbsXlsxSections = globalThis.__mikuprojectWbsXlsxSections;
    const wbsXlsxCells = globalThis.__mikuprojectWbsXlsxCells;
    const wbsXlsxBase = globalThis.__mikuprojectWbsXlsxBase;
    const wbsXlsxTaskmeta = globalThis.__mikuprojectWbsXlsxTaskmeta;
    const wbsXlsxExport = globalThis.__mikuprojectWbsXlsxExport;
    const wbsXlsxPublic = globalThis.__mikuprojectWbsXlsxPublic;
    if (!wbsXlsxLayout) {
        throw new Error("mikuproject WBS xlsx layout module is not loaded");
    }
    if (!wbsDateband) {
        throw new Error("mikuproject WBS dateband module is not loaded");
    }
    if (!wbsXlsxSections) {
        throw new Error("mikuproject WBS xlsx sections module is not loaded");
    }
    if (!wbsXlsxCells) {
        throw new Error("mikuproject WBS xlsx cells module is not loaded");
    }
    if (!wbsXlsxBase) {
        throw new Error("mikuproject WBS xlsx base module is not loaded");
    }
    if (!wbsXlsxTaskmeta) {
        throw new Error("mikuproject WBS xlsx taskmeta module is not loaded");
    }
    if (!wbsXlsxExport) {
        throw new Error("mikuproject WBS xlsx export module is not loaded");
    }
    if (!wbsXlsxPublic) {
        throw new Error("mikuproject WBS xlsx public module is not loaded");
    }
    const WBS_LAYOUT = wbsXlsxLayout.createWbsSheetLayoutHelper();
    const WBS_DATEBAND = wbsDateband.createWbsDatebandHelper();
    const WBS_SECTIONS = wbsXlsxSections.createWbsXlsxSectionsHelper({
        layout: WBS_LAYOUT,
        fills: {
            headerId: HEADER_ID_FILL,
            headerFill: HEADER_FILL,
            headerStructure: HEADER_STRUCTURE_FILL,
            headerSchedule: HEADER_SCHEDULE_FILL,
            headerStatus: HEADER_STATUS_FILL,
            headerAssignment: HEADER_ASSIGNMENT_FILL,
            summarySchedule: SUMMARY_SCHEDULE_FILL,
            summaryAssignment: SUMMARY_ASSIGNMENT_FILL,
            phase: PHASE_FILL,
            milestone: MILESTONE_FILL,
            placeholder: PLACEHOLDER_FILL,
            divider: DIVIDER_FILL,
            referenceColumn: REFERENCE_COLUMN_FILL
        }
    });
    const WBS_TASKMETA = wbsXlsxTaskmeta.createWbsXlsxTaskmetaHelper();
    const WBS_CELLS = wbsXlsxCells.createWbsXlsxCellsHelper({
        layout: WBS_LAYOUT,
        fills: {
            divider: DIVIDER_FILL,
            headerId: HEADER_ID_FILL,
            headerFill: HEADER_FILL,
            headerStructure: HEADER_STRUCTURE_FILL,
            headerSchedule: HEADER_SCHEDULE_FILL,
            headerStatus: HEADER_STATUS_FILL,
            headerAssignment: HEADER_ASSIGNMENT_FILL,
            phase: PHASE_FILL,
            taskKind: TASK_KIND_FILL,
            milestone: MILESTONE_FILL,
            identifier: IDENTIFIER_FILL,
            placeholder: PLACEHOLDER_FILL,
            activeBand: ACTIVE_BAND_FILL,
            progressBand: PROGRESS_BAND_FILL,
            weekendBand: WEEKEND_BAND_FILL,
            weekStartBand: WEEK_START_BAND_FILL,
            monthBoundaryWeek: MONTH_BOUNDARY_WEEK_FILL,
            monthStartHeader: MONTH_START_HEADER_FILL,
            saturdayHeader: SATURDAY_HEADER_FILL,
            sundayHeader: SUNDAY_HEADER_FILL,
            todayBand: TODAY_BAND_FILL,
            todayActiveBand: TODAY_ACTIVE_BAND_FILL,
            todayProgressBand: TODAY_PROGRESS_BAND_FILL,
            holidayBand: HOLIDAY_BAND_FILL,
            nameColumn: NAME_COLUMN_FILL,
            scheduleColumn: SCHEDULE_COLUMN_FILL,
            progressColumn: PROGRESS_COLUMN_FILL
        },
        formatTaskLabel: (task) => WBS_TASKMETA.formatTaskLabel(task),
        classifyTaskKind: (task) => WBS_TASKMETA.classifyTaskKind(task),
        buildDateBand: (startDate, finishDate) => WBS_DATEBAND.buildDateBand(startDate, finishDate),
        parseDateOnly: (value) => WBS_DATEBAND.parseDateOnly(value),
        formatDateOnly: (value) => WBS_DATEBAND.formatDateOnly(value),
        isWeeklyNonWorkingDay: (day, nonWorkingDayTypes) => WBS_DATEBAND.isWeeklyNonWorkingDay(day, nonWorkingDayTypes)
    });
    const WBS_BASE = wbsXlsxBase.createWbsXlsxBaseHelper();
    const pxWidth = (pixels) => Math.round((pixels / 7) * 100) / 100;
    const collectWbsHolidayDates = (model) => WBS_DATEBAND.collectWbsHolidayDates(model);
    const collectProjectNonWorkingDayTypes = (model) => WBS_DATEBAND.collectProjectNonWorkingDayTypes(model);
    const buildDateBand = (startDate, finishDate) => WBS_DATEBAND.buildDateBand(startDate, finishDate);
    const buildDisplayDateBand = (startDate, finishDate, baseDate, displayDaysBeforeBaseDate, displayDaysAfterBaseDate, holidaySet, nonWorkingDayTypes, useBusinessDaysForDisplayRange) => WBS_DATEBAND.buildDisplayDateBand(startDate, finishDate, baseDate, displayDaysBeforeBaseDate, displayDaysAfterBaseDate, holidaySet, nonWorkingDayTypes, useBusinessDaysForDisplayRange);
    const isWeeklyNonWorkingDay = (day, nonWorkingDayTypes) => WBS_DATEBAND.isWeeklyNonWorkingDay(day, nonWorkingDayTypes);
    const parseDateOnly = (value) => WBS_DATEBAND.parseDateOnly(value);
    const formatDateOnly = (value) => WBS_DATEBAND.formatDateOnly(value);
    const firstResourceName = (resourceNames) => WBS_SECTIONS.firstResourceName(resourceNames);
    const formatCalendarLabel = (calendarUID, calendarNameByUid) => WBS_SECTIONS.formatCalendarLabel(calendarUID, calendarNameByUid);
    const truncateWbsReference = (value, maxLength) => WBS_SECTIONS.truncateWbsReference(value, maxLength);
    const referenceCell = (task, value, horizontalAlign = "center") => WBS_SECTIONS.referenceCell(task, value, horizontalAlign);
    const projectInfoRows = (project, calendarNameByUid, holidayCount, columnCount, startColumnIndex, startRowNumber) => WBS_SECTIONS.projectInfoRows(project, calendarNameByUid, holidayCount, columnCount, startColumnIndex, startRowNumber);
    const legendRows = (columnCount, startRowNumber) => WBS_SECTIONS.legendRows(columnCount, startRowNumber);
    const displaySummaryRows = (displayDays, businessDays, baseDate, taskCount, resourceCount, assignmentCount, calendarCount, columnCount, startColumnIndex = 5, startRowNumber = 5, displayDaysBeforeBaseDate, displayDaysAfterBaseDate, useBusinessDaysForDisplayRange, useBusinessDaysForProgressBand) => WBS_SECTIONS.displaySummaryRows(displayDays, businessDays, baseDate, taskCount, resourceCount, assignmentCount, calendarCount, columnCount, startColumnIndex, startRowNumber, displayDaysBeforeBaseDate, displayDaysAfterBaseDate, useBusinessDaysForDisplayRange, useBusinessDaysForProgressBand);
    const formatWbsDate = (value) => WBS_SECTIONS.formatWbsDate(value);
    const formatTaskLabel = (task) => WBS_TASKMETA.formatTaskLabel(task);
    const classifyTaskKind = (task) => WBS_TASKMETA.classifyTaskKind(task);
    const weekBandRow = (fixedColumnCount, weekBandRanges, dateBandLength) => WBS_CELLS.weekBandRow(fixedColumnCount, weekBandRanges, dateBandLength);
    const headerRow = (labels) => WBS_CELLS.headerRow(labels);
    const weekdayRow = (fixedColumnCount, dateBand, currentDate, holidaySet, nonWorkingDayTypes) => WBS_CELLS.weekdayRow(fixedColumnCount, dateBand, currentDate, holidaySet, nonWorkingDayTypes);
    const dateBandHeaderRow = (fixedColumnCount, dateBand, currentDate, holidaySet, nonWorkingDayTypes) => WBS_CELLS.dateBandHeaderRow(fixedColumnCount, dateBand, currentDate, holidaySet, nonWorkingDayTypes);
    const weekdayHeaderRow = (fixedHeaders, dateBand, currentDate, holidaySet, nonWorkingDayTypes) => WBS_CELLS.weekdayHeaderRow(fixedHeaders, dateBand, currentDate, holidaySet, nonWorkingDayTypes);
    const dividerCell = () => WBS_CELLS.dividerCell();
    const taskCell = (task, value, horizontalAlign = "left") => WBS_CELLS.taskCell(task, value, horizontalAlign);
    const detailCell = (task, value) => WBS_CELLS.detailCell(task, value);
    const taskRowHeight = (task) => WBS_CELLS.taskRowHeight(task);
    const kindCell = (task) => WBS_CELLS.kindCell(task);
    const identifierCell = (task, value) => WBS_CELLS.identifierCell(task, value);
    const flagCell = (task, enabled, marker) => WBS_CELLS.flagCell(task, enabled, marker);
    const progressCell = (task, value) => WBS_CELLS.progressCell(task, value);
    const formatDurationLabel = (task, holidaySet, nonWorkingDayTypes, useBusinessDaysForProgressBand) => WBS_CELLS.formatDurationLabel(task, holidaySet, nonWorkingDayTypes, useBusinessDaysForProgressBand);
    const dateBandCell = (task, day, currentDate, holidaySet, nonWorkingDayTypes, useBusinessDaysForProgressBand) => WBS_CELLS.dateBandCell(task, day, currentDate, holidaySet, nonWorkingDayTypes, useBusinessDaysForProgressBand);
    const countBusinessDays = (dateBand, holidaySet, nonWorkingDayTypes) => WBS_CELLS.countBusinessDays(dateBand, holidaySet, nonWorkingDayTypes);
    const buildWeekBandRanges = (dateBand, startColumnIndex, rowNumber) => WBS_CELLS.buildWeekBandRanges(dateBand, startColumnIndex, rowNumber);
    const emptyRow = (columnCount, height = 22) => WBS_BASE.emptyRow(columnCount, height);
    const overlayRows = (rows, startIndex, blockRows, columnCount) => WBS_BASE.overlayRows(rows, startIndex, blockRows, columnCount);
    const sheetTitleRow = (title, columnCount) => WBS_BASE.sheetTitleRow(title, columnCount);
    const infoRow = (text, columnCount) => WBS_BASE.infoRow(text, columnCount);
    const formatWbsExportTimestamp = (value) => WBS_BASE.formatWbsExportTimestamp(value);
    const WBS_EXPORT = wbsXlsxExport.createWbsXlsxExportHelper({
        pxWidth,
        collectWbsHolidayDates: (model) => collectWbsHolidayDates(model),
        collectProjectNonWorkingDayTypes: (model) => collectProjectNonWorkingDayTypes(model),
        buildDisplayDateBand: (startDate, finishDate, baseDate, displayDaysBeforeBaseDate, displayDaysAfterBaseDate, holidaySet, nonWorkingDayTypes, useBusinessDaysForDisplayRange) => buildDisplayDateBand(startDate, finishDate, baseDate, displayDaysBeforeBaseDate, displayDaysAfterBaseDate, holidaySet, nonWorkingDayTypes, useBusinessDaysForDisplayRange),
        firstResourceName: (resourceNames) => firstResourceName(resourceNames),
        formatCalendarLabel: (calendarUID, calendarNameByUid) => formatCalendarLabel(calendarUID, calendarNameByUid),
        truncateWbsReference: (value, maxLength) => truncateWbsReference(value, maxLength),
        referenceCell: (task, value, horizontalAlign) => referenceCell(task, value, horizontalAlign),
        projectInfoRows: (project, calendarNameByUid, holidayCount, columnCount, startColumnIndex, startRowNumber) => projectInfoRows(project, calendarNameByUid, holidayCount, columnCount, startColumnIndex, startRowNumber),
        legendRows: (columnCount, startRowNumber) => legendRows(columnCount, startRowNumber),
        displaySummaryRows: (displayDays, businessDays, baseDate, taskCount, resourceCount, assignmentCount, calendarCount, columnCount, startColumnIndex, startRowNumber, displayDaysBeforeBaseDate, displayDaysAfterBaseDate, useBusinessDaysForDisplayRange, useBusinessDaysForProgressBand) => displaySummaryRows(displayDays, businessDays, baseDate, taskCount, resourceCount, assignmentCount, calendarCount, columnCount, startColumnIndex, startRowNumber, displayDaysBeforeBaseDate, displayDaysAfterBaseDate, useBusinessDaysForDisplayRange, useBusinessDaysForProgressBand),
        formatWbsDate: (value) => formatWbsDate(value),
        dateBandHeaderRow: (fixedColumnCount, dateBand, currentDate, holidaySet, nonWorkingDayTypes) => dateBandHeaderRow(fixedColumnCount, dateBand, currentDate, holidaySet, nonWorkingDayTypes),
        weekdayHeaderRow: (fixedHeaders, dateBand, currentDate, holidaySet, nonWorkingDayTypes) => weekdayHeaderRow(fixedHeaders, dateBand, currentDate, holidaySet, nonWorkingDayTypes),
        dividerCell: () => dividerCell(),
        taskCell: (task, value, horizontalAlign) => taskCell(task, value, horizontalAlign),
        detailCell: (task, value) => detailCell(task, value),
        taskRowHeight: (task) => taskRowHeight(task),
        kindCell: (task) => kindCell(task),
        identifierCell: (task, value) => identifierCell(task, value),
        flagCell: (task, enabled, marker) => flagCell(task, enabled, marker),
        progressCell: (task, value) => progressCell(task, value),
        formatDurationLabel: (task, holidaySet, nonWorkingDayTypes, useBusinessDaysForProgressBand) => formatDurationLabel(task, holidaySet, nonWorkingDayTypes, useBusinessDaysForProgressBand),
        dateBandCell: (task, day, currentDate, holidaySet, nonWorkingDayTypes, useBusinessDaysForProgressBand) => dateBandCell(task, day, currentDate, holidaySet, nonWorkingDayTypes, useBusinessDaysForProgressBand),
        countBusinessDays: (dateBand, holidaySet, nonWorkingDayTypes) => countBusinessDays(dateBand, holidaySet, nonWorkingDayTypes),
        emptyRow: (columnCount, height) => emptyRow(columnCount, height),
        overlayRows: (rows, startIndex, blockRows, columnCount) => overlayRows(rows, startIndex, blockRows, columnCount),
        formatWbsExportTimestamp: (value) => formatWbsExportTimestamp(value),
        formatTaskLabel: (task) => formatTaskLabel(task)
    });
    function exportWbsWorkbook(model, options = {}) {
        return WBS_EXPORT.exportWbsWorkbook(model, options);
    }
    const WBS_XLSX_FACADE = wbsXlsxPublic.createWbsXlsxFacade({
        collectWbsHolidayDates: (model) => collectWbsHolidayDates(model),
        exportWbsWorkbook: (model, options) => exportWbsWorkbook(model, options),
        layout: WBS_LAYOUT
    });
    globalThis.__mikuprojectWbsXlsx = WBS_XLSX_FACADE;
})();
