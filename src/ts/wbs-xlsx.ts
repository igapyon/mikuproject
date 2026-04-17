/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  type WbsXlsxCellLike = {
    value?: string | number | boolean;
    numberFormat?: "general" | "integer" | "decimal" | "date" | "datetime" | "percent";
    horizontalAlign?: "left" | "center" | "right";
    verticalAlign?: "top" | "center" | "bottom";
    wrapText?: boolean;
    bold?: boolean;
    fontSize?: number;
    fillColor?: string;
    border?: "thin";
  };

  type WbsXlsxWorkbookLike = {
    sheets: Array<{
      name: string;
      columns?: Array<{ width?: number; hidden?: boolean }>;
      mergedRanges?: string[];
      rows: Array<{
        height?: number;
        cells: WbsXlsxCellLike[];
      }>;
    }>;
  };

  type WbsExportOptions = {
    holidayDates?: string[];
    displayDaysBeforeBaseDate?: number;
    displayDaysAfterBaseDate?: number;
    useBusinessDaysForDisplayRange?: boolean;
    useBusinessDaysForProgressBand?: boolean;
  };

  type WbsSheetLayoutHelper = {
    columnName(columnIndex: number): string;
    columnIndex(columnReference: string): number;
    reference(rowNumber: number, columnIndex: number): string;
    parseCellReference(reference: string): {
      reference: string;
      rowNumber: number;
      rowIndex: number;
      columnName: string;
      columnIndex: number;
    };
    range(startReference: string, endReference: string): string;
    describeCell(reference: string): string;
    logCell(reference: string, label?: string, logger?: (...args: unknown[]) => void): string;
  };

  type WbsDatebandHelper = {
    parseDateOnly(value: string | undefined): Date | null;
    formatDateOnly(value: Date | null): string;
    buildDateBand(startDate: string | undefined, finishDate: string | undefined): string[];
    collectWbsHolidayDates(model: ProjectModel): string[];
    collectProjectNonWorkingDayTypes(model: ProjectModel): Set<number>;
    buildDisplayDateBand(
      startDate: string | undefined,
      finishDate: string | undefined,
      baseDate: string | undefined,
      displayDaysBeforeBaseDate: number | undefined,
      displayDaysAfterBaseDate: number | undefined,
      holidaySet: Set<string>,
      nonWorkingDayTypes: Set<number>,
      useBusinessDaysForDisplayRange: boolean | undefined
    ): string[];
    isWeeklyNonWorkingDay(day: string, nonWorkingDayTypes: Set<number>): boolean;
  };

  type WbsXlsxSectionsHelper = {
    firstResourceName(resourceNames: string[] | undefined): string;
    formatCalendarLabel(calendarUID: string | undefined, calendarNameByUid: Map<string, string>): string;
    truncateWbsReference(value: string | undefined, maxLength: number): string;
    referenceCell(
      task: TaskModel,
      value: string | undefined,
      horizontalAlign?: "left" | "center" | "right"
    ): WbsXlsxCellLike;
    projectInfoRows(
      project: ProjectModel["project"],
      calendarNameByUid: Map<string, string>,
      holidayCount: number,
      columnCount: number,
      startColumnIndex: number,
      startRowNumber: number
    ): {
      mergedRanges: string[];
      rows: Array<{ height?: number; cells: WbsXlsxCellLike[] }>;
    };
    legendRows(
      columnCount: number,
      startRowNumber: number
    ): {
      mergedRanges: string[];
      rows: Array<{ height?: number; cells: WbsXlsxCellLike[] }>;
    };
    displaySummaryRows(
      displayDays: number,
      businessDays: number,
      baseDate: string | undefined,
      taskCount: number,
      resourceCount: number,
      assignmentCount: number,
      calendarCount: number,
      columnCount: number,
      startColumnIndex?: number,
      startRowNumber?: number,
      displayDaysBeforeBaseDate?: number,
      displayDaysAfterBaseDate?: number,
      useBusinessDaysForDisplayRange?: boolean,
      useBusinessDaysForProgressBand?: boolean
    ): {
      mergedRanges: string[];
      rows: Array<{ height?: number; cells: WbsXlsxCellLike[] }>;
    };
    formatWbsDate(value: string | undefined): string;
  };

  type WbsXlsxCellsHelper = {
    dividerCell(): WbsXlsxCellLike;
    weekBandRow(
      fixedColumnCount: number,
      weekBandRanges: Array<{ startIndex: number; label: string; hasMonthBoundary: boolean }>,
      dateBandLength: number
    ): { height: number; cells: WbsXlsxCellLike[] };
    headerRow(labels: Array<string | WbsXlsxCellLike>): { height: number; cells: WbsXlsxCellLike[] };
    weekdayRow(
      fixedColumnCount: number,
      dateBand: string[],
      currentDate: string | undefined,
      holidaySet: Set<string>,
      nonWorkingDayTypes: Set<number>
    ): { height: number; cells: WbsXlsxCellLike[] };
    dateBandHeaderRow(
      fixedColumnCount: number,
      dateBand: string[],
      currentDate: string | undefined,
      holidaySet: Set<string>,
      nonWorkingDayTypes: Set<number>
    ): { height: number; cells: WbsXlsxCellLike[] };
    weekdayHeaderRow(
      fixedHeaders: string[],
      dateBand: string[],
      currentDate: string | undefined,
      holidaySet: Set<string>,
      nonWorkingDayTypes: Set<number>
    ): { height: number; cells: WbsXlsxCellLike[] };
    taskCell(
      task: TaskModel,
      value: string | number | boolean | undefined,
      horizontalAlign?: "left" | "center" | "right"
    ): WbsXlsxCellLike;
    detailCell(task: TaskModel, value: string | undefined): WbsXlsxCellLike;
    taskRowHeight(task: TaskModel): number | undefined;
    kindCell(task: TaskModel): WbsXlsxCellLike;
    identifierCell(task: TaskModel, value: string | number | boolean | undefined): WbsXlsxCellLike;
    flagCell(task: TaskModel, enabled: boolean | undefined, marker: string): WbsXlsxCellLike;
    progressCell(task: TaskModel, value: number | undefined): WbsXlsxCellLike;
    formatDurationLabel(
      task: TaskModel,
      holidaySet: Set<string>,
      nonWorkingDayTypes: Set<number>,
      useBusinessDaysForProgressBand: boolean | undefined
    ): string;
    dateBandCell(
      task: TaskModel,
      day: string,
      currentDate: string | undefined,
      holidaySet: Set<string>,
      nonWorkingDayTypes: Set<number>,
      useBusinessDaysForProgressBand: boolean | undefined
    ): WbsXlsxCellLike;
    countBusinessDays(dateBand: string[], holidaySet: Set<string>, nonWorkingDayTypes: Set<number>): number;
    buildWeekBandRanges(
      dateBand: string[],
      startColumnIndex: number,
      rowNumber: number
    ): Array<{ range: string; startIndex: number; label: string; hasMonthBoundary: boolean }>;
  };

  type WbsXlsxBaseHelper = {
    emptyRow(columnCount: number, height?: number): { height?: number; cells: WbsXlsxCellLike[] };
    overlayRows(
      rows: Array<{ height?: number; cells: WbsXlsxCellLike[] }>,
      startIndex: number,
      blockRows: Array<{ height?: number; cells: WbsXlsxCellLike[] }>,
      columnCount: number
    ): void;
    sheetTitleRow(title: string, columnCount: number): { height?: number; cells: WbsXlsxCellLike[] };
    infoRow(text: string, columnCount: number): { height?: number; cells: WbsXlsxCellLike[] };
    formatWbsExportTimestamp(value: Date): string;
  };

  type WbsXlsxTaskmetaHelper = {
    formatTaskLabel(task: TaskModel): string;
    classifyTaskKind(task: TaskModel): string;
  };
  type WbsXlsxExportHelper = {
    exportWbsWorkbook(model: ProjectModel, options?: WbsExportOptions): WbsXlsxWorkbookLike;
  };
  type WbsXlsxFacade = {
    collectWbsHolidayDates(model: ProjectModel): string[];
    exportWbsWorkbook(model: ProjectModel, options?: WbsExportOptions): WbsXlsxWorkbookLike;
    layout: WbsSheetLayoutHelper;
  };

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
  const wbsXlsxLayout = (globalThis as typeof globalThis & {
    __mikuprojectWbsXlsxLayout?: {
      createWbsSheetLayoutHelper: () => WbsSheetLayoutHelper;
    };
  }).__mikuprojectWbsXlsxLayout;
  const wbsDateband = (globalThis as typeof globalThis & {
    __mikuprojectWbsDateband?: {
      createWbsDatebandHelper: () => WbsDatebandHelper;
    };
  }).__mikuprojectWbsDateband;
  const wbsXlsxSections = (globalThis as typeof globalThis & {
    __mikuprojectWbsXlsxSections?: {
      createWbsXlsxSectionsHelper: (config: {
        layout: WbsSheetLayoutHelper;
        fills: Record<string, string>;
      }) => WbsXlsxSectionsHelper;
    };
  }).__mikuprojectWbsXlsxSections;
  const wbsXlsxCells = (globalThis as typeof globalThis & {
    __mikuprojectWbsXlsxCells?: {
      createWbsXlsxCellsHelper: (config: {
        layout: WbsSheetLayoutHelper;
        fills: Record<string, string>;
        formatTaskLabel(task: TaskModel): string;
        classifyTaskKind(task: TaskModel): string;
        buildDateBand(startDate: string | undefined, finishDate: string | undefined): string[];
        parseDateOnly(value: string | undefined): Date | null;
        formatDateOnly(value: Date | null): string;
        isWeeklyNonWorkingDay(day: string, nonWorkingDayTypes: Set<number>): boolean;
      }) => WbsXlsxCellsHelper;
    };
  }).__mikuprojectWbsXlsxCells;
  const wbsXlsxBase = (globalThis as typeof globalThis & {
    __mikuprojectWbsXlsxBase?: {
      createWbsXlsxBaseHelper: () => WbsXlsxBaseHelper;
    };
  }).__mikuprojectWbsXlsxBase;
  const wbsXlsxTaskmeta = (globalThis as typeof globalThis & {
    __mikuprojectWbsXlsxTaskmeta?: {
      createWbsXlsxTaskmetaHelper: () => WbsXlsxTaskmetaHelper;
    };
  }).__mikuprojectWbsXlsxTaskmeta;
  const wbsXlsxExport = (globalThis as typeof globalThis & {
    __mikuprojectWbsXlsxExport?: {
      createWbsXlsxExportHelper: (config: {
        pxWidth(pixels: number): number;
        collectWbsHolidayDates(model: ProjectModel): string[];
        collectProjectNonWorkingDayTypes(model: ProjectModel): Set<number>;
        buildDisplayDateBand(
          startDate: string | undefined,
          finishDate: string | undefined,
          baseDate: string | undefined,
          displayDaysBeforeBaseDate: number | undefined,
          displayDaysAfterBaseDate: number | undefined,
          holidaySet: Set<string>,
          nonWorkingDayTypes: Set<number>,
          useBusinessDaysForDisplayRange: boolean | undefined
        ): string[];
        firstResourceName(resourceNames: string[] | undefined): string;
        formatCalendarLabel(calendarUID: string | undefined, calendarNameByUid: Map<string, string>): string;
        truncateWbsReference(value: string | undefined, maxLength: number): string;
        referenceCell(
          task: TaskModel,
          value: string | undefined,
          horizontalAlign?: "left" | "center" | "right"
        ): WbsXlsxCellLike;
        projectInfoRows(
          project: ProjectModel["project"],
          calendarNameByUid: Map<string, string>,
          holidayCount: number,
          columnCount: number,
          startColumnIndex: number,
          startRowNumber: number
        ): {
          mergedRanges: string[];
          rows: Array<{ height?: number; cells: WbsXlsxCellLike[] }>;
        };
        legendRows(
          columnCount: number,
          startRowNumber: number
        ): {
          mergedRanges: string[];
          rows: Array<{ height?: number; cells: WbsXlsxCellLike[] }>;
        };
        displaySummaryRows(
          displayDays: number,
          businessDays: number,
          baseDate: string | undefined,
          taskCount: number,
          resourceCount: number,
          assignmentCount: number,
          calendarCount: number,
          columnCount: number,
          startColumnIndex?: number,
          startRowNumber?: number,
          displayDaysBeforeBaseDate?: number,
          displayDaysAfterBaseDate?: number,
          useBusinessDaysForDisplayRange?: boolean,
          useBusinessDaysForProgressBand?: boolean
        ): {
          mergedRanges: string[];
          rows: Array<{ height?: number; cells: WbsXlsxCellLike[] }>;
        };
        formatWbsDate(value: string | undefined): string;
        dateBandHeaderRow(
          fixedColumnCount: number,
          dateBand: string[],
          currentDate: string | undefined,
          holidaySet: Set<string>,
          nonWorkingDayTypes: Set<number>
        ): { height: number; cells: WbsXlsxCellLike[] };
        weekdayHeaderRow(
          fixedHeaders: string[],
          dateBand: string[],
          currentDate: string | undefined,
          holidaySet: Set<string>,
          nonWorkingDayTypes: Set<number>
        ): { height: number; cells: WbsXlsxCellLike[] };
        dividerCell(): WbsXlsxCellLike;
        taskCell(
          task: TaskModel,
          value: string | number | boolean | undefined,
          horizontalAlign?: "left" | "center" | "right"
        ): WbsXlsxCellLike;
        detailCell(task: TaskModel, value: string | undefined): WbsXlsxCellLike;
        taskRowHeight(task: TaskModel): number | undefined;
        kindCell(task: TaskModel): WbsXlsxCellLike;
        identifierCell(task: TaskModel, value: string | number | boolean | undefined): WbsXlsxCellLike;
        flagCell(task: TaskModel, enabled: boolean | undefined, marker: string): WbsXlsxCellLike;
        progressCell(task: TaskModel, value: number | undefined): WbsXlsxCellLike;
        formatDurationLabel(
          task: TaskModel,
          holidaySet: Set<string>,
          nonWorkingDayTypes: Set<number>,
          useBusinessDaysForProgressBand: boolean | undefined
        ): string;
        dateBandCell(
          task: TaskModel,
          day: string,
          currentDate: string | undefined,
          holidaySet: Set<string>,
          nonWorkingDayTypes: Set<number>,
          useBusinessDaysForProgressBand: boolean | undefined
        ): WbsXlsxCellLike;
        countBusinessDays(dateBand: string[], holidaySet: Set<string>, nonWorkingDayTypes: Set<number>): number;
        emptyRow(columnCount: number, height?: number): { height?: number; cells: WbsXlsxCellLike[] };
        overlayRows(
          rows: Array<{ height?: number; cells: WbsXlsxCellLike[] }>,
          startIndex: number,
          blockRows: Array<{ height?: number; cells: WbsXlsxCellLike[] }>,
          columnCount: number
        ): void;
        formatWbsExportTimestamp(value: Date): string;
        formatTaskLabel(task: TaskModel): string;
      }) => WbsXlsxExportHelper;
    };
  }).__mikuprojectWbsXlsxExport;
  const wbsXlsxPublic = (globalThis as typeof globalThis & {
    __mikuprojectWbsXlsxPublic?: {
      createWbsXlsxFacade: (config: {
        collectWbsHolidayDates(model: ProjectModel): string[];
        exportWbsWorkbook(model: ProjectModel, options?: WbsExportOptions): WbsXlsxWorkbookLike;
        layout: WbsSheetLayoutHelper;
      }) => WbsXlsxFacade;
    };
  }).__mikuprojectWbsXlsxPublic;

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
  const pxWidth = (pixels: number) => Math.round((pixels / 7) * 100) / 100;
  const collectWbsHolidayDates = (model: ProjectModel): string[] => WBS_DATEBAND.collectWbsHolidayDates(model);
  const collectProjectNonWorkingDayTypes = (model: ProjectModel): Set<number> => WBS_DATEBAND.collectProjectNonWorkingDayTypes(model);
  const buildDateBand = (startDate: string | undefined, finishDate: string | undefined): string[] =>
    WBS_DATEBAND.buildDateBand(startDate, finishDate);
  const buildDisplayDateBand = (
    startDate: string | undefined,
    finishDate: string | undefined,
    baseDate: string | undefined,
    displayDaysBeforeBaseDate: number | undefined,
    displayDaysAfterBaseDate: number | undefined,
    holidaySet: Set<string>,
    nonWorkingDayTypes: Set<number>,
    useBusinessDaysForDisplayRange: boolean | undefined
  ): string[] => WBS_DATEBAND.buildDisplayDateBand(
    startDate,
    finishDate,
    baseDate,
    displayDaysBeforeBaseDate,
    displayDaysAfterBaseDate,
    holidaySet,
    nonWorkingDayTypes,
    useBusinessDaysForDisplayRange
  );
  const isWeeklyNonWorkingDay = (day: string, nonWorkingDayTypes: Set<number>): boolean =>
    WBS_DATEBAND.isWeeklyNonWorkingDay(day, nonWorkingDayTypes);
  const parseDateOnly = (value: string | undefined): Date | null => WBS_DATEBAND.parseDateOnly(value);
  const formatDateOnly = (value: Date | null): string => WBS_DATEBAND.formatDateOnly(value);
  const firstResourceName = (resourceNames: string[] | undefined): string => WBS_SECTIONS.firstResourceName(resourceNames);
  const formatCalendarLabel = (calendarUID: string | undefined, calendarNameByUid: Map<string, string>): string =>
    WBS_SECTIONS.formatCalendarLabel(calendarUID, calendarNameByUid);
  const truncateWbsReference = (value: string | undefined, maxLength: number): string =>
    WBS_SECTIONS.truncateWbsReference(value, maxLength);
  const referenceCell = (
    task: TaskModel,
    value: string | undefined,
    horizontalAlign: "left" | "center" | "right" = "center"
  ): WbsXlsxCellLike => WBS_SECTIONS.referenceCell(task, value, horizontalAlign);
  const projectInfoRows = (
    project: ProjectModel["project"],
    calendarNameByUid: Map<string, string>,
    holidayCount: number,
    columnCount: number,
    startColumnIndex: number,
    startRowNumber: number
  ) => WBS_SECTIONS.projectInfoRows(project, calendarNameByUid, holidayCount, columnCount, startColumnIndex, startRowNumber);
  const legendRows = (columnCount: number, startRowNumber: number) => WBS_SECTIONS.legendRows(columnCount, startRowNumber);
  const displaySummaryRows = (
    displayDays: number,
    businessDays: number,
    baseDate: string | undefined,
    taskCount: number,
    resourceCount: number,
    assignmentCount: number,
    calendarCount: number,
    columnCount: number,
    startColumnIndex = 5,
    startRowNumber = 5,
    displayDaysBeforeBaseDate?: number,
    displayDaysAfterBaseDate?: number,
    useBusinessDaysForDisplayRange?: boolean,
    useBusinessDaysForProgressBand?: boolean
  ) => WBS_SECTIONS.displaySummaryRows(
    displayDays,
    businessDays,
    baseDate,
    taskCount,
    resourceCount,
    assignmentCount,
    calendarCount,
    columnCount,
    startColumnIndex,
    startRowNumber,
    displayDaysBeforeBaseDate,
    displayDaysAfterBaseDate,
    useBusinessDaysForDisplayRange,
    useBusinessDaysForProgressBand
  );
  const formatWbsDate = (value: string | undefined): string => WBS_SECTIONS.formatWbsDate(value);
  const formatTaskLabel = (task: TaskModel): string => WBS_TASKMETA.formatTaskLabel(task);
  const classifyTaskKind = (task: TaskModel): string => WBS_TASKMETA.classifyTaskKind(task);
  const weekBandRow = (
    fixedColumnCount: number,
    weekBandRanges: Array<{ startIndex: number; label: string; hasMonthBoundary: boolean }>,
    dateBandLength: number
  ) => WBS_CELLS.weekBandRow(fixedColumnCount, weekBandRanges, dateBandLength);
  const headerRow = (labels: Array<string | WbsXlsxCellLike>) => WBS_CELLS.headerRow(labels);
  const weekdayRow = (
    fixedColumnCount: number,
    dateBand: string[],
    currentDate: string | undefined,
    holidaySet: Set<string>,
    nonWorkingDayTypes: Set<number>
  ) => WBS_CELLS.weekdayRow(fixedColumnCount, dateBand, currentDate, holidaySet, nonWorkingDayTypes);
  const dateBandHeaderRow = (
    fixedColumnCount: number,
    dateBand: string[],
    currentDate: string | undefined,
    holidaySet: Set<string>,
    nonWorkingDayTypes: Set<number>
  ) => WBS_CELLS.dateBandHeaderRow(fixedColumnCount, dateBand, currentDate, holidaySet, nonWorkingDayTypes);
  const weekdayHeaderRow = (
    fixedHeaders: string[],
    dateBand: string[],
    currentDate: string | undefined,
    holidaySet: Set<string>,
    nonWorkingDayTypes: Set<number>
  ) => WBS_CELLS.weekdayHeaderRow(fixedHeaders, dateBand, currentDate, holidaySet, nonWorkingDayTypes);
  const dividerCell = () => WBS_CELLS.dividerCell();
  const taskCell = (
    task: TaskModel,
    value: string | number | boolean | undefined,
    horizontalAlign: "left" | "center" | "right" = "left"
  ) => WBS_CELLS.taskCell(task, value, horizontalAlign);
  const detailCell = (task: TaskModel, value: string | undefined) => WBS_CELLS.detailCell(task, value);
  const taskRowHeight = (task: TaskModel) => WBS_CELLS.taskRowHeight(task);
  const kindCell = (task: TaskModel) => WBS_CELLS.kindCell(task);
  const identifierCell = (task: TaskModel, value: string | number | boolean | undefined) => WBS_CELLS.identifierCell(task, value);
  const flagCell = (task: TaskModel, enabled: boolean | undefined, marker: string) => WBS_CELLS.flagCell(task, enabled, marker);
  const progressCell = (task: TaskModel, value: number | undefined) => WBS_CELLS.progressCell(task, value);
  const formatDurationLabel = (
    task: TaskModel,
    holidaySet: Set<string>,
    nonWorkingDayTypes: Set<number>,
    useBusinessDaysForProgressBand: boolean | undefined
  ) => WBS_CELLS.formatDurationLabel(task, holidaySet, nonWorkingDayTypes, useBusinessDaysForProgressBand);
  const dateBandCell = (
    task: TaskModel,
    day: string,
    currentDate: string | undefined,
    holidaySet: Set<string>,
    nonWorkingDayTypes: Set<number>,
    useBusinessDaysForProgressBand: boolean | undefined
  ) => WBS_CELLS.dateBandCell(task, day, currentDate, holidaySet, nonWorkingDayTypes, useBusinessDaysForProgressBand);
  const countBusinessDays = (dateBand: string[], holidaySet: Set<string>, nonWorkingDayTypes: Set<number>) =>
    WBS_CELLS.countBusinessDays(dateBand, holidaySet, nonWorkingDayTypes);
  const buildWeekBandRanges = (dateBand: string[], startColumnIndex: number, rowNumber: number) =>
    WBS_CELLS.buildWeekBandRanges(dateBand, startColumnIndex, rowNumber);
  const emptyRow = (columnCount: number, height = 22) => WBS_BASE.emptyRow(columnCount, height);
  const overlayRows = (
    rows: Array<{ height?: number; cells: WbsXlsxCellLike[] }>,
    startIndex: number,
    blockRows: Array<{ height?: number; cells: WbsXlsxCellLike[] }>,
    columnCount: number
  ) => WBS_BASE.overlayRows(rows, startIndex, blockRows, columnCount);
  const sheetTitleRow = (title: string, columnCount: number) => WBS_BASE.sheetTitleRow(title, columnCount);
  const infoRow = (text: string, columnCount: number) => WBS_BASE.infoRow(text, columnCount);
  const formatWbsExportTimestamp = (value: Date) => WBS_BASE.formatWbsExportTimestamp(value);
  const WBS_EXPORT = wbsXlsxExport.createWbsXlsxExportHelper({
    pxWidth,
    collectWbsHolidayDates: (model) => collectWbsHolidayDates(model),
    collectProjectNonWorkingDayTypes: (model) => collectProjectNonWorkingDayTypes(model),
    buildDisplayDateBand: (
      startDate,
      finishDate,
      baseDate,
      displayDaysBeforeBaseDate,
      displayDaysAfterBaseDate,
      holidaySet,
      nonWorkingDayTypes,
      useBusinessDaysForDisplayRange
    ) => buildDisplayDateBand(
      startDate,
      finishDate,
      baseDate,
      displayDaysBeforeBaseDate,
      displayDaysAfterBaseDate,
      holidaySet,
      nonWorkingDayTypes,
      useBusinessDaysForDisplayRange
    ),
    firstResourceName: (resourceNames) => firstResourceName(resourceNames),
    formatCalendarLabel: (calendarUID, calendarNameByUid) => formatCalendarLabel(calendarUID, calendarNameByUid),
    truncateWbsReference: (value, maxLength) => truncateWbsReference(value, maxLength),
    referenceCell: (task, value, horizontalAlign) => referenceCell(task, value, horizontalAlign),
    projectInfoRows: (project, calendarNameByUid, holidayCount, columnCount, startColumnIndex, startRowNumber) =>
      projectInfoRows(project, calendarNameByUid, holidayCount, columnCount, startColumnIndex, startRowNumber),
    legendRows: (columnCount, startRowNumber) => legendRows(columnCount, startRowNumber),
    displaySummaryRows: (
      displayDays,
      businessDays,
      baseDate,
      taskCount,
      resourceCount,
      assignmentCount,
      calendarCount,
      columnCount,
      startColumnIndex,
      startRowNumber,
      displayDaysBeforeBaseDate,
      displayDaysAfterBaseDate,
      useBusinessDaysForDisplayRange,
      useBusinessDaysForProgressBand
    ) => displaySummaryRows(
      displayDays,
      businessDays,
      baseDate,
      taskCount,
      resourceCount,
      assignmentCount,
      calendarCount,
      columnCount,
      startColumnIndex,
      startRowNumber,
      displayDaysBeforeBaseDate,
      displayDaysAfterBaseDate,
      useBusinessDaysForDisplayRange,
      useBusinessDaysForProgressBand
    ),
    formatWbsDate: (value) => formatWbsDate(value),
    dateBandHeaderRow: (fixedColumnCount, dateBand, currentDate, holidaySet, nonWorkingDayTypes) =>
      dateBandHeaderRow(fixedColumnCount, dateBand, currentDate, holidaySet, nonWorkingDayTypes),
    weekdayHeaderRow: (fixedHeaders, dateBand, currentDate, holidaySet, nonWorkingDayTypes) =>
      weekdayHeaderRow(fixedHeaders, dateBand, currentDate, holidaySet, nonWorkingDayTypes),
    dividerCell: () => dividerCell(),
    taskCell: (task, value, horizontalAlign) => taskCell(task, value, horizontalAlign),
    detailCell: (task, value) => detailCell(task, value),
    taskRowHeight: (task) => taskRowHeight(task),
    kindCell: (task) => kindCell(task),
    identifierCell: (task, value) => identifierCell(task, value),
    flagCell: (task, enabled, marker) => flagCell(task, enabled, marker),
    progressCell: (task, value) => progressCell(task, value),
    formatDurationLabel: (task, holidaySet, nonWorkingDayTypes, useBusinessDaysForProgressBand) =>
      formatDurationLabel(task, holidaySet, nonWorkingDayTypes, useBusinessDaysForProgressBand),
    dateBandCell: (task, day, currentDate, holidaySet, nonWorkingDayTypes, useBusinessDaysForProgressBand) =>
      dateBandCell(task, day, currentDate, holidaySet, nonWorkingDayTypes, useBusinessDaysForProgressBand),
    countBusinessDays: (dateBand, holidaySet, nonWorkingDayTypes) => countBusinessDays(dateBand, holidaySet, nonWorkingDayTypes),
    emptyRow: (columnCount, height) => emptyRow(columnCount, height),
    overlayRows: (rows, startIndex, blockRows, columnCount) => overlayRows(rows, startIndex, blockRows, columnCount),
    formatWbsExportTimestamp: (value) => formatWbsExportTimestamp(value),
    formatTaskLabel: (task) => formatTaskLabel(task)
  });

  function exportWbsWorkbook(model: ProjectModel, options: WbsExportOptions = {}): WbsXlsxWorkbookLike {
    return WBS_EXPORT.exportWbsWorkbook(model, options);
  }

  const WBS_XLSX_FACADE = wbsXlsxPublic.createWbsXlsxFacade({
    collectWbsHolidayDates: (model) => collectWbsHolidayDates(model),
    exportWbsWorkbook: (model, options) => exportWbsWorkbook(model, options),
    layout: WBS_LAYOUT
  });


  (globalThis as typeof globalThis & {
    __mikuprojectWbsXlsx?: {
      collectWbsHolidayDates: typeof collectWbsHolidayDates;
      exportWbsWorkbook: typeof exportWbsWorkbook;
      layout: WbsSheetLayoutHelper;
    };
  }).__mikuprojectWbsXlsx = WBS_XLSX_FACADE;
})();
