/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
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

  type NativeSvgExportOptions = {
    holidayDates?: string[];
    displayDaysBeforeBaseDate?: number;
    displayDaysAfterBaseDate?: number;
    useBusinessDaysForDisplayRange?: boolean;
    useBusinessDaysForProgressBand?: boolean;
    labelMode?: "near" | "list";
  };

  type NativeSvgTaskRow = {
    task: TaskModel;
    label: string;
    kind: "phase" | "task" | "milestone";
    startIndex: number | null;
    endIndex: number | null;
    y: number;
  };

  type NativeSvgLabelPlacement = {
    x: number;
    anchor: "start" | "end";
    width: number;
  };

  type NativeSvgLabelPlacementDecision = NativeSvgLabelPlacement & {
    reason: string;
    metrics: Record<string, number | string | boolean | null>;
  };

  type NativeSvgPlacementGeometry = {
    chartOriginX: number;
    shapeStartX: number;
    shapeEndX: number;
    shapeCenterX: number;
    chartMidX: number;
    chartEndX: number;
    leftRoom: number;
    rightRoom: number;
  };

  type NativeSvgShapeGeometry = {
    uid: string;
    kind: "phase" | "task" | "milestone";
    startX: number;
    endX: number;
    midY: number;
  };

  type NativeSvgPlacementConfig = {
    cellWidth: number;
    gap: number;
    rangeInset: number;
    milestoneHalfWidth: number;
    preferredColumns: number;
    missingRangeX: number;
    missingRangeReason: string;
  };

  type NativeSvgLabelsHelper = {
    formatTaskLabel(task: TaskModel, labelMode: "near" | "list"): string;
    resolveLabelPlacement(
      row: NativeSvgTaskRow,
      chartOriginX: number,
      chartWidth: number,
      labelMode: "near" | "list"
    ): NativeSvgLabelPlacement;
    resolveWeeklyLabelPlacement(
      row: NativeSvgTaskRow,
      chartOriginX: number,
      chartWidth: number
    ): NativeSvgLabelPlacement;
    shouldUseDailyOnBarLabel(
      row: NativeSvgTaskRow,
      chartOriginX: number,
      chartWidth: number,
      bandLength: number,
      labelMode: "near" | "list"
    ): boolean;
    buildDependencyConnectorElements(
      tasks: TaskModel[],
      rows: NativeSvgTaskRow[],
      chartOriginX: number,
      chartOriginY: number,
      config: {
        cellWidth: number;
        rangeInset: number;
        milestoneHalfWidth: number;
        routeOffset: number;
        routeSpacing: number;
        targetInset: number;
        cornerRadius: number;
      }
    ): string[];
  };

  type ZipEntry = {
    name: string;
    data: Uint8Array;
  };

  type MonthlyCalendarSvgArchive = {
    entries: Array<{
      fileName: string;
      svg: string;
    }>;
    zipBytes: Uint8Array;
  };

  type WbsSvgCalendarHelper = {
    exportMonthlyWbsCalendarSvgArchive(model: ProjectModel): MonthlyCalendarSvgArchive;
    exportMonthlyWbsCalendarSvg(
      model: ProjectModel,
      monthKey: string,
      holidaySet: Set<string>,
      nonWorkingDayTypes: Set<number>
    ): string;
  };

  type WbsSvgTimelineHelper = {
    buildTaskRows(tasks: TaskModel[], dateBand: string[]): NativeSvgTaskRow[];
    buildWeeklyTaskRows(tasks: TaskModel[], weeklyBand: WeeklyBand[]): NativeSvgTaskRow[];
    formatSvgAxisDate(day: string): string;
    formatWeeklyAxisLabel(day: string): string;
    indexOfDate(dateBand: string[], value: string | undefined): number | null;
    indexOfWeekForDate(weeklyBand: WeeklyBand[], value: string | undefined): number | null;
    buildWeeklyBand(startDate: string | undefined, finishDate: string | undefined): WeeklyBand[];
    buildWeeklyMonthSpans(weeklyBand: WeeklyBand[]): Array<{ monthKey: string; startIndex: number; endIndex: number }>;
  };

  type WeeklyBand = {
    startDay: string;
    endDay: string;
    monthKey: string;
  };

  type WbsSvgBarsHelper = {
    renderTaskBar(
      row: NativeSvgTaskRow,
      chartOriginX: number,
      rowY: number,
      cellWidth: number,
      offsets: {
        horizontalInset: number;
        verticalInset: number;
        milestoneHalfWidth: number;
      }
    ): string[];
  };

  type WbsSvgViewportHelper = {
    computeDailyViewport(config: {
      labelPlacements: NativeSvgLabelPlacement[];
      chartOriginXBase: number;
      chartWidth: number;
      rightLabelWidth: number;
      rightPadding: number;
      labelMode: "near" | "list";
    }): {
      shiftX: number;
      chartOriginX: number;
      svgWidth: number;
    };
    computeWeeklyViewport(config: {
      labelPlacements: NativeSvgLabelPlacement[];
      chartOriginXBase: number;
      chartWidth: number;
      trimPadding: number;
      rightPadding: number;
    }): {
      shiftX: number;
      chartOriginX: number;
      svgWidth: number;
    };
  };

  type WbsSvgAxisHelper = {
    renderDailyAxis(config: {
      dateBand: string[];
      chartOriginX: number;
      svgHeight: number;
      topPadding: number;
      bottomPadding: number;
      dayWidth: number;
      holidaySet: Set<string>;
      nonWorkingDayTypes: Set<number>;
      isWeeklyNonWorkingDay(day: string, nonWorkingDayTypes: Set<number>): boolean;
      todayIndex: number | null;
      escapeXml(value: string): string;
      formatSvgAxisDate(day: string): string;
    }): string[];
    renderWeeklyAxis(config: {
      weeklyBand: WeeklyBand[];
      monthSpans: Array<{ monthKey: string; startIndex: number; endIndex: number }>;
      chartOriginX: number;
      svgHeight: number;
      topPadding: number;
      bottomPadding: number;
      weekWidth: number;
      todayWeekIndex: number | null;
      escapeXml(value: string): string;
      formatWeeklyAxisLabel(day: string): string;
    }): string[];
  };

  type WbsSvgScaffoldHelper = {
    createDailyScaffold(config: {
      svgWidth: number;
      svgHeight: number;
      chartOriginX: number;
      chartWidth: number;
      topPadding: number;
      projectName: string | undefined;
      escapeXml(value: string): string;
    }): string[];
    createWeeklyScaffold(config: {
      svgWidth: number;
      svgHeight: number;
      chartOriginX: number;
      chartWidth: number;
      topPadding: number;
      projectName: string | undefined;
      projectStartDate: string | undefined;
      projectFinishDate: string | undefined;
      escapeXml(value: string): string;
    }): string[];
  };

  type WbsSvgRenderHelper = {
    exportNativeSvg(model: ProjectModel, options?: NativeSvgExportOptions): string;
    exportWeeklyNativeSvg(model: ProjectModel, options?: NativeSvgExportOptions): string;
  };
  type NativeSvgFacade = {
    exportNativeSvg(model: ProjectModel, options?: NativeSvgExportOptions): string;
    exportWeeklyNativeSvg(model: ProjectModel, options?: NativeSvgExportOptions): string;
    exportMonthlyWbsCalendarSvgArchive(model: ProjectModel): MonthlyCalendarSvgArchive;
    collectWbsHolidayDates(model: ProjectModel): string[];
  };

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
  const wbsDateband = (globalThis as typeof globalThis & {
    __mikuprojectWbsDateband?: {
      createWbsDatebandHelper: () => WbsDatebandHelper;
    };
  }).__mikuprojectWbsDateband;
  const wbsSvgLabels = (globalThis as typeof globalThis & {
    __mikuprojectWbsSvgLabels?: {
      createWbsSvgLabelsHelper: (config: {
        leftPadding: number;
        dayWidth: number;
        weekWidth: number;
        rowHeight: number;
      }) => NativeSvgLabelsHelper;
    };
  }).__mikuprojectWbsSvgLabels;
  const wbsSvgCalendar = (globalThis as typeof globalThis & {
    __mikuprojectWbsSvgCalendar?: {
      createWbsSvgCalendarHelper: (config: {
        monthlyCellWidth: number;
        monthlyCellHeight: number;
        monthlyHeaderHeight: number;
        monthlyWeekdayHeight: number;
        monthlyLeftPadding: number;
        monthlyTopPadding: number;
        monthlyRightPadding: number;
        monthlyBottomPadding: number;
        monthlyMaxItemsPerDay: number;
        monthlyMaxLabelChars: number;
        collectWbsHolidayDates(model: ProjectModel): string[];
        collectProjectNonWorkingDayTypes(model: ProjectModel): Set<number>;
        parseDateOnly(value: string | undefined): Date | null;
        formatDateOnly(value: Date | null): string;
        buildDateBand(startDate: string | undefined, finishDate: string | undefined): string[];
        isWeeklyNonWorkingDay(day: string, nonWorkingDayTypes: Set<number>): boolean;
        escapeXml(value: string): string;
        encodeUtf8(value: string): Uint8Array;
        packZip(entries: ZipEntry[]): Uint8Array;
      }) => WbsSvgCalendarHelper;
    };
  }).__mikuprojectWbsSvgCalendar;
  const wbsSvgTimeline = (globalThis as typeof globalThis & {
    __mikuprojectWbsSvgTimeline?: {
      createWbsSvgTimelineHelper: (config: {
        rowHeight: number;
        parseDateOnly(value: string | undefined): Date | null;
        formatDateOnly(value: Date | null): string;
      }) => WbsSvgTimelineHelper;
    };
  }).__mikuprojectWbsSvgTimeline;
  const wbsSvgZip = (globalThis as typeof globalThis & {
    __mikuprojectWbsSvgZip?: {
      createWbsSvgZipHelper: (config: {
        fixedModTime: number;
        fixedModDate: number;
      }) => {
        escapeXml(value: string): string;
        encodeUtf8(value: string): Uint8Array;
        packZip(entries: ZipEntry[]): Uint8Array;
      };
    };
  }).__mikuprojectWbsSvgZip;
  const wbsSvgBars = (globalThis as typeof globalThis & {
    __mikuprojectWbsSvgBars?: {
      createWbsSvgBarsHelper: () => WbsSvgBarsHelper;
    };
  }).__mikuprojectWbsSvgBars;
  const wbsSvgViewport = (globalThis as typeof globalThis & {
    __mikuprojectWbsSvgViewport?: {
      createWbsSvgViewportHelper: () => WbsSvgViewportHelper;
    };
  }).__mikuprojectWbsSvgViewport;
  const wbsSvgAxis = (globalThis as typeof globalThis & {
    __mikuprojectWbsSvgAxis?: {
      createWbsSvgAxisHelper: () => WbsSvgAxisHelper;
    };
  }).__mikuprojectWbsSvgAxis;
  const wbsSvgScaffold = (globalThis as typeof globalThis & {
    __mikuprojectWbsSvgScaffold?: {
      createWbsSvgScaffoldHelper: () => WbsSvgScaffoldHelper;
    };
  }).__mikuprojectWbsSvgScaffold;
  const wbsSvgRender = (globalThis as typeof globalThis & {
    __mikuprojectWbsSvgRender?: {
      createWbsSvgRenderHelper: (config: {
        constants: {
          dayWidth: number;
          weekWidth: number;
          listLabelWidth: number;
          nearLeftLabelWidth: number;
          nearRightLabelWidth: number;
          topPadding: number;
          headerHeight: number;
          bottomPadding: number;
          leftPadding: number;
          rightPadding: number;
          rowHeight: number;
          weeklyTopPadding: number;
          weeklyHeaderHeight: number;
          weeklyBottomPadding: number;
          weeklyLeftPadding: number;
          weeklyRightPadding: number;
          weeklyTrimPadding: number;
        };
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
        buildTaskRows(tasks: TaskModel[], dateBand: string[]): NativeSvgTaskRow[];
        buildWeeklyTaskRows(tasks: TaskModel[], weeklyBand: WeeklyBand[]): NativeSvgTaskRow[];
        buildWeeklyBand(startDate: string | undefined, finishDate: string | undefined): WeeklyBand[];
        buildWeeklyMonthSpans(weeklyBand: WeeklyBand[]): Array<{ monthKey: string; startIndex: number; endIndex: number }>;
        formatTaskLabel(task: TaskModel, labelMode: "near" | "list"): string;
        resolveLabelPlacement(
          row: NativeSvgTaskRow,
          chartOriginX: number,
          chartWidth: number,
          labelMode: "near" | "list"
        ): NativeSvgLabelPlacement;
        resolveWeeklyLabelPlacement(
          row: NativeSvgTaskRow,
          chartOriginX: number,
          chartWidth: number
        ): NativeSvgLabelPlacement;
        shouldUseDailyOnBarLabel(
          row: NativeSvgTaskRow,
          chartOriginX: number,
          chartWidth: number,
          bandLength: number,
          labelMode: "near" | "list"
        ): boolean;
        indexOfDate(dateBand: string[], value: string | undefined): number | null;
        indexOfWeekForDate(weeklyBand: WeeklyBand[], value: string | undefined): number | null;
        formatSvgAxisDate(day: string): string;
        formatWeeklyAxisLabel(day: string): string;
        isWeeklyNonWorkingDay(day: string, nonWorkingDayTypes: Set<number>): boolean;
        buildDependencyConnectorElements(
          tasks: TaskModel[],
          rows: NativeSvgTaskRow[],
          chartOriginX: number,
          chartOriginY: number,
          config: {
            cellWidth: number;
            rangeInset: number;
            milestoneHalfWidth: number;
            routeOffset: number;
            routeSpacing: number;
            targetInset: number;
            cornerRadius: number;
          }
        ): string[];
        escapeXml(value: string): string;
        bars: WbsSvgBarsHelper;
        viewport: WbsSvgViewportHelper;
        axis: WbsSvgAxisHelper;
        scaffold: WbsSvgScaffoldHelper;
      }) => WbsSvgRenderHelper;
    };
  }).__mikuprojectWbsSvgRender;
  const wbsSvgPublic = (globalThis as typeof globalThis & {
    __mikuprojectWbsSvgPublic?: {
      createNativeSvgFacade: (config: {
        collectWbsHolidayDates(model: ProjectModel): string[];
        exportNativeSvg(model: ProjectModel, options?: NativeSvgExportOptions): string;
        exportWeeklyNativeSvg(model: ProjectModel, options?: NativeSvgExportOptions): string;
        exportMonthlyWbsCalendarSvgArchive(model: ProjectModel): MonthlyCalendarSvgArchive;
      }) => NativeSvgFacade;
    };
  }).__mikuprojectWbsSvgPublic;

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
    buildTaskRows: (tasks, dateBand) => buildTaskRows(tasks, dateBand),
    buildWeeklyTaskRows: (tasks, weeklyBand) => buildWeeklyTaskRows(tasks, weeklyBand),
    buildWeeklyBand: (startDate, finishDate) => buildWeeklyBand(startDate, finishDate),
    buildWeeklyMonthSpans: (weeklyBand) => buildWeeklyMonthSpans(weeklyBand),
    formatTaskLabel: (task, labelMode) => formatTaskLabel(task, labelMode),
    resolveLabelPlacement: (row, chartOriginX, chartWidth, labelMode) =>
      resolveLabelPlacement(row, chartOriginX, chartWidth, labelMode),
    resolveWeeklyLabelPlacement: (row, chartOriginX, chartWidth) =>
      resolveWeeklyLabelPlacement(row, chartOriginX, chartWidth),
    shouldUseDailyOnBarLabel: (row, chartOriginX, chartWidth, bandLength, labelMode) =>
      shouldUseDailyOnBarLabel(row, chartOriginX, chartWidth, bandLength, labelMode),
    indexOfDate: (dateBand, value) => indexOfDate(dateBand, value),
    indexOfWeekForDate: (weeklyBand, value) => indexOfWeekForDate(weeklyBand, value),
    formatSvgAxisDate: (day) => formatSvgAxisDate(day),
    formatWeeklyAxisLabel: (day) => formatWeeklyAxisLabel(day),
    isWeeklyNonWorkingDay: (day, nonWorkingDayTypes) => isWeeklyNonWorkingDay(day, nonWorkingDayTypes),
    buildDependencyConnectorElements: (tasks, rows, chartOriginX, chartOriginY, config) =>
      buildDependencyConnectorElements(tasks, rows, chartOriginX, chartOriginY, config),
    escapeXml: (value) => escapeXml(value),
    bars: WBS_SVG_BARS,
    viewport: WBS_SVG_VIEWPORT,
    axis: WBS_SVG_AXIS,
    scaffold: WBS_SVG_SCAFFOLD
  });
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
  const formatTaskLabel = (task: TaskModel, labelMode: "near" | "list"): string =>
    WBS_SVG_LABELS.formatTaskLabel(task, labelMode);
  const resolveLabelPlacement = (
    row: NativeSvgTaskRow,
    chartOriginX: number,
    chartWidth: number,
    labelMode: "near" | "list"
  ): NativeSvgLabelPlacement => WBS_SVG_LABELS.resolveLabelPlacement(row, chartOriginX, chartWidth, labelMode);
  const resolveWeeklyLabelPlacement = (
    row: NativeSvgTaskRow,
    chartOriginX: number,
    chartWidth: number
  ): NativeSvgLabelPlacement => WBS_SVG_LABELS.resolveWeeklyLabelPlacement(row, chartOriginX, chartWidth);
  const shouldUseDailyOnBarLabel = (
    row: NativeSvgTaskRow,
    chartOriginX: number,
    chartWidth: number,
    bandLength: number,
    labelMode: "near" | "list"
  ): boolean => WBS_SVG_LABELS.shouldUseDailyOnBarLabel(row, chartOriginX, chartWidth, bandLength, labelMode);
  const buildDependencyConnectorElements = (
    tasks: TaskModel[],
    rows: NativeSvgTaskRow[],
    chartOriginX: number,
    chartOriginY: number,
    config: {
      cellWidth: number;
      rangeInset: number;
      milestoneHalfWidth: number;
      routeOffset: number;
      routeSpacing: number;
      targetInset: number;
      cornerRadius: number;
    }
  ): string[] => WBS_SVG_LABELS.buildDependencyConnectorElements(tasks, rows, chartOriginX, chartOriginY, config);
  const exportMonthlyWbsCalendarSvgArchive = (model: ProjectModel): MonthlyCalendarSvgArchive =>
    WBS_SVG_CALENDAR.exportMonthlyWbsCalendarSvgArchive(model);
  const exportMonthlyWbsCalendarSvg = (
    model: ProjectModel,
    monthKey: string,
    holidaySet: Set<string>,
    nonWorkingDayTypes: Set<number>
  ): string => WBS_SVG_CALENDAR.exportMonthlyWbsCalendarSvg(model, monthKey, holidaySet, nonWorkingDayTypes);
  const buildTaskRows = (tasks: TaskModel[], dateBand: string[]): NativeSvgTaskRow[] =>
    WBS_SVG_TIMELINE.buildTaskRows(tasks, dateBand);
  const buildWeeklyTaskRows = (tasks: TaskModel[], weeklyBand: WeeklyBand[]): NativeSvgTaskRow[] =>
    WBS_SVG_TIMELINE.buildWeeklyTaskRows(tasks, weeklyBand);
  const formatSvgAxisDate = (day: string): string => WBS_SVG_TIMELINE.formatSvgAxisDate(day);
  const formatWeeklyAxisLabel = (day: string): string => WBS_SVG_TIMELINE.formatWeeklyAxisLabel(day);
  const indexOfDate = (dateBand: string[], value: string | undefined): number | null =>
    WBS_SVG_TIMELINE.indexOfDate(dateBand, value);
  const indexOfWeekForDate = (weeklyBand: WeeklyBand[], value: string | undefined): number | null =>
    WBS_SVG_TIMELINE.indexOfWeekForDate(weeklyBand, value);
  const buildWeeklyBand = (startDate: string | undefined, finishDate: string | undefined): WeeklyBand[] =>
    WBS_SVG_TIMELINE.buildWeeklyBand(startDate, finishDate);
  const buildWeeklyMonthSpans = (weeklyBand: WeeklyBand[]): Array<{ monthKey: string; startIndex: number; endIndex: number }> =>
    WBS_SVG_TIMELINE.buildWeeklyMonthSpans(weeklyBand);
  const escapeXml = (value: string): string => WBS_SVG_ZIP.escapeXml(value);
  const encodeUtf8 = (value: string): Uint8Array => WBS_SVG_ZIP.encodeUtf8(value);
  const packZip = (entries: ZipEntry[]): Uint8Array => WBS_SVG_ZIP.packZip(entries);

  function exportNativeSvg(model: ProjectModel, options: NativeSvgExportOptions = {}): string {
    return WBS_SVG_RENDER.exportNativeSvg(model, options);
  }

  function exportWeeklyNativeSvg(model: ProjectModel, options: NativeSvgExportOptions = {}): string {
    return WBS_SVG_RENDER.exportWeeklyNativeSvg(model, options);
  }

  const NATIVE_SVG_FACADE = wbsSvgPublic.createNativeSvgFacade({
    collectWbsHolidayDates: (model) => collectWbsHolidayDates(model),
    exportNativeSvg: (model, options) => exportNativeSvg(model, options),
    exportWeeklyNativeSvg: (model, options) => exportWeeklyNativeSvg(model, options),
    exportMonthlyWbsCalendarSvgArchive: (model) => exportMonthlyWbsCalendarSvgArchive(model)
  });

  (globalThis as typeof globalThis & {
    __mikuprojectNativeSvg?: {
      collectWbsHolidayDates(model: ProjectModel): string[];
      exportNativeSvg: typeof exportNativeSvg;
      exportWeeklyNativeSvg: typeof exportWeeklyNativeSvg;
      exportMonthlyWbsCalendarSvgArchive: typeof exportMonthlyWbsCalendarSvgArchive;
    };
  }).__mikuprojectNativeSvg = NATIVE_SVG_FACADE;
})();
