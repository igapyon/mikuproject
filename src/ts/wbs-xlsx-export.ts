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

  type CreateWbsXlsxExportHelperConfig = {
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
  };

  type WbsXlsxExportHelper = {
    exportWbsWorkbook(model: ProjectModel, options?: WbsExportOptions): WbsXlsxWorkbookLike;
  };

  function createWbsXlsxExportHelper(config: CreateWbsXlsxExportHelperConfig): WbsXlsxExportHelper {
    const {
      pxWidth,
      collectWbsHolidayDates,
      collectProjectNonWorkingDayTypes,
      buildDisplayDateBand,
      firstResourceName,
      formatCalendarLabel,
      truncateWbsReference,
      referenceCell,
      projectInfoRows,
      legendRows,
      displaySummaryRows,
      formatWbsDate,
      dateBandHeaderRow,
      weekdayHeaderRow,
      dividerCell,
      taskCell,
      detailCell,
      taskRowHeight,
      kindCell,
      identifierCell,
      flagCell,
      progressCell,
      formatDurationLabel,
      dateBandCell,
      countBusinessDays,
      emptyRow,
      overlayRows,
      formatWbsExportTimestamp,
      formatTaskLabel
    } = config;

    function exportWbsWorkbook(model: ProjectModel, options: WbsExportOptions = {}): WbsXlsxWorkbookLike {
      const nonWorkingDayTypes = collectProjectNonWorkingDayTypes(model);
      const resourceNameByUid = new Map(model.resources.map((resource) => [resource.uid, resource.name]));
      const predecessorNameByUid = new Map(model.tasks.map((task) => [task.uid, task.name]));
      const calendarNameByUid = new Map(model.calendars.map((calendar) => [calendar.uid, calendar.name]));
      const resourceNamesByTaskUid = new Map<string, string[]>();
      const holidaySet = new Set([
        ...collectWbsHolidayDates(model),
        ...(options.holidayDates || []).map((day) => day.slice(0, 10))
      ]);

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

      const dateBand = buildDisplayDateBand(
        model.project.startDate,
        model.project.finishDate,
        model.project.currentDate,
        options.displayDaysBeforeBaseDate,
        options.displayDaysAfterBaseDate,
        holidaySet,
        nonWorkingDayTypes,
        options.useBusinessDaysForDisplayRange
      );
      const fixedHeaders = [
        "UID",
        "ID",
        "WBS",
        "種別",
        "階層",
        "名称",
        "開始",
        "終了",
        "期間",
        "タスク詳細",
        "進捗",
        "作業進捗",
        "マイル",
        "サマリ",
        "クリティカル",
        "担当",
        "カレンダ",
        "リソース",
        "先行"
      ];
      const totalColumns = fixedHeaders.length + 1 + dateBand.length;
      const rows: WbsXlsxWorkbookLike["sheets"][number]["rows"] = [];
      const mergedRanges = [];
      const projectInfoBlock = projectInfoRows(
        model.project,
        calendarNameByUid,
        holidaySet.size,
        totalColumns,
        0,
        rows.length + 1
      );
      overlayRows(rows, 0, projectInfoBlock.rows, totalColumns);
      const exportTimestampRow = rows[1] || (rows[1] = emptyRow(totalColumns));
      exportTimestampRow.cells[9] = {
        value: formatWbsExportTimestamp(new Date()),
        horizontalAlign: "left",
        verticalAlign: "center"
      };
      mergedRanges.push(...projectInfoBlock.mergedRanges);
      rows.push(dateBandHeaderRow(fixedHeaders.length + 1, dateBand, model.project.currentDate, holidaySet, nonWorkingDayTypes));
      rows.push(weekdayHeaderRow(fixedHeaders, dateBand, model.project.currentDate, holidaySet, nonWorkingDayTypes));
      rows.push(...model.tasks.map((task) => ({
        height: taskRowHeight(task),
        cells: [
          identifierCell(task, task.uid),
          identifierCell(task, task.id),
          identifierCell(task, task.wbs || task.outlineNumber),
          kindCell(task),
          identifierCell(task, task.outlineLevel),
          taskCell(task, formatTaskLabel(task), "left"),
          taskCell(task, formatWbsDate(task.start), "center"),
          taskCell(task, formatWbsDate(task.finish), "center"),
          taskCell(task, formatDurationLabel(task, holidaySet, nonWorkingDayTypes, options.useBusinessDaysForProgressBand), "center"),
          detailCell(task, task.notes),
          progressCell(task, task.percentComplete),
          progressCell(task, task.percentWorkComplete),
          flagCell(task, task.milestone, "Mil"),
          flagCell(task, task.summary, "Sum"),
          flagCell(task, task.critical, "Crit"),
          referenceCell(task, truncateWbsReference(firstResourceName(resourceNamesByTaskUid.get(task.uid)), 14), "center"),
          referenceCell(task, formatCalendarLabel(task.calendarUID || model.project.calendarUID, calendarNameByUid), "center"),
          referenceCell(task, truncateWbsReference((resourceNamesByTaskUid.get(task.uid) || []).join(", "), 18)),
          referenceCell(task, truncateWbsReference(task.predecessors.map((item) => predecessorNameByUid.get(item.predecessorUid) || item.predecessorUid).join(", "), 18)),
          dividerCell(),
          ...dateBand.map((day) => dateBandCell(task, day, model.project.currentDate, holidaySet, nonWorkingDayTypes, options.useBusinessDaysForProgressBand))
        ]
      })));
      rows.push(emptyRow(totalColumns, 28));
      const legendBlock = legendRows(totalColumns, rows.length + 1);
      rows.push(...legendBlock.rows);
      mergedRanges.push(...legendBlock.mergedRanges);
      rows.push(emptyRow(totalColumns, 28));
      const summaryBlock = displaySummaryRows(
        dateBand.length,
        countBusinessDays(dateBand, holidaySet, nonWorkingDayTypes),
        model.project.currentDate,
        model.tasks.length,
        model.resources.length,
        model.assignments.length,
        model.calendars.length,
        totalColumns,
        0,
        rows.length + 1,
        options.displayDaysBeforeBaseDate,
        options.displayDaysAfterBaseDate,
        options.useBusinessDaysForDisplayRange,
        options.useBusinessDaysForProgressBand
      );
      rows.push(...summaryBlock.rows);
      mergedRanges.push(...summaryBlock.mergedRanges);

      return {
        sheets: [
          {
            name: "WBS",
            columns: [
              { width: pxWidth(45) }, { width: pxWidth(45) }, { width: pxWidth(65) }, { width: pxWidth(60) }, { width: pxWidth(45) }, { width: 42 },
              { width: pxWidth(85) }, { width: pxWidth(85) }, { width: pxWidth(65) }, { width: 28 }, { width: 14 },
              { width: 18, hidden: true }, { width: 12, hidden: true }, { width: 12, hidden: true }, { width: 12, hidden: true },
              { width: pxWidth(85) }, { width: 12, hidden: true }, { width: 20, hidden: true }, { width: 18, hidden: true }, { width: 3 },
              ...dateBand.map(() => ({ width: 6 }))
            ],
            mergedRanges,
            rows
          }
        ]
      };
    }

    return {
      exportWbsWorkbook
    };
  }

  (globalThis as typeof globalThis & {
    __mikuprojectWbsXlsxExport?: {
      createWbsXlsxExportHelper: (config: CreateWbsXlsxExportHelperConfig) => WbsXlsxExportHelper;
    };
  }).__mikuprojectWbsXlsxExport = {
    createWbsXlsxExportHelper
  };
})();
