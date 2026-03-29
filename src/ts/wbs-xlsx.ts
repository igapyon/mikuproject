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
  const ACTIVE_BAND_FILL = "#9FD5C9";
  const PROGRESS_BAND_FILL = "#5BAE9C";
  const WEEKEND_BAND_FILL = "#C9D3E1";
  const WEEK_START_BAND_FILL = "#E3EEF9";
  const MONTH_BOUNDARY_WEEK_FILL = "#D6E7F8";
  const MONTH_START_HEADER_FILL = "#DCEAF7";
  const TODAY_BAND_FILL = "#FFE6A7";
  const TODAY_ACTIVE_BAND_FILL = "#F3C96B";
  const TODAY_PROGRESS_BAND_FILL = "#D89A2B";
  const HOLIDAY_BAND_FILL = "#FCE4EC";
  const DIVIDER_FILL = "#D9E2EA";
  const BASEDATE_GUIDE_TAIL_FILL = "#FFF8E1";
  const NAME_COLUMN_FILL = "#FBFCFE";
  const SCHEDULE_COLUMN_FILL = "#FCFAF7";
  const PROGRESS_COLUMN_FILL = "#FCF8FB";
  const REFERENCE_COLUMN_FILL = "#F8FBFB";
  const WBS_LAYOUT = createWbsSheetLayoutHelper();

  function collectWbsHolidayDates(model: ProjectModel): string[] {
    const holidaySet = new Set<string>();
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

  function resolveProjectCalendar(model: ProjectModel): CalendarModel | undefined {
    if (model.project.calendarUID) {
      const projectCalendar = model.calendars.find((calendar) => calendar.uid === model.project.calendarUID);
      if (projectCalendar) {
        return projectCalendar;
      }
    }
    return model.calendars.find((calendar) => calendar.isBaseCalendar) || model.calendars[0];
  }

  function resolveCalendarDayWorking(
    calendarByUid: Map<string, CalendarModel>,
    calendar: CalendarModel | undefined,
    dayType: number,
    visiting = new Set<string>()
  ): boolean | undefined {
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

  function collectProjectNonWorkingDayTypes(model: ProjectModel): Set<number> {
    const calendarByUid = new Map(model.calendars.map((calendar) => [calendar.uid, calendar]));
    const projectCalendar = resolveProjectCalendar(model);
    const nonWorkingDayTypes = new Set<number>();
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
    const dividerColumnIndex = fixedHeaders.length + 1;
    const dateBandStartColumnIndex = dividerColumnIndex;
    const totalColumns = fixedHeaders.length + 1 + dateBand.length;
    const rows = [
      sheetTitleRow("WBS", totalColumns),
      emptyRow(totalColumns, 22)
    ];
    const mergedRanges = [];
    const projectInfoBlock = projectInfoRows(
      model.project,
      calendarNameByUid,
      holidaySet.size,
      totalColumns,
      0,
      rows.length + 1
    );
    overlayRows(rows, 2, projectInfoBlock.rows, totalColumns);
    mergedRanges.push(...projectInfoBlock.mergedRanges);
    const summaryBlock = displaySummaryRows(
      dateBand.length,
      countBusinessDays(dateBand, holidaySet, nonWorkingDayTypes),
      model.project.currentDate,
      model.tasks.length,
      model.resources.length,
      model.assignments.length,
      model.calendars.length,
      totalColumns,
      5,
      3,
      options.displayDaysBeforeBaseDate,
      options.displayDaysAfterBaseDate,
      options.useBusinessDaysForDisplayRange,
      options.useBusinessDaysForProgressBand
    );
    overlayRows(rows, 2, summaryBlock.rows, totalColumns);
    mergedRanges.push(...summaryBlock.mergedRanges);
    const weekBandRanges = buildWeekBandRanges(dateBand, dateBandStartColumnIndex, rows.length + 1);
    rows.push(weekBandRow(fixedHeaders.length + 1, weekBandRanges, dateBand.length));
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
        taskCell(task, formatTaskLabel(task), "center"),
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
    mergedRanges.push(...weekBandRanges.map((item) => item.range), ...legendBlock.mergedRanges);

    return {
      sheets: [
        {
          name: "WBS",
          columns: [
            { width: 8 }, { width: 8 }, { width: 12 }, { width: 10 }, { width: 10 }, { width: 42 },
            { width: 20 }, { width: 18 }, { width: 12 }, { width: 28 }, { width: 14 },
            { width: 18, hidden: true }, { width: 12, hidden: true }, { width: 12, hidden: true }, { width: 12, hidden: true },
            { width: 16 }, { width: 12, hidden: true }, { width: 20, hidden: true }, { width: 18, hidden: true }, { width: 3 },
            ...dateBand.map(() => ({ width: 6 }))
          ],
          mergedRanges,
          rows
        }
      ]
    };
  }

  function emptyRow(columnCount: number, height = 22) {
    return {
      height,
      cells: Array.from({ length: columnCount }, () => ({}))
    };
  }

  function overlayRows(
    rows: Array<{ height?: number; cells: WbsXlsxCellLike[] }>,
    startIndex: number,
    blockRows: Array<{ height?: number; cells: WbsXlsxCellLike[] }>,
    columnCount: number
  ) {
    blockRows.forEach((blockRow, offset) => {
      const rowIndex = startIndex + offset;
      if (!rows[rowIndex]) {
        rows[rowIndex] = emptyRow(columnCount);
      }
      const targetRow = rows[rowIndex];
      if ((blockRow.height || 0) > (targetRow.height || 0)) {
        targetRow.height = blockRow.height;
      }
      blockRow.cells.forEach((cell, cellIndex) => {
        if (hasCellContent(cell)) {
          targetRow.cells[cellIndex] = cell;
        }
      });
    });
  }

  function hasCellContent(cell: WbsXlsxCellLike | undefined): boolean {
    return !!cell && Object.keys(cell).length > 0;
  }

  function formatTaskLabel(task: TaskModel): string {
    const prefix = task.summary ? "> " : (task.milestone ? "* " : "- ");
    return `${"  ".repeat(Math.max(0, task.outlineLevel - 1))}${prefix}${task.name}`;
  }

  function classifyTaskKind(task: TaskModel): string {
    if (task.summary) {
      return "フェーズ";
    }
    if (task.milestone) {
      return "マイル";
    }
    return "タスク";
  }

  function firstResourceName(resourceNames: string[] | undefined): string {
    if (!resourceNames || resourceNames.length === 0) {
      return "";
    }
    return resourceNames[0];
  }

  function formatCalendarLabel(
    calendarUID: string | undefined,
    calendarNameByUid: Map<string, string>
  ): string {
    if (!calendarUID) {
      return "-";
    }
    const calendarName = calendarNameByUid.get(calendarUID);
    return calendarName ? `${calendarUID} ${truncateWbsReference(calendarName, 9)}` : calendarUID;
  }

  function displayReferenceValue(value: string | undefined): string {
    return value && value.trim() ? value : "-";
  }

  function truncateWbsReference(value: string | undefined, maxLength: number): string {
    const normalized = value?.trim() || "";
    if (!normalized) {
      return "";
    }
    if (normalized.length <= maxLength) {
      return normalized;
    }
    return `${normalized.slice(0, Math.max(1, maxLength - 3))}...`;
  }

  function referenceCell(
    task: TaskModel,
    value: string | undefined,
    horizontalAlign: "left" | "center" | "right" = "center"
  ): WbsXlsxCellLike {
    const displayValue = displayReferenceValue(value);
    const placeholder = displayValue === "-";
    return {
      value: displayValue,
      border: "thin",
      horizontalAlign: placeholder ? "center" : horizontalAlign,
      verticalAlign: "center",
      bold: task.summary || task.milestone || false,
      fillColor: placeholder
        ? PLACEHOLDER_FILL
        : (task.summary
          ? PHASE_FILL
          : (task.milestone ? MILESTONE_FILL : REFERENCE_COLUMN_FILL))
    };
  }

  function sheetTitleRow(title: string, columnCount: number) {
    return {
      height: 24,
      cells: [
        {
          value: title,
          bold: true,
          fontSize: 16,
          horizontalAlign: "left"
        },
        ...Array.from({ length: Math.max(0, columnCount - 1) }, () => ({
          fillColor: "#EEF4FA"
        }))
      ]
    };
  }

  function infoRow(text: string, columnCount: number) {
    return {
      height: 24,
      cells: [
        {
          value: text,
          border: "thin",
          horizontalAlign: "left"
        },
        ...Array.from({ length: Math.max(0, columnCount - 1) }, () => ({}))
      ]
    };
  }

  function projectInfoRows(
    project: ProjectModel["project"],
    calendarNameByUid: Map<string, string>,
    holidayCount: number,
    columnCount: number,
    startColumnIndex: number,
    startRowNumber: number
  ) {
    const scheduleMode = project.scheduleFromStart ? "開始基準" : "終了基準";
    const items: Array<{ label: string; value: string | number; fillColor: string }> = [
      { label: "題名", value: truncateWbsReference(project.title || "-", 18) || "-", fillColor: SUMMARY_ASSIGNMENT_FILL },
      { label: "カレンダ", value: formatCalendarLabel(project.calendarUID, calendarNameByUid), fillColor: SUMMARY_ASSIGNMENT_FILL },
      { label: "基準", value: scheduleMode, fillColor: SUMMARY_ASSIGNMENT_FILL },
      { label: "開始日", value: formatWbsDate(project.startDate), fillColor: SUMMARY_SCHEDULE_FILL },
      { label: "終了日", value: formatWbsDate(project.finishDate), fillColor: SUMMARY_SCHEDULE_FILL },
      { label: "現在日", value: formatWbsDate(project.currentDate), fillColor: SUMMARY_SCHEDULE_FILL },
      { label: "祝日", value: holidayCount, fillColor: SUMMARY_SCHEDULE_FILL }
    ];
    return {
      mergedRanges: [
        WBS_LAYOUT.range(
          WBS_LAYOUT.reference(startRowNumber, startColumnIndex),
          WBS_LAYOUT.reference(startRowNumber, startColumnIndex + 3)
        ),
        ...items.map((_, index) => {
          const rowNumber = startRowNumber + index + 1;
          return [
            WBS_LAYOUT.range(
              WBS_LAYOUT.reference(rowNumber, startColumnIndex),
              WBS_LAYOUT.reference(rowNumber, startColumnIndex + 1)
            ),
            WBS_LAYOUT.range(
              WBS_LAYOUT.reference(rowNumber, startColumnIndex + 2),
              WBS_LAYOUT.reference(rowNumber, startColumnIndex + 3)
            )
          ];
        }).flat()
      ],
      rows: [
        projectBlockHeaderRow(columnCount, startColumnIndex, "プロジェクト"),
        ...items.map((item) => projectPairRow(columnCount, startColumnIndex, item.label, item.value, item.fillColor))
      ]
    };
  }

  function legendRows(columnCount: number, startRowNumber: number) {
    const items: Array<{ value: string; fillColor: string }> = [
      { value: "進捗済み", fillColor: PROGRESS_BAND_FILL },
      { value: "予定帯", fillColor: ACTIVE_BAND_FILL },
      { value: "当日", fillColor: TODAY_BAND_FILL },
      { value: "週頭", fillColor: WEEK_START_BAND_FILL },
      { value: "週末", fillColor: WEEKEND_BAND_FILL },
      { value: "祝日", fillColor: HOLIDAY_BAND_FILL },
      { value: "━:フェーズ", fillColor: PHASE_FILL },
      { value: "■:進捗済みタスク", fillColor: PROGRESS_BAND_FILL },
      { value: "□:予定タスク", fillColor: ACTIVE_BAND_FILL },
      { value: "◆:マイルストーン", fillColor: MILESTONE_FILL },
      { value: "Mil:マイルストーン", fillColor: "#FBE4EC" },
      { value: "Sum:サマリ", fillColor: "#F7EAF0" },
      { value: "Crit:クリティカル", fillColor: "#F3E1E9" },
      { value: "-:未設定", fillColor: PLACEHOLDER_FILL }
    ];
    const startColumnRef = WBS_LAYOUT.reference(startRowNumber, WBS_LAYOUT.columnIndex("A"));
    const endColumnRef = WBS_LAYOUT.reference(startRowNumber, WBS_LAYOUT.columnIndex("C"));
    return {
      mergedRanges: [
        WBS_LAYOUT.range(startColumnRef, endColumnRef),
        ...items.map((_, index) => WBS_LAYOUT.range(
          WBS_LAYOUT.reference(startRowNumber + index + 1, WBS_LAYOUT.columnIndex("A")),
          WBS_LAYOUT.reference(startRowNumber + index + 1, WBS_LAYOUT.columnIndex("C"))
        ))
      ],
      rows: [
        blockHeaderRow(columnCount, 0, "凡例"),
        ...items.map((item) => mergedLabelRow(columnCount, 0, item.value, item.fillColor))
      ]
    };
  }

  function weekBandRow(
    fixedColumnCount: number,
    weekBandRanges: Array<{ startIndex: number; label: string; hasMonthBoundary: boolean }>,
    dateBandLength: number
  ) {
    const weekLabelColumnIndex = WBS_LAYOUT.columnIndex("S");
    const dividerColumnIndex = WBS_LAYOUT.columnIndex("T");
    const bandCells = Array.from({ length: dateBandLength }, () => ({} as WbsXlsxCellLike));
    weekBandRanges.forEach((item, index) => {
      bandCells[item.startIndex] = {
        value: item.label,
        bold: true,
        fontSize: 14,
        border: "thin" as const,
        horizontalAlign: "center" as const,
        fillColor: item.hasMonthBoundary ? MONTH_BOUNDARY_WEEK_FILL : (index % 2 === 0 ? "#EDF4FB" : "#EAF1F9")
      };
    });
    return {
      height: 24,
      cells: [
        ...Array.from({ length: fixedColumnCount }, (_, index) => {
          if (index === weekLabelColumnIndex) {
            return {
              value: "週",
              bold: true,
              fontSize: 14,
              border: "thin" as const,
              horizontalAlign: "center" as const,
              fillColor: "#E3EEF9"
            };
          }
          if (index === dividerColumnIndex) {
            return dividerCell();
          }
          if (index < weekLabelColumnIndex) {
            return {};
          }
          return {};
        }),
        ...bandCells
      ]
    };
  }

  function displaySummaryRows(
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
  ) {
    const displayWeeks = displayDays > 0 ? Math.ceil(displayDays / 7) : 0;
    const scheduleItems: Array<{ label: string; value: string | number; fillColor: string }> = [
      { label: "表示日", value: displayDays, fillColor: SUMMARY_SCHEDULE_FILL },
      { label: "表示週", value: displayWeeks, fillColor: SUMMARY_SCHEDULE_FILL },
      { label: "営業日", value: businessDays, fillColor: SUMMARY_SCHEDULE_FILL },
      { label: "前日数", value: displayDaysBeforeBaseDate ?? "-", fillColor: SUMMARY_SCHEDULE_FILL },
      { label: "後日数", value: displayDaysAfterBaseDate ?? "-", fillColor: SUMMARY_SCHEDULE_FILL },
      { label: "表示", value: useBusinessDaysForDisplayRange ? "営業日" : "暦日", fillColor: SUMMARY_SCHEDULE_FILL },
      { label: "進捗", value: useBusinessDaysForProgressBand ? "営業日" : "暦日", fillColor: SUMMARY_SCHEDULE_FILL },
      { label: "基準日", value: (baseDate || "-").slice(0, 10), fillColor: SUMMARY_SCHEDULE_FILL }
    ];
    const countItems: Array<{ label: string; value: string | number; fillColor: string }> = [
      { label: "タスク", value: taskCount, fillColor: SUMMARY_ASSIGNMENT_FILL },
      { label: "リソース", value: resourceCount, fillColor: SUMMARY_ASSIGNMENT_FILL },
      { label: "割当", value: assignmentCount, fillColor: SUMMARY_ASSIGNMENT_FILL },
      { label: "カレンダ", value: calendarCount, fillColor: SUMMARY_ASSIGNMENT_FILL }
    ];
    const blockRows = [blockHeaderRow(columnCount, startColumnIndex, "サマリ")];
    for (const item of scheduleItems) {
      blockRows.push(summaryPairRow(columnCount, startColumnIndex, item.label, item.value, item.fillColor));
    }
    for (let index = 0; index < countItems.length; index += 1) {
      const rowIndex = 1 + index;
      while (blockRows.length <= rowIndex) {
        blockRows.push(emptyRow(columnCount, 22));
      }
      overlaySummaryPair(blockRows[rowIndex], startColumnIndex + 3, countItems[index].label, countItems[index].value, countItems[index].fillColor);
    }
    return {
      mergedRanges: [WBS_LAYOUT.range(
        WBS_LAYOUT.reference(startRowNumber, startColumnIndex),
        WBS_LAYOUT.reference(startRowNumber, startColumnIndex + 1)
      )],
      rows: blockRows
    };
  }

  function blockHeaderRow(columnCount: number, startColumnIndex: number, title: string) {
    const cells = Array.from({ length: columnCount }, () => ({} as WbsXlsxCellLike));
    cells[startColumnIndex] = {
      value: title,
      border: "thin",
      horizontalAlign: "left",
      bold: true,
      fontSize: 14,
      fillColor: HEADER_ID_FILL
    };
    cells[startColumnIndex + 1] = {
      value: "",
      border: "thin",
      fillColor: HEADER_ID_FILL
    };
    return { height: 24, cells };
  }

  function projectBlockHeaderRow(columnCount: number, startColumnIndex: number, title: string) {
    const cells = Array.from({ length: columnCount }, () => ({} as WbsXlsxCellLike));
    cells[startColumnIndex] = {
      value: title,
      border: "thin",
      horizontalAlign: "left",
      bold: true,
      fontSize: 14,
      fillColor: HEADER_ID_FILL
    };
    for (let offset = 1; offset < 4; offset += 1) {
      cells[startColumnIndex + offset] = {
        value: "",
        border: "thin",
        fillColor: HEADER_ID_FILL
      };
    }
    return { height: 24, cells };
  }

  function projectPairRow(
    columnCount: number,
    startColumnIndex: number,
    label: string,
    value: string | number,
    fillColor: string
  ) {
    const cells = Array.from({ length: columnCount }, () => ({} as WbsXlsxCellLike));
    cells[startColumnIndex] = {
      value: label,
      border: "thin",
      horizontalAlign: "right",
      bold: true,
      fillColor
    };
    cells[startColumnIndex + 1] = {
      value: "",
      border: "thin",
      fillColor
    };
    cells[startColumnIndex + 2] = {
      value,
      border: "thin",
      horizontalAlign: typeof value === "number" ? "center" : "left",
      bold: true,
      fillColor
    };
    cells[startColumnIndex + 3] = {
      value: "",
      border: "thin",
      fillColor
    };
    return { height: 22, cells };
  }

  function summaryPairRow(
    columnCount: number,
    startColumnIndex: number,
    label: string,
    value: string | number,
    fillColor: string
  ) {
    const cells = Array.from({ length: columnCount }, () => ({} as WbsXlsxCellLike));
    cells[startColumnIndex] = summaryStatCell(label, fillColor, false);
    cells[startColumnIndex + 1] = summaryStatCell(value, fillColor, true);
    return { height: 22, cells };
  }

  function overlaySummaryPair(
    row: { height?: number; cells: WbsXlsxCellLike[] },
    startColumnIndex: number,
    label: string,
    value: string | number,
    fillColor: string
  ) {
    row.cells[startColumnIndex] = summaryStatCell(label, fillColor, false);
    row.cells[startColumnIndex + 1] = summaryStatCell(value, fillColor, true);
    row.height = Math.max(row.height || 22, 22);
  }

  function mergedLabelRow(
    columnCount: number,
    startColumnIndex: number,
    value: string,
    fillColor: string
  ) {
    const cells = Array.from({ length: columnCount }, () => ({} as WbsXlsxCellLike));
    cells[startColumnIndex] = {
      value,
      border: "thin",
      horizontalAlign: "center",
      bold: true,
      fillColor
    };
    cells[startColumnIndex + 1] = {
      value: "",
      border: "thin",
      fillColor
    };
    return { height: 24, cells };
  }

  function summaryStatCell(value: string | number, fillColor: string, isValueCell: boolean): WbsXlsxCellLike {
    const valueAlign = typeof value === "number" ? "center" : "left";
    return {
      value,
      border: "thin",
      horizontalAlign: isValueCell ? valueAlign : "right",
      bold: true,
      fillColor
    };
  }

  function headerRow(labels: Array<string | WbsXlsxCellLike>) {
    return {
      height: 24,
      cells: labels.map((label) => {
        if (typeof label === "string") {
          return {
            value: label,
            bold: true,
            fillColor: headerFillForLabel(label),
            border: "thin" as const,
            horizontalAlign: "center" as const,
            verticalAlign: "center" as const
          };
        }
        return {
          border: "thin" as const,
          horizontalAlign: "center" as const,
          verticalAlign: "center" as const,
          ...label
        };
      })
    };
  }

  function weekdayRow(
    fixedColumnCount: number,
    dateBand: string[],
    currentDate: string | undefined,
    holidaySet: Set<string>,
    nonWorkingDayTypes: Set<number>
  ) {
    return {
      height: 24,
      cells: [
        ...Array.from({ length: fixedColumnCount }, () => ({} as WbsXlsxCellLike)),
        ...dateBand.map((day) => weekdayCell(day, currentDate, holidaySet, nonWorkingDayTypes))
      ]
    };
  }

  function dateBandHeaderRow(
    fixedColumnCount: number,
    dateBand: string[],
    currentDate: string | undefined,
    holidaySet: Set<string>,
    nonWorkingDayTypes: Set<number>
  ) {
    return {
      height: 24,
      cells: [
        ...Array.from({ length: fixedColumnCount }, () => ({} as WbsXlsxCellLike)),
        ...dateBand.map((day) => dateNumberCell(day, currentDate, holidaySet, nonWorkingDayTypes))
      ]
    };
  }

  function weekdayHeaderRow(
    fixedHeaders: string[],
    dateBand: string[],
    currentDate: string | undefined,
    holidaySet: Set<string>,
    nonWorkingDayTypes: Set<number>
  ) {
    return headerRow([
      ...fixedHeaders,
      dividerCell(),
      ...dateBand.map((day) => weekdayCell(day, currentDate, holidaySet, nonWorkingDayTypes))
    ]);
  }

  function dividerCell(): WbsXlsxCellLike {
    return {
      value: "",
      fillColor: DIVIDER_FILL,
      border: "thin",
      horizontalAlign: "center",
      verticalAlign: "center"
    };
  }

  function headerFillForLabel(label: string): string {
    if (label === "UID" || label === "ID") {
      return HEADER_ID_FILL;
    }
    if (label === "WBS" || label === "種別" || label === "階層" || label === "名称") {
      return HEADER_STRUCTURE_FILL;
    }
    if (label === "開始" || label === "終了" || label === "期間") {
      return HEADER_SCHEDULE_FILL;
    }
    if (label === "タスク詳細") {
      return HEADER_FILL;
    }
    if (label === "進捗" || label === "作業進捗" || label === "マイル" || label === "サマリ" || label === "クリティカル") {
      return HEADER_STATUS_FILL;
    }
    if (label === "担当" || label === "カレンダ" || label === "リソース" || label === "先行") {
      return HEADER_ASSIGNMENT_FILL;
    }
    return HEADER_FILL;
  }

  function cell(value: string | number | boolean | undefined): WbsXlsxCellLike {
    if (value === undefined || value === "") {
      return {};
    }
    return {
      value,
      border: "thin"
    };
  }

  function taskCell(
    task: TaskModel,
    value: string | number | boolean | undefined,
    horizontalAlign: "left" | "center" | "right" = "left"
  ): WbsXlsxCellLike {
    if (value === undefined || value === "") {
      return {};
    }
    return {
      value,
      border: "thin",
      horizontalAlign,
      verticalAlign: "center",
      wrapText: typeof value === "string" ? true : undefined,
      bold: task.summary || task.milestone || false,
      fillColor: task.summary
        ? PHASE_FILL
        : (task.milestone
          ? MILESTONE_FILL
          : (horizontalAlign === "left"
            ? NAME_COLUMN_FILL
            : (horizontalAlign === "center" ? SCHEDULE_COLUMN_FILL : undefined)))
    };
  }

  function detailCell(task: TaskModel, value: string | undefined): WbsXlsxCellLike {
    const normalized = value?.trim() || "";
    const placeholder = !normalized;
    return {
      value: placeholder ? "-" : normalized,
      border: "thin",
      horizontalAlign: placeholder ? "center" : "left",
      verticalAlign: "center",
      wrapText: placeholder ? undefined : true,
      bold: task.summary || task.milestone || false,
      fillColor: placeholder
        ? PLACEHOLDER_FILL
        : (task.summary
          ? PHASE_FILL
          : (task.milestone ? MILESTONE_FILL : NAME_COLUMN_FILL))
    };
  }

  function taskRowHeight(task: TaskModel): number | undefined {
    const labelLength = formatTaskLabel(task).length;
    const notesLength = (task.notes || "").trim().length;
    const maxLength = Math.max(labelLength, notesLength);
    if (maxLength > 72) {
      return 46;
    }
    if (maxLength > 36) {
      return 34;
    }
    return 34;
  }

  function kindCell(task: TaskModel): WbsXlsxCellLike {
    return {
      value: classifyTaskKind(task),
      border: "thin",
      horizontalAlign: "center",
      verticalAlign: "center",
      bold: true,
      fillColor: task.summary ? PHASE_FILL : (task.milestone ? MILESTONE_FILL : TASK_KIND_FILL)
    };
  }

  function identifierCell(task: TaskModel, value: string | number | boolean | undefined): WbsXlsxCellLike {
    if (value === undefined || value === "") {
      return {};
    }
    return {
      value,
      border: "thin",
      horizontalAlign: "center",
      verticalAlign: "center",
      bold: task.summary || task.milestone || false,
      fillColor: task.summary ? PHASE_FILL : (task.milestone ? MILESTONE_FILL : IDENTIFIER_FILL)
    };
  }

  function flagCell(task: TaskModel, enabled: boolean | undefined, marker: string): WbsXlsxCellLike {
    return {
      value: enabled ? marker : "",
      border: "thin",
      horizontalAlign: "center",
      verticalAlign: "center",
      bold: !!enabled,
      fillColor: task.summary ? PHASE_FILL : (task.milestone ? MILESTONE_FILL : undefined)
    };
  }

  function progressCell(task: TaskModel, value: number | undefined): WbsXlsxCellLike {
    const progressValue = formatProgressLabel(value);
    return {
      value: progressValue,
      border: "thin",
      horizontalAlign: "center",
      verticalAlign: "center",
      wrapText: true,
      bold: task.summary || task.milestone || false,
      fillColor: task.summary ? PHASE_FILL : (task.milestone ? MILESTONE_FILL : PROGRESS_COLUMN_FILL)
    };
  }

  function formatProgressLabel(value: number | undefined): string {
    if (value === undefined || value === null || !Number.isFinite(value)) {
      return "";
    }
    const clamped = Math.max(0, Math.min(100, Math.round(value)));
    const filled = Math.round(clamped / 10);
    const bar = `${"#".repeat(filled)}${"-".repeat(10 - filled)}`;
    return `${String(clamped).padStart(3, " ")}%\n[${bar}]`;
  }

  function formatDurationLabel(
    task: TaskModel,
    holidaySet: Set<string>,
    nonWorkingDayTypes: Set<number>,
    useBusinessDaysForProgressBand: boolean | undefined
  ): string {
    if (useBusinessDaysForProgressBand) {
      const businessDays = enumerateBusinessDays(task.start, task.finish, holidaySet, nonWorkingDayTypes).length;
      return businessDays > 0 ? `${businessDays}営業日` : "-";
    }
    const calendarDays = buildDateBand(task.start, task.finish).length;
    return calendarDays > 0 ? `${calendarDays}日` : "-";
  }

  function formatWbsDate(value: string | undefined): string {
    return value ? value.slice(0, 10) : "-";
  }

  function dateNumberCell(
    day: string,
    currentDate: string | undefined,
    holidaySet: Set<string>,
    nonWorkingDayTypes: Set<number>
  ): WbsXlsxCellLike {
    const isToday = isSameDay(day, currentDate);
    const isWeekendDay = isWeeklyNonWorkingDay(day, nonWorkingDayTypes);
    const isHoliday = holidaySet.has(day);
    const weekStart = isWeekStart(day);
    const monthStart = isMonthStart(day);
    return {
      value: formatDateValue(day),
      bold: true,
      border: "thin",
      horizontalAlign: "center",
      verticalAlign: "center",
      fillColor: isToday ? TODAY_BAND_FILL : (isHoliday ? HOLIDAY_BAND_FILL : (isWeekendDay ? WEEKEND_BAND_FILL : (monthStart ? MONTH_START_HEADER_FILL : (weekStart ? WEEK_START_BAND_FILL : HEADER_FILL))))
    };
  }

  function weekdayCell(
    day: string,
    currentDate: string | undefined,
    holidaySet: Set<string>,
    nonWorkingDayTypes: Set<number>
  ): WbsXlsxCellLike {
    const isToday = isSameDay(day, currentDate);
    const isWeekendDay = isWeeklyNonWorkingDay(day, nonWorkingDayTypes);
    const isHoliday = holidaySet.has(day);
    const weekStart = isWeekStart(day);
    const monthStart = isMonthStart(day);
    return {
      value: formatWeekdayValue(day),
      bold: true,
      border: "thin",
      horizontalAlign: "center",
      verticalAlign: "center",
      fillColor: isToday ? TODAY_BAND_FILL : (isHoliday ? HOLIDAY_BAND_FILL : (isWeekendDay ? WEEKEND_BAND_FILL : (monthStart ? MONTH_START_HEADER_FILL : (weekStart ? WEEK_START_BAND_FILL : HEADER_FILL))))
    };
  }

  function dateBandCell(
    task: TaskModel,
    day: string,
    currentDate: string | undefined,
    holidaySet: Set<string>,
    nonWorkingDayTypes: Set<number>,
    useBusinessDaysForProgressBand: boolean | undefined
  ): WbsXlsxCellLike {
    const isWeekendDay = isWeeklyNonWorkingDay(day, nonWorkingDayTypes);
    const isHoliday = holidaySet.has(day);
    const isNonWorkingDay = isWeekendDay || isHoliday;
    const isTaskStart = isSameDay(day, task.start);
    const isTaskFinish = isSameDay(day, task.finish);
    const active = includesDay(task.start, task.finish, day) && (!isNonWorkingDay || isTaskStart || isTaskFinish);
    const isToday = isSameDay(day, currentDate);
    const weekStart = isWeekStart(day);
    const complete = active && isCompletedDay(task, day, holidaySet, nonWorkingDayTypes, useBusinessDaysForProgressBand);
    return {
      value: active ? activeBandMarker(task, complete) : "",
      border: "thin",
      horizontalAlign: "center",
      verticalAlign: "center",
      fillColor: active
        ? (complete
          ? (isToday ? TODAY_PROGRESS_BAND_FILL : PROGRESS_BAND_FILL)
          : (isToday ? TODAY_ACTIVE_BAND_FILL : ACTIVE_BAND_FILL))
        : (isToday ? TODAY_BAND_FILL : (isHoliday ? HOLIDAY_BAND_FILL : (isWeekendDay ? WEEKEND_BAND_FILL : (weekStart ? WEEK_START_BAND_FILL : BAND_FILL))))
    };
  }

  function activeBandMarker(task: TaskModel, complete: boolean): string {
    if (task.summary) {
      return "━";
    }
    if (task.milestone) {
      return "◆";
    }
    return complete ? "■" : "□";
  }

  function buildDateBand(startDate: string | undefined, finishDate: string | undefined): string[] {
    const start = parseDateOnly(startDate);
    const finish = parseDateOnly(finishDate);
    if (!start || !finish || start.getTime() > finish.getTime()) {
      return [];
    }
    const days: string[] = [];
    const cursor = new Date(start.getTime());
    while (cursor.getTime() <= finish.getTime()) {
      days.push(formatDateOnly(cursor));
      cursor.setDate(cursor.getDate() + 1);
    }
    return days;
  }

  function buildDisplayDateBand(
    startDate: string | undefined,
    finishDate: string | undefined,
    baseDate: string | undefined,
    displayDaysBeforeBaseDate: number | undefined,
    displayDaysAfterBaseDate: number | undefined,
    holidaySet: Set<string>,
    nonWorkingDayTypes: Set<number>,
    useBusinessDaysForDisplayRange: boolean | undefined
  ): string[] {
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

  function normalizeDisplayDayCount(value: number | undefined): number | null {
    if (value === undefined || value === null || !Number.isFinite(value)) {
      return null;
    }
    return Math.max(0, Math.floor(value));
  }

  function countBusinessDays(dateBand: string[], holidaySet: Set<string>, nonWorkingDayTypes: Set<number>): number {
    return dateBand.filter((day) => !isWeeklyNonWorkingDay(day, nonWorkingDayTypes) && !holidaySet.has(day)).length;
  }

  function shiftCalendarDays(base: Date, offset: number): Date {
    const result = new Date(base.getTime());
    result.setDate(result.getDate() + offset);
    return result;
  }

  function shiftBusinessDays(base: Date, offset: number, holidaySet: Set<string>, nonWorkingDayTypes: Set<number>): Date {
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

  function buildWeekBandRanges(dateBand: string[], startColumnIndex: number, rowNumber: number) {
    const ranges: Array<{ range: string; startIndex: number; label: string; hasMonthBoundary: boolean }> = [];
    if (dateBand.length === 0) {
      return ranges;
    }
    let chunkStart = 0;
    while (chunkStart < dateBand.length) {
      const weekStart = formatWeekKey(dateBand[chunkStart]);
      let chunkEnd = chunkStart;
      while (chunkEnd + 1 < dateBand.length && formatWeekKey(dateBand[chunkEnd + 1]) === weekStart) {
        chunkEnd += 1;
      }
      const chunkDays = dateBand.slice(chunkStart, chunkEnd + 1);
      ranges.push({
        range: WBS_LAYOUT.range(
          WBS_LAYOUT.reference(rowNumber, startColumnIndex + chunkStart),
          WBS_LAYOUT.reference(rowNumber, startColumnIndex + chunkEnd)
        ),
        startIndex: chunkStart,
        label: formatWeekLabel(weekStart, chunkDays),
        hasMonthBoundary: chunkDays.some((day) => isMonthStart(day))
      });
      chunkStart = chunkEnd + 1;
    }
    return ranges;
  }

  function includesDay(startDate: string | undefined, finishDate: string | undefined, day: string): boolean {
    const start = parseDateOnly(startDate);
    const finish = parseDateOnly(finishDate);
    const target = parseDateOnly(day);
    if (!start || !finish || !target) {
      return false;
    }
    return start.getTime() <= target.getTime() && target.getTime() <= finish.getTime();
  }

  function isCompletedDay(
    task: TaskModel,
    day: string,
    holidaySet: Set<string>,
    nonWorkingDayTypes: Set<number>,
    useBusinessDaysForProgressBand: boolean | undefined
  ): boolean {
    const start = parseDateOnly(task.start);
    const finish = parseDateOnly(task.finish);
    const target = parseDateOnly(day);
    if (!start || !finish || !target) {
      return false;
    }
    if (useBusinessDaysForProgressBand) {
      const activeBusinessDays = enumerateBusinessDays(task.start, task.finish, holidaySet, nonWorkingDayTypes);
      if (activeBusinessDays.length === 0) {
        return false;
      }
      const percent = Math.max(0, Math.min(100, task.percentComplete || 0));
      const completedDays = Math.floor(activeBusinessDays.length * (percent / 100));
      const dayKey = formatDateOnly(target);
      const dayIndex = activeBusinessDays.indexOf(dayKey);
      return dayIndex >= 0 && dayIndex < completedDays;
    }
    const totalDays = Math.floor((finish.getTime() - start.getTime()) / 86400000) + 1;
    if (totalDays <= 0) {
      return false;
    }
    const percent = Math.max(0, Math.min(100, task.percentComplete || 0));
    const completedDays = Math.floor(totalDays * (percent / 100));
    const dayIndex = Math.floor((target.getTime() - start.getTime()) / 86400000);
    return dayIndex >= 0 && dayIndex < completedDays;
  }

  function enumerateBusinessDays(
    startDate: string | undefined,
    finishDate: string | undefined,
    holidaySet: Set<string>,
    nonWorkingDayTypes: Set<number>
  ): string[] {
    return buildDateBand(startDate, finishDate).filter((day) => !isWeeklyNonWorkingDay(day, nonWorkingDayTypes) && !holidaySet.has(day));
  }

  function isSameDay(day: string, other: string | undefined): boolean {
    return day === (other || "").slice(0, 10);
  }

  function isWeeklyNonWorkingDay(day: string, nonWorkingDayTypes: Set<number>): boolean {
    const target = parseDateOnly(day);
    if (!target) {
      return false;
    }
    const dayType = target.getDay() === 0 ? 1 : target.getDay() + 1;
    return nonWorkingDayTypes.has(dayType);
  }

  function isWeekStart(day: string): boolean {
    const target = parseDateOnly(day);
    if (!target) {
      return false;
    }
    return target.getDay() === 0;
  }

  function isMonthStart(day: string): boolean {
    const target = parseDateOnly(day);
    if (!target) {
      return false;
    }
    return target.getDate() === 1;
  }

  function parseDateOnly(value: string | undefined): Date | null {
    if (!value || value.length < 10) {
      return null;
    }
    const dateOnly = value.slice(0, 10);
    const [yearText, monthText, dayText] = dateOnly.split("-");
    const year = Number(yearText);
    const month = Number(monthText);
    const day = Number(dayText);
    if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) {
      return null;
    }
    return new Date(year, month - 1, day);
  }

  function expandExceptionDays(exception: CalendarExceptionModel): string[] {
    const from = (exception.fromDate || "").slice(0, 10);
    const to = (exception.toDate || "").slice(0, 10);
    if (!from) {
      return [];
    }
    if (!to || to === from) {
      return [from];
    }
    const start = parseDateOnly(from);
    const finish = parseDateOnly(to);
    if (!start || !finish || start.getTime() > finish.getTime()) {
      return [from];
    }
    const days: string[] = [];
    const cursor = new Date(start.getTime());
    while (cursor.getTime() <= finish.getTime()) {
      days.push(formatDateOnly(cursor));
      cursor.setDate(cursor.getDate() + 1);
    }
    return days;
  }

  function formatDateOnly(value: Date): string {
    return [
      value.getFullYear(),
      String(value.getMonth() + 1).padStart(2, "0"),
      String(value.getDate()).padStart(2, "0")
    ].join("-");
  }

  function formatDateValue(day: string): string {
    const target = parseDateOnly(day);
    if (!target) {
      return day;
    }
    return `${target.getMonth() + 1}/${target.getDate()}`;
  }

  function formatWeekdayValue(day: string): string {
    const target = parseDateOnly(day);
    if (!target) {
      return day;
    }
    const weekdays = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
    return weekdays[target.getDay()];
  }

  function formatWeekKey(day: string): string {
    const target = parseDateOnly(day);
    if (!target) {
      return day;
    }
    const sunday = new Date(target.getTime());
    const offset = sunday.getDay();
    sunday.setDate(sunday.getDate() - offset);
    return formatDateOnly(sunday);
  }

  function formatWeekLabel(weekKey: string, days: string[]): string {
    if (days.length === 0) {
      return "週";
    }
    const start = parseDateOnly(weekKey);
    if (!start) {
      return weekKey;
    }
    const monthSet = new Set(days.map((day) => {
      const target = parseDateOnly(day);
      return target ? target.getMonth() : -1;
    }));
    const startLabel = `${String(start.getMonth() + 1).padStart(2, "0")}/${String(start.getDate()).padStart(2, "0")}`;
    if (monthSet.size <= 1) {
      return `週 ${startLabel}`;
    }
    const tailMonths = Array.from(monthSet)
      .filter((monthIndex) => monthIndex >= 0 && monthIndex !== start.getMonth())
      .map((monthIndex) => String(monthIndex + 1).padStart(2, "0"));
    return `週 ${startLabel} / ${tailMonths.join(" / ")}`;
  }

  function createWbsSheetLayoutHelper(): WbsSheetLayoutHelper {
    return {
      columnName(columnIndex: number): string {
        let current = columnIndex + 1;
        let name = "";
        while (current > 0) {
          const remainder = (current - 1) % 26;
          name = String.fromCharCode(65 + remainder) + name;
          current = Math.floor((current - 1) / 26);
        }
        return name;
      },
      columnIndex(columnReference: string): number {
        const normalized = (columnReference || "").trim().toUpperCase();
        if (!/^[A-Z]+$/.test(normalized)) {
          throw new Error(`Invalid column reference: ${columnReference}`);
        }
        let value = 0;
        for (const character of normalized) {
          value = (value * 26) + (character.charCodeAt(0) - 64);
        }
        return value - 1;
      },
      reference(rowNumber: number, columnIndex: number): string {
        if (!Number.isInteger(rowNumber) || rowNumber <= 0) {
          throw new Error(`Invalid row number: ${rowNumber}`);
        }
        if (!Number.isInteger(columnIndex) || columnIndex < 0) {
          throw new Error(`Invalid column index: ${columnIndex}`);
        }
        return `${this.columnName(columnIndex)}${rowNumber}`;
      },
      parseCellReference(reference: string) {
        const match = /^([A-Z]+)(\d+)$/i.exec((reference || "").trim());
        if (!match) {
          throw new Error(`Invalid cell reference: ${reference}`);
        }
        const rowNumber = Number(match[2]);
        const columnName = match[1].toUpperCase();
        const columnIndex = this.columnIndex(columnName);
        return {
          reference: `${columnName}${rowNumber}`,
          rowNumber,
          rowIndex: rowNumber - 1,
          columnName,
          columnIndex
        };
      },
      range(startReference: string, endReference: string): string {
        return `${startReference}:${endReference}`;
      },
      describeCell(reference: string): string {
        const cell = this.parseCellReference(reference);
        return `${cell.reference} (row ${cell.rowNumber}, rowIndex ${cell.rowIndex}, column ${cell.columnName}, columnIndex ${cell.columnIndex})`;
      },
      logCell(reference: string, label?: string, logger?: (...args: unknown[]) => void): string {
        const message = label
          ? `${label}: ${this.describeCell(reference)}`
          : this.describeCell(reference);
        (logger || console.log)(message);
        return message;
      }
    };
  }

  (globalThis as typeof globalThis & {
    __mikuprojectWbsXlsx?: {
      collectWbsHolidayDates: typeof collectWbsHolidayDates;
      exportWbsWorkbook: typeof exportWbsWorkbook;
      layout: WbsSheetLayoutHelper;
    };
  }).__mikuprojectWbsXlsx = {
    collectWbsHolidayDates,
    exportWbsWorkbook,
    layout: WBS_LAYOUT
  };
})();
