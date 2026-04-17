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

  type WbsSheetLayoutHelper = {
    reference(rowNumber: number, columnIndex: number): string;
    range(startReference: string, endReference: string): string;
    columnIndex(columnReference: string): number;
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

  type WbsXlsxCellsConfig = {
    layout: WbsSheetLayoutHelper;
    fills: Record<string, string>;
    formatTaskLabel(task: TaskModel): string;
    classifyTaskKind(task: TaskModel): string;
    buildDateBand(startDate: string | undefined, finishDate: string | undefined): string[];
    parseDateOnly(value: string | undefined): Date | null;
    formatDateOnly(value: Date | null): string;
    isWeeklyNonWorkingDay(day: string, nonWorkingDayTypes: Set<number>): boolean;
  };

  function createWbsXlsxCellsHelper(config: WbsXlsxCellsConfig): WbsXlsxCellsHelper {
    const { layout, fills, formatTaskLabel, classifyTaskKind, buildDateBand, parseDateOnly, formatDateOnly, isWeeklyNonWorkingDay } = config;

    function dividerCell(): WbsXlsxCellLike {
      return {
        value: "",
        fillColor: fills.divider,
        border: "thin",
        horizontalAlign: "center",
        verticalAlign: "center"
      };
    }

    function headerFillForLabel(label: string): string {
      if (label === "UID" || label === "ID") {
        return fills.headerId;
      }
      if (label === "WBS" || label === "種別" || label === "階層" || label === "名称") {
        return fills.headerStructure;
      }
      if (label === "開始" || label === "終了" || label === "期間") {
        return fills.headerSchedule;
      }
      if (label === "タスク詳細") {
        return fills.headerFill;
      }
      if (label === "進捗" || label === "作業進捗" || label === "マイル" || label === "サマリ" || label === "クリティカル") {
        return fills.headerStatus;
      }
      if (label === "担当" || label === "カレンダ" || label === "リソース" || label === "先行") {
        return fills.headerAssignment;
      }
      return fills.headerFill;
    }

    function stringifyCellValue(value: string | number | boolean): string {
      return typeof value === "string" ? value : String(value);
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
        value: stringifyCellValue(value),
        border: "thin",
        horizontalAlign,
        verticalAlign: "center",
        wrapText: typeof value === "string" ? true : undefined,
        bold: task.summary || task.milestone || false,
        fillColor: task.summary
          ? fills.phase
          : (task.milestone
            ? fills.milestone
            : (horizontalAlign === "left"
              ? fills.nameColumn
              : (horizontalAlign === "center" ? fills.scheduleColumn : undefined)))
      };
    }

    function detailCell(task: TaskModel, value: string | undefined): WbsXlsxCellLike {
      const normalized = value?.trim() || "";
      const placeholder = !normalized;
      return {
        value: placeholder ? "-" : normalized,
        border: "thin",
        horizontalAlign: "left",
        verticalAlign: "center",
        wrapText: placeholder ? undefined : true,
        bold: task.summary || task.milestone || false,
        fillColor: placeholder
          ? fills.placeholder
          : (task.summary
            ? fills.phase
            : (task.milestone ? fills.milestone : fills.nameColumn))
      };
    }

    function estimateWrappedLineCount(value: string, charactersPerLine: number): number {
      const normalized = value.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
      if (!normalized) {
        return 1;
      }
      return normalized
        .split("\n")
        .reduce((count, line) => count + Math.max(1, Math.ceil(line.length / charactersPerLine)), 0);
    }

    function taskRowHeight(task: TaskModel): number | undefined {
      const labelLineCount = estimateWrappedLineCount(formatTaskLabel(task), 22);
      const notesLineCount = estimateWrappedLineCount((task.notes || "").trim(), 18);
      const maxLineCount = Math.max(labelLineCount, notesLineCount, 1);
      if (maxLineCount >= 5) {
        return 82;
      }
      if (maxLineCount === 4) {
        return 70;
      }
      if (maxLineCount === 3) {
        return 58;
      }
      if (maxLineCount === 2) {
        return 46;
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
        fillColor: task.summary ? fills.phase : (task.milestone ? fills.milestone : fills.taskKind)
      };
    }

    function identifierCell(task: TaskModel, value: string | number | boolean | undefined): WbsXlsxCellLike {
      if (value === undefined || value === "") {
        return {};
      }
      return {
        value: stringifyCellValue(value),
        border: "thin",
        horizontalAlign: "center",
        verticalAlign: "center",
        bold: task.summary || task.milestone || false,
        fillColor: task.summary ? fills.phase : (task.milestone ? fills.milestone : fills.identifier)
      };
    }

    function flagCell(task: TaskModel, enabled: boolean | undefined, marker: string): WbsXlsxCellLike {
      return {
        value: enabled ? marker : "",
        border: "thin",
        horizontalAlign: "center",
        verticalAlign: "center",
        bold: !!enabled,
        fillColor: task.summary ? fills.phase : (task.milestone ? fills.milestone : undefined)
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

    function progressCell(task: TaskModel, value: number | undefined): WbsXlsxCellLike {
      const progressValue = formatProgressLabel(value);
      return {
        value: progressValue,
        border: "thin",
        horizontalAlign: "center",
        verticalAlign: "center",
        wrapText: true,
        bold: task.summary || task.milestone || false,
        fillColor: task.summary ? fills.phase : (task.milestone ? fills.milestone : fills.progressColumn)
      };
    }

    function enumerateBusinessDays(
      startDate: string | undefined,
      finishDate: string | undefined,
      holidaySet: Set<string>,
      nonWorkingDayTypes: Set<number>
    ): string[] {
      return buildDateBand(startDate, finishDate).filter((day) => !isWeeklyNonWorkingDay(day, nonWorkingDayTypes) && !holidaySet.has(day));
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

    function isSameDay(day: string, other: string | undefined): boolean {
      return day === (other || "").slice(0, 10);
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
        fillColor: isToday ? fills.todayBand : (isHoliday ? fills.holidayBand : (isWeekendDay ? fills.weekendBand : (monthStart ? fills.monthStartHeader : (weekStart ? fills.weekStartBand : fills.headerFill))))
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
      const target = parseDateOnly(day);
      const dayType = target ? (target.getDay() === 0 ? 1 : target.getDay() + 1) : 0;
      const weekendHeaderFill = dayType === 7 ? fills.saturdayHeader : (dayType === 1 ? fills.sundayHeader : fills.weekendBand);
      return {
        value: formatWeekdayValue(day),
        bold: true,
        border: "thin",
        horizontalAlign: "center",
        verticalAlign: "center",
        fillColor: isHoliday ? fills.holidayBand : (isWeekendDay ? weekendHeaderFill : (isToday ? fills.todayBand : (monthStart ? fills.monthStartHeader : (weekStart ? fills.weekStartBand : fills.headerFill))))
      };
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

    function activeBandMarker(task: TaskModel, complete: boolean): string {
      if (task.summary) {
        return "━";
      }
      if (task.milestone) {
        return "◆";
      }
      return complete ? "■" : "□";
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
            ? (isToday ? fills.todayProgressBand : fills.progressBand)
            : (isToday ? fills.todayActiveBand : fills.activeBand))
          : (isToday ? fills.todayBand : (isHoliday ? fills.holidayBand : (isWeekendDay ? fills.weekendBand : (weekStart ? fills.weekStartBand : undefined))))
      };
    }

    function countBusinessDays(dateBand: string[], holidaySet: Set<string>, nonWorkingDayTypes: Set<number>): number {
      return dateBand.filter((day) => !isWeeklyNonWorkingDay(day, nonWorkingDayTypes) && !holidaySet.has(day)).length;
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
          range: layout.range(
            layout.reference(rowNumber, startColumnIndex + chunkStart),
            layout.reference(rowNumber, startColumnIndex + chunkEnd)
          ),
          startIndex: chunkStart,
          label: formatWeekLabel(weekStart, chunkDays),
          hasMonthBoundary: chunkDays.some((day) => isMonthStart(day))
        });
        chunkStart = chunkEnd + 1;
      }
      return ranges;
    }

    function weekBandRow(
      fixedColumnCount: number,
      weekBandRanges: Array<{ startIndex: number; label: string; hasMonthBoundary: boolean }>,
      dateBandLength: number
    ) {
      const weekLabelColumnIndex = layout.columnIndex("S");
      const dividerColumnIndex = layout.columnIndex("T");
      const bandCells = Array.from({ length: dateBandLength }, () => ({} as WbsXlsxCellLike));
      weekBandRanges.forEach((item, index) => {
        bandCells[item.startIndex] = {
          value: item.label,
          bold: true,
          fontSize: 14,
          border: "thin" as const,
          horizontalAlign: "center" as const,
          fillColor: item.hasMonthBoundary ? fills.monthBoundaryWeek : (index % 2 === 0 ? "#EDF4FB" : "#EAF1F9")
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
            return {};
          }),
          ...bandCells
        ]
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

    return {
      dividerCell,
      weekBandRow,
      headerRow,
      weekdayRow,
      dateBandHeaderRow,
      weekdayHeaderRow,
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
      buildWeekBandRanges
    };
  }

  (globalThis as typeof globalThis & {
    __mikuprojectWbsXlsxCells?: {
      createWbsXlsxCellsHelper: (config: WbsXlsxCellsConfig) => WbsXlsxCellsHelper;
    };
  }).__mikuprojectWbsXlsxCells = {
    createWbsXlsxCellsHelper
  };
})();
