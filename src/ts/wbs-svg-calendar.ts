/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
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

  type WbsSvgCalendarConfig = {
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
  };

  function createWbsSvgCalendarHelper(config: WbsSvgCalendarConfig): WbsSvgCalendarHelper {
    function truncateLabel(value: string): string {
      const text = String(value || "").trim() || "-";
      if (text.length <= config.monthlyMaxLabelChars) {
        return text;
      }
      return `${text.slice(0, config.monthlyMaxLabelChars - 3)}...`;
    }

    function uniqueDayList(days: string[]): string[] {
      return Array.from(new Set(days.filter(Boolean)));
    }

    function buildProjectMonthKeys(startDate: string | undefined, finishDate: string | undefined): string[] {
      const start = config.parseDateOnly(startDate);
      const finish = config.parseDateOnly(finishDate);
      if (!start || !finish) {
        return [];
      }
      const cursor = new Date(start.getFullYear(), start.getMonth(), 1);
      const limit = new Date(finish.getFullYear(), finish.getMonth(), 1);
      const months: string[] = [];
      while (cursor.getTime() <= limit.getTime()) {
        months.push(`${cursor.getFullYear()}-${String(cursor.getMonth() + 1).padStart(2, "0")}`);
        cursor.setMonth(cursor.getMonth() + 1);
      }
      return months;
    }

    function startOfWeekSunday(date: Date): Date {
      const result = new Date(date.getTime());
      result.setDate(result.getDate() - result.getDay());
      return result;
    }

    function endOfWeekSaturday(date: Date): Date {
      const result = new Date(date.getTime());
      result.setDate(result.getDate() + (6 - result.getDay()));
      return result;
    }

    function formatMonthLabel(date: Date): string {
      return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
    }

    function buildMonthlyDayItemMap(
      tasks: TaskModel[],
      calendarDays: string[],
      holidaySet: Set<string>,
      nonWorkingDayTypes: Set<number>
    ): Map<string, Array<{ label: string; className: string }>> {
      const dayItems = new Map<string, Array<{ label: string; className: string }>>();
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
            dayItems.get(startDay)?.push({ label: truncateLabel(task.name || "-"), className: "milestoneText" });
          }
          continue;
        }
        if (task.summary) {
          for (const day of uniqueDayList([startDay, finishDay])) {
            if (calendarDaySet.has(day)) {
              dayItems.get(day)?.push({ label: truncateLabel(task.name || "-"), className: "summaryText" });
            }
          }
          continue;
        }
        for (const day of config.buildDateBand(startDay, finishDay)) {
          if (!calendarDaySet.has(day)) {
            continue;
          }
          const isBoundaryDay = day === startDay || day === finishDay;
          const isHoliday = holidaySet.has(day);
          const isWeekend = config.isWeeklyNonWorkingDay(day, nonWorkingDayTypes);
          if ((isHoliday || isWeekend) && !isBoundaryDay) {
            continue;
          }
          dayItems.get(day)?.push({ label: truncateLabel(task.name || "-"), className: "itemText" });
        }
      }
      return dayItems;
    }

    function exportMonthlyWbsCalendarSvg(
      model: ProjectModel,
      monthKey: string,
      holidaySet: Set<string>,
      nonWorkingDayTypes: Set<number>
    ): string {
      const monthStart = config.parseDateOnly(`${monthKey}-01`);
      if (!monthStart) {
        throw new Error(`Invalid month key: ${monthKey}`);
      }
      const monthEnd = new Date(monthStart.getFullYear(), monthStart.getMonth() + 1, 0);
      const gridStart = startOfWeekSunday(monthStart);
      const gridEnd = endOfWeekSaturday(monthEnd);
      const calendarDays = config.buildDateBand(config.formatDateOnly(gridStart), config.formatDateOnly(gridEnd));
      const weeks: string[][] = [];
      for (let index = 0; index < calendarDays.length; index += 7) {
        weeks.push(calendarDays.slice(index, index + 7));
      }
      const dayItems = buildMonthlyDayItemMap(model.tasks, calendarDays, holidaySet, nonWorkingDayTypes);
      const svgWidth = config.monthlyLeftPadding + (7 * config.monthlyCellWidth) + config.monthlyRightPadding;
      const svgHeight = config.monthlyTopPadding + config.monthlyHeaderHeight + config.monthlyWeekdayHeight + (weeks.length * config.monthlyCellHeight) + config.monthlyBottomPadding;
      const title = `${model.project.name || "-"} ${formatMonthLabel(monthStart)}`;
      const weekdayLabels = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];

      const parts: string[] = [
        `<svg xmlns="http://www.w3.org/2000/svg" width="${svgWidth}" height="${svgHeight}" viewBox="0 0 ${svgWidth} ${svgHeight}" role="img" aria-label="${config.escapeXml(title)}">`,
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
        `<text class="title" x="${config.monthlyLeftPadding}" y="${config.monthlyTopPadding + 28}">${config.escapeXml(model.project.name || "-")}</text>`,
        `<text class="meta" x="${config.monthlyLeftPadding}" y="${config.monthlyTopPadding + 56}">${config.escapeXml(formatMonthLabel(monthStart))}</text>`,
        `<text class="meta" x="${config.monthlyLeftPadding}" y="${config.monthlyTopPadding + 78}">${config.escapeXml(`project range ${String(model.project.startDate || "").slice(0, 10)} - ${String(model.project.finishDate || "").slice(0, 10)}`)}</text>`
      ];

      const weekdayY = config.monthlyTopPadding + config.monthlyHeaderHeight;
      for (let dayIndex = 0; dayIndex < weekdayLabels.length; dayIndex += 1) {
        const x = config.monthlyLeftPadding + (dayIndex * config.monthlyCellWidth);
        parts.push(`<text class="weekday" x="${x + 10}" y="${weekdayY}">${weekdayLabels[dayIndex]}</text>`);
      }

      const gridOriginY = config.monthlyTopPadding + config.monthlyHeaderHeight + config.monthlyWeekdayHeight;
      for (let weekIndex = 0; weekIndex < weeks.length; weekIndex += 1) {
        const week = weeks[weekIndex];
        for (let dayIndex = 0; dayIndex < week.length; dayIndex += 1) {
          const day = week[dayIndex];
          const x = config.monthlyLeftPadding + (dayIndex * config.monthlyCellWidth);
          const y = gridOriginY + (weekIndex * config.monthlyCellHeight);
          const date = config.parseDateOnly(day);
          const inMonth = !!date && date.getMonth() === monthStart.getMonth() && date.getFullYear() === monthStart.getFullYear();
          const isHoliday = holidaySet.has(day);
          const isWeekend = config.isWeeklyNonWorkingDay(day, nonWorkingDayTypes);
          const background = !inMonth
            ? "#f8fafc"
            : (isHoliday ? "#fce7ef" : (isWeekend ? "#eef4f8" : "#ffffff"));
          parts.push(`<rect x="${x}" y="${y}" width="${config.monthlyCellWidth}" height="${config.monthlyCellHeight}" fill="${background}" class="cellBorder"/>`);
          parts.push(`<text class="${inMonth ? "dayNumber" : "dayNumberMuted"}" x="${x + 10}" y="${y + 22}">${config.escapeXml(String(date ? date.getDate() : ""))}</text>`);
          const items = dayItems.get(day) || [];
          const visibleItems = items.slice(0, config.monthlyMaxItemsPerDay);
          for (let itemIndex = 0; itemIndex < visibleItems.length; itemIndex += 1) {
            const item = visibleItems[itemIndex];
            const itemY = y + 42 + (itemIndex * 24);
            parts.push(`<text class="${item.className}" x="${x + 10}" y="${itemY}">${config.escapeXml(item.label)}</text>`);
          }
          if (items.length > config.monthlyMaxItemsPerDay) {
            parts.push(`<text class="moreText" x="${x + 10}" y="${y + config.monthlyCellHeight - 10}">${config.escapeXml(`+${items.length - config.monthlyMaxItemsPerDay} more`)}</text>`);
          }
        }
      }

      parts.push("</svg>");
      return parts.join("");
    }

    function exportMonthlyWbsCalendarSvgArchive(model: ProjectModel): MonthlyCalendarSvgArchive {
      const holidaySet = new Set(config.collectWbsHolidayDates(model));
      const nonWorkingDayTypes = config.collectProjectNonWorkingDayTypes(model);
      const months = buildProjectMonthKeys(model.project.startDate, model.project.finishDate);
      const entries = months.map((monthKey) => ({
        fileName: `${monthKey}.svg`,
        svg: exportMonthlyWbsCalendarSvg(model, monthKey, holidaySet, nonWorkingDayTypes)
      }));
      const zipEntries: ZipEntry[] = entries.map((entry) => ({
        name: `monthly-calendar/${entry.fileName}`,
        data: config.encodeUtf8(entry.svg)
      }));
      const zipBytes = config.packZip(zipEntries);
      return {
        entries,
        zipBytes
      };
    }

    return {
      exportMonthlyWbsCalendarSvgArchive,
      exportMonthlyWbsCalendarSvg
    };
  }

  (globalThis as typeof globalThis & {
    __mikuprojectWbsSvgCalendar?: {
      createWbsSvgCalendarHelper: (config: WbsSvgCalendarConfig) => WbsSvgCalendarHelper;
    };
  }).__mikuprojectWbsSvgCalendar = {
    createWbsSvgCalendarHelper
  };
})();
