/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  type NativeSvgTaskRow = {
    task: TaskModel;
    label: string;
    kind: "phase" | "task" | "milestone";
    startIndex: number | null;
    endIndex: number | null;
    y: number;
  };

  type WeeklyBand = {
    startDay: string;
    endDay: string;
    monthKey: string;
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

  type WbsSvgTimelineConfig = {
    rowHeight: number;
    parseDateOnly(value: string | undefined): Date | null;
    formatDateOnly(value: Date | null): string;
  };

  function createWbsSvgTimelineHelper(config: WbsSvgTimelineConfig): WbsSvgTimelineHelper {
    function buildTaskRows(tasks: TaskModel[], dateBand: string[]): NativeSvgTaskRow[] {
      return tasks.map((task, index) => ({
        task,
        label: task.name || "-",
        kind: task.summary ? "phase" : (task.milestone ? "milestone" : "task"),
        startIndex: indexOfDate(dateBand, task.start),
        endIndex: indexOfDate(dateBand, task.finish),
        y: index * config.rowHeight
      }));
    }

    function buildWeeklyTaskRows(tasks: TaskModel[], weeklyBand: WeeklyBand[]): NativeSvgTaskRow[] {
      return tasks.map((task, index) => ({
        task,
        label: task.name || "-",
        kind: task.summary ? "phase" : (task.milestone ? "milestone" : "task"),
        startIndex: indexOfWeekForDate(weeklyBand, task.start),
        endIndex: task.milestone ? indexOfWeekForDate(weeklyBand, task.start) : indexOfWeekForDate(weeklyBand, task.finish),
        y: index * config.rowHeight
      }));
    }

    function formatSvgAxisDate(day: string): string {
      const match = day.match(/^(\d{4})-(\d{2})-(\d{2})$/);
      if (!match) {
        return day;
      }
      return `${Number(match[2])}/${Number(match[3])}`;
    }

    function formatWeeklyAxisLabel(day: string): string {
      const match = day.match(/^(\d{4})-(\d{2})-(\d{2})$/);
      if (!match) {
        return day;
      }
      return `${Number(match[2])}/${Number(match[3])}`;
    }

    function indexOfDate(dateBand: string[], value: string | undefined): number | null {
      const key = String(value || "").slice(0, 10);
      if (!key) {
        return null;
      }
      const index = dateBand.indexOf(key);
      return index >= 0 ? index : null;
    }

    function indexOfWeekForDate(weeklyBand: WeeklyBand[], value: string | undefined): number | null {
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

    function buildWeeklyBand(startDate: string | undefined, finishDate: string | undefined): WeeklyBand[] {
      const start = config.parseDateOnly(startDate);
      const finish = config.parseDateOnly(finishDate);
      if (!start || !finish || start.getTime() > finish.getTime()) {
        return [];
      }
      const cursor = startOfWeekSunday(start);
      const limit = endOfWeekSaturday(finish);
      const weeks: WeeklyBand[] = [];
      while (cursor.getTime() <= limit.getTime()) {
        const weekStart = new Date(cursor.getTime());
        const weekEnd = new Date(cursor.getTime());
        weekEnd.setDate(weekEnd.getDate() + 6);
        weeks.push({
          startDay: config.formatDateOnly(weekStart),
          endDay: config.formatDateOnly(weekEnd),
          monthKey: formatMonthLabel(new Date(weekStart.getFullYear(), weekStart.getMonth(), 1))
        });
        cursor.setDate(cursor.getDate() + 7);
      }
      return weeks;
    }

    function buildWeeklyMonthSpans(weeklyBand: WeeklyBand[]): Array<{ monthKey: string; startIndex: number; endIndex: number }> {
      const spans: Array<{ monthKey: string; startIndex: number; endIndex: number }> = [];
      for (let index = 0; index < weeklyBand.length; index += 1) {
        const current = weeklyBand[index];
        const last = spans[spans.length - 1];
        if (last && last.monthKey === current.monthKey) {
          last.endIndex = index;
        } else {
          spans.push({
            monthKey: current.monthKey,
            startIndex: index,
            endIndex: index
          });
        }
      }
      return spans;
    }

    return {
      buildTaskRows,
      buildWeeklyTaskRows,
      formatSvgAxisDate,
      formatWeeklyAxisLabel,
      indexOfDate,
      indexOfWeekForDate,
      buildWeeklyBand,
      buildWeeklyMonthSpans
    };
  }

  (globalThis as typeof globalThis & {
    __mikuprojectWbsSvgTimeline?: {
      createWbsSvgTimelineHelper: (config: WbsSvgTimelineConfig) => WbsSvgTimelineHelper;
    };
  }).__mikuprojectWbsSvgTimeline = {
    createWbsSvgTimelineHelper
  };
})();
