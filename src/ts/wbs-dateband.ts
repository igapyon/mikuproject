/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  type WbsDatebandHelper = {
    parseDateOnly(value: string | undefined): Date | null;
    formatDateOnly(value: Date | null): string;
    buildDateBand(startDate: string | undefined, finishDate: string | undefined): string[];
    expandExceptionDays(exception: CalendarExceptionModel): string[];
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

  function formatDateOnly(value: Date | null): string {
    if (!value) {
      return "";
    }
    return [
      value.getFullYear(),
      String(value.getMonth() + 1).padStart(2, "0"),
      String(value.getDate()).padStart(2, "0")
    ].join("-");
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

  function normalizeDisplayDayCount(value: number | undefined): number | null {
    if (value === undefined || value === null || !Number.isFinite(value)) {
      return null;
    }
    return Math.max(0, Math.floor(value));
  }

  function shiftCalendarDays(base: Date, offset: number): Date {
    const result = new Date(base.getTime());
    result.setDate(result.getDate() + offset);
    return result;
  }

  function isWeeklyNonWorkingDay(day: string, nonWorkingDayTypes: Set<number>): boolean {
    const target = parseDateOnly(day);
    if (!target) {
      return false;
    }
    const dayType = target.getDay() === 0 ? 1 : target.getDay() + 1;
    return nonWorkingDayTypes.has(dayType);
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

  function createWbsDatebandHelper(): WbsDatebandHelper {
    return {
      parseDateOnly,
      formatDateOnly,
      buildDateBand,
      expandExceptionDays,
      collectWbsHolidayDates,
      collectProjectNonWorkingDayTypes,
      buildDisplayDateBand,
      isWeeklyNonWorkingDay
    };
  }

  (globalThis as typeof globalThis & {
    __mikuprojectWbsDateband?: {
      createWbsDatebandHelper: () => WbsDatebandHelper;
    };
  }).__mikuprojectWbsDateband = {
    createWbsDatebandHelper
  };
})();
