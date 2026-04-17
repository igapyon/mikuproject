/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    function parseDateOnly(value) {
        const text = String(value || "").trim().slice(0, 10);
        if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) {
            return null;
        }
        const parsed = new Date(`${text}T00:00:00`);
        return Number.isNaN(parsed.getTime()) ? null : parsed;
    }
    function formatDateOnly(value) {
        const year = value.getFullYear();
        const month = String(value.getMonth() + 1).padStart(2, "0");
        const day = String(value.getDate()).padStart(2, "0");
        return `${year}-${month}-${day}`;
    }
    function addDateDays(base, days) {
        const next = new Date(base.getTime());
        next.setDate(next.getDate() + days);
        return next;
    }
    function buildNthWeekdayOfMonth(year, monthIndex, jsWeekday, nth) {
        const first = new Date(year, monthIndex, 1);
        const offset = (jsWeekday - first.getDay() + 7) % 7;
        return new Date(year, monthIndex, 1 + offset + ((nth - 1) * 7));
    }
    function calculateVernalEquinoxDay(year) {
        return Math.floor(20.8431 + (0.242194 * (year - 1980)) - Math.floor((year - 1980) / 4));
    }
    function calculateAutumnalEquinoxDay(year) {
        return Math.floor(23.2488 + (0.242194 * (year - 1980)) - Math.floor((year - 1980) / 4));
    }
    function buildJapaneseHolidayMapForYear(year) {
        const holidays = new Map();
        const addHoliday = (date, name) => {
            holidays.set(formatDateOnly(date), name);
        };
        addHoliday(new Date(year, 0, 1), "元日");
        addHoliday(buildNthWeekdayOfMonth(year, 0, 1, 2), "成人の日");
        addHoliday(new Date(year, 1, 11), "建国記念の日");
        addHoliday(new Date(year, 1, 23), "天皇誕生日");
        addHoliday(new Date(year, 2, calculateVernalEquinoxDay(year)), "春分の日");
        addHoliday(new Date(year, 3, 29), "昭和の日");
        addHoliday(new Date(year, 4, 3), "憲法記念日");
        addHoliday(new Date(year, 4, 4), "みどりの日");
        addHoliday(new Date(year, 4, 5), "こどもの日");
        addHoliday(buildNthWeekdayOfMonth(year, 6, 1, 3), "海の日");
        addHoliday(new Date(year, 7, 11), "山の日");
        addHoliday(buildNthWeekdayOfMonth(year, 8, 1, 3), "敬老の日");
        addHoliday(new Date(year, 8, calculateAutumnalEquinoxDay(year)), "秋分の日");
        addHoliday(buildNthWeekdayOfMonth(year, 9, 1, 2), "スポーツの日");
        addHoliday(new Date(year, 10, 3), "文化の日");
        addHoliday(new Date(year, 10, 23), "勤労感謝の日");
        const baseHolidayDates = Array.from(holidays.keys()).sort();
        for (const dateText of baseHolidayDates) {
            const date = parseDateOnly(dateText);
            if (!date || date.getDay() !== 0) {
                continue;
            }
            let substitute = addDateDays(date, 1);
            while (holidays.has(formatDateOnly(substitute))) {
                substitute = addDateDays(substitute, 1);
            }
            holidays.set(formatDateOnly(substitute), "休日");
        }
        const sortedDates = Array.from(holidays.keys()).sort();
        for (let index = 0; index < sortedDates.length - 1; index += 1) {
            const current = parseDateOnly(sortedDates[index]);
            const next = parseDateOnly(sortedDates[index + 1]);
            if (!current || !next) {
                continue;
            }
            const gapDays = Math.floor((next.getTime() - current.getTime()) / 86400000);
            if (gapDays !== 2) {
                continue;
            }
            const between = addDateDays(current, 1);
            const betweenText = formatDateOnly(between);
            if (holidays.has(betweenText) || between.getDay() === 0) {
                continue;
            }
            holidays.set(betweenText, "休日");
        }
        return new Map(Array.from(holidays.entries()).sort((left, right) => left[0].localeCompare(right[0])));
    }
    function buildDefaultWorkingTimes(project) {
        const start = project.defaultStartTime || "09:00:00";
        const finish = project.defaultFinishTime || "18:00:00";
        if (start < "12:00:00" && finish > "13:00:00") {
            return [
                { fromTime: start, toTime: "12:00:00" },
                { fromTime: "13:00:00", toTime: finish }
            ];
        }
        return [{ fromTime: start, toTime: finish }];
    }
    function buildDefaultStandardWeekDays(project) {
        const workingTimes = buildDefaultWorkingTimes(project);
        return Array.from({ length: 7 }, (_, index) => {
            const dayType = index + 1;
            const dayWorking = dayType !== 1 && dayType !== 7;
            return {
                dayType,
                dayWorking,
                workingTimes: dayWorking ? workingTimes.map((item) => ({ ...item })) : []
            };
        });
    }
    function buildDefaultJapaneseHolidayExceptions(project) {
        const start = parseDateOnly(project.startDate) || parseDateOnly(project.finishDate) || new Date();
        const finishLimit = parseDateOnly(project.finishDate) || start;
        const rangeStart = start.getTime() <= finishLimit.getTime() ? start : finishLimit;
        const rangeFinish = start.getTime() <= finishLimit.getTime() ? finishLimit : start;
        const exceptions = [];
        for (let year = rangeStart.getFullYear(); year <= rangeFinish.getFullYear(); year += 1) {
            const holidays = buildJapaneseHolidayMapForYear(year);
            for (const [dateText, name] of holidays.entries()) {
                const date = parseDateOnly(dateText);
                if (!date || date.getTime() < rangeStart.getTime() || date.getTime() > rangeFinish.getTime()) {
                    continue;
                }
                exceptions.push({
                    name,
                    fromDate: `${dateText}T00:00:00`,
                    toDate: `${dateText}T23:59:59`,
                    dayWorking: false,
                    workingTimes: []
                });
            }
        }
        return exceptions;
    }
    function allocateDefaultCalendarUid(model) {
        const usedUids = new Set(model.calendars.map((calendar) => String(calendar.uid)));
        let candidate = 1;
        while (usedUids.has(String(candidate))) {
            candidate += 1;
        }
        return String(candidate);
    }
    globalThis.__mikuprojectMsprojectCalendar = {
        ensureDefaultProjectCalendar(model) {
            if (model.calendars.length === 0) {
                const uid = allocateDefaultCalendarUid(model);
                model.calendars.push({
                    uid,
                    name: "Standard",
                    isBaseCalendar: true,
                    isBaselineCalendar: true,
                    weekDays: buildDefaultStandardWeekDays(model.project),
                    exceptions: buildDefaultJapaneseHolidayExceptions(model.project),
                    workWeeks: []
                });
                model.project.calendarUID = uid;
            }
            return model;
        }
    };
})();
