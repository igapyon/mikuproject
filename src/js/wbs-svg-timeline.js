/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    function createWbsSvgTimelineHelper(config) {
        function buildTaskRows(tasks, dateBand) {
            return tasks.map((task, index) => ({
                task,
                label: task.name || "-",
                kind: task.summary ? "phase" : (task.milestone ? "milestone" : "task"),
                startIndex: indexOfDate(dateBand, task.start),
                endIndex: indexOfDate(dateBand, task.finish),
                y: index * config.rowHeight
            }));
        }
        function buildWeeklyTaskRows(tasks, weeklyBand) {
            return tasks.map((task, index) => ({
                task,
                label: task.name || "-",
                kind: task.summary ? "phase" : (task.milestone ? "milestone" : "task"),
                startIndex: indexOfWeekForDate(weeklyBand, task.start),
                endIndex: task.milestone ? indexOfWeekForDate(weeklyBand, task.start) : indexOfWeekForDate(weeklyBand, task.finish),
                y: index * config.rowHeight
            }));
        }
        function formatSvgAxisDate(day) {
            const match = day.match(/^(\d{4})-(\d{2})-(\d{2})$/);
            if (!match) {
                return day;
            }
            return `${Number(match[2])}/${Number(match[3])}`;
        }
        function formatWeeklyAxisLabel(day) {
            const match = day.match(/^(\d{4})-(\d{2})-(\d{2})$/);
            if (!match) {
                return day;
            }
            return `${Number(match[2])}/${Number(match[3])}`;
        }
        function indexOfDate(dateBand, value) {
            const key = String(value || "").slice(0, 10);
            if (!key) {
                return null;
            }
            const index = dateBand.indexOf(key);
            return index >= 0 ? index : null;
        }
        function indexOfWeekForDate(weeklyBand, value) {
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
        function startOfWeekSunday(date) {
            const result = new Date(date.getTime());
            result.setDate(result.getDate() - result.getDay());
            return result;
        }
        function endOfWeekSaturday(date) {
            const result = new Date(date.getTime());
            result.setDate(result.getDate() + (6 - result.getDay()));
            return result;
        }
        function formatMonthLabel(date) {
            return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
        }
        function buildWeeklyBand(startDate, finishDate) {
            const start = config.parseDateOnly(startDate);
            const finish = config.parseDateOnly(finishDate);
            if (!start || !finish || start.getTime() > finish.getTime()) {
                return [];
            }
            const cursor = startOfWeekSunday(start);
            const limit = endOfWeekSaturday(finish);
            const weeks = [];
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
        function buildWeeklyMonthSpans(weeklyBand) {
            const spans = [];
            for (let index = 0; index < weeklyBand.length; index += 1) {
                const current = weeklyBand[index];
                const last = spans[spans.length - 1];
                if (last && last.monthKey === current.monthKey) {
                    last.endIndex = index;
                }
                else {
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
    globalThis.__mikuprojectWbsSvgTimeline = {
        createWbsSvgTimelineHelper
    };
})();
