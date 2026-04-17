/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    function createWbsSvgAxisHelper() {
        function renderDailyAxis(config) {
            const parts = [];
            const { dateBand, chartOriginX, svgHeight, topPadding, bottomPadding, dayWidth, holidaySet, nonWorkingDayTypes, isWeeklyNonWorkingDay, todayIndex, escapeXml, formatSvgAxisDate } = config;
            for (let index = 0; index < dateBand.length; index += 1) {
                const day = dateBand[index];
                const x = chartOriginX + (index * dayWidth);
                const isHoliday = holidaySet.has(day);
                const isWeekend = isWeeklyNonWorkingDay(day, nonWorkingDayTypes);
                const fill = isHoliday ? "#fce4ec" : (isWeekend ? "#eef3f8" : "#ffffff");
                parts.push(`<rect x="${x}" y="${topPadding + 30}" width="${dayWidth}" height="${svgHeight - topPadding - bottomPadding - 8}" fill="${fill}"/>`);
                parts.push(`<line class="grid" x1="${x}" y1="${topPadding + 26}" x2="${x}" y2="${svgHeight - bottomPadding}" />`);
                parts.push(`<text class="axis" x="${x + (dayWidth / 2)}" y="${topPadding + 54}" text-anchor="middle">${escapeXml(formatSvgAxisDate(day))}</text>`);
            }
            parts.push(`<line class="grid" x1="${chartOriginX + (dateBand.length * dayWidth)}" y1="${topPadding + 26}" x2="${chartOriginX + (dateBand.length * dayWidth)}" y2="${svgHeight - bottomPadding}" />`);
            if (todayIndex !== null && todayIndex >= 0) {
                const todayX = chartOriginX + (todayIndex * dayWidth) + (dayWidth / 2);
                parts.push(`<line class="today" x1="${todayX}" y1="${topPadding + 26}" x2="${todayX}" y2="${svgHeight - bottomPadding}" />`);
            }
            return parts;
        }
        function renderWeeklyAxis(config) {
            const parts = [];
            const { weeklyBand, monthSpans, chartOriginX, svgHeight, topPadding, bottomPadding, weekWidth, todayWeekIndex, escapeXml, formatWeeklyAxisLabel } = config;
            for (const span of monthSpans) {
                const x = chartOriginX + (span.startIndex * weekWidth);
                const width = (span.endIndex - span.startIndex + 1) * weekWidth;
                parts.push(`<text class="monthAxis" x="${x + (width / 2)}" y="${topPadding + 66}" text-anchor="middle">${escapeXml(span.monthKey)}</text>`);
                parts.push(`<line class="monthBoundary" x1="${x}" y1="${topPadding + 72}" x2="${x}" y2="${svgHeight - bottomPadding}" />`);
            }
            parts.push(`<line class="monthBoundary" x1="${chartOriginX + (weeklyBand.length * weekWidth)}" y1="${topPadding + 72}" x2="${chartOriginX + (weeklyBand.length * weekWidth)}" y2="${svgHeight - bottomPadding}" />`);
            for (let index = 0; index < weeklyBand.length; index += 1) {
                const week = weeklyBand[index];
                const x = chartOriginX + (index * weekWidth);
                parts.push(`<line class="grid" x1="${x}" y1="${topPadding + 72}" x2="${x}" y2="${svgHeight - bottomPadding}" />`);
                parts.push(`<text class="weekAxis" x="${x + (weekWidth / 2)}" y="${topPadding + 86}" text-anchor="middle">${escapeXml(formatWeeklyAxisLabel(week.startDay))}</text>`);
            }
            if (todayWeekIndex !== null && todayWeekIndex >= 0) {
                const todayX = chartOriginX + (todayWeekIndex * weekWidth) + (weekWidth / 2);
                parts.push(`<line class="today" x1="${todayX}" y1="${topPadding + 72}" x2="${todayX}" y2="${svgHeight - bottomPadding}" />`);
            }
            return parts;
        }
        return {
            renderDailyAxis,
            renderWeeklyAxis
        };
    }
    globalThis.__mikuprojectWbsSvgAxis = {
        createWbsSvgAxisHelper
    };
})();
