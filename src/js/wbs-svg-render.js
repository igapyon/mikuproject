/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    function createWbsSvgRenderHelper(config) {
        const { constants, collectWbsHolidayDates, collectProjectNonWorkingDayTypes, buildDisplayDateBand, buildTaskRows, buildWeeklyTaskRows, buildWeeklyBand, buildWeeklyMonthSpans, formatTaskLabel, resolveLabelPlacement, resolveWeeklyLabelPlacement, shouldUseDailyOnBarLabel, indexOfDate, indexOfWeekForDate, formatSvgAxisDate, formatWeeklyAxisLabel, isWeeklyNonWorkingDay, buildDependencyConnectorElements, escapeXml, bars, viewport, axis, scaffold } = config;
        function exportNativeSvg(model, options = {}) {
            const labelMode = options.labelMode || "near";
            const holidaySet = new Set([
                ...collectWbsHolidayDates(model),
                ...(options.holidayDates || []).map((day) => String(day || "").slice(0, 10)).filter(Boolean)
            ]);
            const nonWorkingDayTypes = collectProjectNonWorkingDayTypes(model);
            const dateBand = buildDisplayDateBand(model.project.startDate, model.project.finishDate, model.project.currentDate, options.displayDaysBeforeBaseDate, options.displayDaysAfterBaseDate, holidaySet, nonWorkingDayTypes, options.useBusinessDaysForDisplayRange);
            const rows = buildTaskRows(model.tasks, dateBand);
            const chartWidth = dateBand.length * constants.dayWidth;
            const leftLabelWidth = labelMode === "list" ? constants.listLabelWidth : constants.nearLeftLabelWidth;
            const rightLabelWidth = labelMode === "list" ? 0 : constants.nearRightLabelWidth;
            const chartOriginXBase = constants.leftPadding + leftLabelWidth;
            const dailyOnBarFlags = rows.map((row) => shouldUseDailyOnBarLabel(row, chartOriginXBase, chartWidth, dateBand.length, labelMode));
            const labelPlacements = rows.map((row, index) => {
                if (dailyOnBarFlags[index]) {
                    return { x: chartOriginXBase, anchor: "start", width: 0 };
                }
                return resolveLabelPlacement(row, chartOriginXBase, chartWidth, labelMode);
            });
            const dailyViewport = viewport.computeDailyViewport({
                labelPlacements,
                chartOriginXBase,
                chartWidth,
                rightLabelWidth,
                rightPadding: constants.rightPadding,
                labelMode
            });
            const shiftX = dailyViewport.shiftX;
            const chartOriginX = dailyViewport.chartOriginX;
            const svgWidth = dailyViewport.svgWidth;
            const svgHeight = constants.topPadding + constants.headerHeight + (rows.length * constants.rowHeight) + constants.bottomPadding;
            const todayIndex = indexOfDate(dateBand, model.project.currentDate);
            const parts = scaffold.createDailyScaffold({
                svgWidth,
                svgHeight,
                chartOriginX,
                chartWidth,
                topPadding: constants.topPadding,
                projectName: model.project.name,
                escapeXml
            });
            const chartOriginY = constants.topPadding + constants.headerHeight;
            parts.push(...axis.renderDailyAxis({
                dateBand,
                chartOriginX,
                svgHeight,
                topPadding: constants.topPadding,
                bottomPadding: constants.bottomPadding,
                dayWidth: constants.dayWidth,
                holidaySet,
                nonWorkingDayTypes,
                isWeeklyNonWorkingDay,
                todayIndex,
                escapeXml,
                formatSvgAxisDate
            }));
            parts.push(...buildDependencyConnectorElements(model.tasks, rows, chartOriginX, chartOriginY, {
                cellWidth: constants.dayWidth,
                rangeInset: 6,
                milestoneHalfWidth: 13,
                routeOffset: 14,
                routeSpacing: 6,
                targetInset: 4,
                cornerRadius: 8
            }));
            for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
                const row = rows[rowIndex];
                const rowY = chartOriginY + row.y;
                const useOnBarLabel = dailyOnBarFlags[rowIndex];
                if (row.startIndex !== null && row.endIndex !== null) {
                    const barX = chartOriginX + (row.startIndex * constants.dayWidth) + 6;
                    const barWidth = Math.max(12, ((row.endIndex - row.startIndex + 1) * constants.dayWidth) - 12);
                    parts.push(...bars.renderTaskBar(row, chartOriginX, rowY, constants.dayWidth, {
                        horizontalInset: 6,
                        verticalInset: 8,
                        milestoneHalfWidth: 13
                    }));
                    if (useOnBarLabel) {
                        const labelText = escapeXml(formatTaskLabel(row.task, labelMode));
                        const labelCenterX = barX + (barWidth / 2);
                        parts.push(`<text class="${row.kind === "phase" ? "phaseLabel" : "label"}" x="${labelCenterX}" y="${rowY + 24}" text-anchor="middle">${labelText}</text>`);
                        continue;
                    }
                }
                const labelPlacement = labelPlacements[rowIndex];
                parts.push(`<text class="${row.kind === "phase" ? "phaseLabel" : "label"}" x="${labelPlacement.x + shiftX}" y="${rowY + 24}" text-anchor="${labelPlacement.anchor}">${escapeXml(formatTaskLabel(row.task, labelMode))}</text>`);
            }
            parts.push("</svg>");
            return parts.join("");
        }
        function exportWeeklyNativeSvg(model, options = {}) {
            const holidaySet = new Set([
                ...collectWbsHolidayDates(model),
                ...(options.holidayDates || []).map((day) => String(day || "").slice(0, 10)).filter(Boolean)
            ]);
            const nonWorkingDayTypes = collectProjectNonWorkingDayTypes(model);
            const dateBand = buildDisplayDateBand(model.project.startDate, model.project.finishDate, model.project.currentDate, options.displayDaysBeforeBaseDate, options.displayDaysAfterBaseDate, holidaySet, nonWorkingDayTypes, options.useBusinessDaysForDisplayRange);
            if (dateBand.length === 0) {
                return exportNativeSvg(model, { ...options, labelMode: "list" });
            }
            const weeklyBand = buildWeeklyBand(dateBand[0], dateBand[dateBand.length - 1]);
            const rows = buildWeeklyTaskRows(model.tasks, weeklyBand);
            const chartWidth = weeklyBand.length * constants.weekWidth;
            const chartOriginXBase = constants.weeklyLeftPadding;
            const labelPlacements = rows.map((row) => resolveWeeklyLabelPlacement(row, chartOriginXBase, chartWidth));
            const weeklyViewport = viewport.computeWeeklyViewport({
                labelPlacements,
                chartOriginXBase,
                chartWidth,
                trimPadding: constants.weeklyTrimPadding,
                rightPadding: constants.weeklyRightPadding
            });
            const shiftX = weeklyViewport.shiftX;
            const chartOriginX = weeklyViewport.chartOriginX;
            const svgWidth = weeklyViewport.svgWidth;
            const svgHeight = constants.weeklyTopPadding + constants.weeklyHeaderHeight + (rows.length * constants.rowHeight) + constants.weeklyBottomPadding;
            const todayWeekIndex = indexOfWeekForDate(weeklyBand, model.project.currentDate);
            const parts = scaffold.createWeeklyScaffold({
                svgWidth,
                svgHeight,
                chartOriginX,
                chartWidth,
                topPadding: constants.weeklyTopPadding,
                projectName: model.project.name,
                projectStartDate: model.project.startDate,
                projectFinishDate: model.project.finishDate,
                escapeXml
            });
            const chartOriginY = constants.weeklyTopPadding + constants.weeklyHeaderHeight;
            const monthSpans = buildWeeklyMonthSpans(weeklyBand);
            parts.push(...axis.renderWeeklyAxis({
                weeklyBand,
                monthSpans,
                chartOriginX,
                svgHeight,
                topPadding: constants.weeklyTopPadding,
                bottomPadding: constants.weeklyBottomPadding,
                weekWidth: constants.weekWidth,
                todayWeekIndex,
                escapeXml,
                formatWeeklyAxisLabel
            }));
            parts.push(...buildDependencyConnectorElements(model.tasks, rows, chartOriginX, chartOriginY, {
                cellWidth: constants.weekWidth,
                rangeInset: 4,
                milestoneHalfWidth: 11,
                routeOffset: 12,
                routeSpacing: 5,
                targetInset: 4,
                cornerRadius: 7
            }));
            for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
                const row = rows[rowIndex];
                const rowY = chartOriginY + row.y;
                if (row.startIndex === null || row.endIndex === null) {
                    const fallbackLabelPlacement = labelPlacements[rowIndex];
                    parts.push(`<text class="${row.kind === "phase" ? "phaseLabel" : "label"}" x="${fallbackLabelPlacement.x + shiftX}" y="${rowY + 24}" text-anchor="${fallbackLabelPlacement.anchor}">${escapeXml(formatTaskLabel(row.task, "near"))}</text>`);
                    continue;
                }
                parts.push(...bars.renderTaskBar(row, chartOriginX, rowY, constants.weekWidth, {
                    horizontalInset: 4,
                    verticalInset: 8,
                    milestoneHalfWidth: 11
                }));
                const labelPlacement = labelPlacements[rowIndex];
                parts.push(`<text class="${row.kind === "phase" ? "phaseLabel" : "label"}" x="${labelPlacement.x + shiftX}" y="${rowY + 24}" text-anchor="${labelPlacement.anchor}">${escapeXml(formatTaskLabel(row.task, "near"))}</text>`);
            }
            parts.push("</svg>");
            return parts.join("");
        }
        return {
            exportNativeSvg,
            exportWeeklyNativeSvg
        };
    }
    globalThis.__mikuprojectWbsSvgRender = {
        createWbsSvgRenderHelper
    };
})();
