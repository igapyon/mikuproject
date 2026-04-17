/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    function createWbsXlsxSectionsHelper(config) {
        const { layout, fills } = config;
        function stringifyCellValue(value) {
            return typeof value === "string" ? value : String(value);
        }
        function blockHeaderRow(columnCount, startColumnIndex, title) {
            const cells = Array.from({ length: columnCount }, () => ({}));
            cells[startColumnIndex] = {
                value: title,
                border: "thin",
                horizontalAlign: "left",
                bold: true,
                fontSize: 14,
                fillColor: fills.headerId
            };
            cells[startColumnIndex + 1] = {
                value: "",
                border: "thin",
                fillColor: fills.headerId
            };
            cells[startColumnIndex + 2] = {
                value: "",
                border: "thin",
                fillColor: fills.headerId
            };
            return { height: 24, cells };
        }
        function projectBlockHeaderRow(columnCount, startColumnIndex, title) {
            const cells = Array.from({ length: columnCount }, () => ({}));
            cells[startColumnIndex] = {
                value: title,
                border: "thin",
                horizontalAlign: "left",
                bold: true,
                fontSize: 14,
                fillColor: fills.headerId
            };
            for (let offset = 1; offset < 5; offset += 1) {
                cells[startColumnIndex + offset] = {
                    value: "",
                    border: "thin",
                    fillColor: fills.headerId
                };
            }
            return { height: 24, cells };
        }
        function summaryStatCell(value, fillColor, isValueCell) {
            const valueAlign = typeof value === "number" ? "center" : "left";
            return {
                value: stringifyCellValue(value),
                border: "thin",
                horizontalAlign: isValueCell ? valueAlign : "right",
                bold: true,
                fillColor
            };
        }
        function projectPairRow(columnCount, startColumnIndex, label, value, fillColor) {
            const cells = Array.from({ length: columnCount }, () => ({}));
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
                value: stringifyCellValue(value),
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
            cells[startColumnIndex + 4] = {
                value: "",
                border: "thin",
                fillColor
            };
            return { height: 22, cells };
        }
        function summaryPairRow(columnCount, startColumnIndex, label, value, fillColor) {
            const cells = Array.from({ length: columnCount }, () => ({}));
            cells[startColumnIndex] = summaryStatCell(label, fillColor, false);
            cells[startColumnIndex + 1] = summaryStatCell(value, fillColor, true);
            cells[startColumnIndex + 2] = {
                value: "",
                border: "thin",
                fillColor
            };
            return { height: 22, cells };
        }
        function mergedLabelRow(columnCount, startColumnIndex, value, fillColor) {
            const cells = Array.from({ length: columnCount }, () => ({}));
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
            cells[startColumnIndex + 2] = {
                value: "",
                border: "thin",
                fillColor
            };
            return { height: 24, cells };
        }
        function displayReferenceValue(value) {
            return value && value.trim() ? value : "-";
        }
        function truncateWbsReference(value, maxLength) {
            const normalized = (value === null || value === void 0 ? void 0 : value.trim()) || "";
            if (!normalized) {
                return "";
            }
            if (normalized.length <= maxLength) {
                return normalized;
            }
            return `${normalized.slice(0, Math.max(1, maxLength - 3))}...`;
        }
        function formatCalendarLabel(calendarUID, calendarNameByUid) {
            if (!calendarUID) {
                return "-";
            }
            const calendarName = calendarNameByUid.get(calendarUID);
            return calendarName ? `${calendarUID} ${truncateWbsReference(calendarName, 9)}` : calendarUID;
        }
        function firstResourceName(resourceNames) {
            if (!resourceNames || resourceNames.length === 0) {
                return "";
            }
            return resourceNames[0];
        }
        function referenceCell(task, value, horizontalAlign = "center") {
            const displayValue = displayReferenceValue(value);
            const placeholder = displayValue === "-";
            return {
                value: displayValue,
                border: "thin",
                horizontalAlign: placeholder ? "center" : horizontalAlign,
                verticalAlign: "center",
                bold: task.summary || task.milestone || false,
                fillColor: placeholder
                    ? fills.placeholder
                    : (task.summary
                        ? fills.phase
                        : (task.milestone ? fills.milestone : fills.referenceColumn))
            };
        }
        function projectInfoRows(project, calendarNameByUid, holidayCount, columnCount, startColumnIndex, startRowNumber) {
            const items = [
                { label: "プロジェクト名", value: truncateWbsReference(project.name || "-", 18) || "-", fillColor: fills.summaryAssignment },
                { label: "カレンダ", value: formatCalendarLabel(project.calendarUID, calendarNameByUid), fillColor: fills.summaryAssignment },
                { label: "開始日", value: formatWbsDate(project.startDate), fillColor: fills.summarySchedule },
                { label: "終了日", value: formatWbsDate(project.finishDate), fillColor: fills.summarySchedule },
                { label: "現在日", value: formatWbsDate(project.currentDate), fillColor: fills.summarySchedule },
                { label: "祝日", value: holidayCount, fillColor: fills.summarySchedule }
            ];
            return {
                mergedRanges: [
                    layout.range(layout.reference(startRowNumber, startColumnIndex), layout.reference(startRowNumber, startColumnIndex + 4)),
                    ...items.map((_, index) => {
                        const rowNumber = startRowNumber + index + 1;
                        return [
                            layout.range(layout.reference(rowNumber, startColumnIndex), layout.reference(rowNumber, startColumnIndex + 1)),
                            layout.range(layout.reference(rowNumber, startColumnIndex + 2), layout.reference(rowNumber, startColumnIndex + 4))
                        ];
                    }).flat()
                ],
                rows: [
                    projectBlockHeaderRow(columnCount, startColumnIndex, "プロジェクト情報"),
                    ...items.map((item) => projectPairRow(columnCount, startColumnIndex, item.label, item.value, item.fillColor))
                ]
            };
        }
        function legendRows(columnCount, startRowNumber) {
            const items = [
                { value: "進捗済み", fillColor: "#8EB9EA" },
                { value: "予定帯", fillColor: "#D9EFFF" },
                { value: "当日", fillColor: "#FFE6A7" },
                { value: "週頭", fillColor: "#E3EEF9" },
                { value: "週末", fillColor: "#EEF3F8" },
                { value: "祝日", fillColor: "#FCE4EC" },
                { value: "━:フェーズ", fillColor: fills.phase },
                { value: "■:進捗済みタスク", fillColor: "#8EB9EA" },
                { value: "□:予定タスク", fillColor: "#D9EFFF" },
                { value: "◆:マイルストーン", fillColor: fills.milestone },
                { value: "Mil:マイルストーン", fillColor: "#FBE4EC" },
                { value: "Sum:サマリ", fillColor: "#F7EAF0" },
                { value: "Crit:クリティカル", fillColor: "#F3E1E9" },
                { value: "-:未設定", fillColor: fills.placeholder }
            ];
            const startColumnRef = layout.reference(startRowNumber, layout.columnIndex("A"));
            const endColumnRef = layout.reference(startRowNumber, layout.columnIndex("C"));
            return {
                mergedRanges: [
                    layout.range(startColumnRef, endColumnRef),
                    ...items.map((_, index) => layout.range(layout.reference(startRowNumber + index + 1, layout.columnIndex("A")), layout.reference(startRowNumber + index + 1, layout.columnIndex("C"))))
                ],
                rows: [
                    blockHeaderRow(columnCount, 0, "凡例"),
                    ...items.map((item) => mergedLabelRow(columnCount, 0, item.value, item.fillColor))
                ]
            };
        }
        function displaySummaryRows(displayDays, businessDays, baseDate, taskCount, resourceCount, assignmentCount, calendarCount, columnCount, startColumnIndex = 5, startRowNumber = 5, displayDaysBeforeBaseDate, displayDaysAfterBaseDate, useBusinessDaysForDisplayRange, useBusinessDaysForProgressBand) {
            const displayWeeks = displayDays > 0 ? Math.ceil(displayDays / 7) : 0;
            const scheduleItems = [
                { label: "表示日", value: displayDays, fillColor: fills.summarySchedule },
                { label: "表示週", value: displayWeeks, fillColor: fills.summarySchedule },
                { label: "営業日", value: businessDays, fillColor: fills.summarySchedule },
                { label: "前日数", value: displayDaysBeforeBaseDate !== null && displayDaysBeforeBaseDate !== void 0 ? displayDaysBeforeBaseDate : "-", fillColor: fills.summarySchedule },
                { label: "後日数", value: displayDaysAfterBaseDate !== null && displayDaysAfterBaseDate !== void 0 ? displayDaysAfterBaseDate : "-", fillColor: fills.summarySchedule },
                { label: "表示", value: useBusinessDaysForDisplayRange ? "営業日" : "暦日", fillColor: fills.summarySchedule },
                { label: "進捗", value: useBusinessDaysForProgressBand ? "営業日" : "暦日", fillColor: fills.summarySchedule },
                { label: "基準日", value: (baseDate || "-").slice(0, 10), fillColor: fills.summarySchedule }
            ];
            const countItems = [
                { label: "タスク", value: taskCount, fillColor: fills.summaryAssignment },
                { label: "リソース", value: resourceCount, fillColor: fills.summaryAssignment },
                { label: "割当", value: assignmentCount, fillColor: fills.summaryAssignment },
                { label: "カレンダ", value: calendarCount, fillColor: fills.summaryAssignment }
            ];
            const blockRows = [blockHeaderRow(columnCount, startColumnIndex, "サマリ")];
            for (const item of scheduleItems) {
                blockRows.push(summaryPairRow(columnCount, startColumnIndex, item.label, item.value, item.fillColor));
            }
            for (const item of countItems) {
                blockRows.push(summaryPairRow(columnCount, startColumnIndex, item.label, item.value, item.fillColor));
            }
            const mergedRanges = [
                layout.range(layout.reference(startRowNumber, startColumnIndex), layout.reference(startRowNumber, startColumnIndex + 2))
            ];
            for (let index = 1; index < blockRows.length; index += 1) {
                const rowNumber = startRowNumber + index;
                mergedRanges.push(layout.range(layout.reference(rowNumber, startColumnIndex + 1), layout.reference(rowNumber, startColumnIndex + 2)));
            }
            return {
                mergedRanges,
                rows: blockRows
            };
        }
        function formatWbsDate(value) {
            return value ? value.slice(0, 10) : "-";
        }
        return {
            firstResourceName,
            formatCalendarLabel,
            truncateWbsReference,
            referenceCell,
            projectInfoRows,
            legendRows,
            displaySummaryRows,
            formatWbsDate
        };
    }
    globalThis.__mikuprojectWbsXlsxSections = {
        createWbsXlsxSectionsHelper
    };
})();
