(() => {
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
    function collectWbsHolidayDates(model) {
        const holidaySet = new Set();
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
    function exportWbsWorkbook(model, options = {}) {
        const resourceNameByUid = new Map(model.resources.map((resource) => [resource.uid, resource.name]));
        const predecessorNameByUid = new Map(model.tasks.map((task) => [task.uid, task.name]));
        const calendarNameByUid = new Map(model.calendars.map((calendar) => [calendar.uid, calendar.name]));
        const resourceNamesByTaskUid = new Map();
        const holidaySet = new Set((options.holidayDates || []).map((day) => day.slice(0, 10)));
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
        const dateBand = buildDisplayDateBand(model.project.startDate, model.project.finishDate, model.project.currentDate, options.displayDaysBeforeBaseDate, options.displayDaysAfterBaseDate, holidaySet, options.useBusinessDaysForDisplayRange);
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
        const dateBandStartColumnIndex = dividerColumnIndex + 1;
        const totalColumns = fixedHeaders.length + 1 + dateBand.length;
        const lastColumnRef = columnName(totalColumns);
        const rows = [
            sheetTitleRow("WBS", totalColumns),
            sheetSubtitleRow(model.project.name || "Project", totalColumns)
        ];
        const mergedRanges = [
            `A1:${lastColumnRef}1`,
            `A2:${lastColumnRef}2`
        ];
        const projectInfoBlock = projectInfoRows(model.project, calendarNameByUid, holidaySet.size, totalColumns, 0, rows.length + 1);
        overlayRows(rows, 2, projectInfoBlock.rows, totalColumns);
        mergedRanges.push(...projectInfoBlock.mergedRanges);
        const summaryBlock = displaySummaryRows(dateBand.length, countBusinessDays(dateBand, holidaySet), model.project.currentDate, model.tasks.length, model.resources.length, model.assignments.length, model.calendars.length, totalColumns, 5, 3, options.displayDaysBeforeBaseDate, options.displayDaysAfterBaseDate, options.useBusinessDaysForDisplayRange, options.useBusinessDaysForProgressBand);
        overlayRows(rows, 2, summaryBlock.rows, totalColumns);
        mergedRanges.push(...summaryBlock.mergedRanges);
        rows.push(emptyRow(totalColumns, 28));
        rows.push(taskViewRow((model.project.currentDate || "-").slice(0, 10) || "-", totalColumns));
        const weekBandRanges = buildWeekBandRanges(dateBand, dateBandStartColumnIndex, rows.length + 1);
        rows.push(weekBandRow(fixedHeaders.length + 1, weekBandRanges, dateBand.length));
        rows.push(todayGuideRow(fixedHeaders.length + 1, dateBand, model.project.currentDate, holidaySet));
        rows.push(headerRow([
            ...fixedHeaders.map((label) => label === "名称"
                ? {
                    value: label,
                    bold: true,
                    fillColor: headerFillForLabel(label),
                    border: "thin",
                    horizontalAlign: "left"
                }
                : label),
            dividerCell(),
            ...dateBand.map((day) => dateNumberCell(day, model.project.currentDate, holidaySet))
        ]));
        rows.push(weekdayRow(fixedHeaders.length + 1, dateBand, model.project.currentDate, holidaySet));
        rows.push(...model.tasks.map((task) => ({
            height: taskRowHeight(task),
            cells: [
                identifierCell(task, task.uid),
                identifierCell(task, task.id),
                identifierCell(task, task.wbs || task.outlineNumber),
                kindCell(task),
                identifierCell(task, task.outlineLevel),
                taskCell(task, formatTaskLabel(task)),
                taskCell(task, formatWbsDate(task.start), "center"),
                taskCell(task, formatWbsDate(task.finish), "center"),
                taskCell(task, formatDurationLabel(task, holidaySet, options.useBusinessDaysForProgressBand), "center"),
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
                ...dateBand.map((day) => dateBandCell(task, day, model.project.currentDate, holidaySet, options.useBusinessDaysForProgressBand))
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
                        { width: 20 }, { width: 18 }, { width: 12 }, { width: 14 },
                        { width: 18 }, { width: 12 }, { width: 12 }, { width: 12 },
                        { width: 16 }, { width: 12 }, { width: 20 }, { width: 18 }, { width: 3 },
                        ...dateBand.map(() => ({ width: 6 }))
                    ],
                    mergedRanges,
                    rows
                }
            ]
        };
    }
    function emptyRow(columnCount, height = 22) {
        return {
            height,
            cells: Array.from({ length: columnCount }, () => ({}))
        };
    }
    function overlayRows(rows, startIndex, blockRows, columnCount) {
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
    function hasCellContent(cell) {
        return !!cell && Object.keys(cell).length > 0;
    }
    function formatTaskLabel(task) {
        const prefix = task.summary ? "> " : (task.milestone ? "* " : "- ");
        return `${"  ".repeat(Math.max(0, task.outlineLevel - 1))}${prefix}${task.name}`;
    }
    function classifyTaskKind(task) {
        if (task.summary) {
            return "フェーズ";
        }
        if (task.milestone) {
            return "マイル";
        }
        return "タスク";
    }
    function firstResourceName(resourceNames) {
        if (!resourceNames || resourceNames.length === 0) {
            return "";
        }
        return resourceNames[0];
    }
    function formatCalendarLabel(calendarUID, calendarNameByUid) {
        if (!calendarUID) {
            return "-";
        }
        const calendarName = calendarNameByUid.get(calendarUID);
        return calendarName ? `${calendarUID} ${truncateWbsReference(calendarName, 9)}` : calendarUID;
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
    function referenceCell(task, value, horizontalAlign = "left") {
        const displayValue = displayReferenceValue(value);
        const placeholder = displayValue === "-";
        return {
            value: displayValue,
            border: "thin",
            horizontalAlign: placeholder ? "center" : horizontalAlign,
            bold: task.summary || task.milestone || false,
            fillColor: placeholder
                ? PLACEHOLDER_FILL
                : (task.summary
                    ? PHASE_FILL
                    : (task.milestone ? MILESTONE_FILL : REFERENCE_COLUMN_FILL))
        };
    }
    function sheetTitleRow(title, columnCount) {
        return {
            height: 24,
            cells: [
                {
                    value: title,
                    bold: true,
                    fillColor: "#EEF4FA",
                    border: "thin",
                    horizontalAlign: "left"
                },
                ...Array.from({ length: Math.max(0, columnCount - 1) }, () => ({}))
            ]
        };
    }
    function sheetSubtitleRow(title, columnCount) {
        return {
            height: 22,
            cells: [
                {
                    value: title,
                    bold: true,
                    fillColor: "#F6F9FC",
                    border: "thin",
                    horizontalAlign: "left"
                },
                ...Array.from({ length: Math.max(0, columnCount - 1) }, () => ({}))
            ]
        };
    }
    function taskViewRow(baseDate, columnCount) {
        return {
            height: 26,
            cells: Array.from({ length: columnCount }, (_, index) => {
                if (index === 5) {
                    return {
                        value: "タスク表",
                        bold: true,
                        fillColor: "#E6F1FB",
                        border: "thin",
                        horizontalAlign: "center"
                    };
                }
                if (index === 6) {
                    return {
                        value: `基準日 ${baseDate}`,
                        bold: true,
                        fillColor: "#E6F1FB",
                        border: "thin",
                        horizontalAlign: "center"
                    };
                }
                if (index > 5 && index < 12) {
                    return {
                        value: "",
                        fillColor: "#E6F1FB",
                        border: "thin"
                    };
                }
                return {};
            })
        };
    }
    function infoRow(text, columnCount) {
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
    function projectInfoRows(project, calendarNameByUid, holidayCount, columnCount, startColumnIndex, startRowNumber) {
        const scheduleMode = project.scheduleFromStart ? "開始基準" : "終了基準";
        const items = [
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
                `${columnName(startColumnIndex + 1)}${startRowNumber}:${columnName(startColumnIndex + 4)}${startRowNumber}`,
                ...items.map((_, index) => {
                    const rowNumber = startRowNumber + index + 1;
                    return [
                        `${columnName(startColumnIndex + 1)}${rowNumber}:${columnName(startColumnIndex + 2)}${rowNumber}`,
                        `${columnName(startColumnIndex + 3)}${rowNumber}:${columnName(startColumnIndex + 4)}${rowNumber}`
                    ];
                }).flat()
            ],
            rows: [
                projectBlockHeaderRow(columnCount, startColumnIndex, "プロジェクト"),
                ...items.map((item) => projectPairRow(columnCount, startColumnIndex, item.label, item.value, item.fillColor))
            ]
        };
    }
    function legendRows(columnCount, startRowNumber) {
        const items = [
            { value: "進捗済み", fillColor: PROGRESS_BAND_FILL },
            { value: "予定帯", fillColor: ACTIVE_BAND_FILL },
            { value: "当日", fillColor: TODAY_BAND_FILL },
            { value: "週頭", fillColor: WEEK_START_BAND_FILL },
            { value: "週末", fillColor: WEEKEND_BAND_FILL },
            { value: "祝日", fillColor: HOLIDAY_BAND_FILL },
            { value: "━:フェーズ", fillColor: PHASE_FILL },
            { value: "■:タスク", fillColor: ACTIVE_BAND_FILL },
            { value: "◆:マイルストーン", fillColor: MILESTONE_FILL },
            { value: "Mil:マイルストーン", fillColor: "#FBE4EC" },
            { value: "Sum:サマリ", fillColor: "#F7EAF0" },
            { value: "Crit:クリティカル", fillColor: "#F3E1E9" },
            { value: "-:未設定", fillColor: PLACEHOLDER_FILL }
        ];
        const startColumnRef = columnName(6);
        const endColumnRef = columnName(7);
        return {
            mergedRanges: [
                `${startColumnRef}${startRowNumber}:${endColumnRef}${startRowNumber}`,
                ...items.map((_, index) => `${startColumnRef}${startRowNumber + index + 1}:${endColumnRef}${startRowNumber + index + 1}`)
            ],
            rows: [
                blockHeaderRow(columnCount, 5, "凡例"),
                ...items.map((item) => mergedLabelRow(columnCount, 5, item.value, item.fillColor))
            ]
        };
    }
    function weekBandRow(fixedColumnCount, weekBandRanges, dateBandLength) {
        const bandCells = Array.from({ length: dateBandLength }, () => ({}));
        weekBandRanges.forEach((item, index) => {
            bandCells[item.startIndex] = {
                value: item.label,
                bold: true,
                border: "thin",
                horizontalAlign: "center",
                fillColor: item.hasMonthBoundary ? MONTH_BOUNDARY_WEEK_FILL : (index % 2 === 0 ? "#EDF4FB" : "#EAF1F9")
            };
        });
        return {
            height: 24,
            cells: [
                ...Array.from({ length: fixedColumnCount }, (_, index) => {
                    if (index === 5) {
                        return {
                            value: "週",
                            bold: true,
                            border: "thin",
                            horizontalAlign: "right",
                            fillColor: "#E3EEF9"
                        };
                    }
                    if (index === 6) {
                        return {
                            value: "",
                            border: "thin",
                            fillColor: "#E3EEF9"
                        };
                    }
                    if (index > 5 && index < 9) {
                        return {
                            value: "",
                            border: "thin",
                            fillColor: "#E3EEF9"
                        };
                    }
                    if (index === 5 || index === 9) {
                        return {
                            value: "",
                            border: "thin",
                            fillColor: "#E3EEF9"
                        };
                    }
                    return {};
                }),
                ...bandCells
            ]
        };
    }
    function displaySummaryRows(displayDays, businessDays, baseDate, taskCount, resourceCount, assignmentCount, calendarCount, columnCount, startColumnIndex = 5, startRowNumber = 5, displayDaysBeforeBaseDate, displayDaysAfterBaseDate, useBusinessDaysForDisplayRange, useBusinessDaysForProgressBand) {
        const displayWeeks = displayDays > 0 ? Math.ceil(displayDays / 7) : 0;
        const items = [
            { label: "表示日", value: displayDays, fillColor: SUMMARY_SCHEDULE_FILL },
            { label: "表示週", value: displayWeeks, fillColor: SUMMARY_SCHEDULE_FILL },
            { label: "営業日", value: businessDays, fillColor: SUMMARY_SCHEDULE_FILL },
            { label: "前日数", value: displayDaysBeforeBaseDate !== null && displayDaysBeforeBaseDate !== void 0 ? displayDaysBeforeBaseDate : "-", fillColor: SUMMARY_SCHEDULE_FILL },
            { label: "後日数", value: displayDaysAfterBaseDate !== null && displayDaysAfterBaseDate !== void 0 ? displayDaysAfterBaseDate : "-", fillColor: SUMMARY_SCHEDULE_FILL },
            { label: "表示", value: useBusinessDaysForDisplayRange ? "営業日" : "暦日", fillColor: SUMMARY_SCHEDULE_FILL },
            { label: "進捗", value: useBusinessDaysForProgressBand ? "営業日" : "暦日", fillColor: SUMMARY_SCHEDULE_FILL },
            { label: "タスク", value: taskCount, fillColor: SUMMARY_ASSIGNMENT_FILL },
            { label: "リソース", value: resourceCount, fillColor: SUMMARY_ASSIGNMENT_FILL },
            { label: "割当", value: assignmentCount, fillColor: SUMMARY_ASSIGNMENT_FILL },
            { label: "カレンダ", value: calendarCount, fillColor: SUMMARY_ASSIGNMENT_FILL },
            { label: "基準日", value: (baseDate || "-").slice(0, 10), fillColor: SUMMARY_SCHEDULE_FILL }
        ];
        return {
            mergedRanges: [`${columnName(startColumnIndex + 1)}${startRowNumber}:${columnName(startColumnIndex + 2)}${startRowNumber}`],
            rows: [
                blockHeaderRow(columnCount, startColumnIndex, "サマリ"),
                ...items.map((item) => summaryPairRow(columnCount, startColumnIndex, item.label, item.value, item.fillColor))
            ]
        };
    }
    function blockHeaderRow(columnCount, startColumnIndex, title) {
        const cells = Array.from({ length: columnCount }, () => ({}));
        cells[startColumnIndex] = {
            value: title,
            border: "thin",
            horizontalAlign: "left",
            bold: true,
            fillColor: HEADER_ID_FILL
        };
        cells[startColumnIndex + 1] = {
            value: "",
            border: "thin",
            fillColor: HEADER_ID_FILL
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
    function summaryPairRow(columnCount, startColumnIndex, label, value, fillColor) {
        const cells = Array.from({ length: columnCount }, () => ({}));
        cells[startColumnIndex] = summaryStatCell(label, fillColor, false);
        cells[startColumnIndex + 1] = summaryStatCell(value, fillColor, true);
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
        return { height: 24, cells };
    }
    function summaryStatCell(value, fillColor, isValueCell) {
        const valueAlign = typeof value === "number" ? "center" : "left";
        return {
            value,
            border: "thin",
            horizontalAlign: isValueCell ? valueAlign : "right",
            bold: true,
            fillColor
        };
    }
    function headerRow(labels) {
        return {
            height: 24,
            cells: labels.map((label) => {
                if (typeof label === "string") {
                    return {
                        value: label,
                        bold: true,
                        fillColor: headerFillForLabel(label),
                        border: "thin",
                        horizontalAlign: "center"
                    };
                }
                return {
                    border: "thin",
                    horizontalAlign: "center",
                    ...label
                };
            })
        };
    }
    function weekdayRow(fixedColumnCount, dateBand, currentDate, holidaySet) {
        return {
            height: 24,
            cells: [
                ...Array.from({ length: fixedColumnCount }, () => ({})),
                ...dateBand.map((day) => weekdayCell(day, currentDate, holidaySet))
            ]
        };
    }
    function dividerCell() {
        return {
            value: "",
            fillColor: DIVIDER_FILL,
            border: "thin",
            horizontalAlign: "center"
        };
    }
    function headerFillForLabel(label) {
        if (label === "UID" || label === "ID") {
            return HEADER_ID_FILL;
        }
        if (label === "WBS" || label === "種別" || label === "階層" || label === "名称") {
            return HEADER_STRUCTURE_FILL;
        }
        if (label === "開始" || label === "終了" || label === "期間") {
            return HEADER_SCHEDULE_FILL;
        }
        if (label === "進捗" || label === "作業進捗" || label === "マイル" || label === "サマリ" || label === "クリティカル") {
            return HEADER_STATUS_FILL;
        }
        if (label === "担当" || label === "カレンダ" || label === "リソース" || label === "先行") {
            return HEADER_ASSIGNMENT_FILL;
        }
        return HEADER_FILL;
    }
    function todayGuideRow(fixedColumnCount, dateBand, currentDate, holidaySet) {
        return {
            height: 24,
            cells: [
                ...Array.from({ length: fixedColumnCount }, (_, index) => {
                    if (index === 5) {
                        return {
                            value: "基準日",
                            bold: true,
                            border: "thin",
                            horizontalAlign: "right",
                            fillColor: "#FFEFC2"
                        };
                    }
                    if (index === 6) {
                        return {
                            value: "",
                            border: "thin",
                            fillColor: "#FFEFC2"
                        };
                    }
                    if (index > 5 && index < 9) {
                        return {
                            value: "",
                            border: "thin",
                            fillColor: "#FFEFC2"
                        };
                    }
                    if (index === 5 || index === 9) {
                        return {
                            value: "",
                            border: "thin",
                            fillColor: "#FFEFC2"
                        };
                    }
                    return {};
                }),
                ...dateBand.map((day, index) => ({
                    value: isSameDay(day, currentDate) ? "▼基準日" : "",
                    bold: true,
                    border: "thin",
                    horizontalAlign: "center",
                    fillColor: isSameDay(day, currentDate)
                        ? TODAY_BAND_FILL
                        : (holidaySet.has(day)
                            ? HOLIDAY_BAND_FILL
                            : (isWeekStart(day)
                                ? WEEK_START_BAND_FILL
                                : (index < 3 ? BASEDATE_GUIDE_TAIL_FILL : BAND_FILL)))
                }))
            ]
        };
    }
    function cell(value) {
        if (value === undefined || value === "") {
            return {};
        }
        return {
            value,
            border: "thin"
        };
    }
    function taskCell(task, value, horizontalAlign = "left") {
        if (value === undefined || value === "") {
            return {};
        }
        return {
            value,
            border: "thin",
            horizontalAlign,
            wrapText: horizontalAlign === "left" ? true : undefined,
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
    function taskRowHeight(task) {
        const labelLength = formatTaskLabel(task).length;
        if (labelLength > 36) {
            return 34;
        }
        if (labelLength > 24) {
            return 28;
        }
        return undefined;
    }
    function kindCell(task) {
        return {
            value: classifyTaskKind(task),
            border: "thin",
            horizontalAlign: "center",
            bold: true,
            fillColor: task.summary ? PHASE_FILL : (task.milestone ? MILESTONE_FILL : TASK_KIND_FILL)
        };
    }
    function identifierCell(task, value) {
        if (value === undefined || value === "") {
            return {};
        }
        return {
            value,
            border: "thin",
            horizontalAlign: "center",
            bold: task.summary || task.milestone || false,
            fillColor: task.summary ? PHASE_FILL : (task.milestone ? MILESTONE_FILL : IDENTIFIER_FILL)
        };
    }
    function flagCell(task, enabled, marker) {
        return {
            value: enabled ? marker : "",
            border: "thin",
            horizontalAlign: "center",
            bold: !!enabled,
            fillColor: task.summary ? PHASE_FILL : (task.milestone ? MILESTONE_FILL : undefined)
        };
    }
    function progressCell(task, value) {
        const progressValue = formatProgressLabel(value);
        return {
            value: progressValue,
            border: "thin",
            horizontalAlign: "center",
            bold: task.summary || task.milestone || false,
            fillColor: task.summary ? PHASE_FILL : (task.milestone ? MILESTONE_FILL : PROGRESS_COLUMN_FILL)
        };
    }
    function formatProgressLabel(value) {
        if (value === undefined || value === null || !Number.isFinite(value)) {
            return "";
        }
        const clamped = Math.max(0, Math.min(100, Math.round(value)));
        const filled = Math.round(clamped / 10);
        const bar = `${"#".repeat(filled)}${"-".repeat(10 - filled)}`;
        return `${String(clamped).padStart(3, " ")}% [${bar}]`;
    }
    function formatDurationLabel(task, holidaySet, useBusinessDaysForProgressBand) {
        if (useBusinessDaysForProgressBand) {
            const businessDays = enumerateBusinessDays(task.start, task.finish, holidaySet).length;
            return businessDays > 0 ? `${businessDays}営業日` : "-";
        }
        const calendarDays = buildDateBand(task.start, task.finish).length;
        return calendarDays > 0 ? `${calendarDays}日` : "-";
    }
    function formatWbsDate(value) {
        return value ? value.slice(0, 10) : "-";
    }
    function dateNumberCell(day, currentDate, holidaySet) {
        const isToday = isSameDay(day, currentDate);
        const isWeekendDay = isWeekend(day);
        const isHoliday = holidaySet.has(day);
        const weekStart = isWeekStart(day);
        const monthStart = isMonthStart(day);
        return {
            value: formatDateValue(day),
            bold: true,
            border: "thin",
            horizontalAlign: "center",
            fillColor: isToday ? TODAY_BAND_FILL : (isHoliday ? HOLIDAY_BAND_FILL : (isWeekendDay ? WEEKEND_BAND_FILL : (monthStart ? MONTH_START_HEADER_FILL : (weekStart ? WEEK_START_BAND_FILL : HEADER_FILL))))
        };
    }
    function weekdayCell(day, currentDate, holidaySet) {
        const isToday = isSameDay(day, currentDate);
        const isWeekendDay = isWeekend(day);
        const isHoliday = holidaySet.has(day);
        const weekStart = isWeekStart(day);
        const monthStart = isMonthStart(day);
        return {
            value: formatWeekdayValue(day),
            bold: true,
            border: "thin",
            horizontalAlign: "center",
            fillColor: isToday ? TODAY_BAND_FILL : (isHoliday ? HOLIDAY_BAND_FILL : (isWeekendDay ? WEEKEND_BAND_FILL : (monthStart ? MONTH_START_HEADER_FILL : (weekStart ? WEEK_START_BAND_FILL : HEADER_FILL))))
        };
    }
    function dateBandCell(task, day, currentDate, holidaySet, useBusinessDaysForProgressBand) {
        const active = includesDay(task.start, task.finish, day);
        const isToday = isSameDay(day, currentDate);
        const isWeekendDay = isWeekend(day);
        const isHoliday = holidaySet.has(day);
        const weekStart = isWeekStart(day);
        const complete = active && isCompletedDay(task, day, holidaySet, useBusinessDaysForProgressBand);
        return {
            value: active ? activeBandMarker(task) : "",
            border: "thin",
            horizontalAlign: "center",
            fillColor: active
                ? (complete
                    ? (isToday ? TODAY_PROGRESS_BAND_FILL : PROGRESS_BAND_FILL)
                    : (isToday ? TODAY_ACTIVE_BAND_FILL : ACTIVE_BAND_FILL))
                : (isToday ? TODAY_BAND_FILL : (isHoliday ? HOLIDAY_BAND_FILL : (isWeekendDay ? WEEKEND_BAND_FILL : (weekStart ? WEEK_START_BAND_FILL : BAND_FILL))))
        };
    }
    function activeBandMarker(task) {
        if (task.summary) {
            return "━";
        }
        if (task.milestone) {
            return "◆";
        }
        return "■";
    }
    function buildDateBand(startDate, finishDate) {
        const start = parseDateOnly(startDate);
        const finish = parseDateOnly(finishDate);
        if (!start || !finish || start.getTime() > finish.getTime()) {
            return [];
        }
        const days = [];
        const cursor = new Date(start.getTime());
        while (cursor.getTime() <= finish.getTime()) {
            days.push(formatDateOnly(cursor));
            cursor.setDate(cursor.getDate() + 1);
        }
        return days;
    }
    function buildDisplayDateBand(startDate, finishDate, baseDate, displayDaysBeforeBaseDate, displayDaysAfterBaseDate, holidaySet, useBusinessDaysForDisplayRange) {
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
            ? shiftBusinessDays(base, -(before || 0), holidaySet)
            : shiftCalendarDays(base, -(before || 0));
        const to = useBusinessDaysForDisplayRange
            ? shiftBusinessDays(base, after || 0, holidaySet)
            : shiftCalendarDays(base, after || 0);
        const clampedStart = from.getTime() < projectStart.getTime() ? projectStart : from;
        const clampedFinish = to.getTime() > projectFinish.getTime() ? projectFinish : to;
        if (clampedStart.getTime() > clampedFinish.getTime()) {
            return fullBand;
        }
        return buildDateBand(formatDateOnly(clampedStart), formatDateOnly(clampedFinish));
    }
    function normalizeDisplayDayCount(value) {
        if (value === undefined || value === null || !Number.isFinite(value)) {
            return null;
        }
        return Math.max(0, Math.floor(value));
    }
    function countBusinessDays(dateBand, holidaySet) {
        return dateBand.filter((day) => !isWeekend(day) && !holidaySet.has(day)).length;
    }
    function shiftCalendarDays(base, offset) {
        const result = new Date(base.getTime());
        result.setDate(result.getDate() + offset);
        return result;
    }
    function shiftBusinessDays(base, offset, holidaySet) {
        const result = new Date(base.getTime());
        const direction = offset < 0 ? -1 : 1;
        let remaining = Math.abs(offset);
        while (remaining > 0) {
            result.setDate(result.getDate() + direction);
            const day = formatDateOnly(result);
            if (isWeekend(day) || holidaySet.has(day)) {
                continue;
            }
            remaining -= 1;
        }
        return result;
    }
    function buildWeekBandRanges(dateBand, startColumnIndex, rowNumber) {
        const ranges = [];
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
            const startColumn = columnName(startColumnIndex + chunkStart);
            const endColumn = columnName(startColumnIndex + chunkEnd);
            const chunkDays = dateBand.slice(chunkStart, chunkEnd + 1);
            ranges.push({
                range: `${startColumn}${rowNumber}:${endColumn}${rowNumber}`,
                startIndex: chunkStart,
                label: formatWeekLabel(weekStart, chunkDays),
                hasMonthBoundary: chunkDays.some((day) => isMonthStart(day))
            });
            chunkStart = chunkEnd + 1;
        }
        return ranges;
    }
    function includesDay(startDate, finishDate, day) {
        const start = parseDateOnly(startDate);
        const finish = parseDateOnly(finishDate);
        const target = parseDateOnly(day);
        if (!start || !finish || !target) {
            return false;
        }
        return start.getTime() <= target.getTime() && target.getTime() <= finish.getTime();
    }
    function isCompletedDay(task, day, holidaySet, useBusinessDaysForProgressBand) {
        const start = parseDateOnly(task.start);
        const finish = parseDateOnly(task.finish);
        const target = parseDateOnly(day);
        if (!start || !finish || !target) {
            return false;
        }
        if (useBusinessDaysForProgressBand) {
            const activeBusinessDays = enumerateBusinessDays(task.start, task.finish, holidaySet);
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
    function enumerateBusinessDays(startDate, finishDate, holidaySet) {
        return buildDateBand(startDate, finishDate).filter((day) => !isWeekend(day) && !holidaySet.has(day));
    }
    function isSameDay(day, other) {
        return day === (other || "").slice(0, 10);
    }
    function isWeekend(day) {
        const target = parseDateOnly(day);
        if (!target) {
            return false;
        }
        const weekday = target.getDay();
        return weekday === 0 || weekday === 6;
    }
    function isWeekStart(day) {
        const target = parseDateOnly(day);
        if (!target) {
            return false;
        }
        return target.getDay() === 0;
    }
    function isMonthStart(day) {
        const target = parseDateOnly(day);
        if (!target) {
            return false;
        }
        return target.getDate() === 1;
    }
    function parseDateOnly(value) {
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
    function expandExceptionDays(exception) {
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
        const days = [];
        const cursor = new Date(start.getTime());
        while (cursor.getTime() <= finish.getTime()) {
            days.push(formatDateOnly(cursor));
            cursor.setDate(cursor.getDate() + 1);
        }
        return days;
    }
    function formatDateOnly(value) {
        return [
            value.getFullYear(),
            String(value.getMonth() + 1).padStart(2, "0"),
            String(value.getDate()).padStart(2, "0")
        ].join("-");
    }
    function formatDateValue(day) {
        const target = parseDateOnly(day);
        if (!target) {
            return day;
        }
        return `${target.getMonth() + 1}/${target.getDate()}`;
    }
    function formatWeekdayValue(day) {
        const target = parseDateOnly(day);
        if (!target) {
            return day;
        }
        const weekdays = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
        return weekdays[target.getDay()];
    }
    function formatWeekKey(day) {
        const target = parseDateOnly(day);
        if (!target) {
            return day;
        }
        const sunday = new Date(target.getTime());
        const offset = sunday.getDay();
        sunday.setDate(sunday.getDate() - offset);
        return formatDateOnly(sunday);
    }
    function formatWeekLabel(weekKey, days) {
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
    function columnName(index) {
        let current = index;
        let name = "";
        while (current > 0) {
            const remainder = (current - 1) % 26;
            name = String.fromCharCode(65 + remainder) + name;
            current = Math.floor((current - 1) / 26);
        }
        return name;
    }
    globalThis.__mikuprojectWbsXlsx = {
        collectWbsHolidayDates,
        exportWbsWorkbook
    };
})();
