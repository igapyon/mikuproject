/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const workbookSchema = globalThis.__mikuprojectProjectWorkbookSchema;
    if (!workbookSchema) {
        throw new Error("mikuproject Project Workbook Schema module is not loaded");
    }
    const { PROJECT_EDITABLE_FIELDS } = workbookSchema;
    const HEADER_FILL = "#D9EAF7";
    const SECTION_FILL = "#BFD7EA";
    const LABEL_FILL = "#EDF5FB";
    const ALT_ROW_FILL = "#F9FBFD";
    const DATE_FILL = "#FFF4E8";
    const PERCENT_FILL = "#FCECF3";
    const REFERENCE_FILL = "#EEF7F4";
    const COUNT_FILL = "#F2F5F8";
    const EDITABLE_FILL = "#FDE7C7";
    const DURATION_FILL = "#FBF6ED";
    const NOTES_FILL = "#FFFBEA";
    const NAME_FILL = "#FAF6FF";
    const WORK_FILL = "#F1F8FD";
    const BOOLEAN_TRUE_LABEL = "○";
    const BOOLEAN_FALSE_LABEL = "ー";
    const OPTIONS_SHEET_NAME = "Options";
    const SUMMARY_FILL = "#E6F2E0";
    const MILESTONE_FILL = "#FFF0CF";
    const CRITICAL_FILL = "#F8DDE6";
    const SHEET_THEMES = {
        project: { section: "#BFD7EA", header: "#D9EAF7", label: "#EDF5FB" },
        tasks: { section: "#D4E0EC", header: "#E6EDF4", label: "#F2F6FA" },
        resources: { section: "#C8E3D8", header: "#DDF0E8", label: "#EFF8F4" },
        assignments: { section: "#D7D2EC", header: "#E7E3F5", label: "#F2F0FA" },
        calendars: { section: "#D7E3C4", header: "#E7F0DA", label: "#F2F7EA" },
        nonWorkingDays: { section: "#E9C7D5", header: "#F4DDE6", label: "#FBEEF3" }
    };
    function headerRow(labels, fillColor = HEADER_FILL) {
        return {
            height: 24,
            cells: labels.map((label) => ({
                value: label,
                bold: true,
                fillColor,
                border: "thin",
                horizontalAlign: "center"
            }))
        };
    }
    function titleRow(title, fillColor = SECTION_FILL) {
        return {
            height: 28,
            cells: [
                {
                    value: title,
                    bold: true,
                    fontSize: 16,
                    fillColor,
                    horizontalAlign: "left"
                },
                {
                    fillColor
                }
            ]
        };
    }
    function sectionTitleRow(title, columnCount, fillColor = SECTION_FILL) {
        return {
            height: 26,
            cells: [
                {
                    value: title,
                    bold: true,
                    fontSize: 14,
                    fillColor,
                    horizontalAlign: "left"
                },
                ...Array.from({ length: Math.max(0, columnCount - 1) }, () => ({
                    fillColor
                }))
            ]
        };
    }
    function keyValueRow(label, value, labelFill = LABEL_FILL) {
        return {
            cells: [
                {
                    value: label,
                    bold: true,
                    fillColor: labelFill,
                    border: "thin"
                },
                keyValueCell(label, value)
            ]
        };
    }
    function cell(value) {
        if (value === undefined) {
            return {};
        }
        return {
            value: stringifyCellValue(value),
            border: "thin"
        };
    }
    function stringifyCellValue(value) {
        if (typeof value === "boolean") {
            return value ? BOOLEAN_TRUE_LABEL : BOOLEAN_FALSE_LABEL;
        }
        return typeof value === "string" ? value : String(value);
    }
    function keyValueCell(label, value) {
        if (isEditableProjectLabel(label)) {
            return editableCell(buildProjectValueCell(label, value));
        }
        return buildProjectValueCell(label, value);
    }
    function buildProjectValueCell(label, value) {
        if (isDateTimeLabel(label)) {
            return {
                ...cell(formatDateTimeDisplay(value)),
                fillColor: DATE_FILL
            };
        }
        if (label === "Name" || label === "Title") {
            return {
                ...cell(value),
                fillColor: NAME_FILL,
                bold: true
            };
        }
        if (label === "Author" || label === "Company") {
            return {
                ...cell(value),
                fillColor: NAME_FILL
            };
        }
        if (label === "CalendarUID") {
            return {
                ...cell(value),
                fillColor: REFERENCE_FILL,
                horizontalAlign: "center"
            };
        }
        if (label === "ScheduleFromStart") {
            return {
                ...cell(value),
                fillColor: COUNT_FILL,
                horizontalAlign: "center"
            };
        }
        return {
            ...cell(value),
            fillColor: isNumericSummaryLabel(label) ? COUNT_FILL : undefined
        };
    }
    function textCell(value, rowIndex) {
        return styledCell(value, rowIndex);
    }
    function taskNameCell(task, rowIndex) {
        const fillColor = task.summary ? SUMMARY_FILL : (task.critical ? CRITICAL_FILL : undefined);
        return {
            ...styledCell(task.name, rowIndex, { fillColor }),
            bold: task.summary || task.milestone
        };
    }
    function taskFlagCell(value, _rowIndex, _activeFillColor) {
        const displayValue = value === undefined ? undefined : stringifyCellValue(value);
        return {
            value: displayValue,
            border: "thin",
            fillColor: EDITABLE_FILL,
            horizontalAlign: "center"
        };
    }
    function referenceCell(value, rowIndex) {
        return styledCell(value, rowIndex, {
            fillColor: REFERENCE_FILL,
            horizontalAlign: "center"
        });
    }
    function countCell(value, rowIndex) {
        return styledCell(value, rowIndex, {
            fillColor: COUNT_FILL,
            horizontalAlign: "center"
        });
    }
    function percentCell(value, rowIndex) {
        return styledCell(value, rowIndex, {
            fillColor: PERCENT_FILL,
            horizontalAlign: "center"
        });
    }
    function durationCell(value, rowIndex) {
        return styledCell(value, rowIndex, {
            fillColor: DURATION_FILL,
            horizontalAlign: "center"
        });
    }
    function predecessorCell(value, rowIndex) {
        return styledCell(value, rowIndex, {
            fillColor: REFERENCE_FILL
        });
    }
    function notesCell(value, rowIndex) {
        return styledCell(value, rowIndex, {
            fillColor: NOTES_FILL
        });
    }
    function entityNameCell(value, rowIndex) {
        return styledCell(value, rowIndex, {
            fillColor: NAME_FILL
        });
    }
    function workCell(value, rowIndex) {
        return styledCell(value, rowIndex, {
            fillColor: WORK_FILL
        });
    }
    function editableCell(cellLike) {
        return {
            ...cellLike,
            border: cellLike.border || "thin",
            fillColor: EDITABLE_FILL
        };
    }
    function dateTimeCell(value, rowIndex) {
        return styledCell(formatDateTimeDisplay(value), rowIndex, {
            fillColor: DATE_FILL
        });
    }
    function dateOnlyCell(value, rowIndex) {
        return styledCell(value, rowIndex, {
            fillColor: DATE_FILL,
            horizontalAlign: "center"
        });
    }
    function styledCell(value, rowIndex, overrides = {}) {
        const base = cell(value);
        if (base.value === undefined) {
            return base;
        }
        return {
            ...base,
            fillColor: overrides.fillColor || (rowIndex % 2 === 0 ? ALT_ROW_FILL : undefined),
            horizontalAlign: overrides.horizontalAlign,
            numberFormat: overrides.numberFormat
        };
    }
    function formatDateTimeDisplay(value) {
        if (typeof value !== "string") {
            return value;
        }
        return value.replace("T", " ");
    }
    function isDateTimeLabel(label) {
        return ["StartDate", "FinishDate", "CurrentDate", "StatusDate"].includes(label);
    }
    function isNumericSummaryLabel(label) {
        return ["OutlineCodes", "WBSMasks", "ExtendedAttributes", "MinutesPerDay", "MinutesPerWeek", "DaysPerMonth"].includes(label);
    }
    function isEditableProjectLabel(label) {
        return PROJECT_EDITABLE_FIELDS.includes(label);
    }
    function readProjectFieldValue(project, field) {
        switch (field) {
            case "Name":
                return project.name;
            case "Title":
                return project.title;
            case "Author":
                return project.author;
            case "Company":
                return project.company;
            case "StartDate":
                return project.startDate;
            case "FinishDate":
                return project.finishDate;
            case "CurrentDate":
                return project.currentDate;
            case "StatusDate":
                return project.statusDate;
            case "CalendarUID":
                return project.calendarUID;
            case "MinutesPerDay":
                return project.minutesPerDay;
            case "MinutesPerWeek":
                return project.minutesPerWeek;
            case "DaysPerMonth":
                return project.daysPerMonth;
            case "ScheduleFromStart":
                return project.scheduleFromStart;
            case "OutlineCodes":
                return project.outlineCodes.length;
            case "WBSMasks":
                return project.wbsMasks.length;
            case "ExtendedAttributes":
                return project.extendedAttributes.length;
            default:
                return undefined;
        }
    }
    function formatExceptionDate(exception) {
        var _a, _b;
        const fromDate = (_a = exception.fromDate) === null || _a === void 0 ? void 0 : _a.slice(0, 10);
        const toDate = (_b = exception.toDate) === null || _b === void 0 ? void 0 : _b.slice(0, 10);
        if (!fromDate || !toDate) {
            return undefined;
        }
        return fromDate === toDate ? fromDate : undefined;
    }
    function formatExceptionBoundaryDate(value) {
        return value === null || value === void 0 ? void 0 : value.slice(0, 10);
    }
    function buildOptionsSheet() {
        return {
            name: OPTIONS_SHEET_NAME,
            columns: [{ width: 18 }, { width: 14 }],
            mergedRanges: [],
            rows: [
                headerRow(["BooleanChoice", "Meaning"], HEADER_FILL),
                { cells: [textCell(BOOLEAN_TRUE_LABEL, 0), textCell("true", 0)] },
                { cells: [textCell(BOOLEAN_FALSE_LABEL, 1), textCell("false", 1)] }
            ]
        };
    }
    function buildBooleanDataValidations(ranges) {
        const sqref = ranges.filter((value) => Boolean(value)).join(" ");
        if (!sqref) {
            return undefined;
        }
        return [{
                type: "list",
                sqref,
                formula1: `${OPTIONS_SHEET_NAME}!$A$2:$A$3`,
                allowBlank: true
            }];
    }
    function buildColumnRange(columnIndex, startRow, endRow) {
        if (endRow < startRow) {
            return undefined;
        }
        const columnName = encodeColumnName(columnIndex);
        return `${columnName}${startRow}:${columnName}${endRow}`;
    }
    function buildSingleCellReference(columnIndex, rowNumber) {
        if (!rowNumber) {
            return undefined;
        }
        return `${encodeColumnName(columnIndex)}${rowNumber}`;
    }
    function encodeColumnName(columnIndex) {
        let current = columnIndex + 1;
        let name = "";
        while (current > 0) {
            const remainder = (current - 1) % 26;
            name = String.fromCharCode(65 + remainder) + name;
            current = Math.floor((current - 1) / 26);
        }
        return name;
    }
    globalThis.__mikuprojectProjectXlsxExportUtil = {
        BOOLEAN_TRUE_LABEL,
        BOOLEAN_FALSE_LABEL,
        OPTIONS_SHEET_NAME,
        SUMMARY_FILL,
        MILESTONE_FILL,
        CRITICAL_FILL,
        SHEET_THEMES,
        headerRow,
        titleRow,
        sectionTitleRow,
        keyValueRow,
        countCell,
        editableCell,
        referenceCell,
        textCell,
        taskNameCell,
        taskFlagCell,
        dateTimeCell,
        durationCell,
        percentCell,
        predecessorCell,
        notesCell,
        entityNameCell,
        workCell,
        dateOnlyCell,
        readProjectFieldValue,
        formatExceptionDate,
        formatExceptionBoundaryDate,
        buildOptionsSheet,
        buildBooleanDataValidations,
        buildColumnRange,
        buildSingleCellReference
    };
})();
