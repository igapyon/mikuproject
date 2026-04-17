/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  const workbookSchema = (globalThis as typeof globalThis & {
    __mikuprojectProjectWorkbookSchema?: {
      PROJECT_EDITABLE_FIELDS: readonly string[];
    };
  }).__mikuprojectProjectWorkbookSchema;

  if (!workbookSchema) {
    throw new Error("mikuproject Project Workbook Schema module is not loaded");
  }

  const { PROJECT_EDITABLE_FIELDS } = workbookSchema;

  type XlsxCellLike = {
    value?: string | number | boolean;
    numberFormat?: "general" | "integer" | "decimal" | "date" | "datetime" | "percent";
    horizontalAlign?: "left" | "center" | "right";
    bold?: boolean;
    fontSize?: number;
    fillColor?: string;
    border?: "thin";
  };

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
  } as const;

  function headerRow(labels: string[], fillColor = HEADER_FILL) {
    return {
      height: 24,
      cells: labels.map((label) => ({
        value: label,
        bold: true,
        fillColor,
        border: "thin" as const,
        horizontalAlign: "center" as const
      }))
    };
  }

  function titleRow(title: string, fillColor = SECTION_FILL) {
    return {
      height: 28,
      cells: [
        {
          value: title,
          bold: true,
          fontSize: 16,
          fillColor,
          horizontalAlign: "left" as const
        },
        {
          fillColor
        }
      ]
    };
  }

  function sectionTitleRow(title: string, columnCount: number, fillColor = SECTION_FILL) {
    return {
      height: 26,
      cells: [
        {
          value: title,
          bold: true,
          fontSize: 14,
          fillColor,
          horizontalAlign: "left" as const
        },
        ...Array.from({ length: Math.max(0, columnCount - 1) }, () => ({
          fillColor
        }))
      ]
    };
  }

  function keyValueRow(label: string, value: string | number | boolean | undefined, labelFill = LABEL_FILL) {
    return {
      cells: [
        {
          value: label,
          bold: true,
          fillColor: labelFill,
          border: "thin" as const
        },
        keyValueCell(label, value)
      ]
    };
  }

  function cell(value: string | number | boolean | undefined): XlsxCellLike {
    if (value === undefined) {
      return {};
    }
    return {
      value: stringifyCellValue(value),
      border: "thin"
    };
  }

  function stringifyCellValue(value: string | number | boolean): string {
    if (typeof value === "boolean") {
      return value ? BOOLEAN_TRUE_LABEL : BOOLEAN_FALSE_LABEL;
    }
    return typeof value === "string" ? value : String(value);
  }

  function keyValueCell(label: string, value: string | number | boolean | undefined): XlsxCellLike {
    if (isEditableProjectLabel(label)) {
      return editableCell(buildProjectValueCell(label, value));
    }
    return buildProjectValueCell(label, value);
  }

  function buildProjectValueCell(label: string, value: string | number | boolean | undefined): XlsxCellLike {
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

  function textCell(value: string | number | boolean | undefined, rowIndex: number): XlsxCellLike {
    return styledCell(value, rowIndex);
  }

  function taskNameCell(task: TaskModel, rowIndex: number): XlsxCellLike {
    const fillColor = task.summary ? SUMMARY_FILL : (task.critical ? CRITICAL_FILL : undefined);
    return {
      ...styledCell(task.name, rowIndex, { fillColor }),
      bold: task.summary || task.milestone
    };
  }

  function taskFlagCell(value: string | number | boolean | undefined, _rowIndex: number, _activeFillColor: string): XlsxCellLike {
    const displayValue = value === undefined ? undefined : stringifyCellValue(value);
    return {
      value: displayValue,
      border: "thin",
      fillColor: EDITABLE_FILL,
      horizontalAlign: "center"
    };
  }

  function referenceCell(value: string | number | boolean | undefined, rowIndex: number): XlsxCellLike {
    return styledCell(value, rowIndex, {
      fillColor: REFERENCE_FILL,
      horizontalAlign: "center"
    });
  }

  function countCell(value: string | number | boolean | undefined, rowIndex: number): XlsxCellLike {
    return styledCell(value, rowIndex, {
      fillColor: COUNT_FILL,
      horizontalAlign: "center"
    });
  }

  function percentCell(value: string | number | boolean | undefined, rowIndex: number): XlsxCellLike {
    return styledCell(value, rowIndex, {
      fillColor: PERCENT_FILL,
      horizontalAlign: "center"
    });
  }

  function durationCell(value: string | number | boolean | undefined, rowIndex: number): XlsxCellLike {
    return styledCell(value, rowIndex, {
      fillColor: DURATION_FILL,
      horizontalAlign: "center"
    });
  }

  function predecessorCell(value: string | number | boolean | undefined, rowIndex: number): XlsxCellLike {
    return styledCell(value, rowIndex, {
      fillColor: REFERENCE_FILL
    });
  }

  function notesCell(value: string | number | boolean | undefined, rowIndex: number): XlsxCellLike {
    return styledCell(value, rowIndex, {
      fillColor: NOTES_FILL
    });
  }

  function entityNameCell(value: string | number | boolean | undefined, rowIndex: number): XlsxCellLike {
    return styledCell(value, rowIndex, {
      fillColor: NAME_FILL
    });
  }

  function workCell(value: string | number | boolean | undefined, rowIndex: number): XlsxCellLike {
    return styledCell(value, rowIndex, {
      fillColor: WORK_FILL
    });
  }

  function editableCell(cellLike: XlsxCellLike): XlsxCellLike {
    return {
      ...cellLike,
      border: cellLike.border || "thin",
      fillColor: EDITABLE_FILL
    };
  }

  function dateTimeCell(value: string | number | boolean | undefined, rowIndex: number): XlsxCellLike {
    return styledCell(formatDateTimeDisplay(value), rowIndex, {
      fillColor: DATE_FILL
    });
  }

  function dateOnlyCell(value: string | number | boolean | undefined, rowIndex: number): XlsxCellLike {
    return styledCell(value, rowIndex, {
      fillColor: DATE_FILL,
      horizontalAlign: "center"
    });
  }

  function styledCell(
    value: string | number | boolean | undefined,
    rowIndex: number,
    overrides: Partial<XlsxCellLike> = {}
  ): XlsxCellLike {
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

  function formatDateTimeDisplay(value: string | number | boolean | undefined): string | number | boolean | undefined {
    if (typeof value !== "string") {
      return value;
    }
    return value.replace("T", " ");
  }

  function isDateTimeLabel(label: string): boolean {
    return ["StartDate", "FinishDate", "CurrentDate", "StatusDate"].includes(label);
  }

  function isNumericSummaryLabel(label: string): boolean {
    return ["OutlineCodes", "WBSMasks", "ExtendedAttributes", "MinutesPerDay", "MinutesPerWeek", "DaysPerMonth"].includes(label);
  }

  function isEditableProjectLabel(label: string): boolean {
    return PROJECT_EDITABLE_FIELDS.includes(label);
  }

  function readProjectFieldValue(project: ProjectModel["project"], field: string): string | number | boolean | undefined {
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

  function formatExceptionDate(exception: CalendarExceptionModel): string | undefined {
    const fromDate = exception.fromDate?.slice(0, 10);
    const toDate = exception.toDate?.slice(0, 10);
    if (!fromDate || !toDate) {
      return undefined;
    }
    return fromDate === toDate ? fromDate : undefined;
  }

  function formatExceptionBoundaryDate(value: string | undefined): string | undefined {
    return value?.slice(0, 10);
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

  function buildBooleanDataValidations(ranges: Array<string | undefined>) {
    const sqref = ranges.filter((value): value is string => Boolean(value)).join(" ");
    if (!sqref) {
      return undefined;
    }
    return [{
      type: "list" as const,
      sqref,
      formula1: `${OPTIONS_SHEET_NAME}!$A$2:$A$3`,
      allowBlank: true
    }];
  }

  function buildColumnRange(columnIndex: number, startRow: number, endRow: number): string | undefined {
    if (endRow < startRow) {
      return undefined;
    }
    const columnName = encodeColumnName(columnIndex);
    return `${columnName}${startRow}:${columnName}${endRow}`;
  }

  function buildSingleCellReference(columnIndex: number, rowNumber: number | undefined): string | undefined {
    if (!rowNumber) {
      return undefined;
    }
    return `${encodeColumnName(columnIndex)}${rowNumber}`;
  }

  function encodeColumnName(columnIndex: number): string {
    let current = columnIndex + 1;
    let name = "";
    while (current > 0) {
      const remainder = (current - 1) % 26;
      name = String.fromCharCode(65 + remainder) + name;
      current = Math.floor((current - 1) / 26);
    }
    return name;
  }

  (globalThis as typeof globalThis & {
    __mikuprojectProjectXlsxExportUtil?: {
      BOOLEAN_TRUE_LABEL: string;
      BOOLEAN_FALSE_LABEL: string;
      OPTIONS_SHEET_NAME: string;
      SUMMARY_FILL: string;
      MILESTONE_FILL: string;
      CRITICAL_FILL: string;
      SHEET_THEMES: typeof SHEET_THEMES;
      headerRow: typeof headerRow;
      titleRow: typeof titleRow;
      sectionTitleRow: typeof sectionTitleRow;
      keyValueRow: typeof keyValueRow;
      countCell: typeof countCell;
      editableCell: typeof editableCell;
      referenceCell: typeof referenceCell;
      textCell: typeof textCell;
      taskNameCell: typeof taskNameCell;
      taskFlagCell: typeof taskFlagCell;
      dateTimeCell: typeof dateTimeCell;
      durationCell: typeof durationCell;
      percentCell: typeof percentCell;
      predecessorCell: typeof predecessorCell;
      notesCell: typeof notesCell;
      entityNameCell: typeof entityNameCell;
      workCell: typeof workCell;
      dateOnlyCell: typeof dateOnlyCell;
      readProjectFieldValue: typeof readProjectFieldValue;
      formatExceptionDate: typeof formatExceptionDate;
      formatExceptionBoundaryDate: typeof formatExceptionBoundaryDate;
      buildOptionsSheet: typeof buildOptionsSheet;
      buildBooleanDataValidations: typeof buildBooleanDataValidations;
      buildColumnRange: typeof buildColumnRange;
      buildSingleCellReference: typeof buildSingleCellReference;
    };
  }).__mikuprojectProjectXlsxExportUtil = {
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
