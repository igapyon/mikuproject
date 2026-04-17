/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  type XlsxCellLike = {
    value?: string | number | boolean;
    numberFormat?: "general" | "integer" | "decimal" | "date" | "datetime" | "percent";
    horizontalAlign?: "left" | "center" | "right";
    bold?: boolean;
    fontSize?: number;
    fillColor?: string;
    border?: "thin";
  };

  type XlsxWorkbookLike = {
    sheets: Array<{
      name: string;
      columns?: Array<{ width?: number }>;
      mergedRanges?: string[];
      dataValidations?: Array<{
        type: "list";
        sqref: string;
        formula1: string;
        allowBlank?: boolean;
      }>;
      rows: Array<{
        height?: number;
        cells: XlsxCellLike[];
      }>;
    }>;
  };

  type ImportChange = {
    scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
    uid: string;
    label: string;
    field: string;
    before: string | number | boolean | undefined;
    after: string | number | boolean;
  };

  const BOOLEAN_TRUE_LABEL = "○";
  const BOOLEAN_FALSE_LABEL = "ー";

  function cloneProjectModel(model: ProjectModel): ProjectModel {
    return JSON.parse(JSON.stringify(model)) as ProjectModel;
  }

  function buildDefaultWorkingTimes(project: ProjectInfo): WorkingTimeModel[] {
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

  function buildDefaultStandardWeekDays(project: ProjectInfo): WeekDayModel[] {
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

  function normalizeDateOnly(value: string): string | undefined {
    const trimmed = value.trim();
    const match = trimmed.match(/^(\d{4}-\d{2}-\d{2})/);
    return match ? match[1] : undefined;
  }

  function normalizeDateTimeInput(value: string | undefined): string | undefined {
    if (!value) {
      return value;
    }
    const trimmed = value.trim();
    if (/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/.test(trimmed)) {
      return trimmed.replace(" ", "T");
    }
    return trimmed;
  }

  function normalizeExceptionBoundaryInput(value: string | undefined, isEndOfDay: boolean): string | undefined {
    const normalizedDate = value ? normalizeDateOnly(value) : undefined;
    if (normalizedDate && normalizedDate === value?.trim()) {
      return `${normalizedDate}${isEndOfDay ? "T23:59:59" : "T00:00:00"}`;
    }
    return normalizeDateTimeInput(value);
  }

  function readHeaderMap(sheet: XlsxWorkbookLike["sheets"][number], headerRowIndex: number): Map<string, number> {
    const headerRow = sheet.rows[headerRowIndex];
    const columnIndexByLabel = new Map<string, number>();
    if (!headerRow) {
      return columnIndexByLabel;
    }
    headerRow.cells.forEach((cell, index) => {
      if (typeof cell.value === "string" && cell.value) {
        columnIndexByLabel.set(cell.value, index);
      }
    });
    return columnIndexByLabel;
  }

  function readStringCellAt(cells: XlsxCellLike[], index: number | undefined): string | undefined {
    if (index === undefined) {
      return undefined;
    }
    return readStringCell(cells[index]);
  }

  function readNumberCellAt(cells: XlsxCellLike[], index: number | undefined): number | undefined {
    if (index === undefined) {
      return undefined;
    }
    return readNumberCell(cells[index]);
  }

  function readBooleanCellAt(cells: XlsxCellLike[], index: number | undefined): boolean | undefined {
    if (index === undefined) {
      return undefined;
    }
    return readBooleanCell(cells[index]);
  }

  function readStringCell(cell: XlsxCellLike | undefined): string | undefined {
    if (!cell || cell.value === undefined) {
      return undefined;
    }
    if (typeof cell.value === "string") {
      return cell.value;
    }
    if (typeof cell.value === "number" || typeof cell.value === "boolean") {
      return String(cell.value);
    }
    return undefined;
  }

  function readNumberCell(cell: XlsxCellLike | undefined): number | undefined {
    if (!cell || cell.value === undefined) {
      return undefined;
    }
    if (typeof cell.value === "number" && Number.isFinite(cell.value)) {
      return cell.value;
    }
    if (typeof cell.value === "string" && cell.value.trim() !== "") {
      const parsed = Number(cell.value);
      return Number.isFinite(parsed) ? parsed : undefined;
    }
    return undefined;
  }

  function readBooleanCell(cell: XlsxCellLike | undefined): boolean | undefined {
    if (!cell || cell.value === undefined) {
      return undefined;
    }
    if (typeof cell.value === "boolean") {
      return cell.value;
    }
    if (typeof cell.value === "number") {
      return cell.value !== 0;
    }
    if (typeof cell.value === "string") {
      if (cell.value === "true" || cell.value === "TRUE" || cell.value === "1" || cell.value === BOOLEAN_TRUE_LABEL || cell.value === "〇" || cell.value === "○") {
        return true;
      }
      if (cell.value === "false" || cell.value === "FALSE" || cell.value === "0" || cell.value === BOOLEAN_FALSE_LABEL || cell.value === "-" || cell.value === "－") {
        return false;
      }
    }
    return undefined;
  }

  function assignIfChanged<T extends object, K extends keyof T>(
    changes: ImportChange[],
    scope: ImportChange["scope"],
    uid: string,
    label: string,
    target: T,
    key: K,
    field: string,
    value: T[K] | undefined
  ): void {
    if (value === undefined) {
      return;
    }
    const before = target[key];
    if (before === value) {
      return;
    }
    target[key] = value;
    changes.push({
      scope,
      uid,
      label,
      field,
      before: before as string | number | boolean | undefined,
      after: value as string | number | boolean
    });
  }

  function parsePredecessorCell(value: string | undefined): PredecessorModel[] | undefined {
    if (value === undefined) {
      return undefined;
    }
    const normalized = value
      .split(",")
      .map((item) => item.trim())
      .filter(Boolean);
    return normalized.map((predecessorUid) => ({ predecessorUid }));
  }

  function formatPredecessorList(predecessors: PredecessorModel[]): string {
    return predecessors.map((item) => item.predecessorUid).join(", ");
  }

  function assignPredecessorsIfChanged(
    changes: ImportChange[],
    uid: string,
    label: string,
    task: TaskModel,
    predecessors: PredecessorModel[] | undefined
  ): void {
    if (predecessors === undefined) {
      return;
    }
    const beforeText = formatPredecessorList(task.predecessors);
    const afterText = formatPredecessorList(predecessors);
    if (beforeText === afterText) {
      return;
    }
    task.predecessors = predecessors;
    changes.push({
      scope: "tasks",
      uid,
      label,
      field: "Predecessors",
      before: beforeText || undefined,
      after: afterText
    });
  }

  globalThis.__mikuprojectProjectXlsxImportUtil = {
    cloneProjectModel,
    buildDefaultWorkingTimes,
    buildDefaultStandardWeekDays,
    formatExceptionDate,
    formatExceptionBoundaryDate,
    normalizeDateOnly,
    normalizeDateTimeInput,
    normalizeExceptionBoundaryInput,
    readHeaderMap,
    readStringCellAt,
    readNumberCellAt,
    readBooleanCellAt,
    readStringCell,
    readNumberCell,
    readBooleanCell,
    assignIfChanged,
    parsePredecessorCell,
    assignPredecessorsIfChanged
  };
})();
