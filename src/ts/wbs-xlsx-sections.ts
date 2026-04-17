/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  type WbsXlsxCellLike = {
    value?: string | number | boolean;
    numberFormat?: "general" | "integer" | "decimal" | "date" | "datetime" | "percent";
    horizontalAlign?: "left" | "center" | "right";
    verticalAlign?: "top" | "center" | "bottom";
    wrapText?: boolean;
    bold?: boolean;
    fontSize?: number;
    fillColor?: string;
    border?: "thin";
  };

  type WbsSheetLayoutHelper = {
    columnName(columnIndex: number): string;
    columnIndex(columnReference: string): number;
    reference(rowNumber: number, columnIndex: number): string;
    parseCellReference(reference: string): {
      reference: string;
      rowNumber: number;
      rowIndex: number;
      columnName: string;
      columnIndex: number;
    };
    range(startReference: string, endReference: string): string;
    describeCell(reference: string): string;
    logCell(reference: string, label?: string, logger?: (...args: unknown[]) => void): string;
  };

  type WbsXlsxSectionsConfig = {
    layout: WbsSheetLayoutHelper;
    fills: {
      headerId: string;
      headerFill: string;
      headerStructure: string;
      headerSchedule: string;
      headerStatus: string;
      headerAssignment: string;
      summarySchedule: string;
      summaryAssignment: string;
      phase: string;
      milestone: string;
      placeholder: string;
      divider: string;
      referenceColumn: string;
    };
  };

  type WbsXlsxSectionsHelper = {
    firstResourceName(resourceNames: string[] | undefined): string;
    formatCalendarLabel(calendarUID: string | undefined, calendarNameByUid: Map<string, string>): string;
    truncateWbsReference(value: string | undefined, maxLength: number): string;
    referenceCell(
      task: TaskModel,
      value: string | undefined,
      horizontalAlign?: "left" | "center" | "right"
    ): WbsXlsxCellLike;
    projectInfoRows(
      project: ProjectModel["project"],
      calendarNameByUid: Map<string, string>,
      holidayCount: number,
      columnCount: number,
      startColumnIndex: number,
      startRowNumber: number
    ): {
      mergedRanges: string[];
      rows: Array<{ height?: number; cells: WbsXlsxCellLike[] }>;
    };
    legendRows(
      columnCount: number,
      startRowNumber: number
    ): {
      mergedRanges: string[];
      rows: Array<{ height?: number; cells: WbsXlsxCellLike[] }>;
    };
    displaySummaryRows(
      displayDays: number,
      businessDays: number,
      baseDate: string | undefined,
      taskCount: number,
      resourceCount: number,
      assignmentCount: number,
      calendarCount: number,
      columnCount: number,
      startColumnIndex?: number,
      startRowNumber?: number,
      displayDaysBeforeBaseDate?: number,
      displayDaysAfterBaseDate?: number,
      useBusinessDaysForDisplayRange?: boolean,
      useBusinessDaysForProgressBand?: boolean
    ): {
      mergedRanges: string[];
      rows: Array<{ height?: number; cells: WbsXlsxCellLike[] }>;
    };
    formatWbsDate(value: string | undefined): string;
  };

  function createWbsXlsxSectionsHelper(config: WbsXlsxSectionsConfig): WbsXlsxSectionsHelper {
    const { layout, fills } = config;

    function stringifyCellValue(value: string | number | boolean): string {
      return typeof value === "string" ? value : String(value);
    }

    function blockHeaderRow(columnCount: number, startColumnIndex: number, title: string) {
      const cells = Array.from({ length: columnCount }, () => ({} as WbsXlsxCellLike));
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

    function projectBlockHeaderRow(columnCount: number, startColumnIndex: number, title: string) {
      const cells = Array.from({ length: columnCount }, () => ({} as WbsXlsxCellLike));
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

    function summaryStatCell(value: string | number, fillColor: string, isValueCell: boolean): WbsXlsxCellLike {
      const valueAlign = typeof value === "number" ? "center" : "left";
      return {
        value: stringifyCellValue(value),
        border: "thin",
        horizontalAlign: isValueCell ? valueAlign : "right",
        bold: true,
        fillColor
      };
    }

    function projectPairRow(
      columnCount: number,
      startColumnIndex: number,
      label: string,
      value: string | number,
      fillColor: string
    ) {
      const cells = Array.from({ length: columnCount }, () => ({} as WbsXlsxCellLike));
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

    function summaryPairRow(
      columnCount: number,
      startColumnIndex: number,
      label: string,
      value: string | number,
      fillColor: string
    ) {
      const cells = Array.from({ length: columnCount }, () => ({} as WbsXlsxCellLike));
      cells[startColumnIndex] = summaryStatCell(label, fillColor, false);
      cells[startColumnIndex + 1] = summaryStatCell(value, fillColor, true);
      cells[startColumnIndex + 2] = {
        value: "",
        border: "thin",
        fillColor
      };
      return { height: 22, cells };
    }

    function mergedLabelRow(
      columnCount: number,
      startColumnIndex: number,
      value: string,
      fillColor: string
    ) {
      const cells = Array.from({ length: columnCount }, () => ({} as WbsXlsxCellLike));
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

    function displayReferenceValue(value: string | undefined): string {
      return value && value.trim() ? value : "-";
    }

    function truncateWbsReference(value: string | undefined, maxLength: number): string {
      const normalized = value?.trim() || "";
      if (!normalized) {
        return "";
      }
      if (normalized.length <= maxLength) {
        return normalized;
      }
      return `${normalized.slice(0, Math.max(1, maxLength - 3))}...`;
    }

    function formatCalendarLabel(calendarUID: string | undefined, calendarNameByUid: Map<string, string>): string {
      if (!calendarUID) {
        return "-";
      }
      const calendarName = calendarNameByUid.get(calendarUID);
      return calendarName ? `${calendarUID} ${truncateWbsReference(calendarName, 9)}` : calendarUID;
    }

    function firstResourceName(resourceNames: string[] | undefined): string {
      if (!resourceNames || resourceNames.length === 0) {
        return "";
      }
      return resourceNames[0];
    }

    function referenceCell(
      task: TaskModel,
      value: string | undefined,
      horizontalAlign: "left" | "center" | "right" = "center"
    ): WbsXlsxCellLike {
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

    function projectInfoRows(
      project: ProjectModel["project"],
      calendarNameByUid: Map<string, string>,
      holidayCount: number,
      columnCount: number,
      startColumnIndex: number,
      startRowNumber: number
    ) {
      const items: Array<{ label: string; value: string | number; fillColor: string }> = [
        { label: "プロジェクト名", value: truncateWbsReference(project.name || "-", 18) || "-", fillColor: fills.summaryAssignment },
        { label: "カレンダ", value: formatCalendarLabel(project.calendarUID, calendarNameByUid), fillColor: fills.summaryAssignment },
        { label: "開始日", value: formatWbsDate(project.startDate), fillColor: fills.summarySchedule },
        { label: "終了日", value: formatWbsDate(project.finishDate), fillColor: fills.summarySchedule },
        { label: "現在日", value: formatWbsDate(project.currentDate), fillColor: fills.summarySchedule },
        { label: "祝日", value: holidayCount, fillColor: fills.summarySchedule }
      ];
      return {
        mergedRanges: [
          layout.range(
            layout.reference(startRowNumber, startColumnIndex),
            layout.reference(startRowNumber, startColumnIndex + 4)
          ),
          ...items.map((_, index) => {
            const rowNumber = startRowNumber + index + 1;
            return [
              layout.range(
                layout.reference(rowNumber, startColumnIndex),
                layout.reference(rowNumber, startColumnIndex + 1)
              ),
              layout.range(
                layout.reference(rowNumber, startColumnIndex + 2),
                layout.reference(rowNumber, startColumnIndex + 4)
              )
            ];
          }).flat()
        ],
        rows: [
          projectBlockHeaderRow(columnCount, startColumnIndex, "プロジェクト情報"),
          ...items.map((item) => projectPairRow(columnCount, startColumnIndex, item.label, item.value, item.fillColor))
        ]
      };
    }

    function legendRows(columnCount: number, startRowNumber: number) {
      const items: Array<{ value: string; fillColor: string }> = [
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
          ...items.map((_, index) => layout.range(
            layout.reference(startRowNumber + index + 1, layout.columnIndex("A")),
            layout.reference(startRowNumber + index + 1, layout.columnIndex("C"))
          ))
        ],
        rows: [
          blockHeaderRow(columnCount, 0, "凡例"),
          ...items.map((item) => mergedLabelRow(columnCount, 0, item.value, item.fillColor))
        ]
      };
    }

    function displaySummaryRows(
      displayDays: number,
      businessDays: number,
      baseDate: string | undefined,
      taskCount: number,
      resourceCount: number,
      assignmentCount: number,
      calendarCount: number,
      columnCount: number,
      startColumnIndex = 5,
      startRowNumber = 5,
      displayDaysBeforeBaseDate?: number,
      displayDaysAfterBaseDate?: number,
      useBusinessDaysForDisplayRange?: boolean,
      useBusinessDaysForProgressBand?: boolean
    ) {
      const displayWeeks = displayDays > 0 ? Math.ceil(displayDays / 7) : 0;
      const scheduleItems: Array<{ label: string; value: string | number; fillColor: string }> = [
        { label: "表示日", value: displayDays, fillColor: fills.summarySchedule },
        { label: "表示週", value: displayWeeks, fillColor: fills.summarySchedule },
        { label: "営業日", value: businessDays, fillColor: fills.summarySchedule },
        { label: "前日数", value: displayDaysBeforeBaseDate ?? "-", fillColor: fills.summarySchedule },
        { label: "後日数", value: displayDaysAfterBaseDate ?? "-", fillColor: fills.summarySchedule },
        { label: "表示", value: useBusinessDaysForDisplayRange ? "営業日" : "暦日", fillColor: fills.summarySchedule },
        { label: "進捗", value: useBusinessDaysForProgressBand ? "営業日" : "暦日", fillColor: fills.summarySchedule },
        { label: "基準日", value: (baseDate || "-").slice(0, 10), fillColor: fills.summarySchedule }
      ];
      const countItems: Array<{ label: string; value: string | number; fillColor: string }> = [
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
        layout.range(
          layout.reference(startRowNumber, startColumnIndex),
          layout.reference(startRowNumber, startColumnIndex + 2)
        )
      ];
      for (let index = 1; index < blockRows.length; index += 1) {
        const rowNumber = startRowNumber + index;
        mergedRanges.push(layout.range(
          layout.reference(rowNumber, startColumnIndex + 1),
          layout.reference(rowNumber, startColumnIndex + 2)
        ));
      }
      return {
        mergedRanges,
        rows: blockRows
      };
    }

    function formatWbsDate(value: string | undefined): string {
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

  (globalThis as typeof globalThis & {
    __mikuprojectWbsXlsxSections?: {
      createWbsXlsxSectionsHelper: (config: WbsXlsxSectionsConfig) => WbsXlsxSectionsHelper;
    };
  }).__mikuprojectWbsXlsxSections = {
    createWbsXlsxSectionsHelper
  };
})();
