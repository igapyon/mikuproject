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

  type WbsXlsxRowLike = {
    height?: number;
    cells: WbsXlsxCellLike[];
  };

  type WbsXlsxBaseHelper = {
    emptyRow(columnCount: number, height?: number): WbsXlsxRowLike;
    overlayRows(rows: WbsXlsxRowLike[], startIndex: number, blockRows: WbsXlsxRowLike[], columnCount: number): void;
    sheetTitleRow(title: string, columnCount: number): WbsXlsxRowLike;
    infoRow(text: string, columnCount: number): WbsXlsxRowLike;
    formatWbsExportTimestamp(value: Date): string;
  };

  function createWbsXlsxBaseHelper(): WbsXlsxBaseHelper {
    function hasCellContent(cell: WbsXlsxCellLike | undefined): boolean {
      return !!cell && Object.keys(cell).length > 0;
    }

    function emptyRow(columnCount: number, height = 22): WbsXlsxRowLike {
      return {
        height,
        cells: Array.from({ length: columnCount }, () => ({}))
      };
    }

    function overlayRows(
      rows: WbsXlsxRowLike[],
      startIndex: number,
      blockRows: WbsXlsxRowLike[],
      columnCount: number
    ) {
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

    function sheetTitleRow(title: string, columnCount: number): WbsXlsxRowLike {
      return {
        height: 24,
        cells: [
          {
            value: title,
            bold: true,
            fontSize: 16,
            horizontalAlign: "left"
          },
          ...Array.from({ length: Math.max(0, columnCount - 1) }, () => ({
            fillColor: "#EEF4FA"
          }))
        ]
      };
    }

    function infoRow(text: string, columnCount: number): WbsXlsxRowLike {
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

    function formatWbsExportTimestamp(value: Date): string {
      const year = value.getFullYear();
      const month = String(value.getMonth() + 1).padStart(2, "0");
      const day = String(value.getDate()).padStart(2, "0");
      const hours = String(value.getHours()).padStart(2, "0");
      const minutes = String(value.getMinutes()).padStart(2, "0");
      return `出力日時 ${year}-${month}-${day} ${hours}:${minutes}`;
    }

    return {
      emptyRow,
      overlayRows,
      sheetTitleRow,
      infoRow,
      formatWbsExportTimestamp
    };
  }

  (globalThis as typeof globalThis & {
    __mikuprojectWbsXlsxBase?: {
      createWbsXlsxBaseHelper: () => WbsXlsxBaseHelper;
    };
  }).__mikuprojectWbsXlsxBase = {
    createWbsXlsxBaseHelper
  };
})();
