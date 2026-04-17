/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
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

  function createWbsSheetLayoutHelper(): WbsSheetLayoutHelper {
    return {
      columnName(columnIndex: number): string {
        let current = columnIndex + 1;
        let name = "";
        while (current > 0) {
          const remainder = (current - 1) % 26;
          name = String.fromCharCode(65 + remainder) + name;
          current = Math.floor((current - 1) / 26);
        }
        return name;
      },
      columnIndex(columnReference: string): number {
        const normalized = (columnReference || "").trim().toUpperCase();
        if (!/^[A-Z]+$/.test(normalized)) {
          throw new Error(`Invalid column reference: ${columnReference}`);
        }
        let value = 0;
        for (const character of normalized) {
          value = (value * 26) + (character.charCodeAt(0) - 64);
        }
        return value - 1;
      },
      reference(rowNumber: number, columnIndex: number): string {
        if (!Number.isInteger(rowNumber) || rowNumber <= 0) {
          throw new Error(`Invalid row number: ${rowNumber}`);
        }
        if (!Number.isInteger(columnIndex) || columnIndex < 0) {
          throw new Error(`Invalid column index: ${columnIndex}`);
        }
        return `${this.columnName(columnIndex)}${rowNumber}`;
      },
      parseCellReference(reference: string) {
        const match = /^([A-Z]+)(\d+)$/i.exec((reference || "").trim());
        if (!match) {
          throw new Error(`Invalid cell reference: ${reference}`);
        }
        const rowNumber = Number(match[2]);
        const columnName = match[1].toUpperCase();
        const columnIndex = this.columnIndex(columnName);
        return {
          reference: `${columnName}${rowNumber}`,
          rowNumber,
          rowIndex: rowNumber - 1,
          columnName,
          columnIndex
        };
      },
      range(startReference: string, endReference: string): string {
        return `${startReference}:${endReference}`;
      },
      describeCell(reference: string): string {
        const cell = this.parseCellReference(reference);
        return `${cell.reference} (row ${cell.rowNumber}, rowIndex ${cell.rowIndex}, column ${cell.columnName}, columnIndex ${cell.columnIndex})`;
      },
      logCell(reference: string, label?: string, logger?: (...args: unknown[]) => void): string {
        const message = label
          ? `${label}: ${this.describeCell(reference)}`
          : this.describeCell(reference);
        (logger || console.log)(message);
        return message;
      }
    };
  }

  (globalThis as typeof globalThis & {
    __mikuprojectWbsXlsxLayout?: {
      createWbsSheetLayoutHelper: () => WbsSheetLayoutHelper;
    };
  }).__mikuprojectWbsXlsxLayout = {
    createWbsSheetLayoutHelper
  };
})();
