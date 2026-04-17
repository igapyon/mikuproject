/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    function createWbsXlsxBaseHelper() {
        function hasCellContent(cell) {
            return !!cell && Object.keys(cell).length > 0;
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
        function sheetTitleRow(title, columnCount) {
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
        function formatWbsExportTimestamp(value) {
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
    globalThis.__mikuprojectWbsXlsxBase = {
        createWbsXlsxBaseHelper
    };
})();
