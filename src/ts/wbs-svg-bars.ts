/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  type NativeSvgTaskRow = {
    task: TaskModel;
    label: string;
    kind: "phase" | "task" | "milestone";
    startIndex: number | null;
    endIndex: number | null;
    y: number;
  };

  type WbsSvgBarsHelper = {
    renderTaskBar(
      row: NativeSvgTaskRow,
      chartOriginX: number,
      rowY: number,
      cellWidth: number,
      offsets: {
        horizontalInset: number;
        verticalInset: number;
        milestoneHalfWidth: number;
      }
    ): string[];
  };

  function createWbsSvgBarsHelper(): WbsSvgBarsHelper {
    function renderTaskBar(
      row: NativeSvgTaskRow,
      chartOriginX: number,
      rowY: number,
      cellWidth: number,
      offsets: {
        horizontalInset: number;
        verticalInset: number;
        milestoneHalfWidth: number;
      }
    ): string[] {
      if (row.startIndex === null || row.endIndex === null) {
        return [];
      }

      const barX = chartOriginX + (row.startIndex * cellWidth) + offsets.horizontalInset;
      const barWidth = Math.max(10, ((row.endIndex - row.startIndex + 1) * cellWidth) - (offsets.horizontalInset * 2));
      const barY = rowY + offsets.verticalInset;

      if (row.kind === "milestone") {
        const centerX = chartOriginX + (row.startIndex * cellWidth) + (cellWidth / 2);
        const centerY = rowY + 19;
        const isCompleted = (row.task.percentComplete || 0) >= 100;
        const fill = isCompleted ? "#d9efff" : "#ffffff";
        const half = offsets.milestoneHalfWidth;
        return [
          `<polygon points="${centerX},${centerY - half} ${centerX + half},${centerY} ${centerX},${centerY + half} ${centerX - half},${centerY}" fill="${fill}" stroke="#4f95d6" stroke-width="3"/>`
        ];
      }

      if (row.kind === "phase") {
        const lineY = rowY + 19;
        const startX = barX;
        const endX = barX + barWidth;
        const trackStroke = "#8eb9ea";
        const progressStroke = "#2f79d0";
        const phaseStrokeWidth = 3;
        const progressEndX = startX + Math.max(0, Math.min(barWidth, Math.round(barWidth * (Math.max(0, Math.min(100, row.task.percentComplete || 0)) / 100))));
        const parts = [
          `<line x1="${startX}" y1="${lineY}" x2="${endX}" y2="${lineY}" stroke="${trackStroke}" stroke-width="${phaseStrokeWidth}" stroke-linecap="round"/>`
        ];
        if (progressEndX > startX) {
          parts.push(`<line x1="${startX}" y1="${lineY}" x2="${progressEndX}" y2="${lineY}" stroke="${progressStroke}" stroke-width="${phaseStrokeWidth}" stroke-linecap="round"/>`);
        }
        return parts;
      }

      const trackFill = "#d9efff";
      const trackStroke = "#4f95d6";
      const progressFill = "#3f86d8";
      const progressWidth = Math.max(0, Math.min(barWidth, Math.round(barWidth * (Math.max(0, Math.min(100, row.task.percentComplete || 0)) / 100))));
      const parts = [
        `<rect x="${barX}" y="${barY}" width="${barWidth}" height="22" rx="4" fill="${trackFill}" stroke="${trackStroke}" stroke-width="3"/>`
      ];
      if (progressWidth > 0) {
        parts.push(`<rect x="${barX}" y="${barY}" width="${progressWidth}" height="22" rx="4" fill="${progressFill}" stroke="none"/>`);
      }
      return parts;
    }

    return {
      renderTaskBar
    };
  }

  (globalThis as typeof globalThis & {
    __mikuprojectWbsSvgBars?: {
      createWbsSvgBarsHelper: () => WbsSvgBarsHelper;
    };
  }).__mikuprojectWbsSvgBars = {
    createWbsSvgBarsHelper
  };
})();
