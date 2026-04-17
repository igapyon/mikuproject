/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  type SvgPreviewMode = "daily" | "weekly" | "monthly";

  type SvgPreviewState = {
    currentNativeSvg: string;
    currentWeeklyPreviewSvg: string;
    currentMonthlyPreviewSvg: string;
    currentSvgPreviewMode: SvgPreviewMode;
  };

  const mainPreview = {
    createEmptyState(): SvgPreviewState {
      return {
        currentNativeSvg: "",
        currentWeeklyPreviewSvg: "",
        currentMonthlyPreviewSvg: "",
        currentSvgPreviewMode: "daily"
      };
    },

    renderPreviewMarkup(state: SvgPreviewState): string {
      if (state.currentSvgPreviewMode === "weekly") {
        return state.currentWeeklyPreviewSvg || `<div class="md-preview-empty">Weekly SVG を生成すると、ここにプレビューを表示します。</div>`;
      }
      if (state.currentSvgPreviewMode === "monthly") {
        return state.currentMonthlyPreviewSvg || `<div class="md-preview-empty">Monthly Calendar SVG を生成すると、ここにプレビューを表示します。</div>`;
      }
      return state.currentNativeSvg || `<div class="md-preview-empty">Daily SVG を生成すると、ここにプレビューを表示します。</div>`;
    },

    applyPreviewModeButtonClasses(doc: Document, mode: SvgPreviewMode): void {
      const dailyButton = doc.getElementById("previewDailySvgBtn");
      const weeklyButton = doc.getElementById("previewWeeklySvgBtn");
      const monthlyButton = doc.getElementById("previewMonthlySvgBtn");
      if (!(dailyButton instanceof HTMLButtonElement) || !(weeklyButton instanceof HTMLButtonElement) || !(monthlyButton instanceof HTMLButtonElement)) {
        throw new Error("SVG preview mode buttons are not available");
      }
      dailyButton.classList.toggle("md-button--primary", mode === "daily");
      dailyButton.classList.toggle("md-button--tonal", mode !== "daily");
      weeklyButton.classList.toggle("md-button--primary", mode === "weekly");
      weeklyButton.classList.toggle("md-button--tonal", mode !== "weekly");
      monthlyButton.classList.toggle("md-button--primary", mode === "monthly");
      monthlyButton.classList.toggle("md-button--tonal", mode !== "monthly");
    },

    updateDownloadButtons(doc: Document, hasModel: boolean): void {
      const nativeSvgButton = doc.getElementById("downloadSvgBtn");
      const weeklySvgButton = doc.getElementById("downloadWeeklySvgBtn");
      const monthlyWbsButton = doc.getElementById("downloadMonthlyCalendarSvgBtn");
      if (!(nativeSvgButton instanceof HTMLButtonElement) || !(weeklySvgButton instanceof HTMLButtonElement) || !(monthlyWbsButton instanceof HTMLButtonElement)) {
        throw new Error("SVG download buttons are not available");
      }
      nativeSvgButton.disabled = !hasModel;
      weeklySvgButton.disabled = !hasModel;
      monthlyWbsButton.disabled = !hasModel;
    },

    buildRenderedState(input: {
      model: ProjectModel | null;
      previousState: SvgPreviewState;
      buildWbsOptions: (model: ProjectModel) => {
        holidayDates: string[];
        displayDaysBeforeBaseDate?: number;
        displayDaysAfterBaseDate?: number;
        useBusinessDaysForDisplayRange?: boolean;
        useBusinessDaysForProgressBand?: boolean;
      };
      exportNativeSvg: (
        model: ProjectModel,
        options: {
          holidayDates: string[];
          displayDaysBeforeBaseDate?: number;
          displayDaysAfterBaseDate?: number;
          useBusinessDaysForDisplayRange?: boolean;
          useBusinessDaysForProgressBand?: boolean;
        }
      ) => string;
      exportWeeklyNativeSvg: (
        model: ProjectModel,
        options: {
          holidayDates: string[];
          displayDaysBeforeBaseDate?: number;
          displayDaysAfterBaseDate?: number;
          useBusinessDaysForDisplayRange?: boolean;
          useBusinessDaysForProgressBand?: boolean;
        }
      ) => string;
      exportMonthlyWbsCalendarSvgArchive: (model: ProjectModel) => {
        entries: Array<{ fileName: string; svg: string }>;
      };
    }): SvgPreviewState {
      if (!input.model) {
        return {
          currentNativeSvg: "",
          currentWeeklyPreviewSvg: "",
          currentMonthlyPreviewSvg: "",
          currentSvgPreviewMode: input.previousState.currentSvgPreviewMode
        };
      }
      const wbsOptions = input.buildWbsOptions(input.model);
      const monthlyArchive = input.exportMonthlyWbsCalendarSvgArchive(input.model);
      return {
        currentNativeSvg: input.exportNativeSvg(input.model, wbsOptions),
        currentWeeklyPreviewSvg: input.exportWeeklyNativeSvg(input.model, wbsOptions),
        currentMonthlyPreviewSvg: monthlyArchive.entries.length > 0 ? monthlyArchive.entries.map((entry) => entry.svg).join("") : "",
        currentSvgPreviewMode: input.previousState.currentSvgPreviewMode
      };
    },

    setMode(state: SvgPreviewState, mode: SvgPreviewMode): SvgPreviewState {
      return {
        ...state,
        currentSvgPreviewMode: mode
      };
    }
  };

  (globalThis as typeof globalThis & {
    __mikuprojectMainPreview?: typeof mainPreview;
  }).__mikuprojectMainPreview = mainPreview;
})();
