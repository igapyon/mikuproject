/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  const mainPreviewActions = {
    updateSvgPreviewModeButtons(input: {
      document: Document;
      mode: "daily" | "weekly" | "monthly";
      applyPreviewModeButtonClasses: (doc: Document, mode: "daily" | "weekly" | "monthly") => void;
    }): void {
      input.applyPreviewModeButtonClasses(input.document, input.mode);
    },

    renderCurrentSvgPreviewMarkup(input: {
      state: {
        currentNativeSvg: string;
        currentWeeklyPreviewSvg: string;
        currentMonthlyPreviewSvg: string;
        currentSvgPreviewMode: "daily" | "weekly" | "monthly";
      };
      renderPreviewMarkup: (state: {
        currentNativeSvg: string;
        currentWeeklyPreviewSvg: string;
        currentMonthlyPreviewSvg: string;
        currentSvgPreviewMode: "daily" | "weekly" | "monthly";
      }) => string;
      setSvgPreviewMarkup: (markup: string) => void;
    }): void {
      input.setSvgPreviewMarkup(input.renderPreviewMarkup(input.state));
    },

    updateSvgButton(input: {
      document: Document;
      hasModel: boolean;
      updateDownloadButtons: (doc: Document, hasModel: boolean) => void;
    }): void {
      input.updateDownloadButtons(input.document, input.hasModel);
    },

    async renderSvgPreview(input: {
      currentModel: ProjectModel | null;
      currentState: {
        currentNativeSvg: string;
        currentWeeklyPreviewSvg: string;
        currentMonthlyPreviewSvg: string;
        currentSvgPreviewMode: "daily" | "weekly" | "monthly";
      };
      setState: (state: {
        currentNativeSvg: string;
        currentWeeklyPreviewSvg: string;
        currentMonthlyPreviewSvg: string;
        currentSvgPreviewMode: "daily" | "weekly" | "monthly";
      }) => void;
      buildRenderedState: typeof globalThis.__mikuprojectMainPreview.buildRenderedState;
      buildWbsOptions: (model: ProjectModel) => {
        holidayDates: string[];
        displayDaysBeforeBaseDate?: number;
        displayDaysAfterBaseDate?: number;
        useBusinessDaysForDisplayRange?: boolean;
        useBusinessDaysForProgressBand?: boolean;
      };
      exportNativeSvg: typeof globalThis.__mikuprojectNativeSvg.exportNativeSvg;
      exportWeeklyNativeSvg: typeof globalThis.__mikuprojectNativeSvg.exportWeeklyNativeSvg;
      exportMonthlyWbsCalendarSvgArchive: typeof globalThis.__mikuprojectNativeSvg.exportMonthlyWbsCalendarSvgArchive;
      renderCurrentSvgPreviewMarkup: () => void;
      updateSvgButton: () => void;
    }): Promise<void> {
      input.setState(input.buildRenderedState({
        model: input.currentModel,
        previousState: input.currentState,
        buildWbsOptions: input.buildWbsOptions,
        exportNativeSvg: input.exportNativeSvg,
        exportWeeklyNativeSvg: input.exportWeeklyNativeSvg,
        exportMonthlyWbsCalendarSvgArchive: input.exportMonthlyWbsCalendarSvgArchive
      }));
      input.renderCurrentSvgPreviewMarkup();
      input.updateSvgButton();
    },

    setSvgPreviewMode(input: {
      currentState: {
        currentNativeSvg: string;
        currentWeeklyPreviewSvg: string;
        currentMonthlyPreviewSvg: string;
        currentSvgPreviewMode: "daily" | "weekly" | "monthly";
      };
      mode: "daily" | "weekly" | "monthly";
      setMode: (state: {
        currentNativeSvg: string;
        currentWeeklyPreviewSvg: string;
        currentMonthlyPreviewSvg: string;
        currentSvgPreviewMode: "daily" | "weekly" | "monthly";
      }, mode: "daily" | "weekly" | "monthly") => {
        currentNativeSvg: string;
        currentWeeklyPreviewSvg: string;
        currentMonthlyPreviewSvg: string;
        currentSvgPreviewMode: "daily" | "weekly" | "monthly";
      };
      setState: (state: {
        currentNativeSvg: string;
        currentWeeklyPreviewSvg: string;
        currentMonthlyPreviewSvg: string;
        currentSvgPreviewMode: "daily" | "weekly" | "monthly";
      }) => void;
      updateSvgPreviewModeButtons: () => void;
      renderCurrentSvgPreviewMarkup: () => void;
    }): void {
      input.setState(input.setMode(input.currentState, input.mode));
      input.updateSvgPreviewModeButtons();
      input.renderCurrentSvgPreviewMarkup();
    }
  };

  (globalThis as typeof globalThis & {
    __mikuprojectMainPreviewActions?: typeof mainPreviewActions;
  }).__mikuprojectMainPreviewActions = mainPreviewActions;
})();
