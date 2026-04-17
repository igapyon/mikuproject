/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const mainPreviewActions = {
        updateSvgPreviewModeButtons(input) {
            input.applyPreviewModeButtonClasses(input.document, input.mode);
        },
        renderCurrentSvgPreviewMarkup(input) {
            input.setSvgPreviewMarkup(input.renderPreviewMarkup(input.state));
        },
        updateSvgButton(input) {
            input.updateDownloadButtons(input.document, input.hasModel);
        },
        async renderSvgPreview(input) {
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
        setSvgPreviewMode(input) {
            input.setState(input.setMode(input.currentState, input.mode));
            input.updateSvgPreviewModeButtons();
            input.renderCurrentSvgPreviewMarkup();
        }
    };
    globalThis.__mikuprojectMainPreviewActions = mainPreviewActions;
})();
