/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    function createNativeSvgFacade(config) {
        return {
            collectWbsHolidayDates: (model) => config.collectWbsHolidayDates(model),
            exportNativeSvg: (model, options) => config.exportNativeSvg(model, options),
            exportWeeklyNativeSvg: (model, options) => config.exportWeeklyNativeSvg(model, options),
            exportMonthlyWbsCalendarSvgArchive: (model) => config.exportMonthlyWbsCalendarSvgArchive(model)
        };
    }
    globalThis.__mikuprojectWbsSvgPublic = {
        createNativeSvgFacade
    };
})();
