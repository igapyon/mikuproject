/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    function createWbsXlsxFacade(config) {
        return {
            collectWbsHolidayDates: (model) => config.collectWbsHolidayDates(model),
            exportWbsWorkbook: (model, options) => config.exportWbsWorkbook(model, options),
            layout: config.layout
        };
    }
    globalThis.__mikuprojectWbsXlsxPublic = {
        createWbsXlsxFacade
    };
})();
