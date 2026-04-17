/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const mikuprojectCoreApiReportAdapters = globalThis.__mikuprojectCoreApiReportAdapters;
    if (!mikuprojectCoreApiReportAdapters) {
        throw new Error("mikuproject core api report adapters module is not loaded");
    }
    globalThis.__mikuprojectCoreApiReportPublic = {
        report: mikuprojectCoreApiReportAdapters.report
    };
})();
