/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const mikuprojectXml = globalThis.__mikuprojectXml;
    if (!mikuprojectXml) {
        throw new Error("mikuproject XML module is not loaded");
    }
    const mikuprojectCoreApiMsprojectAi = globalThis.__mikuprojectCoreApiMsprojectAi;
    if (!mikuprojectCoreApiMsprojectAi) {
        throw new Error("mikuproject core api msproject ai module is not loaded");
    }
    globalThis.__mikuprojectCoreApiMsproject = {
        samples: {
            getSampleXml: () => mikuprojectXml.SAMPLE_XML,
            getSampleProjectDraftView: () => mikuprojectXml.SAMPLE_PROJECT_DRAFT_VIEW
        },
        projectModel: {
            normalize: mikuprojectXml.normalizeProjectModel,
            validate: mikuprojectXml.validateProjectModel
        },
        msProject: {
            parseXmlDocument: mikuprojectXml.parseXmlDocument,
            importFromXml: mikuprojectXml.importMsProjectXml,
            exportToXml: mikuprojectXml.exportMsProjectXml,
            importFromCsvParentId: mikuprojectXml.importCsvParentId,
            exportToCsvParentId: mikuprojectXml.exportCsvParentId
        },
        aiViews: mikuprojectCoreApiMsprojectAi.aiViews,
        mermaid: mikuprojectCoreApiMsprojectAi.mermaid
    };
})();
