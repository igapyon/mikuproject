/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const mikuprojectCoreApiMsproject = globalThis.__mikuprojectCoreApiMsproject;
    if (!mikuprojectCoreApiMsproject) {
        throw new Error("mikuproject core api msproject module is not loaded");
    }
    const mikuprojectCoreApiWorkbook = globalThis.__mikuprojectCoreApiWorkbook;
    if (!mikuprojectCoreApiWorkbook) {
        throw new Error("mikuproject core api workbook module is not loaded");
    }
    function importMsProjectXml(sourceText, mode) {
        if (mode !== "replace") {
            throw new Error("MS Project XML は replace import のみ対応です");
        }
        return {
            kind: "ms_project_xml",
            mode,
            model: mikuprojectCoreApiMsproject.msProject.importFromXml(sourceText),
            warnings: []
        };
    }
    function importXlsx(sourceBytes, mode, baseModel) {
        if (mode === "patch") {
            throw new Error("XLSX は replace または merge import のみ対応です");
        }
        const workbook = mikuprojectCoreApiWorkbook.xlsx.decodeWorkbook(sourceBytes);
        if (mode === "replace") {
            return {
                kind: "xlsx",
                mode,
                model: mikuprojectCoreApiWorkbook.xlsx.importAsProjectModel(workbook),
                warnings: []
            };
        }
        if (!baseModel) {
            throw new Error("XLSX merge import には baseModel が必要です");
        }
        if (typeof mikuprojectCoreApiWorkbook.xlsx.importIntoProjectModelDetailed === "function") {
            return {
                kind: "xlsx",
                mode,
                ...mikuprojectCoreApiWorkbook.xlsx.importIntoProjectModelDetailed(workbook, baseModel),
                warnings: []
            };
        }
        return {
            kind: "xlsx",
            mode,
            model: mikuprojectCoreApiWorkbook.xlsx.importIntoProjectModel(workbook, baseModel),
            changes: [],
            warnings: []
        };
    }
    globalThis.__mikuprojectCoreApiExternalBinary = {
        importMsProjectXml,
        importXlsx
    };
})();
