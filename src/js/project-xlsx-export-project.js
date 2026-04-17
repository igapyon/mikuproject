/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const workbookSchema = globalThis.__mikuprojectProjectWorkbookSchema;
    if (!workbookSchema) {
        throw new Error("mikuproject Project Workbook Schema module is not loaded");
    }
    const { PROJECT_FIELD_ORDER } = workbookSchema;
    const projectXlsxExportUtil = globalThis.__mikuprojectProjectXlsxExportUtil;
    if (!projectXlsxExportUtil) {
        throw new Error("mikuproject Project XLSX export util module is not loaded");
    }
    function buildProjectSheet(model) {
        const project = model.project;
        const rows = [
            projectXlsxExportUtil.titleRow("Project", projectXlsxExportUtil.SHEET_THEMES.project.section),
            projectXlsxExportUtil.titleRow("Basic Info", projectXlsxExportUtil.SHEET_THEMES.project.section),
            projectXlsxExportUtil.headerRow(["Field", "Value"], projectXlsxExportUtil.SHEET_THEMES.project.header),
            ...PROJECT_FIELD_ORDER.slice(0, 8).map((field) => projectXlsxExportUtil.keyValueRow(field, projectXlsxExportUtil.readProjectFieldValue(project, field), projectXlsxExportUtil.SHEET_THEMES.project.label)),
            projectXlsxExportUtil.titleRow("Settings", projectXlsxExportUtil.SHEET_THEMES.project.section),
            ...PROJECT_FIELD_ORDER.slice(8).map((field) => projectXlsxExportUtil.keyValueRow(field, projectXlsxExportUtil.readProjectFieldValue(project, field), projectXlsxExportUtil.SHEET_THEMES.project.label))
        ];
        const rowNumberByField = new Map();
        rows.forEach((row, index) => {
            var _a;
            const field = typeof ((_a = row.cells[0]) === null || _a === void 0 ? void 0 : _a.value) === "string" ? row.cells[0].value : undefined;
            if (field) {
                rowNumberByField.set(field, index + 1);
            }
        });
        return {
            name: "Project",
            columns: [{ width: 26 }, { width: 42 }],
            mergedRanges: ["A11:B11"],
            dataValidations: projectXlsxExportUtil.buildBooleanDataValidations([
                projectXlsxExportUtil.buildSingleCellReference(1, rowNumberByField.get("ScheduleFromStart"))
            ]),
            rows
        };
    }
    globalThis.__mikuprojectProjectXlsxExportProject = {
        buildProjectSheet
    };
})();
