/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const workbookSchema = globalThis.__mikuprojectProjectWorkbookSchema;
    if (!workbookSchema) {
        throw new Error("mikuproject Project Workbook Schema module is not loaded");
    }
    const { HEADER_ROW_INDEX, DATA_ROW_START_INDEX, PROJECT_FIELD_ORDER, SHEET_HEADERS } = workbookSchema;
    const projectXlsxExportUtil = globalThis.__mikuprojectProjectXlsxExportUtil;
    if (!projectXlsxExportUtil) {
        throw new Error("mikuproject Project XLSX export util module is not loaded");
    }
    const projectXlsxExportEntities = globalThis.__mikuprojectProjectXlsxExportEntities;
    if (!projectXlsxExportEntities) {
        throw new Error("mikuproject Project XLSX export entities module is not loaded");
    }
    const projectXlsxExportProject = globalThis.__mikuprojectProjectXlsxExportProject;
    if (!projectXlsxExportProject) {
        throw new Error("mikuproject Project XLSX export project module is not loaded");
    }
    const projectXlsxExportCalendars = globalThis.__mikuprojectProjectXlsxExportCalendars;
    if (!projectXlsxExportCalendars) {
        throw new Error("mikuproject Project XLSX export calendars module is not loaded");
    }
    function exportProjectWorkbook(model) {
        const projectSheet = projectXlsxExportProject.buildProjectSheet(model);
        const tasksSheet = projectXlsxExportEntities.buildTasksSheet(model);
        const resourcesSheet = projectXlsxExportEntities.buildResourcesSheet(model);
        const assignmentsSheet = projectXlsxExportEntities.buildAssignmentsSheet(model);
        const calendarsSheet = projectXlsxExportCalendars.buildCalendarsSheet(model);
        const nonWorkingDaysSheet = projectXlsxExportCalendars.buildNonWorkingDaysSheet(model);
        return {
            sheets: [
                projectSheet,
                tasksSheet,
                resourcesSheet,
                assignmentsSheet,
                calendarsSheet,
                nonWorkingDaysSheet,
                projectXlsxExportUtil.buildOptionsSheet()
            ]
        };
    }
    globalThis.__mikuprojectProjectXlsxExport = {
        exportProjectWorkbook
    };
})();
