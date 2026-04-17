/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const workbookSchema = globalThis.__mikuprojectProjectWorkbookSchema;
    if (!workbookSchema) {
        throw new Error("mikuproject Project Workbook Schema module is not loaded");
    }
    const { HEADER_ROW_INDEX, DATA_ROW_START_INDEX } = workbookSchema;
    const projectXlsxImportUtil = globalThis.__mikuprojectProjectXlsxImportUtil;
    if (!projectXlsxImportUtil) {
        throw new Error("mikuproject Project XLSX import util module is not loaded");
    }
    const projectXlsxImportCalendars = globalThis.__mikuprojectProjectXlsxImportCalendars;
    if (!projectXlsxImportCalendars) {
        throw new Error("mikuproject Project XLSX import calendars module is not loaded");
    }
    const projectXlsxImportEntities = globalThis.__mikuprojectProjectXlsxImportEntities;
    if (!projectXlsxImportEntities) {
        throw new Error("mikuproject Project XLSX import entities module is not loaded");
    }
    const projectXlsxImportProject = globalThis.__mikuprojectProjectXlsxImportProject;
    if (!projectXlsxImportProject) {
        throw new Error("mikuproject Project XLSX import project module is not loaded");
    }
    function importProjectWorkbook(workbook, baseModel) {
        return importProjectWorkbookDetailed(workbook, baseModel).model;
    }
    function importProjectWorkbookAsProjectModel(workbook) {
        const project = projectXlsxImportProject.importProjectSheetAsProjectInfo(workbook);
        const tasks = projectXlsxImportEntities.importTasksSheetAsTasks(workbook);
        const resources = projectXlsxImportEntities.importResourcesSheetAsResources(workbook);
        const assignments = projectXlsxImportEntities.importAssignmentsSheetAsAssignments(workbook);
        const calendars = projectXlsxImportCalendars.importCalendarsSheetAsCalendars(workbook, project);
        projectXlsxImportCalendars.importNonWorkingDaysSheetAsCalendarExceptions(workbook, calendars);
        return {
            project,
            tasks,
            resources,
            assignments,
            calendars
        };
    }
    function importProjectWorkbookDetailed(workbook, baseModel) {
        const nextModel = projectXlsxImportUtil.cloneProjectModel(baseModel);
        const changes = [];
        projectXlsxImportProject.importProjectSheet(workbook, nextModel, changes);
        projectXlsxImportEntities.importTasksSheet(workbook, nextModel, changes);
        projectXlsxImportEntities.importResourcesSheet(workbook, nextModel, changes);
        projectXlsxImportEntities.importAssignmentsSheet(workbook, nextModel, changes);
        projectXlsxImportCalendars.importCalendarsSheet(workbook, nextModel, changes);
        projectXlsxImportCalendars.importNonWorkingDaysSheet(workbook, nextModel, changes);
        return {
            model: nextModel,
            changes
        };
    }
    globalThis.__mikuprojectProjectXlsxImport = {
        importProjectWorkbook,
        importProjectWorkbookAsProjectModel,
        importProjectWorkbookDetailed
    };
})();
