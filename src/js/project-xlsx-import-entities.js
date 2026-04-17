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
    function importTasksSheet(workbook, model, changes) {
        const tasksSheet = workbook.sheets.find((sheet) => sheet.name === "Tasks");
        if (!tasksSheet) {
            return;
        }
        const columnIndexByLabel = projectXlsxImportUtil.readHeaderMap(tasksSheet, HEADER_ROW_INDEX);
        const uidColumnIndex = columnIndexByLabel.get("UID");
        if (uidColumnIndex === undefined) {
            return;
        }
        const taskByUid = new Map(model.tasks.map((task) => [task.uid, task]));
        for (const row of tasksSheet.rows.slice(DATA_ROW_START_INDEX)) {
            const uid = projectXlsxImportUtil.readStringCell(row.cells[uidColumnIndex]);
            if (!uid) {
                continue;
            }
            const task = taskByUid.get(uid);
            if (!task) {
                continue;
            }
            const taskLabel = task.name;
            projectXlsxImportUtil.assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "name", "Name", projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Name")));
            projectXlsxImportUtil.assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "start", "Start", projectXlsxImportUtil.normalizeDateTimeInput(projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Start"))));
            projectXlsxImportUtil.assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "finish", "Finish", projectXlsxImportUtil.normalizeDateTimeInput(projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Finish"))));
            projectXlsxImportUtil.assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "duration", "Duration", projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Duration")));
            projectXlsxImportUtil.assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "percentComplete", "PercentComplete", projectXlsxImportUtil.readNumberCellAt(row.cells, columnIndexByLabel.get("PercentComplete")));
            projectXlsxImportUtil.assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "percentWorkComplete", "PercentWorkComplete", projectXlsxImportUtil.readNumberCellAt(row.cells, columnIndexByLabel.get("PercentWorkComplete")));
            projectXlsxImportUtil.assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "milestone", "Milestone", projectXlsxImportUtil.readBooleanCellAt(row.cells, columnIndexByLabel.get("Milestone")));
            projectXlsxImportUtil.assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "summary", "Summary", projectXlsxImportUtil.readBooleanCellAt(row.cells, columnIndexByLabel.get("Summary")));
            projectXlsxImportUtil.assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "critical", "Critical", projectXlsxImportUtil.readBooleanCellAt(row.cells, columnIndexByLabel.get("Critical")));
            projectXlsxImportUtil.assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "calendarUID", "CalendarUID", projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("CalendarUID")));
            projectXlsxImportUtil.assignPredecessorsIfChanged(changes, task.uid, taskLabel, task, projectXlsxImportUtil.parsePredecessorCell(projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Predecessors"))));
            projectXlsxImportUtil.assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "notes", "Notes", projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Notes")));
        }
    }
    function importTasksSheetAsTasks(workbook) {
        var _a, _b, _c, _d;
        const tasksSheet = workbook.sheets.find((sheet) => sheet.name === "Tasks");
        if (!tasksSheet) {
            return [];
        }
        const columnIndexByLabel = projectXlsxImportUtil.readHeaderMap(tasksSheet, HEADER_ROW_INDEX);
        const uidColumnIndex = columnIndexByLabel.get("UID");
        if (uidColumnIndex === undefined) {
            return [];
        }
        const tasks = [];
        for (const [rowIndex, row] of tasksSheet.rows.slice(DATA_ROW_START_INDEX).entries()) {
            const uid = projectXlsxImportUtil.readStringCell(row.cells[uidColumnIndex]);
            if (!uid) {
                continue;
            }
            tasks.push({
                uid,
                id: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("ID")) || String(rowIndex + 1),
                name: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Name")) || uid,
                outlineLevel: projectXlsxImportUtil.readNumberCellAt(row.cells, columnIndexByLabel.get("OutlineLevel")) || 1,
                outlineNumber: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("OutlineNumber")) || String(rowIndex + 1),
                wbs: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("WBS")) || undefined,
                start: projectXlsxImportUtil.normalizeDateTimeInput(projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Start"))) || "",
                finish: projectXlsxImportUtil.normalizeDateTimeInput(projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Finish"))) || "",
                duration: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Duration")) || "",
                milestone: (_a = projectXlsxImportUtil.readBooleanCellAt(row.cells, columnIndexByLabel.get("Milestone"))) !== null && _a !== void 0 ? _a : false,
                summary: (_b = projectXlsxImportUtil.readBooleanCellAt(row.cells, columnIndexByLabel.get("Summary"))) !== null && _b !== void 0 ? _b : false,
                critical: (_c = projectXlsxImportUtil.readBooleanCellAt(row.cells, columnIndexByLabel.get("Critical"))) !== null && _c !== void 0 ? _c : undefined,
                percentComplete: projectXlsxImportUtil.readNumberCellAt(row.cells, columnIndexByLabel.get("PercentComplete")) || 0,
                percentWorkComplete: (_d = projectXlsxImportUtil.readNumberCellAt(row.cells, columnIndexByLabel.get("PercentWorkComplete"))) !== null && _d !== void 0 ? _d : undefined,
                calendarUID: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("CalendarUID")) || undefined,
                notes: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Notes")) || undefined,
                predecessors: projectXlsxImportUtil.parsePredecessorCell(projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Predecessors"))) || [],
                extendedAttributes: [],
                baselines: [],
                timephasedData: []
            });
        }
        return tasks;
    }
    function importResourcesSheet(workbook, model, changes) {
        const resourcesSheet = workbook.sheets.find((sheet) => sheet.name === "Resources");
        if (!resourcesSheet) {
            return;
        }
        const columnIndexByLabel = projectXlsxImportUtil.readHeaderMap(resourcesSheet, HEADER_ROW_INDEX);
        const uidColumnIndex = columnIndexByLabel.get("UID");
        if (uidColumnIndex === undefined) {
            return;
        }
        const resourceByUid = new Map(model.resources.map((resource) => [resource.uid, resource]));
        for (const row of resourcesSheet.rows.slice(DATA_ROW_START_INDEX)) {
            const uid = projectXlsxImportUtil.readStringCell(row.cells[uidColumnIndex]);
            if (!uid) {
                continue;
            }
            const resource = resourceByUid.get(uid);
            if (!resource) {
                continue;
            }
            const resourceLabel = resource.name;
            projectXlsxImportUtil.assignIfChanged(changes, "resources", resource.uid, resourceLabel, resource, "name", "Name", projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Name")));
            projectXlsxImportUtil.assignIfChanged(changes, "resources", resource.uid, resourceLabel, resource, "group", "Group", projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Group")));
            projectXlsxImportUtil.assignIfChanged(changes, "resources", resource.uid, resourceLabel, resource, "maxUnits", "MaxUnits", projectXlsxImportUtil.readNumberCellAt(row.cells, columnIndexByLabel.get("MaxUnits")));
            projectXlsxImportUtil.assignIfChanged(changes, "resources", resource.uid, resourceLabel, resource, "calendarUID", "CalendarUID", projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("CalendarUID")));
        }
    }
    function importResourcesSheetAsResources(workbook) {
        var _a, _b, _c;
        const resourcesSheet = workbook.sheets.find((sheet) => sheet.name === "Resources");
        if (!resourcesSheet) {
            return [];
        }
        const columnIndexByLabel = projectXlsxImportUtil.readHeaderMap(resourcesSheet, HEADER_ROW_INDEX);
        const uidColumnIndex = columnIndexByLabel.get("UID");
        if (uidColumnIndex === undefined) {
            return [];
        }
        const resources = [];
        for (const [rowIndex, row] of resourcesSheet.rows.slice(DATA_ROW_START_INDEX).entries()) {
            const uid = projectXlsxImportUtil.readStringCell(row.cells[uidColumnIndex]);
            if (!uid) {
                continue;
            }
            resources.push({
                uid,
                id: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("ID")) || String(rowIndex + 1),
                name: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Name")) || uid,
                type: (_a = projectXlsxImportUtil.readNumberCellAt(row.cells, columnIndexByLabel.get("Type"))) !== null && _a !== void 0 ? _a : undefined,
                initials: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Initials")) || undefined,
                group: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Group")) || undefined,
                maxUnits: (_b = projectXlsxImportUtil.readNumberCellAt(row.cells, columnIndexByLabel.get("MaxUnits"))) !== null && _b !== void 0 ? _b : undefined,
                calendarUID: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("CalendarUID")) || undefined,
                standardRate: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("StandardRate")) || undefined,
                overtimeRate: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("OvertimeRate")) || undefined,
                costPerUse: (_c = projectXlsxImportUtil.readNumberCellAt(row.cells, columnIndexByLabel.get("CostPerUse"))) !== null && _c !== void 0 ? _c : undefined,
                work: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Work")) || undefined,
                actualWork: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("ActualWork")) || undefined,
                remainingWork: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("RemainingWork")) || undefined,
                extendedAttributes: [],
                baselines: [],
                timephasedData: []
            });
        }
        return resources;
    }
    function importAssignmentsSheet(workbook, model, changes) {
        const assignmentsSheet = workbook.sheets.find((sheet) => sheet.name === "Assignments");
        if (!assignmentsSheet) {
            return;
        }
        const columnIndexByLabel = projectXlsxImportUtil.readHeaderMap(assignmentsSheet, HEADER_ROW_INDEX);
        const uidColumnIndex = columnIndexByLabel.get("UID");
        if (uidColumnIndex === undefined) {
            return;
        }
        const assignmentByUid = new Map(model.assignments.map((assignment) => [assignment.uid, assignment]));
        for (const row of assignmentsSheet.rows.slice(DATA_ROW_START_INDEX)) {
            const uid = projectXlsxImportUtil.readStringCell(row.cells[uidColumnIndex]);
            if (!uid) {
                continue;
            }
            const assignment = assignmentByUid.get(uid);
            if (!assignment) {
                continue;
            }
            const assignmentLabel = `TaskUID=${assignment.taskUid}`;
            projectXlsxImportUtil.assignIfChanged(changes, "assignments", assignment.uid, assignmentLabel, assignment, "units", "Units", projectXlsxImportUtil.readNumberCellAt(row.cells, columnIndexByLabel.get("Units")));
            projectXlsxImportUtil.assignIfChanged(changes, "assignments", assignment.uid, assignmentLabel, assignment, "work", "Work", projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Work")));
            projectXlsxImportUtil.assignIfChanged(changes, "assignments", assignment.uid, assignmentLabel, assignment, "percentWorkComplete", "PercentWorkComplete", projectXlsxImportUtil.readNumberCellAt(row.cells, columnIndexByLabel.get("PercentWorkComplete")));
        }
    }
    function importAssignmentsSheetAsAssignments(workbook) {
        var _a, _b;
        const assignmentsSheet = workbook.sheets.find((sheet) => sheet.name === "Assignments");
        if (!assignmentsSheet) {
            return [];
        }
        const columnIndexByLabel = projectXlsxImportUtil.readHeaderMap(assignmentsSheet, HEADER_ROW_INDEX);
        const uidColumnIndex = columnIndexByLabel.get("UID");
        if (uidColumnIndex === undefined) {
            return [];
        }
        const assignments = [];
        for (const row of assignmentsSheet.rows.slice(DATA_ROW_START_INDEX)) {
            const uid = projectXlsxImportUtil.readStringCell(row.cells[uidColumnIndex]);
            if (!uid) {
                continue;
            }
            const taskUid = projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("TaskUID"));
            const resourceUid = projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("ResourceUID"));
            if (!taskUid || !resourceUid) {
                continue;
            }
            assignments.push({
                uid,
                taskUid,
                resourceUid,
                start: projectXlsxImportUtil.normalizeDateTimeInput(projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Start"))) || undefined,
                finish: projectXlsxImportUtil.normalizeDateTimeInput(projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Finish"))) || undefined,
                units: (_a = projectXlsxImportUtil.readNumberCellAt(row.cells, columnIndexByLabel.get("Units"))) !== null && _a !== void 0 ? _a : undefined,
                work: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("Work")) || undefined,
                actualWork: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("ActualWork")) || undefined,
                remainingWork: projectXlsxImportUtil.readStringCellAt(row.cells, columnIndexByLabel.get("RemainingWork")) || undefined,
                percentWorkComplete: (_b = projectXlsxImportUtil.readNumberCellAt(row.cells, columnIndexByLabel.get("PercentWorkComplete"))) !== null && _b !== void 0 ? _b : undefined,
                extendedAttributes: [],
                baselines: [],
                timephasedData: []
            });
        }
        return assignments;
    }
    globalThis.__mikuprojectProjectXlsxImportEntities = {
        importTasksSheet,
        importTasksSheetAsTasks,
        importResourcesSheet,
        importResourcesSheetAsResources,
        importAssignmentsSheet,
        importAssignmentsSheetAsAssignments
    };
})();
