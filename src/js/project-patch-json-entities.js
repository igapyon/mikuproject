/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const projectPatchJsonUtil = globalThis.__mikuprojectProjectPatchJsonUtil;
    if (!projectPatchJsonUtil) {
        throw new Error("mikuproject Project Patch JSON util module is not loaded");
    }
    function warnUnsupportedPatchKeys(operation, allowedKeys, opName, uid, label, warnings) {
        warnUnsupportedPatchKeysForScope(operation, allowedKeys, opName, "tasks", uid, label, warnings);
    }
    function warnUnsupportedPatchKeysForScope(operation, allowedKeys, opName, scope, uid, label, warnings) {
        Object.keys(operation).forEach((key) => {
            if (!allowedKeys.has(key)) {
                warnings.push({
                    message: `${opName} の未対応 key は無視します: ${key}`,
                    scope,
                    uid,
                    label
                });
            }
        });
    }
    function applyAddAssignmentOperation(operation, model, changes, warnings, operationIndex) {
        const uid = String(operation.uid || "").trim();
        if (!uid) {
            warnings.push({ message: `add_assignment の uid がありません: operations[${operationIndex}]` });
            return undefined;
        }
        if (model.assignments.some((item) => item.uid === uid)) {
            warnings.push({ message: `add_assignment の uid が既存 assignment と重複しています: ${uid}`, scope: "assignments", uid, label: uid });
            return undefined;
        }
        const taskUid = String(operation.task_uid || "").trim();
        if (!taskUid) {
            warnings.push({ message: `add_assignment.task_uid がありません: ${uid}`, scope: "assignments", uid, label: uid });
            return undefined;
        }
        const task = model.tasks.find((item) => item.uid === taskUid);
        if (!task) {
            warnings.push({ message: `add_assignment.task_uid が既存 task を指していません: ${uid} -> ${taskUid}`, scope: "assignments", uid, label: uid });
            return undefined;
        }
        const resourceUid = String(operation.resource_uid || "").trim();
        if (!resourceUid) {
            warnings.push({ message: `add_assignment.resource_uid がありません: ${uid}`, scope: "assignments", uid, label: uid });
            return undefined;
        }
        const resource = model.resources.find((item) => item.uid === resourceUid);
        if (!resource) {
            warnings.push({ message: `add_assignment.resource_uid が既存 resource を指していません: ${uid} -> ${resourceUid}`, scope: "assignments", uid, label: uid });
            return undefined;
        }
        warnUnsupportedPatchKeysForScope(operation, new Set(["op", "uid", "task_uid", "resource_uid", "start", "finish", "units", "work", "percent_work_complete"]), "add_assignment", "assignments", uid, uid, warnings);
        const normalizedStart = operation.start === undefined
            ? undefined
            : projectPatchJsonUtil.normalizePatchedPlainDateTime(operation.start, "start", model.project);
        if (operation.start !== undefined && normalizedStart === undefined) {
            warnings.push({ message: `add_assignment.start の日付形式が解釈できません: ${uid}`, scope: "assignments", uid, label: uid });
            return undefined;
        }
        const normalizedFinish = operation.finish === undefined
            ? undefined
            : projectPatchJsonUtil.normalizePatchedPlainDateTime(operation.finish, "finish", model.project);
        if (operation.finish !== undefined && normalizedFinish === undefined) {
            warnings.push({ message: `add_assignment.finish の日付形式が解釈できません: ${uid}`, scope: "assignments", uid, label: uid });
            return undefined;
        }
        if (normalizedStart && normalizedFinish && normalizedStart > normalizedFinish) {
            warnings.push({ message: `add_assignment.start が finish より後です: ${uid}`, scope: "assignments", uid, label: uid });
            return undefined;
        }
        if (operation.units !== undefined && (typeof operation.units !== "number" || !Number.isFinite(operation.units) || operation.units < 0)) {
            warnings.push({ message: `add_assignment.units は 0 以上の数値が必要です: ${uid}`, scope: "assignments", uid, label: uid });
            return undefined;
        }
        const work = operation.work === undefined ? undefined : String(operation.work || "").trim();
        if (operation.work !== undefined && !work) {
            warnings.push({ message: `add_assignment.work は空でない文字列が必要です: ${uid}`, scope: "assignments", uid, label: uid });
            return undefined;
        }
        if (operation.percent_work_complete !== undefined) {
            const percentWorkComplete = operation.percent_work_complete;
            if (typeof percentWorkComplete !== "number" || !Number.isFinite(percentWorkComplete) || percentWorkComplete < 0 || percentWorkComplete > 100) {
                warnings.push({ message: `add_assignment.percent_work_complete は 0 以上 100 以下の数値が必要です: ${uid}`, scope: "assignments", uid, label: uid });
                return undefined;
            }
        }
        const addedAssignment = {
            uid,
            taskUid,
            resourceUid,
            start: normalizedStart,
            finish: normalizedFinish,
            units: operation.units,
            work,
            percentWorkComplete: operation.percent_work_complete,
            extendedAttributes: [],
            baselines: [],
            timephasedData: []
        };
        model.assignments.push(addedAssignment);
        changes.push({ scope: "assignments", uid, label: uid, field: "taskUid", before: undefined, after: `${task.uid} (${task.name || task.uid})` });
        changes.push({ scope: "assignments", uid, label: uid, field: "resourceUid", before: undefined, after: `${resource.uid} (${resource.name || resource.uid})` });
        if (normalizedStart !== undefined) {
            changes.push({ scope: "assignments", uid, label: uid, field: "start", before: undefined, after: normalizedStart });
        }
        if (normalizedFinish !== undefined) {
            changes.push({ scope: "assignments", uid, label: uid, field: "finish", before: undefined, after: normalizedFinish });
        }
        if (operation.units !== undefined) {
            changes.push({ scope: "assignments", uid, label: uid, field: "units", before: undefined, after: operation.units });
        }
        if (work !== undefined) {
            changes.push({ scope: "assignments", uid, label: uid, field: "work", before: undefined, after: work });
        }
        if (operation.percent_work_complete !== undefined) {
            changes.push({ scope: "assignments", uid, label: uid, field: "percentWorkComplete", before: undefined, after: operation.percent_work_complete });
        }
        return addedAssignment;
    }
    function applyAddResourceOperation(operation, model, changes, warnings, operationIndex) {
        var _a, _b, _c, _d;
        const uid = String(operation.uid || "").trim();
        if (!uid) {
            warnings.push({ message: `add_resource の uid がありません: operations[${operationIndex}]` });
            return undefined;
        }
        if (model.resources.some((item) => item.uid === uid)) {
            warnings.push({ message: `add_resource の uid が既存 resource と重複しています: ${uid}`, scope: "resources", uid, label: uid });
            return undefined;
        }
        const name = String(operation.name || "").trim();
        if (!name) {
            warnings.push({ message: `add_resource.name は空でない文字列が必要です: ${uid}`, scope: "resources", uid, label: uid });
            return undefined;
        }
        warnUnsupportedPatchKeysForScope(operation, new Set(["op", "uid", "name", "initials", "group", "calendar_uid", "max_units", "standard_rate", "overtime_rate", "cost_per_use", "percent_work_complete"]), "add_resource", "resources", uid, name, warnings);
        let nextCalendarUid;
        if (operation.calendar_uid !== undefined) {
            if (typeof operation.calendar_uid !== "string") {
                warnings.push({ message: `add_resource.calendar_uid は文字列が必要です: ${uid}`, scope: "resources", uid, label: name });
                return undefined;
            }
            nextCalendarUid = operation.calendar_uid.trim() || undefined;
            if (nextCalendarUid && !model.calendars.some((calendar) => calendar.uid === nextCalendarUid)) {
                warnings.push({ message: `add_resource.calendar_uid が既存 calendar を指していません: ${uid} -> ${nextCalendarUid}`, scope: "resources", uid, label: name });
                return undefined;
            }
        }
        if (operation.max_units !== undefined && (typeof operation.max_units !== "number" || !Number.isFinite(operation.max_units) || operation.max_units < 0)) {
            warnings.push({ message: `add_resource.max_units は 0 以上の数値が必要です: ${uid}`, scope: "resources", uid, label: name });
            return undefined;
        }
        if (operation.cost_per_use !== undefined && (typeof operation.cost_per_use !== "number" || !Number.isFinite(operation.cost_per_use) || operation.cost_per_use < 0)) {
            warnings.push({ message: `add_resource.cost_per_use は 0 以上の数値が必要です: ${uid}`, scope: "resources", uid, label: name });
            return undefined;
        }
        if (operation.percent_work_complete !== undefined && (typeof operation.percent_work_complete !== "number" || !Number.isFinite(operation.percent_work_complete) || operation.percent_work_complete < 0 || operation.percent_work_complete > 100)) {
            warnings.push({ message: `add_resource.percent_work_complete は 0 以上 100 以下の数値が必要です: ${uid}`, scope: "resources", uid, label: name });
            return undefined;
        }
        if (operation.initials !== undefined && typeof operation.initials !== "string") {
            warnings.push({ message: `add_resource.initials は文字列が必要です: ${uid}`, scope: "resources", uid, label: name });
            return undefined;
        }
        if (operation.group !== undefined && typeof operation.group !== "string") {
            warnings.push({ message: `add_resource.group は文字列が必要です: ${uid}`, scope: "resources", uid, label: name });
            return undefined;
        }
        if (operation.standard_rate !== undefined && typeof operation.standard_rate !== "string") {
            warnings.push({ message: `add_resource.standard_rate は文字列が必要です: ${uid}`, scope: "resources", uid, label: name });
            return undefined;
        }
        if (operation.overtime_rate !== undefined && typeof operation.overtime_rate !== "string") {
            warnings.push({ message: `add_resource.overtime_rate は文字列が必要です: ${uid}`, scope: "resources", uid, label: name });
            return undefined;
        }
        const addedResource = {
            uid,
            id: "",
            name,
            initials: ((_a = operation.initials) === null || _a === void 0 ? void 0 : _a.trim()) || undefined,
            group: ((_b = operation.group) === null || _b === void 0 ? void 0 : _b.trim()) || undefined,
            maxUnits: operation.max_units,
            calendarUID: nextCalendarUid,
            standardRate: ((_c = operation.standard_rate) === null || _c === void 0 ? void 0 : _c.trim()) || undefined,
            overtimeRate: ((_d = operation.overtime_rate) === null || _d === void 0 ? void 0 : _d.trim()) || undefined,
            costPerUse: operation.cost_per_use,
            percentWorkComplete: operation.percent_work_complete,
            extendedAttributes: [],
            baselines: [],
            timephasedData: []
        };
        model.resources.push(addedResource);
        changes.push({ scope: "resources", uid, label: name, field: "name", before: undefined, after: name });
        if (addedResource.initials !== undefined) {
            changes.push({ scope: "resources", uid, label: name, field: "initials", before: undefined, after: addedResource.initials });
        }
        if (addedResource.group !== undefined) {
            changes.push({ scope: "resources", uid, label: name, field: "group", before: undefined, after: addedResource.group });
        }
        if (addedResource.calendarUID !== undefined) {
            changes.push({ scope: "resources", uid, label: name, field: "calendarUID", before: undefined, after: addedResource.calendarUID });
        }
        if (addedResource.maxUnits !== undefined) {
            changes.push({ scope: "resources", uid, label: name, field: "maxUnits", before: undefined, after: addedResource.maxUnits });
        }
        if (addedResource.standardRate !== undefined) {
            changes.push({ scope: "resources", uid, label: name, field: "standardRate", before: undefined, after: addedResource.standardRate });
        }
        if (addedResource.overtimeRate !== undefined) {
            changes.push({ scope: "resources", uid, label: name, field: "overtimeRate", before: undefined, after: addedResource.overtimeRate });
        }
        if (addedResource.costPerUse !== undefined) {
            changes.push({ scope: "resources", uid, label: name, field: "costPerUse", before: undefined, after: addedResource.costPerUse });
        }
        if (addedResource.percentWorkComplete !== undefined) {
            changes.push({ scope: "resources", uid, label: name, field: "percentWorkComplete", before: undefined, after: addedResource.percentWorkComplete });
        }
        return addedResource;
    }
    function applyAddCalendarOperation(operation, model, changes, warnings, operationIndex) {
        const uid = String(operation.uid || "").trim();
        if (!uid) {
            warnings.push({ message: `add_calendar の uid がありません: operations[${operationIndex}]` });
            return undefined;
        }
        if (model.calendars.some((item) => item.uid === uid)) {
            warnings.push({ message: `add_calendar の uid が既存 calendar と重複しています: ${uid}`, scope: "calendars", uid, label: uid });
            return undefined;
        }
        const name = String(operation.name || "").trim();
        if (!name) {
            warnings.push({ message: `add_calendar.name は空でない文字列が必要です: ${uid}`, scope: "calendars", uid, label: uid });
            return undefined;
        }
        warnUnsupportedPatchKeysForScope(operation, new Set(["op", "uid", "name", "is_base_calendar", "base_calendar_uid"]), "add_calendar", "calendars", uid, name, warnings);
        if (operation.is_base_calendar !== undefined && typeof operation.is_base_calendar !== "boolean") {
            warnings.push({ message: `add_calendar.is_base_calendar は boolean が必要です: ${uid}`, scope: "calendars", uid, label: name });
            return undefined;
        }
        let nextBaseCalendarUid;
        if (operation.base_calendar_uid !== undefined) {
            if (typeof operation.base_calendar_uid !== "string") {
                warnings.push({ message: `add_calendar.base_calendar_uid は文字列が必要です: ${uid}`, scope: "calendars", uid, label: name });
                return undefined;
            }
            nextBaseCalendarUid = operation.base_calendar_uid.trim() || undefined;
            if (nextBaseCalendarUid === uid) {
                warnings.push({ message: `add_calendar.base_calendar_uid は自身を指せません: ${uid}`, scope: "calendars", uid, label: name });
                return undefined;
            }
            if (nextBaseCalendarUid && !model.calendars.some((calendar) => calendar.uid === nextBaseCalendarUid)) {
                warnings.push({ message: `add_calendar.base_calendar_uid が既存 calendar を指していません: ${uid} -> ${nextBaseCalendarUid}`, scope: "calendars", uid, label: name });
                return undefined;
            }
        }
        const addedCalendar = {
            uid,
            name,
            isBaseCalendar: Boolean(operation.is_base_calendar),
            baseCalendarUID: nextBaseCalendarUid,
            weekDays: [],
            exceptions: [],
            workWeeks: []
        };
        model.calendars.push(addedCalendar);
        changes.push({ scope: "calendars", uid, label: name, field: "name", before: undefined, after: name });
        if (addedCalendar.isBaseCalendar) {
            changes.push({ scope: "calendars", uid, label: name, field: "isBaseCalendar", before: undefined, after: true });
        }
        if (addedCalendar.baseCalendarUID !== undefined) {
            changes.push({ scope: "calendars", uid, label: name, field: "baseCalendarUID", before: undefined, after: addedCalendar.baseCalendarUID });
        }
        return addedCalendar;
    }
    function applyDeleteResourceOperation(operation, model, changes, warnings, operationIndex) {
        const uid = String(operation.uid || "").trim();
        if (!uid) {
            warnings.push({ message: `delete_resource の uid がありません: operations[${operationIndex}]` });
            return;
        }
        const resourceIndex = model.resources.findIndex((item) => item.uid === uid);
        if (resourceIndex < 0) {
            warnings.push({ message: `delete_resource の uid が既存 resource を指していません: ${uid}`, scope: "resources", uid, label: uid });
            return;
        }
        const resource = model.resources[resourceIndex];
        const referencingAssignments = model.assignments.filter((assignment) => assignment.resourceUid === uid);
        if (referencingAssignments.length > 0) {
            const assignmentText = referencingAssignments.map((assignment) => assignment.uid).join(", ");
            warnings.push({ message: `delete_resource first cut では assignment がある resource は削除できません: ${uid} (assignments=${assignmentText})`, scope: "resources", uid, label: resource.name || resource.uid });
            return;
        }
        model.resources.splice(resourceIndex, 1);
        changes.push({ scope: "resources", uid, label: resource.name || resource.uid, field: "name", before: resource.name || resource.uid, after: "(deleted)" });
    }
    function applyDeleteCalendarOperation(operation, model, changes, warnings, operationIndex) {
        const uid = String(operation.uid || "").trim();
        if (!uid) {
            warnings.push({ message: `delete_calendar の uid がありません: operations[${operationIndex}]` });
            return;
        }
        const calendarIndex = model.calendars.findIndex((item) => item.uid === uid);
        if (calendarIndex < 0) {
            warnings.push({ message: `delete_calendar の uid が既存 calendar を指していません: ${uid}`, scope: "calendars", uid, label: uid });
            return;
        }
        const calendar = model.calendars[calendarIndex];
        const projectRef = model.project.calendarUID === uid ? 1 : 0;
        const taskRefs = model.tasks.filter((task) => task.calendarUID === uid).map((task) => task.uid);
        const resourceRefs = model.resources.filter((resource) => resource.calendarUID === uid).map((resource) => resource.uid);
        const baseCalendarRefs = model.calendars.filter((item) => item.baseCalendarUID === uid && item.uid !== uid).map((item) => item.uid);
        if (projectRef || taskRefs.length > 0 || resourceRefs.length > 0 || baseCalendarRefs.length > 0) {
            const blockers = [];
            if (projectRef)
                blockers.push("project=1");
            if (taskRefs.length > 0)
                blockers.push(`tasks=${taskRefs.join(", ")}`);
            if (resourceRefs.length > 0)
                blockers.push(`resources=${resourceRefs.join(", ")}`);
            if (baseCalendarRefs.length > 0)
                blockers.push(`baseRefs=${baseCalendarRefs.join(", ")}`);
            warnings.push({ message: `delete_calendar first cut では参照が残っている calendar は削除できません: ${uid} (${blockers.join("; ")})`, scope: "calendars", uid, label: calendar.name || calendar.uid });
            return;
        }
        model.calendars.splice(calendarIndex, 1);
        changes.push({ scope: "calendars", uid, label: calendar.name || calendar.uid, field: "name", before: calendar.name || calendar.uid, after: "(deleted)" });
    }
    function applyDeleteAssignmentOperation(operation, model, changes, warnings, operationIndex) {
        const uid = String(operation.uid || "").trim();
        if (!uid) {
            warnings.push({ message: `delete_assignment の uid がありません: operations[${operationIndex}]` });
            return;
        }
        const assignmentIndex = model.assignments.findIndex((item) => item.uid === uid);
        if (assignmentIndex < 0) {
            warnings.push({ message: `delete_assignment の uid が既存 assignment を指していません: ${uid}`, scope: "assignments", uid, label: uid });
            return;
        }
        const assignment = model.assignments[assignmentIndex];
        model.assignments.splice(assignmentIndex, 1);
        changes.push({ scope: "assignments", uid, label: uid, field: "taskUid", before: assignment.taskUid, after: "(deleted)" });
        changes.push({ scope: "assignments", uid, label: uid, field: "resourceUid", before: assignment.resourceUid, after: "(deleted)" });
    }
    globalThis.__mikuprojectProjectPatchJsonEntities = {
        warnUnsupportedPatchKeys,
        warnUnsupportedPatchKeysForScope,
        applyAddAssignmentOperation,
        applyAddResourceOperation,
        applyAddCalendarOperation,
        applyDeleteResourceOperation,
        applyDeleteCalendarOperation,
        applyDeleteAssignmentOperation
    };
})();
