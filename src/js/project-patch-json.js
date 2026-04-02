/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const mikuprojectXml = globalThis.__mikuprojectXml;
    if (!mikuprojectXml) {
        throw new Error("mikuproject XML module is not loaded");
    }
    function importProjectPatchJson(documentLike, baseModel) {
        const validation = validatePatchDocument(documentLike);
        const nextModel = cloneProjectModel(baseModel);
        const changes = [];
        const warnings = [...validation.warnings];
        const taskByUid = new Map(nextModel.tasks.map((task) => [task.uid, task]));
        const resourceByUid = new Map(nextModel.resources.map((resource) => [resource.uid, resource]));
        const assignmentByUid = new Map(nextModel.assignments.map((assignment) => [assignment.uid, assignment]));
        const calendarByUid = new Map(nextModel.calendars.map((calendar) => [calendar.uid, calendar]));
        validation.document.operations.forEach((operation, index) => {
            const op = String(operation.op || "").trim();
            if (op === "update_project") {
                if (!operation.fields || typeof operation.fields !== "object" || Array.isArray(operation.fields)) {
                    warnings.push({ message: `update_project.fields がオブジェクトではありません: operations[${index}]`, scope: "project", uid: "project", label: nextModel.project.name || "project" });
                    return;
                }
                applyUpdateProjectOperation(nextModel.project, operation.fields, nextModel, changes, warnings, index);
                return;
            }
            if (op === "add_task") {
                const addedTask = applyAddTaskOperation(operation, nextModel.tasks, nextModel.project, changes, warnings, index);
                if (addedTask) {
                    taskByUid.set(addedTask.uid, addedTask);
                }
                return;
            }
            if (op === "add_resource") {
                const addedResource = applyAddResourceOperation(operation, nextModel, changes, warnings, index);
                if (addedResource) {
                    resourceByUid.set(addedResource.uid, addedResource);
                }
                return;
            }
            if (op === "add_calendar") {
                const addedCalendar = applyAddCalendarOperation(operation, nextModel, changes, warnings, index);
                if (addedCalendar) {
                    calendarByUid.set(addedCalendar.uid, addedCalendar);
                }
                return;
            }
            if (op === "add_assignment") {
                const addedAssignment = applyAddAssignmentOperation(operation, nextModel, changes, warnings, index);
                if (addedAssignment) {
                    assignmentByUid.set(addedAssignment.uid, addedAssignment);
                }
                return;
            }
            if (op === "delete_assignment") {
                applyDeleteAssignmentOperation(operation, nextModel, changes, warnings, index);
                assignmentByUid.clear();
                nextModel.assignments.forEach((assignment) => assignmentByUid.set(assignment.uid, assignment));
                return;
            }
            if (op === "delete_resource") {
                applyDeleteResourceOperation(operation, nextModel, changes, warnings, index);
                resourceByUid.clear();
                nextModel.resources.forEach((resource) => resourceByUid.set(resource.uid, resource));
                return;
            }
            if (op === "delete_calendar") {
                applyDeleteCalendarOperation(operation, nextModel, changes, warnings, index);
                calendarByUid.clear();
                nextModel.calendars.forEach((calendar) => calendarByUid.set(calendar.uid, calendar));
                return;
            }
            if (op === "update_task") {
                const uid = String(operation.uid || "").trim();
                if (!uid) {
                    warnings.push({ message: `update_task の uid がありません: operations[${index}]` });
                    return;
                }
                const task = taskByUid.get(uid);
                if (!task) {
                    warnings.push({ message: `update_task の uid が既存 task を指していません: ${uid}`, scope: "tasks", uid });
                    return;
                }
                if (!operation.fields || typeof operation.fields !== "object" || Array.isArray(operation.fields)) {
                    warnings.push({ message: `update_task.fields がオブジェクトではありません: ${uid}`, scope: "tasks", uid, label: task.name || task.uid });
                    return;
                }
                applyUpdateTaskOperation(task, operation.fields, nextModel, changes, warnings, index);
                return;
            }
            if (op === "update_assignment") {
                const uid = String(operation.uid || "").trim();
                if (!uid) {
                    warnings.push({ message: `update_assignment の uid がありません: operations[${index}]` });
                    return;
                }
                const assignment = assignmentByUid.get(uid);
                if (!assignment) {
                    warnings.push({ message: `update_assignment の uid が既存 assignment を指していません: ${uid}`, scope: "assignments", uid });
                    return;
                }
                if (!operation.fields || typeof operation.fields !== "object" || Array.isArray(operation.fields)) {
                    warnings.push({ message: `update_assignment.fields がオブジェクトではありません: ${uid}`, scope: "assignments", uid, label: assignment.uid });
                    return;
                }
                applyUpdateAssignmentOperation(assignment, operation.fields, nextModel.project, changes, warnings, index);
                return;
            }
            if (op === "update_resource") {
                const uid = String(operation.uid || "").trim();
                if (!uid) {
                    warnings.push({ message: `update_resource の uid がありません: operations[${index}]` });
                    return;
                }
                const resource = resourceByUid.get(uid);
                if (!resource) {
                    warnings.push({ message: `update_resource の uid が既存 resource を指していません: ${uid}`, scope: "resources", uid });
                    return;
                }
                if (!operation.fields || typeof operation.fields !== "object" || Array.isArray(operation.fields)) {
                    warnings.push({ message: `update_resource.fields がオブジェクトではありません: ${uid}`, scope: "resources", uid, label: resource.name || resource.uid });
                    return;
                }
                applyUpdateResourceOperation(resource, operation.fields, nextModel, changes, warnings, index);
                return;
            }
            if (op === "update_calendar") {
                const uid = String(operation.uid || "").trim();
                if (!uid) {
                    warnings.push({ message: `update_calendar の uid がありません: operations[${index}]` });
                    return;
                }
                const calendar = calendarByUid.get(uid);
                if (!calendar) {
                    warnings.push({ message: `update_calendar の uid が既存 calendar を指していません: ${uid}`, scope: "calendars", uid });
                    return;
                }
                if (!operation.fields || typeof operation.fields !== "object" || Array.isArray(operation.fields)) {
                    warnings.push({ message: `update_calendar.fields がオブジェクトではありません: ${uid}`, scope: "calendars", uid, label: calendar.name || calendar.uid });
                    return;
                }
                applyUpdateCalendarOperation(calendar, operation.fields, nextModel, changes, warnings, index);
                return;
            }
            if (op === "move_task") {
                applyMoveTaskOperation(operation, nextModel.tasks, changes, warnings, index);
                return;
            }
            if (op === "delete_task") {
                applyDeleteTaskOperation(operation, nextModel, changes, warnings, index);
                taskByUid.clear();
                nextModel.tasks.forEach((task) => taskByUid.set(task.uid, task));
                return;
            }
            if (op === "link_tasks") {
                applyLinkTasksOperation(operation, taskByUid, changes, warnings, index);
                return;
            }
            if (op === "unlink_tasks") {
                applyUnlinkTasksOperation(operation, taskByUid, changes, warnings, index);
                return;
            }
            warnings.push({ message: buildUnsupportedOperationWarningMessage(op, index) });
        });
        return {
            model: mikuprojectXml.normalizeProjectModel(nextModel),
            changes,
            warnings
        };
    }
    function validatePatchDocument(documentLike) {
        if (!documentLike || typeof documentLike !== "object" || Array.isArray(documentLike)) {
            throw new Error("Patch JSON がオブジェクトではありません");
        }
        const document = documentLike;
        if (!Array.isArray(document.operations)) {
            throw new Error("Patch JSON には operations 配列が必要です");
        }
        const warnings = [];
        document.operations.forEach((operation, index) => {
            if (!operation || typeof operation !== "object" || Array.isArray(operation)) {
                throw new Error(`operations[${index}] がオブジェクトではありません`);
            }
        });
        return {
            document: document,
            warnings
        };
    }
    function buildUnsupportedOperationWarningMessage(op, index) {
        const normalizedOp = op || "(empty)";
        if (normalizedOp === "link_tasks" || normalizedOp === "unlink_tasks") {
            return `未対応の op は無視します: operations[${index}].op = ${normalizedOp}。依存関係は ${normalizedOp} で扱う方針ですが、現時点では未実装です`;
        }
        if (normalizedOp === "move_task") {
            return `未対応の op は無視します: operations[${index}].op = ${normalizedOp}。親子や順序変更は move_task で扱う方針ですが、現時点では未実装です`;
        }
        return `未対応の op は無視します: operations[${index}].op = ${normalizedOp}`;
    }
    function applyLinkTasksOperation(operation, taskByUid, changes, warnings, operationIndex) {
        const relation = resolvePatchTaskRelation(operation, "link_tasks", taskByUid, warnings, operationIndex);
        if (!relation) {
            return;
        }
        const { fromTask, toTask } = relation;
        const typeCode = resolveDependencyType(operation.type, warnings, toTask, operationIndex, "link_tasks");
        const linkLag = resolveDependencyLag(operation, warnings, toTask, operationIndex, "link_tasks");
        const existing = toTask.predecessors.find((item) => item.predecessorUid === fromTask.uid);
        if (existing) {
            const beforeText = formatPredecessorList(toTask.predecessors);
            const nextType = typeCode !== null && typeCode !== void 0 ? typeCode : existing.type;
            const nextLag = linkLag !== null && linkLag !== void 0 ? linkLag : existing.linkLag;
            const relationText = formatRequestedDependencyRelation(fromTask.uid, toTask.uid, nextType, nextLag);
            if (existing.type === nextType && existing.linkLag === nextLag) {
                warnings.push({
                    message: `link_tasks の依存関係は既に存在します: ${relationText}`,
                    scope: "tasks",
                    uid: toTask.uid,
                    label: toTask.name || toTask.uid
                });
                return;
            }
            existing.type = nextType;
            existing.linkLag = nextLag;
            changes.push({
                scope: "tasks",
                uid: toTask.uid,
                label: toTask.name || toTask.uid,
                field: "Predecessors",
                before: beforeText || undefined,
                after: formatPredecessorList(toTask.predecessors)
            });
            return;
        }
        const beforeText = formatPredecessorList(toTask.predecessors);
        const relationText = formatRequestedDependencyRelation(fromTask.uid, toTask.uid, typeCode, linkLag);
        if (createsDependencyCycle(taskByUid, fromTask.uid, toTask.uid)) {
            warnings.push({
                message: `link_tasks で循環依存になるため無視します: ${relationText}`,
                scope: "tasks",
                uid: toTask.uid,
                label: toTask.name || toTask.uid
            });
            return;
        }
        toTask.predecessors.push({
            predecessorUid: fromTask.uid,
            type: typeCode,
            linkLag
        });
        changes.push({
            scope: "tasks",
            uid: toTask.uid,
            label: toTask.name || toTask.uid,
            field: "Predecessors",
            before: beforeText || undefined,
            after: formatPredecessorList(toTask.predecessors)
        });
    }
    function applyAddTaskOperation(operation, tasks, project, changes, warnings, operationIndex) {
        var _a;
        const uid = String(operation.uid || "").trim();
        if (!uid) {
            warnings.push({ message: `add_task の uid がありません: operations[${operationIndex}]` });
            return undefined;
        }
        if (uid === "0") {
            warnings.push({ message: `add_task の uid に placeholder 値 0 は使えません: ${uid}`, scope: "tasks", uid });
            return undefined;
        }
        if (tasks.some((item) => item.uid === uid)) {
            warnings.push({ message: `add_task の uid が既存 task と重複しています: ${uid}`, scope: "tasks", uid });
            return undefined;
        }
        const name = String(operation.name || "").trim();
        if (!name) {
            warnings.push({ message: `add_task.name は空でない文字列が必要です: ${uid}`, scope: "tasks", uid });
            return undefined;
        }
        warnUnsupportedPatchKeys(operation, new Set(["op", "uid", "name", "is_summary", "new_parent_uid", "new_index", "is_milestone", "planned_start", "planned_finish", "planned_duration", "planned_duration_hours"]), "add_task", uid, name, warnings);
        if (operation.is_summary !== undefined) {
            if (typeof operation.is_summary !== "boolean") {
                warnings.push({ message: `add_task.is_summary は boolean が必要です: ${uid}`, scope: "tasks", uid, label: name });
                return undefined;
            }
        }
        if (operation.is_milestone !== undefined && typeof operation.is_milestone !== "boolean") {
            warnings.push({ message: `add_task.is_milestone は boolean が必要です: ${uid}`, scope: "tasks", uid, label: name });
            return undefined;
        }
        const isSummary = Boolean(operation.is_summary);
        if (isSummary && operation.is_milestone === true) {
            warnings.push({ message: `add_task では is_summary と is_milestone を同時に true にできません: ${uid}`, scope: "tasks", uid, label: name });
            return undefined;
        }
        const nextParentUid = operation.new_parent_uid == null || operation.new_parent_uid === ""
            ? null
            : String(operation.new_parent_uid).trim();
        let nextParent;
        if (nextParentUid) {
            nextParent = tasks.find((item) => item.uid === nextParentUid);
            if (!nextParent) {
                warnings.push({ message: `add_task.new_parent_uid が既存 task を指していません: ${nextParentUid}`, scope: "tasks", uid, label: name });
                return undefined;
            }
            if (!nextParent.summary) {
                warnings.push({ message: `add_task.new_parent_uid は summary task を指す必要があります: ${nextParentUid}`, scope: "tasks", uid, label: name });
                return undefined;
            }
        }
        if (typeof operation.new_index !== "number" || !Number.isInteger(operation.new_index) || operation.new_index < 0) {
            warnings.push({ message: `add_task.new_index は 0 以上の整数が必要です: ${uid}`, scope: "tasks", uid, label: name });
            return undefined;
        }
        const nextIndex = operation.new_index;
        const parentMap = buildTaskParentMap(tasks);
        const siblingStartIndices = tasks
            .map((item, index) => ({ item, index }))
            .filter(({ item }) => (parentMap.get(item.uid) || null) === nextParentUid)
            .map(({ index }) => index);
        if (nextIndex > siblingStartIndices.length) {
            warnings.push({ message: `add_task.new_index が sibling 範囲外です: ${uid} -> ${nextIndex}`, scope: "tasks", uid, label: name });
            return undefined;
        }
        const isMilestone = Boolean(operation.is_milestone);
        const normalizedStartCandidate = normalizeOptionalPatchedTaskDate(operation.planned_start, "start", { uid, name }, project);
        if (operation.planned_start !== undefined && normalizedStartCandidate === undefined) {
            warnings.push({ message: `add_task.planned_start の日付形式が解釈できません: ${uid}`, scope: "tasks", uid, label: name });
            return undefined;
        }
        const normalizedFinishCandidate = normalizeOptionalPatchedTaskDate(operation.planned_finish, "finish", { uid, name }, project);
        if (operation.planned_finish !== undefined && normalizedFinishCandidate === undefined) {
            warnings.push({ message: `add_task.planned_finish の日付形式が解釈できません: ${uid}`, scope: "tasks", uid, label: name });
            return undefined;
        }
        const normalizedStart = normalizedStartCandidate || project.startDate;
        let normalizedFinish = normalizedFinishCandidate || normalizedStart;
        if (normalizedStart > normalizedFinish) {
            warnings.push({ message: `add_task.planned_start が planned_finish より後です: ${uid}`, scope: "tasks", uid, label: name });
            return undefined;
        }
        let duration = resolveAddedTaskDuration(operation, uid, name, warnings) || "PT0H0M0S";
        if (isMilestone) {
            if (normalizedFinish !== normalizedStart) {
                warnings.push({ message: `add_task.is_milestone=true のため planned_finish は planned_start に揃えます: ${uid}`, scope: "tasks", uid, label: name });
                normalizedFinish = normalizedStart;
            }
            if (!isZeroDuration(duration)) {
                warnings.push({ message: `add_task.is_milestone=true のため planned_duration は 0 に揃えます: ${uid}`, scope: "tasks", uid, label: name });
                duration = "PT0H0M0S";
            }
        }
        const insertionIndex = resolveAddInsertionIndex(tasks, siblingStartIndices, nextIndex, nextParentUid);
        const addedTask = {
            uid,
            id: "",
            name,
            outlineLevel: nextParent ? nextParent.outlineLevel + 1 : 1,
            outlineNumber: "",
            start: normalizedStart,
            finish: normalizedFinish,
            duration,
            milestone: isMilestone,
            summary: isSummary,
            percentComplete: 0,
            extendedAttributes: [],
            baselines: [],
            timephasedData: [],
            predecessors: []
        };
        tasks.splice(insertionIndex, 0, addedTask);
        rebuildTaskHierarchyMetadata(tasks);
        const nextPositionMap = buildTaskPositionMap(tasks, buildTaskParentMap(tasks));
        changes.push({
            scope: "tasks",
            uid,
            label: name,
            field: "name",
            before: undefined,
            after: name
        });
        changes.push({
            scope: "tasks",
            uid,
            label: name,
            field: "parent_uid",
            before: undefined,
            after: nextParentUid || "(root)"
        });
        changes.push({
            scope: "tasks",
            uid,
            label: name,
            field: "position",
            before: undefined,
            after: (_a = nextPositionMap.get(uid)) !== null && _a !== void 0 ? _a : nextIndex
        });
        if (operation.planned_start !== undefined) {
            changes.push({
                scope: "tasks",
                uid,
                label: name,
                field: "planned_start",
                before: undefined,
                after: normalizedStart
            });
        }
        if (operation.planned_finish !== undefined) {
            changes.push({
                scope: "tasks",
                uid,
                label: name,
                field: "planned_finish",
                before: undefined,
                after: normalizedFinish
            });
        }
        return addedTask;
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
            : normalizePatchedPlainDateTime(operation.start, "start", model.project);
        if (operation.start !== undefined && normalizedStart === undefined) {
            warnings.push({ message: `add_assignment.start の日付形式が解釈できません: ${uid}`, scope: "assignments", uid, label: uid });
            return undefined;
        }
        const normalizedFinish = operation.finish === undefined
            ? undefined
            : normalizePatchedPlainDateTime(operation.finish, "finish", model.project);
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
        const percentWorkComplete = operation.percent_work_complete;
        if (percentWorkComplete !== undefined) {
            changes.push({ scope: "assignments", uid, label: uid, field: "percentWorkComplete", before: undefined, after: percentWorkComplete });
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
        if (operation.cost_per_use !== undefined) {
            const costPerUse = operation.cost_per_use;
            if (typeof costPerUse !== "number" || !Number.isFinite(costPerUse) || costPerUse < 0) {
                warnings.push({ message: `add_resource.cost_per_use は 0 以上の数値が必要です: ${uid}`, scope: "resources", uid, label: name });
                return undefined;
            }
        }
        if (operation.percent_work_complete !== undefined) {
            const percentWorkComplete = operation.percent_work_complete;
            if (typeof percentWorkComplete !== "number" || !Number.isFinite(percentWorkComplete) || percentWorkComplete < 0 || percentWorkComplete > 100) {
                warnings.push({ message: `add_resource.percent_work_complete は 0 以上 100 以下の数値が必要です: ${uid}`, scope: "resources", uid, label: name });
                return undefined;
            }
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
            const rawBaseCalendarUid = operation.base_calendar_uid;
            if (typeof rawBaseCalendarUid !== "string") {
                warnings.push({ message: `add_calendar.base_calendar_uid は文字列が必要です: ${uid}`, scope: "calendars", uid, label: name });
                return undefined;
            }
            nextBaseCalendarUid = rawBaseCalendarUid.trim() || undefined;
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
            warnings.push({
                message: `delete_resource first cut では assignment がある resource は削除できません: ${uid} (assignments=${assignmentText})`,
                scope: "resources",
                uid,
                label: resource.name || resource.uid
            });
            return;
        }
        model.resources.splice(resourceIndex, 1);
        changes.push({
            scope: "resources",
            uid,
            label: resource.name || resource.uid,
            field: "name",
            before: resource.name || resource.uid,
            after: "(deleted)"
        });
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
            if (projectRef) {
                blockers.push("project=1");
            }
            if (taskRefs.length > 0) {
                blockers.push(`tasks=${taskRefs.join(", ")}`);
            }
            if (resourceRefs.length > 0) {
                blockers.push(`resources=${resourceRefs.join(", ")}`);
            }
            if (baseCalendarRefs.length > 0) {
                blockers.push(`baseRefs=${baseCalendarRefs.join(", ")}`);
            }
            warnings.push({
                message: `delete_calendar first cut では参照が残っている calendar は削除できません: ${uid} (${blockers.join("; ")})`,
                scope: "calendars",
                uid,
                label: calendar.name || calendar.uid
            });
            return;
        }
        model.calendars.splice(calendarIndex, 1);
        changes.push({
            scope: "calendars",
            uid,
            label: calendar.name || calendar.uid,
            field: "name",
            before: calendar.name || calendar.uid,
            after: "(deleted)"
        });
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
        changes.push({
            scope: "assignments",
            uid,
            label: uid,
            field: "taskUid",
            before: assignment.taskUid,
            after: "(deleted)"
        });
        changes.push({
            scope: "assignments",
            uid,
            label: uid,
            field: "resourceUid",
            before: assignment.resourceUid,
            after: "(deleted)"
        });
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
    function applyMoveTaskOperation(operation, tasks, changes, warnings, operationIndex) {
        var _a, _b;
        const uid = String(operation.uid || "").trim();
        if (!uid) {
            warnings.push({ message: `move_task の uid がありません: operations[${operationIndex}]` });
            return;
        }
        const taskIndex = tasks.findIndex((item) => item.uid === uid);
        if (taskIndex < 0) {
            warnings.push({ message: `move_task の uid が既存 task を指していません: ${uid}`, scope: "tasks", uid });
            return;
        }
        const task = tasks[taskIndex];
        const currentParentMap = buildTaskParentMap(tasks);
        const currentPositionMap = buildTaskPositionMap(tasks, currentParentMap);
        const currentParentUid = currentParentMap.get(uid) || null;
        const nextParentUid = operation.new_parent_uid == null || operation.new_parent_uid === ""
            ? null
            : String(operation.new_parent_uid).trim();
        if (typeof operation.new_index !== "number" || !Number.isInteger(operation.new_index) || operation.new_index < 0) {
            warnings.push({
                message: `move_task.new_index は 0 以上の整数が必要です: ${uid}`,
                scope: "tasks",
                uid,
                label: task.name || task.uid
            });
            return;
        }
        const nextIndex = operation.new_index;
        const currentRange = collectTaskSubtreeRange(tasks, taskIndex);
        const subtree = tasks.slice(currentRange.start, currentRange.end);
        const subtreeUidSet = new Set(subtree.map((item) => item.uid));
        if (nextParentUid && subtreeUidSet.has(nextParentUid)) {
            warnings.push({
                message: `move_task では task を自身または配下へ移動できません: ${uid} -> ${nextParentUid}`,
                scope: "tasks",
                uid,
                label: task.name || task.uid
            });
            return;
        }
        let nextParent;
        if (nextParentUid) {
            nextParent = tasks.find((item) => item.uid === nextParentUid);
            if (!nextParent) {
                warnings.push({
                    message: `move_task.new_parent_uid が既存 task を指していません: ${nextParentUid}`,
                    scope: "tasks",
                    uid,
                    label: task.name || task.uid
                });
                return;
            }
            if (!nextParent.summary) {
                warnings.push({
                    message: `move_task.new_parent_uid は summary task を指す必要があります: ${nextParentUid}`,
                    scope: "tasks",
                    uid,
                    label: task.name || task.uid
                });
                return;
            }
        }
        const remainingTasks = tasks.slice(0, currentRange.start).concat(tasks.slice(currentRange.end));
        const remainingParentMap = buildTaskParentMap(remainingTasks);
        const siblingStartIndices = remainingTasks
            .map((item, index) => ({ item, index }))
            .filter(({ item }) => (remainingParentMap.get(item.uid) || null) === nextParentUid)
            .map(({ index }) => index);
        if (nextIndex > siblingStartIndices.length) {
            warnings.push({
                message: `move_task.new_index が sibling 範囲外です: ${uid} -> ${nextIndex}`,
                scope: "tasks",
                uid,
                label: task.name || task.uid
            });
            return;
        }
        const insertionIndex = resolveMoveInsertionIndex(remainingTasks, siblingStartIndices, nextIndex, nextParentUid);
        const nextOutlineLevel = nextParent ? nextParent.outlineLevel + 1 : 1;
        const levelDelta = nextOutlineLevel - task.outlineLevel;
        if (isNoOpMove(tasks, remainingTasks, subtree, insertionIndex, levelDelta)) {
            warnings.push({
                message: `move_task は結果が変わらないため無視します: ${uid} -> parent=${nextParentUid || "(root)"} index=${nextIndex}`,
                scope: "tasks",
                uid,
                label: task.name || task.uid
            });
            return;
        }
        subtree.forEach((item) => {
            item.outlineLevel += levelDelta;
        });
        remainingTasks.splice(insertionIndex, 0, ...subtree);
        rebuildTaskHierarchyMetadata(remainingTasks);
        tasks.splice(0, tasks.length, ...remainingTasks);
        const nextParentMap = buildTaskParentMap(tasks);
        const nextPositionMap = buildTaskPositionMap(tasks, nextParentMap);
        const nextParentText = nextParentUid || "(root)";
        const currentParentText = currentParentUid || "(root)";
        const nextPosition = (_a = nextPositionMap.get(uid)) !== null && _a !== void 0 ? _a : nextIndex;
        const currentPosition = (_b = currentPositionMap.get(uid)) !== null && _b !== void 0 ? _b : 0;
        if (currentParentText !== nextParentText) {
            changes.push({
                scope: "tasks",
                uid,
                label: task.name || task.uid,
                field: "parent_uid",
                before: currentParentText,
                after: nextParentText
            });
        }
        if (currentPosition !== nextPosition) {
            changes.push({
                scope: "tasks",
                uid,
                label: task.name || task.uid,
                field: "position",
                before: currentPosition,
                after: nextPosition
            });
        }
    }
    function applyDeleteTaskOperation(operation, model, changes, warnings, operationIndex) {
        var _a;
        const uid = String(operation.uid || "").trim();
        if (!uid) {
            warnings.push({ message: `delete_task の uid がありません: operations[${operationIndex}]` });
            return;
        }
        const taskIndex = model.tasks.findIndex((item) => item.uid === uid);
        if (taskIndex < 0) {
            warnings.push({ message: `delete_task の uid が既存 task を指していません: ${uid}`, scope: "tasks", uid });
            return;
        }
        const task = model.tasks[taskIndex];
        const taskRange = collectTaskSubtreeRange(model.tasks, taskIndex);
        if (taskRange.end - taskRange.start > 1) {
            const childCount = taskRange.end - taskRange.start - 1;
            warnings.push({
                message: `delete_task first cut では summary task や子を持つ task は削除できません: ${uid} (children=${childCount})`,
                scope: "tasks",
                uid,
                label: task.name || task.uid
            });
            return;
        }
        const referencingAssignments = model.assignments.filter((assignment) => assignment.taskUid === uid);
        if (referencingAssignments.length > 0) {
            const assignmentText = referencingAssignments.map((assignment) => assignment.uid).join(", ");
            warnings.push({
                message: `delete_task first cut では assignment がある task は削除できません: ${uid} (assignments=${assignmentText})`,
                scope: "tasks",
                uid,
                label: task.name || task.uid
            });
            return;
        }
        const successorTasks = model.tasks.filter((item) => item.predecessors.some((predecessor) => predecessor.predecessorUid === uid));
        if (successorTasks.length > 0) {
            const successorText = successorTasks.map((item) => item.uid).join(", ");
            warnings.push({
                message: `delete_task first cut では後続依存がある task は削除できません: ${uid} (successors=${successorText})`,
                scope: "tasks",
                uid,
                label: task.name || task.uid
            });
            return;
        }
        const parentMap = buildTaskParentMap(model.tasks);
        const positionMap = buildTaskPositionMap(model.tasks, parentMap);
        const parentUid = parentMap.get(uid) || "(root)";
        const position = (_a = positionMap.get(uid)) !== null && _a !== void 0 ? _a : 0;
        model.tasks.splice(taskIndex, 1);
        rebuildTaskHierarchyMetadata(model.tasks);
        changes.push({
            scope: "tasks",
            uid,
            label: task.name || task.uid,
            field: "name",
            before: task.name || task.uid,
            after: "(deleted)"
        });
        changes.push({
            scope: "tasks",
            uid,
            label: task.name || task.uid,
            field: "parent_uid",
            before: parentUid,
            after: "(deleted)"
        });
        changes.push({
            scope: "tasks",
            uid,
            label: task.name || task.uid,
            field: "position",
            before: position,
            after: "(deleted)"
        });
    }
    function normalizeOptionalPatchedTaskDate(rawValue, kind, task, project) {
        if (rawValue === undefined) {
            return undefined;
        }
        return normalizePatchedTaskDate(rawValue, kind, task, project);
    }
    function resolveAddedTaskDuration(operation, uid, name, warnings) {
        if (operation.planned_duration !== undefined && operation.planned_duration_hours !== undefined) {
            warnings.push({ message: `add_task.planned_duration と add_task.planned_duration_hours が同時指定されたため、planned_duration_hours は無視します: ${uid}`, scope: "tasks", uid, label: name });
        }
        if (operation.planned_duration !== undefined) {
            if (typeof operation.planned_duration !== "string" || operation.planned_duration.trim() === "") {
                warnings.push({ message: `add_task.planned_duration は空でない文字列が必要です: ${uid}`, scope: "tasks", uid, label: name });
                return undefined;
            }
            return operation.planned_duration.trim();
        }
        if (operation.planned_duration_hours !== undefined) {
            if (typeof operation.planned_duration_hours !== "number" || !Number.isFinite(operation.planned_duration_hours) || operation.planned_duration_hours < 0) {
                warnings.push({ message: `add_task.planned_duration_hours は 0 以上の数値が必要です: ${uid}`, scope: "tasks", uid, label: name });
                return undefined;
            }
            return formatDurationHours(operation.planned_duration_hours);
        }
        return undefined;
    }
    function applyUnlinkTasksOperation(operation, taskByUid, changes, warnings, operationIndex) {
        const relation = resolvePatchTaskRelation(operation, "unlink_tasks", taskByUid, warnings, operationIndex);
        if (!relation) {
            return;
        }
        const { fromTask, toTask } = relation;
        const requestedType = resolveDependencyType(operation.type, warnings, toTask, operationIndex, "unlink_tasks", { allowEmpty: true });
        const requestedLag = resolveDependencyLag(operation, warnings, toTask, operationIndex, "unlink_tasks", { allowEmpty: true });
        const beforeText = formatPredecessorList(toTask.predecessors);
        const matched = toTask.predecessors.filter((item) => {
            if (item.predecessorUid !== fromTask.uid) {
                return false;
            }
            if (requestedType !== undefined && item.type !== requestedType) {
                return false;
            }
            if (requestedLag !== undefined && normalizeDurationText(item.linkLag) !== normalizeDurationText(requestedLag)) {
                return false;
            }
            return true;
        });
        const filtered = toTask.predecessors.filter((item) => {
            if (item.predecessorUid !== fromTask.uid) {
                return true;
            }
            if (requestedType !== undefined && item.type !== requestedType) {
                return true;
            }
            if (requestedLag !== undefined && normalizeDurationText(item.linkLag) !== normalizeDurationText(requestedLag)) {
                return true;
            }
            return false;
        });
        if (filtered.length === toTask.predecessors.length) {
            warnings.push({
                message: `unlink_tasks の対象依存関係が見つかりません: ${formatRequestedDependencyRelation(fromTask.uid, toTask.uid, requestedType, requestedLag)}`,
                scope: "tasks",
                uid: toTask.uid,
                label: toTask.name || toTask.uid
            });
            return;
        }
        toTask.predecessors = filtered;
        if (matched.length > 1) {
            warnings.push({
                message: `unlink_tasks は一致した依存関係 ${matched.length} 件をすべて解除しました: ${formatRequestedDependencyRelation(fromTask.uid, toTask.uid, requestedType, requestedLag)}`,
                scope: "tasks",
                uid: toTask.uid,
                label: toTask.name || toTask.uid
            });
        }
        changes.push({
            scope: "tasks",
            uid: toTask.uid,
            label: toTask.name || toTask.uid,
            field: "Predecessors",
            before: beforeText || undefined,
            after: formatPredecessorList(toTask.predecessors)
        });
    }
    function resolvePatchTaskRelation(operation, opName, taskByUid, warnings, operationIndex) {
        const fromUid = String(operation.from_uid || "").trim();
        const toUid = String(operation.to_uid || "").trim();
        if (!fromUid || !toUid) {
            warnings.push({ message: `${opName} には from_uid / to_uid が必要です: operations[${operationIndex}]` });
            return undefined;
        }
        const fromTask = taskByUid.get(fromUid);
        if (!fromTask) {
            warnings.push({ message: `${opName}.from_uid が既存 task を指していません: ${fromUid}`, scope: "tasks", uid: fromUid });
            return undefined;
        }
        const toTask = taskByUid.get(toUid);
        if (!toTask) {
            warnings.push({ message: `${opName}.to_uid が既存 task を指していません: ${toUid}`, scope: "tasks", uid: toUid });
            return undefined;
        }
        if (fromUid === toUid) {
            warnings.push({
                message: `${opName} では同一 task を依存元 / 依存先にできません: ${fromUid}`,
                scope: "tasks",
                uid: toTask.uid,
                label: toTask.name || toTask.uid
            });
            return undefined;
        }
        return { fromTask, toTask };
    }
    function collectTaskSubtreeRange(tasks, startIndex) {
        const rootTask = tasks[startIndex];
        let end = startIndex + 1;
        while (end < tasks.length && tasks[end].outlineLevel > rootTask.outlineLevel) {
            end += 1;
        }
        return { start: startIndex, end };
    }
    function buildTaskParentMap(tasks) {
        const parentMap = new Map();
        const stack = [];
        for (const task of tasks) {
            while (stack.length > 0 && task.outlineLevel <= stack[stack.length - 1].outlineLevel) {
                stack.pop();
            }
            parentMap.set(task.uid, stack.length > 0 ? stack[stack.length - 1].uid : null);
            if (task.summary) {
                stack.push(task);
            }
        }
        return parentMap;
    }
    function buildTaskPositionMap(tasks, parentMap) {
        const counters = new Map();
        const positionMap = new Map();
        for (const task of tasks) {
            const parentUid = parentMap.get(task.uid) || "__root__";
            const position = counters.get(parentUid) || 0;
            positionMap.set(task.uid, position);
            counters.set(parentUid, position + 1);
        }
        return positionMap;
    }
    function resolveMoveInsertionIndex(remainingTasks, siblingStartIndices, nextIndex, nextParentUid) {
        if (siblingStartIndices.length === 0) {
            if (!nextParentUid) {
                return 0;
            }
            const parentIndex = remainingTasks.findIndex((item) => item.uid === nextParentUid);
            return parentIndex >= 0 ? parentIndex + 1 : remainingTasks.length;
        }
        if (nextIndex < siblingStartIndices.length) {
            return siblingStartIndices[nextIndex];
        }
        if (!nextParentUid) {
            return remainingTasks.length;
        }
        const parentIndex = remainingTasks.findIndex((item) => item.uid === nextParentUid);
        if (parentIndex < 0) {
            return remainingTasks.length;
        }
        return collectTaskSubtreeRange(remainingTasks, parentIndex).end;
    }
    function resolveAddInsertionIndex(tasks, siblingStartIndices, nextIndex, nextParentUid) {
        if (siblingStartIndices.length === 0) {
            if (!nextParentUid) {
                return 0;
            }
            const parentIndex = tasks.findIndex((item) => item.uid === nextParentUid);
            return parentIndex >= 0 ? parentIndex + 1 : tasks.length;
        }
        if (nextIndex < siblingStartIndices.length) {
            return siblingStartIndices[nextIndex];
        }
        if (!nextParentUid) {
            return tasks.length;
        }
        const parentIndex = tasks.findIndex((item) => item.uid === nextParentUid);
        if (parentIndex < 0) {
            return tasks.length;
        }
        return collectTaskSubtreeRange(tasks, parentIndex).end;
    }
    function isNoOpMove(originalTasks, remainingTasks, subtree, insertionIndex, levelDelta) {
        const currentSignature = originalTasks.map((task) => `${task.uid}@${task.outlineLevel}`).join("|");
        const candidateTasks = remainingTasks
            .map((task) => ({ uid: task.uid, outlineLevel: task.outlineLevel }))
            .slice();
        const subtreeSignature = subtree.map((task) => ({
            uid: task.uid,
            outlineLevel: task.outlineLevel + levelDelta
        }));
        candidateTasks.splice(insertionIndex, 0, ...subtreeSignature);
        const nextSignature = candidateTasks.map((task) => `${task.uid}@${task.outlineLevel}`).join("|");
        return currentSignature === nextSignature;
    }
    function rebuildTaskHierarchyMetadata(tasks) {
        const counters = [];
        tasks.forEach((task, index) => {
            counters.length = Math.max(1, task.outlineLevel);
            counters[task.outlineLevel - 1] = (counters[task.outlineLevel - 1] || 0) + 1;
            task.outlineNumber = counters.slice(0, task.outlineLevel).join(".");
            task.id = String(index + 1);
        });
    }
    function createsDependencyCycle(taskByUid, fromUid, toUid) {
        const queue = [toUid];
        const visited = new Set();
        while (queue.length > 0) {
            const currentUid = queue.shift();
            if (currentUid === fromUid) {
                return true;
            }
            if (visited.has(currentUid)) {
                continue;
            }
            visited.add(currentUid);
            taskByUid.forEach((task) => {
                if (task.predecessors.some((item) => item.predecessorUid === currentUid) && !visited.has(task.uid)) {
                    queue.push(task.uid);
                }
            });
        }
        return false;
    }
    function resolveDependencyType(rawType, warnings, task, operationIndex, opName, options = {}) {
        const trimmed = String(rawType || "").trim();
        if (!trimmed) {
            return options.allowEmpty ? undefined : 1;
        }
        const typeMap = {
            FS: 1,
            SS: 2,
            FF: 3,
            SF: 4
        };
        const mapped = typeMap[trimmed.toUpperCase()];
        if (mapped !== undefined) {
            return mapped;
        }
        warnings.push({
            message: `${opName}.type は FS / SS / FF / SF のいずれかが必要です: operations[${operationIndex}].type = ${trimmed}`,
            scope: "tasks",
            uid: task.uid,
            label: task.name || task.uid
        });
        return undefined;
    }
    function resolveDependencyLag(operation, warnings, task, operationIndex, opName, options = {}) {
        if (operation.lag !== undefined && operation.lag_hours !== undefined) {
            warnings.push({
                message: `${opName}.lag と ${opName}.lag_hours が同時指定されたため、lag_hours は無視します: ${task.uid}`,
                scope: "tasks",
                uid: task.uid,
                label: task.name || task.uid
            });
        }
        if (operation.lag !== undefined) {
            if (typeof operation.lag !== "string" || operation.lag.trim() === "") {
                warnings.push({
                    message: `${opName}.lag は空でない文字列が必要です: ${task.uid}`,
                    scope: "tasks",
                    uid: task.uid,
                    label: task.name || task.uid
                });
                return undefined;
            }
            const normalizedLag = operation.lag.trim();
            if (!isValidDurationText(normalizedLag)) {
                warnings.push({
                    message: `${opName}.lag は ISO 8601 duration 形式が必要です: ${task.uid}`,
                    scope: "tasks",
                    uid: task.uid,
                    label: task.name || task.uid
                });
                return undefined;
            }
            return normalizedLag;
        }
        if (operation.lag_hours !== undefined) {
            if (typeof operation.lag_hours !== "number" || !Number.isFinite(operation.lag_hours)) {
                warnings.push({
                    message: `${opName}.lag_hours は数値が必要です: ${task.uid}`,
                    scope: "tasks",
                    uid: task.uid,
                    label: task.name || task.uid
                });
                return undefined;
            }
            return formatDurationHours(operation.lag_hours);
        }
        return options.allowEmpty ? undefined : undefined;
    }
    function formatDependencyType(type) {
        const typeMap = {
            1: "FS",
            2: "SS",
            3: "FF",
            4: "SF"
        };
        return typeMap[type !== null && type !== void 0 ? type : 1] || `type=${type}`;
    }
    function isZeroDuration(duration) {
        const text = String(duration || "").trim();
        return text === "" || text === "PT0H0M0S" || text === "PT0M0S" || text === "PT0S";
    }
    function normalizeDurationText(duration) {
        const text = String(duration || "").trim();
        return text || undefined;
    }
    function isValidDurationText(duration) {
        const text = String(duration || "").trim();
        if (!text) {
            return false;
        }
        const match = /^-?P(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)$/.exec(text);
        if (!match) {
            return false;
        }
        return Boolean(match[1] || match[2] || match[3]);
    }
    function formatPredecessor(predecessor) {
        const extras = [formatDependencyType(predecessor.type)];
        if (!isZeroDuration(predecessor.linkLag)) {
            extras.push(`lag=${String(predecessor.linkLag || "").trim()}`);
        }
        return `${predecessor.predecessorUid}(${extras.join(", ")})`;
    }
    function formatRequestedDependencyRelation(fromUid, toUid, type, linkLag) {
        const extras = [formatDependencyType(type)];
        if (!isZeroDuration(linkLag)) {
            extras.push(`lag=${String(linkLag || "").trim()}`);
        }
        return `${fromUid} -> ${toUid} (${extras.join(", ")})`;
    }
    function formatPredecessorList(predecessors) {
        return predecessors.map((item) => formatPredecessor(item)).join(", ");
    }
    function applyUpdateTaskOperation(task, rawFields, model, changes, warnings, operationIndex) {
        const project = model.project;
        const fields = { ...rawFields };
        if (Object.keys(fields).length === 0) {
            warnings.push({ message: `update_task に fields がありません: ${task.uid}`, scope: "tasks", uid: task.uid, label: task.name || task.uid });
            return;
        }
        if (fields.planned_duration !== undefined && fields.planned_duration_hours !== undefined) {
            warnings.push({ message: `planned_duration と planned_duration_hours が同時指定されたため、planned_duration_hours は無視します: ${task.uid}`, scope: "tasks", uid: task.uid, label: task.name || task.uid });
            delete fields.planned_duration_hours;
        }
        const handledFields = new Set();
        let nextMilestone = task.milestone;
        if (fields.is_milestone !== undefined) {
            handledFields.add("is_milestone");
            if (typeof fields.is_milestone !== "boolean") {
                warnings.push({ message: `update_task.is_milestone は boolean が必要です: ${task.uid}`, scope: "tasks", uid: task.uid, label: task.name || task.uid });
            }
            else if (task.summary && fields.is_milestone) {
                warnings.push({ message: `update_task では summary task を milestone にできません: ${task.uid}`, scope: "tasks", uid: task.uid, label: task.name || task.uid });
            }
            else if (task.milestone !== fields.is_milestone) {
                changes.push({
                    scope: "tasks",
                    uid: task.uid,
                    label: task.name || task.uid,
                    field: "is_milestone",
                    before: task.milestone,
                    after: fields.is_milestone
                });
                task.milestone = fields.is_milestone;
                nextMilestone = fields.is_milestone;
            }
        }
        if (fields.name !== undefined) {
            handledFields.add("name");
            if (typeof fields.name !== "string" || fields.name.trim() === "") {
                warnings.push({ message: `update_task.name は空でない文字列が必要です: ${task.uid}`, scope: "tasks", uid: task.uid, label: task.name || task.uid });
            }
            else if (task.name !== fields.name.trim()) {
                changes.push({
                    scope: "tasks",
                    uid: task.uid,
                    label: task.name || task.uid,
                    field: "name",
                    before: task.name,
                    after: fields.name.trim()
                });
                task.name = fields.name.trim();
            }
        }
        if (fields.notes !== undefined) {
            handledFields.add("notes");
            if (typeof fields.notes !== "string") {
                warnings.push({ message: `update_task.notes は文字列が必要です: ${task.uid}`, scope: "tasks", uid: task.uid, label: task.name || task.uid });
            }
            else if ((task.notes || "") !== fields.notes) {
                changes.push({
                    scope: "tasks",
                    uid: task.uid,
                    label: task.name || task.uid,
                    field: "notes",
                    before: task.notes,
                    after: fields.notes
                });
                task.notes = fields.notes || undefined;
            }
        }
        if (fields.calendar_uid !== undefined) {
            handledFields.add("calendar_uid");
            if (typeof fields.calendar_uid !== "string") {
                warnings.push({ message: `update_task.calendar_uid は文字列が必要です: ${task.uid}`, scope: "tasks", uid: task.uid, label: task.name || task.uid });
            }
            else {
                const nextCalendarUid = fields.calendar_uid.trim() || undefined;
                if (nextCalendarUid && !model.calendars.some((calendar) => calendar.uid === nextCalendarUid)) {
                    warnings.push({ message: `update_task.calendar_uid が既存 calendar を指していません: ${task.uid} -> ${nextCalendarUid}`, scope: "tasks", uid: task.uid, label: task.name || task.uid });
                }
                else if (task.calendarUID !== nextCalendarUid) {
                    changes.push({
                        scope: "tasks",
                        uid: task.uid,
                        label: task.name || task.uid,
                        field: "calendarUID",
                        before: task.calendarUID,
                        after: nextCalendarUid || "(cleared)"
                    });
                    task.calendarUID = nextCalendarUid;
                }
            }
        }
        if (fields.percent_complete !== undefined) {
            handledFields.add("percent_complete");
            if (typeof fields.percent_complete !== "number" || !Number.isFinite(fields.percent_complete) || fields.percent_complete < 0 || fields.percent_complete > 100) {
                warnings.push({ message: `update_task.percent_complete は 0 以上 100 以下の数値が必要です: ${task.uid}`, scope: "tasks", uid: task.uid, label: task.name || task.uid });
            }
            else if (task.percentComplete !== fields.percent_complete) {
                changes.push({
                    scope: "tasks",
                    uid: task.uid,
                    label: task.name || task.uid,
                    field: "percentComplete",
                    before: task.percentComplete,
                    after: fields.percent_complete
                });
                task.percentComplete = fields.percent_complete;
            }
        }
        if (fields.percent_work_complete !== undefined) {
            handledFields.add("percent_work_complete");
            if (typeof fields.percent_work_complete !== "number" || !Number.isFinite(fields.percent_work_complete) || fields.percent_work_complete < 0 || fields.percent_work_complete > 100) {
                warnings.push({ message: `update_task.percent_work_complete は 0 以上 100 以下の数値が必要です: ${task.uid}`, scope: "tasks", uid: task.uid, label: task.name || task.uid });
            }
            else if (task.percentWorkComplete !== fields.percent_work_complete) {
                changes.push({
                    scope: "tasks",
                    uid: task.uid,
                    label: task.name || task.uid,
                    field: "percentWorkComplete",
                    before: task.percentWorkComplete,
                    after: fields.percent_work_complete
                });
                task.percentWorkComplete = fields.percent_work_complete;
            }
        }
        if (fields.critical !== undefined) {
            handledFields.add("critical");
            if (typeof fields.critical !== "boolean") {
                warnings.push({ message: `update_task.critical は boolean が必要です: ${task.uid}`, scope: "tasks", uid: task.uid, label: task.name || task.uid });
            }
            else if (task.critical !== fields.critical) {
                changes.push({
                    scope: "tasks",
                    uid: task.uid,
                    label: task.name || task.uid,
                    field: "critical",
                    before: task.critical,
                    after: fields.critical
                });
                task.critical = fields.critical;
            }
        }
        if (fields.planned_start !== undefined) {
            handledFields.add("planned_start");
            const normalizedStart = normalizePatchedTaskDate(fields.planned_start, "start", { ...task, milestone: nextMilestone }, project);
            if (normalizedStart === undefined) {
                warnings.push({ message: `update_task.planned_start の日付形式が解釈できません: ${task.uid}`, scope: "tasks", uid: task.uid, label: task.name || task.uid });
            }
            else if (task.start !== normalizedStart) {
                changes.push({
                    scope: "tasks",
                    uid: task.uid,
                    label: task.name || task.uid,
                    field: "planned_start",
                    before: task.start,
                    after: normalizedStart
                });
                task.start = normalizedStart;
            }
        }
        if (fields.planned_finish !== undefined) {
            handledFields.add("planned_finish");
            const normalizedFinish = normalizePatchedTaskDate(fields.planned_finish, "finish", { ...task, milestone: nextMilestone }, project);
            if (normalizedFinish === undefined) {
                warnings.push({ message: `update_task.planned_finish の日付形式が解釈できません: ${task.uid}`, scope: "tasks", uid: task.uid, label: task.name || task.uid });
            }
            else if (task.finish !== normalizedFinish) {
                changes.push({
                    scope: "tasks",
                    uid: task.uid,
                    label: task.name || task.uid,
                    field: "planned_finish",
                    before: task.finish,
                    after: normalizedFinish
                });
                task.finish = normalizedFinish;
            }
        }
        if (fields.planned_duration !== undefined) {
            handledFields.add("planned_duration");
            if (typeof fields.planned_duration !== "string" || fields.planned_duration.trim() === "") {
                warnings.push({ message: `update_task.planned_duration は空でない文字列が必要です: ${task.uid}`, scope: "tasks", uid: task.uid, label: task.name || task.uid });
            }
            else if (task.duration !== fields.planned_duration.trim()) {
                changes.push({
                    scope: "tasks",
                    uid: task.uid,
                    label: task.name || task.uid,
                    field: "planned_duration",
                    before: task.duration,
                    after: fields.planned_duration.trim()
                });
                task.duration = fields.planned_duration.trim();
            }
        }
        if (fields.planned_duration_hours !== undefined) {
            handledFields.add("planned_duration_hours");
            if (typeof fields.planned_duration_hours !== "number" || !Number.isFinite(fields.planned_duration_hours) || fields.planned_duration_hours < 0) {
                warnings.push({ message: `update_task.planned_duration_hours は 0 以上の数値が必要です: ${task.uid}`, scope: "tasks", uid: task.uid, label: task.name || task.uid });
            }
            else {
                const beforeHours = parseDurationHours(task.duration);
                const nextDuration = formatDurationHours(fields.planned_duration_hours);
                if (task.duration !== nextDuration) {
                    changes.push({
                        scope: "tasks",
                        uid: task.uid,
                        label: task.name || task.uid,
                        field: "planned_duration_hours",
                        before: beforeHours,
                        after: fields.planned_duration_hours
                    });
                    task.duration = nextDuration;
                }
            }
        }
        if (nextMilestone) {
            if (task.finish !== task.start) {
                warnings.push({ message: `update_task.is_milestone=true のため planned_finish は planned_start に揃えます: ${task.uid}`, scope: "tasks", uid: task.uid, label: task.name || task.uid });
                if (task.finish !== task.start) {
                    changes.push({
                        scope: "tasks",
                        uid: task.uid,
                        label: task.name || task.uid,
                        field: "planned_finish",
                        before: task.finish,
                        after: task.start
                    });
                    task.finish = task.start;
                }
            }
            if (!isZeroDuration(task.duration)) {
                warnings.push({ message: `update_task.is_milestone=true のため planned_duration は 0 に揃えます: ${task.uid}`, scope: "tasks", uid: task.uid, label: task.name || task.uid });
                changes.push({
                    scope: "tasks",
                    uid: task.uid,
                    label: task.name || task.uid,
                    field: "planned_duration",
                    before: task.duration,
                    after: "PT0H0M0S"
                });
                task.duration = "PT0H0M0S";
            }
        }
        Object.keys(fields).forEach((fieldName) => {
            if (handledFields.has(fieldName)) {
                return;
            }
            warnings.push({ message: `未対応の field は無視します: operations[${operationIndex}].fields.${fieldName}`, scope: "tasks", uid: task.uid, label: task.name || task.uid });
        });
    }
    function applyUpdateProjectOperation(project, rawFields, model, changes, warnings, operationIndex) {
        const fields = { ...rawFields };
        const projectUid = "project";
        const projectLabel = project.name || "project";
        if (Object.keys(fields).length === 0) {
            warnings.push({ message: "update_project に fields がありません", scope: "project", uid: projectUid, label: projectLabel });
            return;
        }
        const handledFields = new Set();
        if (fields.name !== undefined) {
            handledFields.add("name");
            if (typeof fields.name !== "string" || fields.name.trim() === "") {
                warnings.push({ message: "update_project.name は空でない文字列が必要です", scope: "project", uid: projectUid, label: projectLabel });
            }
            else if (project.name !== fields.name.trim()) {
                changes.push({ scope: "project", uid: projectUid, label: projectLabel, field: "name", before: project.name, after: fields.name.trim() });
                project.name = fields.name.trim();
            }
        }
        if (fields.title !== undefined) {
            handledFields.add("title");
            if (typeof fields.title !== "string") {
                warnings.push({ message: "update_project.title は文字列が必要です", scope: "project", uid: projectUid, label: projectLabel });
            }
            else {
                const nextTitle = fields.title.trim() || undefined;
                if (project.title !== nextTitle) {
                    changes.push({ scope: "project", uid: projectUid, label: projectLabel, field: "title", before: project.title, after: nextTitle || "(cleared)" });
                    project.title = nextTitle;
                }
            }
        }
        if (fields.author !== undefined) {
            handledFields.add("author");
            if (typeof fields.author !== "string") {
                warnings.push({ message: "update_project.author は文字列が必要です", scope: "project", uid: projectUid, label: projectLabel });
            }
            else {
                const nextAuthor = fields.author.trim() || undefined;
                if (project.author !== nextAuthor) {
                    changes.push({ scope: "project", uid: projectUid, label: projectLabel, field: "author", before: project.author, after: nextAuthor || "(cleared)" });
                    project.author = nextAuthor;
                }
            }
        }
        if (fields.company !== undefined) {
            handledFields.add("company");
            if (typeof fields.company !== "string") {
                warnings.push({ message: "update_project.company は文字列が必要です", scope: "project", uid: projectUid, label: projectLabel });
            }
            else {
                const nextCompany = fields.company.trim() || undefined;
                if (project.company !== nextCompany) {
                    changes.push({ scope: "project", uid: projectUid, label: projectLabel, field: "company", before: project.company, after: nextCompany || "(cleared)" });
                    project.company = nextCompany;
                }
            }
        }
        if (fields.current_date !== undefined) {
            handledFields.add("current_date");
            const normalizedCurrentDate = normalizePatchedPlainDateTime(fields.current_date, "start", project);
            if (normalizedCurrentDate === undefined) {
                warnings.push({ message: "update_project.current_date の日付形式が解釈できません", scope: "project", uid: projectUid, label: projectLabel });
            }
            else if (project.currentDate !== normalizedCurrentDate) {
                changes.push({ scope: "project", uid: projectUid, label: projectLabel, field: "currentDate", before: project.currentDate, after: normalizedCurrentDate });
                project.currentDate = normalizedCurrentDate;
            }
        }
        let nextStartDate = project.startDate;
        let nextFinishDate = project.finishDate;
        let startDateTouched = false;
        let finishDateTouched = false;
        if (fields.start_date !== undefined) {
            handledFields.add("start_date");
            const normalizedStartDate = normalizePatchedPlainDateTime(fields.start_date, "start", project);
            if (normalizedStartDate === undefined) {
                warnings.push({ message: "update_project.start_date の日付形式が解釈できません", scope: "project", uid: projectUid, label: projectLabel });
            }
            else {
                nextStartDate = normalizedStartDate;
                startDateTouched = true;
            }
        }
        if (fields.finish_date !== undefined) {
            handledFields.add("finish_date");
            const normalizedFinishDate = normalizePatchedPlainDateTime(fields.finish_date, "finish", project);
            if (normalizedFinishDate === undefined) {
                warnings.push({ message: "update_project.finish_date の日付形式が解釈できません", scope: "project", uid: projectUid, label: projectLabel });
            }
            else {
                nextFinishDate = normalizedFinishDate;
                finishDateTouched = true;
            }
        }
        if (nextStartDate > nextFinishDate) {
            if (startDateTouched || finishDateTouched) {
                warnings.push({ message: "update_project.start_date が finish_date より後です", scope: "project", uid: projectUid, label: projectLabel });
            }
            startDateTouched = false;
            finishDateTouched = false;
            nextStartDate = project.startDate;
            nextFinishDate = project.finishDate;
        }
        if (startDateTouched && project.startDate !== nextStartDate) {
            changes.push({ scope: "project", uid: projectUid, label: projectLabel, field: "startDate", before: project.startDate, after: nextStartDate });
            project.startDate = nextStartDate;
        }
        if (finishDateTouched && project.finishDate !== nextFinishDate) {
            changes.push({ scope: "project", uid: projectUid, label: projectLabel, field: "finishDate", before: project.finishDate, after: nextFinishDate });
            project.finishDate = nextFinishDate;
        }
        if (fields.status_date !== undefined) {
            handledFields.add("status_date");
            const normalizedStatusDate = normalizePatchedPlainDateTime(fields.status_date, "start", project);
            if (normalizedStatusDate === undefined) {
                warnings.push({ message: "update_project.status_date の日付形式が解釈できません", scope: "project", uid: projectUid, label: projectLabel });
            }
            else if (project.statusDate !== normalizedStatusDate) {
                changes.push({ scope: "project", uid: projectUid, label: projectLabel, field: "statusDate", before: project.statusDate, after: normalizedStatusDate });
                project.statusDate = normalizedStatusDate;
            }
        }
        if (fields.calendar_uid !== undefined) {
            handledFields.add("calendar_uid");
            if (typeof fields.calendar_uid !== "string") {
                warnings.push({ message: "update_project.calendar_uid は文字列が必要です", scope: "project", uid: projectUid, label: projectLabel });
            }
            else {
                const nextCalendarUid = fields.calendar_uid.trim() || undefined;
                if (nextCalendarUid && !model.calendars.some((calendar) => calendar.uid === nextCalendarUid)) {
                    warnings.push({ message: `update_project.calendar_uid が既存 calendar を指していません: ${nextCalendarUid}`, scope: "project", uid: projectUid, label: projectLabel });
                }
                else if (project.calendarUID !== nextCalendarUid) {
                    changes.push({ scope: "project", uid: projectUid, label: projectLabel, field: "calendarUID", before: project.calendarUID, after: nextCalendarUid || "(cleared)" });
                    project.calendarUID = nextCalendarUid;
                }
            }
        }
        if (fields.minutes_per_day !== undefined) {
            handledFields.add("minutes_per_day");
            if (typeof fields.minutes_per_day !== "number" || !Number.isFinite(fields.minutes_per_day) || fields.minutes_per_day <= 0) {
                warnings.push({ message: "update_project.minutes_per_day は 0 より大きい数値が必要です", scope: "project", uid: projectUid, label: projectLabel });
            }
            else if (project.minutesPerDay !== fields.minutes_per_day) {
                changes.push({ scope: "project", uid: projectUid, label: projectLabel, field: "minutesPerDay", before: project.minutesPerDay, after: fields.minutes_per_day });
                project.minutesPerDay = fields.minutes_per_day;
            }
        }
        if (fields.minutes_per_week !== undefined) {
            handledFields.add("minutes_per_week");
            if (typeof fields.minutes_per_week !== "number" || !Number.isFinite(fields.minutes_per_week) || fields.minutes_per_week <= 0) {
                warnings.push({ message: "update_project.minutes_per_week は 0 より大きい数値が必要です", scope: "project", uid: projectUid, label: projectLabel });
            }
            else if (project.minutesPerWeek !== fields.minutes_per_week) {
                changes.push({ scope: "project", uid: projectUid, label: projectLabel, field: "minutesPerWeek", before: project.minutesPerWeek, after: fields.minutes_per_week });
                project.minutesPerWeek = fields.minutes_per_week;
            }
        }
        if (fields.days_per_month !== undefined) {
            handledFields.add("days_per_month");
            if (typeof fields.days_per_month !== "number" || !Number.isFinite(fields.days_per_month) || fields.days_per_month <= 0) {
                warnings.push({ message: "update_project.days_per_month は 0 より大きい数値が必要です", scope: "project", uid: projectUid, label: projectLabel });
            }
            else if (project.daysPerMonth !== fields.days_per_month) {
                changes.push({ scope: "project", uid: projectUid, label: projectLabel, field: "daysPerMonth", before: project.daysPerMonth, after: fields.days_per_month });
                project.daysPerMonth = fields.days_per_month;
            }
        }
        if (fields.schedule_from_start !== undefined) {
            handledFields.add("schedule_from_start");
            if (typeof fields.schedule_from_start !== "boolean") {
                warnings.push({ message: "update_project.schedule_from_start は boolean が必要です", scope: "project", uid: projectUid, label: projectLabel });
            }
            else if (project.scheduleFromStart !== fields.schedule_from_start) {
                changes.push({ scope: "project", uid: projectUid, label: projectLabel, field: "scheduleFromStart", before: project.scheduleFromStart, after: fields.schedule_from_start });
                project.scheduleFromStart = fields.schedule_from_start;
            }
        }
        Object.keys(fields).forEach((fieldName) => {
            if (handledFields.has(fieldName)) {
                return;
            }
            warnings.push({ message: `未対応の field は無視します: operations[${operationIndex}].fields.${fieldName}`, scope: "project", uid: projectUid, label: projectLabel });
        });
    }
    function applyUpdateAssignmentOperation(assignment, rawFields, project, changes, warnings, operationIndex) {
        const fields = { ...rawFields };
        if (Object.keys(fields).length === 0) {
            warnings.push({ message: `update_assignment に fields がありません: ${assignment.uid}`, scope: "assignments", uid: assignment.uid, label: assignment.uid });
            return;
        }
        const handledFields = new Set();
        let nextStart = assignment.start;
        let nextFinish = assignment.finish;
        let startTouched = false;
        let finishTouched = false;
        if (fields.start !== undefined) {
            handledFields.add("start");
            const normalizedStart = normalizePatchedPlainDateTime(fields.start, "start", project);
            if (normalizedStart === undefined) {
                warnings.push({ message: `update_assignment.start の日付形式が解釈できません: ${assignment.uid}`, scope: "assignments", uid: assignment.uid, label: assignment.uid });
            }
            else {
                nextStart = normalizedStart;
                startTouched = true;
            }
        }
        if (fields.finish !== undefined) {
            handledFields.add("finish");
            const normalizedFinish = normalizePatchedPlainDateTime(fields.finish, "finish", project);
            if (normalizedFinish === undefined) {
                warnings.push({ message: `update_assignment.finish の日付形式が解釈できません: ${assignment.uid}`, scope: "assignments", uid: assignment.uid, label: assignment.uid });
            }
            else {
                nextFinish = normalizedFinish;
                finishTouched = true;
            }
        }
        if (nextStart && nextFinish && nextStart > nextFinish) {
            warnings.push({ message: `update_assignment.start が finish より後です: ${assignment.uid}`, scope: "assignments", uid: assignment.uid, label: assignment.uid });
            startTouched = false;
            finishTouched = false;
            nextStart = assignment.start;
            nextFinish = assignment.finish;
        }
        if (startTouched && assignment.start !== nextStart) {
            changes.push({
                scope: "assignments",
                uid: assignment.uid,
                label: assignment.uid,
                field: "start",
                before: assignment.start,
                after: nextStart
            });
            assignment.start = nextStart;
        }
        if (finishTouched && assignment.finish !== nextFinish) {
            changes.push({
                scope: "assignments",
                uid: assignment.uid,
                label: assignment.uid,
                field: "finish",
                before: assignment.finish,
                after: nextFinish
            });
            assignment.finish = nextFinish;
        }
        if (fields.units !== undefined) {
            handledFields.add("units");
            if (typeof fields.units !== "number" || !Number.isFinite(fields.units) || fields.units < 0) {
                warnings.push({ message: `update_assignment.units は 0 以上の数値が必要です: ${assignment.uid}`, scope: "assignments", uid: assignment.uid, label: assignment.uid });
            }
            else if (assignment.units !== fields.units) {
                changes.push({
                    scope: "assignments",
                    uid: assignment.uid,
                    label: assignment.uid,
                    field: "units",
                    before: assignment.units,
                    after: fields.units
                });
                assignment.units = fields.units;
            }
        }
        if (fields.work !== undefined) {
            handledFields.add("work");
            if (typeof fields.work !== "string" || fields.work.trim() === "") {
                warnings.push({ message: `update_assignment.work は空でない文字列が必要です: ${assignment.uid}`, scope: "assignments", uid: assignment.uid, label: assignment.uid });
            }
            else if (assignment.work !== fields.work.trim()) {
                changes.push({
                    scope: "assignments",
                    uid: assignment.uid,
                    label: assignment.uid,
                    field: "work",
                    before: assignment.work,
                    after: fields.work.trim()
                });
                assignment.work = fields.work.trim();
            }
        }
        if (fields.percent_work_complete !== undefined) {
            handledFields.add("percent_work_complete");
            if (typeof fields.percent_work_complete !== "number" || !Number.isFinite(fields.percent_work_complete) || fields.percent_work_complete < 0 || fields.percent_work_complete > 100) {
                warnings.push({ message: `update_assignment.percent_work_complete は 0 以上 100 以下の数値が必要です: ${assignment.uid}`, scope: "assignments", uid: assignment.uid, label: assignment.uid });
            }
            else if (assignment.percentWorkComplete !== fields.percent_work_complete) {
                changes.push({
                    scope: "assignments",
                    uid: assignment.uid,
                    label: assignment.uid,
                    field: "percentWorkComplete",
                    before: assignment.percentWorkComplete,
                    after: fields.percent_work_complete
                });
                assignment.percentWorkComplete = fields.percent_work_complete;
            }
        }
        Object.keys(fields).forEach((fieldName) => {
            if (handledFields.has(fieldName)) {
                return;
            }
            warnings.push({ message: `未対応の field は無視します: operations[${operationIndex}].fields.${fieldName}`, scope: "assignments", uid: assignment.uid, label: assignment.uid });
        });
    }
    function applyUpdateResourceOperation(resource, rawFields, model, changes, warnings, operationIndex) {
        const fields = { ...rawFields };
        if (Object.keys(fields).length === 0) {
            warnings.push({ message: `update_resource に fields がありません: ${resource.uid}`, scope: "resources", uid: resource.uid, label: resource.name || resource.uid });
            return;
        }
        const handledFields = new Set();
        if (fields.name !== undefined) {
            handledFields.add("name");
            if (typeof fields.name !== "string" || fields.name.trim() === "") {
                warnings.push({ message: `update_resource.name は空でない文字列が必要です: ${resource.uid}`, scope: "resources", uid: resource.uid, label: resource.name || resource.uid });
            }
            else if (resource.name !== fields.name.trim()) {
                changes.push({
                    scope: "resources",
                    uid: resource.uid,
                    label: resource.name || resource.uid,
                    field: "name",
                    before: resource.name,
                    after: fields.name.trim()
                });
                resource.name = fields.name.trim();
            }
        }
        if (fields.initials !== undefined) {
            handledFields.add("initials");
            if (typeof fields.initials !== "string") {
                warnings.push({ message: `update_resource.initials は文字列が必要です: ${resource.uid}`, scope: "resources", uid: resource.uid, label: resource.name || resource.uid });
            }
            else {
                const nextInitials = fields.initials.trim() || undefined;
                if (resource.initials !== nextInitials) {
                    changes.push({
                        scope: "resources",
                        uid: resource.uid,
                        label: resource.name || resource.uid,
                        field: "initials",
                        before: resource.initials,
                        after: nextInitials || "(cleared)"
                    });
                    resource.initials = nextInitials;
                }
            }
        }
        if (fields.group !== undefined) {
            handledFields.add("group");
            if (typeof fields.group !== "string") {
                warnings.push({ message: `update_resource.group は文字列が必要です: ${resource.uid}`, scope: "resources", uid: resource.uid, label: resource.name || resource.uid });
            }
            else {
                const nextGroup = fields.group.trim() || undefined;
                if (resource.group !== nextGroup) {
                    changes.push({
                        scope: "resources",
                        uid: resource.uid,
                        label: resource.name || resource.uid,
                        field: "group",
                        before: resource.group,
                        after: nextGroup || "(cleared)"
                    });
                    resource.group = nextGroup;
                }
            }
        }
        if (fields.calendar_uid !== undefined) {
            handledFields.add("calendar_uid");
            if (typeof fields.calendar_uid !== "string") {
                warnings.push({ message: `update_resource.calendar_uid は文字列が必要です: ${resource.uid}`, scope: "resources", uid: resource.uid, label: resource.name || resource.uid });
            }
            else {
                const nextCalendarUid = fields.calendar_uid.trim() || undefined;
                if (nextCalendarUid && !model.calendars.some((calendar) => calendar.uid === nextCalendarUid)) {
                    warnings.push({ message: `update_resource.calendar_uid が既存 calendar を指していません: ${resource.uid} -> ${nextCalendarUid}`, scope: "resources", uid: resource.uid, label: resource.name || resource.uid });
                }
                else if (resource.calendarUID !== nextCalendarUid) {
                    changes.push({
                        scope: "resources",
                        uid: resource.uid,
                        label: resource.name || resource.uid,
                        field: "calendarUID",
                        before: resource.calendarUID,
                        after: nextCalendarUid || "(cleared)"
                    });
                    resource.calendarUID = nextCalendarUid;
                }
            }
        }
        if (fields.max_units !== undefined) {
            handledFields.add("max_units");
            if (typeof fields.max_units !== "number" || !Number.isFinite(fields.max_units) || fields.max_units < 0) {
                warnings.push({ message: `update_resource.max_units は 0 以上の数値が必要です: ${resource.uid}`, scope: "resources", uid: resource.uid, label: resource.name || resource.uid });
            }
            else if (resource.maxUnits !== fields.max_units) {
                changes.push({
                    scope: "resources",
                    uid: resource.uid,
                    label: resource.name || resource.uid,
                    field: "maxUnits",
                    before: resource.maxUnits,
                    after: fields.max_units
                });
                resource.maxUnits = fields.max_units;
            }
        }
        if (fields.standard_rate !== undefined) {
            handledFields.add("standard_rate");
            if (typeof fields.standard_rate !== "string") {
                warnings.push({ message: `update_resource.standard_rate は文字列が必要です: ${resource.uid}`, scope: "resources", uid: resource.uid, label: resource.name || resource.uid });
            }
            else {
                const nextStandardRate = fields.standard_rate.trim() || undefined;
                if (resource.standardRate !== nextStandardRate) {
                    changes.push({
                        scope: "resources",
                        uid: resource.uid,
                        label: resource.name || resource.uid,
                        field: "standardRate",
                        before: resource.standardRate,
                        after: nextStandardRate || "(cleared)"
                    });
                    resource.standardRate = nextStandardRate;
                }
            }
        }
        if (fields.overtime_rate !== undefined) {
            handledFields.add("overtime_rate");
            if (typeof fields.overtime_rate !== "string") {
                warnings.push({ message: `update_resource.overtime_rate は文字列が必要です: ${resource.uid}`, scope: "resources", uid: resource.uid, label: resource.name || resource.uid });
            }
            else {
                const nextOvertimeRate = fields.overtime_rate.trim() || undefined;
                if (resource.overtimeRate !== nextOvertimeRate) {
                    changes.push({
                        scope: "resources",
                        uid: resource.uid,
                        label: resource.name || resource.uid,
                        field: "overtimeRate",
                        before: resource.overtimeRate,
                        after: nextOvertimeRate || "(cleared)"
                    });
                    resource.overtimeRate = nextOvertimeRate;
                }
            }
        }
        if (fields.cost_per_use !== undefined) {
            handledFields.add("cost_per_use");
            if (typeof fields.cost_per_use !== "number" || !Number.isFinite(fields.cost_per_use) || fields.cost_per_use < 0) {
                warnings.push({ message: `update_resource.cost_per_use は 0 以上の数値が必要です: ${resource.uid}`, scope: "resources", uid: resource.uid, label: resource.name || resource.uid });
            }
            else if (resource.costPerUse !== fields.cost_per_use) {
                changes.push({
                    scope: "resources",
                    uid: resource.uid,
                    label: resource.name || resource.uid,
                    field: "costPerUse",
                    before: resource.costPerUse,
                    after: fields.cost_per_use
                });
                resource.costPerUse = fields.cost_per_use;
            }
        }
        if (fields.percent_work_complete !== undefined) {
            handledFields.add("percent_work_complete");
            if (typeof fields.percent_work_complete !== "number" || !Number.isFinite(fields.percent_work_complete) || fields.percent_work_complete < 0 || fields.percent_work_complete > 100) {
                warnings.push({ message: `update_resource.percent_work_complete は 0 以上 100 以下の数値が必要です: ${resource.uid}`, scope: "resources", uid: resource.uid, label: resource.name || resource.uid });
            }
            else if (resource.percentWorkComplete !== fields.percent_work_complete) {
                changes.push({
                    scope: "resources",
                    uid: resource.uid,
                    label: resource.name || resource.uid,
                    field: "percentWorkComplete",
                    before: resource.percentWorkComplete,
                    after: fields.percent_work_complete
                });
                resource.percentWorkComplete = fields.percent_work_complete;
            }
        }
        Object.keys(fields).forEach((fieldName) => {
            if (handledFields.has(fieldName)) {
                return;
            }
            warnings.push({ message: `未対応の field は無視します: operations[${operationIndex}].fields.${fieldName}`, scope: "resources", uid: resource.uid, label: resource.name || resource.uid });
        });
    }
    function applyUpdateCalendarOperation(calendar, rawFields, model, changes, warnings, operationIndex) {
        const fields = { ...rawFields };
        if (Object.keys(fields).length === 0) {
            warnings.push({ message: `update_calendar に fields がありません: ${calendar.uid}`, scope: "calendars", uid: calendar.uid, label: calendar.name || calendar.uid });
            return;
        }
        const handledFields = new Set();
        if (fields.name !== undefined) {
            handledFields.add("name");
            if (typeof fields.name !== "string" || fields.name.trim() === "") {
                warnings.push({ message: `update_calendar.name は空でない文字列が必要です: ${calendar.uid}`, scope: "calendars", uid: calendar.uid, label: calendar.name || calendar.uid });
            }
            else if (calendar.name !== fields.name.trim()) {
                changes.push({
                    scope: "calendars",
                    uid: calendar.uid,
                    label: calendar.name || calendar.uid,
                    field: "name",
                    before: calendar.name,
                    after: fields.name.trim()
                });
                calendar.name = fields.name.trim();
            }
        }
        if (fields.is_base_calendar !== undefined) {
            handledFields.add("is_base_calendar");
            if (typeof fields.is_base_calendar !== "boolean") {
                warnings.push({ message: `update_calendar.is_base_calendar は boolean が必要です: ${calendar.uid}`, scope: "calendars", uid: calendar.uid, label: calendar.name || calendar.uid });
            }
            else if (calendar.isBaseCalendar !== fields.is_base_calendar) {
                changes.push({
                    scope: "calendars",
                    uid: calendar.uid,
                    label: calendar.name || calendar.uid,
                    field: "isBaseCalendar",
                    before: calendar.isBaseCalendar,
                    after: fields.is_base_calendar
                });
                calendar.isBaseCalendar = fields.is_base_calendar;
            }
        }
        if (fields.base_calendar_uid !== undefined) {
            handledFields.add("base_calendar_uid");
            if (typeof fields.base_calendar_uid !== "string") {
                warnings.push({ message: `update_calendar.base_calendar_uid は文字列が必要です: ${calendar.uid}`, scope: "calendars", uid: calendar.uid, label: calendar.name || calendar.uid });
            }
            else {
                const nextBaseCalendarUid = fields.base_calendar_uid.trim() || undefined;
                if (nextBaseCalendarUid === calendar.uid) {
                    warnings.push({ message: `update_calendar.base_calendar_uid は自身を指せません: ${calendar.uid}`, scope: "calendars", uid: calendar.uid, label: calendar.name || calendar.uid });
                }
                else if (nextBaseCalendarUid && !model.calendars.some((item) => item.uid === nextBaseCalendarUid)) {
                    warnings.push({ message: `update_calendar.base_calendar_uid が既存 calendar を指していません: ${calendar.uid} -> ${nextBaseCalendarUid}`, scope: "calendars", uid: calendar.uid, label: calendar.name || calendar.uid });
                }
                else if (calendar.baseCalendarUID !== nextBaseCalendarUid) {
                    changes.push({
                        scope: "calendars",
                        uid: calendar.uid,
                        label: calendar.name || calendar.uid,
                        field: "baseCalendarUID",
                        before: calendar.baseCalendarUID,
                        after: nextBaseCalendarUid || "(cleared)"
                    });
                    calendar.baseCalendarUID = nextBaseCalendarUid;
                }
            }
        }
        Object.keys(fields).forEach((fieldName) => {
            if (handledFields.has(fieldName)) {
                return;
            }
            warnings.push({ message: `未対応の field は無視します: operations[${operationIndex}].fields.${fieldName}`, scope: "calendars", uid: calendar.uid, label: calendar.name || calendar.uid });
        });
    }
    function normalizePatchedTaskDate(value, kind, task, project) {
        if (typeof value !== "string") {
            return undefined;
        }
        const trimmed = value.trim();
        if (!trimmed) {
            return undefined;
        }
        if (!isDateText(trimmed)) {
            return undefined;
        }
        if (isDateOnlyText(trimmed) && !task.milestone) {
            const timeText = kind === "start"
                ? (project.defaultStartTime || "09:00:00")
                : (project.defaultFinishTime || "18:00:00");
            return `${trimmed}T${timeText}`;
        }
        return trimmed;
    }
    function normalizePatchedPlainDateTime(value, kind, project) {
        if (typeof value !== "string") {
            return undefined;
        }
        const trimmed = value.trim();
        if (!trimmed) {
            return undefined;
        }
        if (!isDateText(trimmed)) {
            return undefined;
        }
        if (isDateOnlyText(trimmed)) {
            const timeText = kind === "start"
                ? (project.defaultStartTime || "09:00:00")
                : (project.defaultFinishTime || "18:00:00");
            return `${trimmed}T${timeText}`;
        }
        return trimmed;
    }
    function isDateOnlyText(value) {
        return /^\d{4}-\d{2}-\d{2}$/.test(value);
    }
    function isDateText(value) {
        if (isDateOnlyText(value)) {
            return true;
        }
        return !Number.isNaN(new Date(value).getTime());
    }
    function formatDurationHours(hours) {
        const totalSeconds = Math.round(hours * 60 * 60);
        const normalizedSeconds = Math.max(0, totalSeconds);
        const durationHours = Math.floor(normalizedSeconds / 3600);
        const durationMinutes = Math.floor((normalizedSeconds % 3600) / 60);
        const durationSeconds = normalizedSeconds % 60;
        return `PT${durationHours}H${durationMinutes}M${durationSeconds}S`;
    }
    function parseDurationHours(duration) {
        const text = String(duration || "").trim();
        const match = text.match(/^PT(?:(\d+(?:\.\d+)?)H)?(?:(\d+(?:\.\d+)?)M)?(?:(\d+(?:\.\d+)?)S)?$/);
        if (!match) {
            return undefined;
        }
        const hours = Number(match[1] || 0);
        const minutes = Number(match[2] || 0);
        const seconds = Number(match[3] || 0);
        return hours + (minutes / 60) + (seconds / 3600);
    }
    function cloneProjectModel(model) {
        return JSON.parse(JSON.stringify(model));
    }
    globalThis.__mikuprojectProjectPatchJson = {
        importProjectPatchJson,
        validatePatchDocument
    };
})();
