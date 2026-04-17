/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    function isPlaceholderUid(value) {
        return String(value || "").trim() === "0";
    }
    function formatDependencyType(type) {
        if (type === undefined) {
            return "FS";
        }
        const typeMap = {
            0: "FF",
            1: "FS",
            2: "FF",
            3: "SF",
            4: "SS"
        };
        return typeMap[type] || String(type);
    }
    function parseDurationHours(duration) {
        const text = String(duration || "").trim();
        if (!text) {
            return undefined;
        }
        const match = /^(-)?P(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)$/.exec(text);
        if (!match) {
            return undefined;
        }
        const sign = match[1] ? -1 : 1;
        const hours = Number(match[2] || 0);
        const minutes = Number(match[3] || 0);
        const seconds = Number(match[4] || 0);
        return sign * (hours + minutes / 60 + seconds / 3600);
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
    function collectTopLevelPhases(tasks) {
        return tasks.filter((task) => !isPlaceholderUid(task.uid) && task.summary && task.outlineLevel === 1);
    }
    function collectPhaseTaskUids(tasks, phaseUid) {
        const phaseIndex = tasks.findIndex((task) => task.uid === phaseUid);
        if (phaseIndex < 0) {
            return new Set();
        }
        const phase = tasks[phaseIndex];
        const uids = new Set();
        for (let index = phaseIndex + 1; index < tasks.length; index += 1) {
            const task = tasks[index];
            if (task.outlineLevel <= phase.outlineLevel) {
                break;
            }
            if (!isPlaceholderUid(task.uid)) {
                uids.add(task.uid);
            }
        }
        return uids;
    }
    function collectTaskSubtreeUids(tasks, rootUid, maxDepth) {
        const rootIndex = tasks.findIndex((task) => task.uid === rootUid);
        if (rootIndex < 0) {
            return new Set();
        }
        const rootTask = tasks[rootIndex];
        const uids = new Set();
        if (!isPlaceholderUid(rootTask.uid)) {
            uids.add(rootTask.uid);
        }
        for (let index = rootIndex + 1; index < tasks.length; index += 1) {
            const task = tasks[index];
            if (task.outlineLevel <= rootTask.outlineLevel) {
                break;
            }
            const relativeDepth = task.outlineLevel - rootTask.outlineLevel;
            if (typeof maxDepth === "number" && relativeDepth > maxDepth) {
                continue;
            }
            if (!isPlaceholderUid(task.uid)) {
                uids.add(task.uid);
            }
        }
        return uids;
    }
    function resolvePhaseUidForTask(taskUid, parentMap, phaseUidSet) {
        let currentUid = taskUid;
        while (currentUid) {
            if (phaseUidSet.has(currentUid)) {
                return currentUid;
            }
            currentUid = parentMap.get(currentUid) || null;
        }
        return undefined;
    }
    function buildDefaultRules(scope) {
        if (scope === "project_overview_view") {
            return {
                allow_patch_ops: ["add_task", "update_task", "move_task"],
                forbid_completed_task_changes: true,
                forbid_summary_task_direct_edit: true
            };
        }
        if (scope === "task_edit_view") {
            return {
                allow_patch_ops: [
                    "update_task",
                    "move_task",
                    "link_tasks",
                    "unlink_tasks",
                    "add_task",
                    "delete_task",
                    "update_assignment",
                    "add_assignment",
                    "delete_assignment"
                ],
                allowed_edit_fields: [
                    "name",
                    "notes",
                    "calendar_uid",
                    "percent_complete",
                    "percent_work_complete",
                    "critical",
                    "planned_start",
                    "planned_finish",
                    "planned_duration",
                    "planned_duration_hours",
                    "is_milestone"
                ],
                forbid_completed_task_changes: true,
                forbid_summary_task_direct_edit: true
            };
        }
        return {
            allow_patch_ops: ["add_task", "update_task", "move_task", "link_tasks", "unlink_tasks"],
            forbid_completed_task_changes: true,
            forbid_summary_task_direct_edit: true
        };
    }
    function toIsoLocalString(value) {
        const year = value.getFullYear();
        const month = String(value.getMonth() + 1).padStart(2, "0");
        const day = String(value.getDate()).padStart(2, "0");
        const hour = String(value.getHours()).padStart(2, "0");
        const minute = String(value.getMinutes()).padStart(2, "0");
        const second = String(value.getSeconds()).padStart(2, "0");
        return `${year}-${month}-${day}T${hour}:${minute}:${second}`;
    }
    function addHoursToDateTime(dateTime, hours) {
        const parsed = new Date(dateTime);
        if (Number.isNaN(parsed.getTime())) {
            return dateTime;
        }
        parsed.setTime(parsed.getTime() + (hours * 60 * 60 * 1000));
        return toIsoLocalString(parsed);
    }
    function isDateOnlyText(value) {
        return typeof value === "string" && /^\d{4}-\d{2}-\d{2}$/.test(value.trim());
    }
    function withTimeOnDate(dateText, timeText) {
        return `${dateText}T${timeText}`;
    }
    const aiViews = {
        buildProjectDraftRequest(input) {
            return {
                view_type: "project_draft_request",
                project: {
                    name: input.name,
                    planned_start: input.plannedStart || undefined
                },
                requirements: {
                    goal: input.goal || undefined,
                    team_count: input.teamCount,
                    must_have_phases: input.mustHavePhases || [],
                    must_have_milestones: input.mustHaveMilestones || []
                }
            };
        },
        importProjectDraftView(input) {
            var _a, _b;
            const { draft } = input;
            if (!draft || typeof draft !== "object") {
                throw new Error("project_draft_view がオブジェクトではありません");
            }
            const data = draft;
            if (data.view_type !== "project_draft_view") {
                throw new Error("view_type が project_draft_view ではありません");
            }
            if (!((_b = (_a = data.project) === null || _a === void 0 ? void 0 : _a.name) === null || _b === void 0 ? void 0 : _b.trim())) {
                throw new Error("project.name がありません");
            }
            const inputTasks = Array.isArray(data.tasks) ? data.tasks : [];
            const inputResources = Array.isArray(data.resources) ? data.resources : [];
            const inputAssignments = Array.isArray(data.assignments) ? data.assignments : [];
            const seenUids = new Set();
            for (const task of inputTasks) {
                const uid = String(task.uid || "").trim();
                if (!uid) {
                    throw new Error("draft task の uid がありません");
                }
                if (seenUids.has(uid)) {
                    throw new Error(`draft task の uid が重複しています: ${uid}`);
                }
                seenUids.add(uid);
                if (!String(task.name || "").trim()) {
                    throw new Error(`draft task の name がありません: ${uid}`);
                }
            }
            for (const task of inputTasks) {
                const parentUid = task.parent_uid == null || task.parent_uid === "" ? null : String(task.parent_uid);
                if (parentUid && !seenUids.has(parentUid)) {
                    throw new Error(`draft task の parent_uid が既存 uid を指していません: ${String(task.uid || "")} -> ${parentUid}`);
                }
            }
            const seenResourceUids = new Set();
            for (const resource of inputResources) {
                const uid = String(resource.uid || "").trim();
                if (!uid) {
                    throw new Error("draft resource の uid がありません");
                }
                if (seenResourceUids.has(uid)) {
                    throw new Error(`draft resource の uid が重複しています: ${uid}`);
                }
                seenResourceUids.add(uid);
                if (!String(resource.name || "").trim()) {
                    throw new Error(`draft resource の name がありません: ${uid}`);
                }
            }
            for (const assignment of inputAssignments) {
                const uid = String(assignment.uid || "").trim();
                if (!uid) {
                    throw new Error("draft assignment の uid がありません");
                }
                const taskUid = String(assignment.task_uid || "").trim();
                const resourceUid = String(assignment.resource_uid || "").trim();
                if (!taskUid || !seenUids.has(taskUid)) {
                    throw new Error(`draft assignment の task_uid が既存 uid を指していません: ${uid} -> ${taskUid}`);
                }
                if (!resourceUid || !seenResourceUids.has(resourceUid)) {
                    throw new Error(`draft assignment の resource_uid が既存 uid を指していません: ${uid} -> ${resourceUid}`);
                }
            }
            const projectStart = data.project.planned_start || data.project.planned_finish || toIsoLocalString(new Date());
            const draftUidMap = new Map();
            inputTasks.forEach((task, index) => {
                draftUidMap.set(String(task.uid || "").trim(), String(index + 1));
            });
            const draftResourceUidMap = new Map();
            inputResources.forEach((resource, index) => {
                draftResourceUidMap.set(String(resource.uid || "").trim(), String(index + 1));
            });
            const normalizedTasks = inputTasks.map((task, index) => ({
                uid: draftUidMap.get(String(task.uid || "").trim()) || String(index + 1),
                name: String(task.name || "").trim(),
                parentUid: task.parent_uid == null || task.parent_uid === "" ? null : (draftUidMap.get(String(task.parent_uid)) || null),
                position: typeof task.position === "number" && Number.isFinite(task.position) ? task.position : index,
                isSummary: Boolean(task.is_summary),
                isMilestone: Boolean(task.is_milestone),
                percentComplete: typeof task.percent_complete === "number" && Number.isFinite(task.percent_complete)
                    ? Math.max(0, Math.min(100, task.percent_complete))
                    : 0,
                plannedDuration: task.planned_duration || undefined,
                plannedDurationHours: typeof task.planned_duration_hours === "number" && Number.isFinite(task.planned_duration_hours)
                    ? task.planned_duration_hours
                    : undefined,
                plannedStart: task.planned_start || undefined,
                plannedFinish: task.planned_finish || undefined,
                predecessorUids: [
                    ...(Array.isArray(task.predecessor_uids) ? task.predecessor_uids : []),
                    ...(Array.isArray(task.predecessors)
                        ? task.predecessors.flatMap((item) => {
                            if (typeof item === "string") {
                                return [item];
                            }
                            return (item === null || item === void 0 ? void 0 : item.task_uid) ? [item.task_uid] : [];
                        })
                        : [])
                ].map((item) => draftUidMap.get(String(item)) || String(item))
            }));
            const byParent = new Map();
            for (const task of normalizedTasks) {
                const siblings = byParent.get(task.parentUid) || [];
                siblings.push(task);
                byParent.set(task.parentUid, siblings);
            }
            for (const siblings of byParent.values()) {
                siblings.sort((left, right) => left.position - right.position || left.uid.localeCompare(right.uid));
            }
            const orderedTasks = [];
            function walk(parentUid, outlinePath) {
                const siblings = byParent.get(parentUid) || [];
                siblings.forEach((task, index) => {
                    const currentPath = [...outlinePath, index + 1];
                    const outlineNumber = currentPath.join(".");
                    let start = task.plannedStart || task.plannedFinish || projectStart;
                    let finish = task.plannedFinish
                        || (typeof task.plannedDurationHours === "number" ? addHoursToDateTime(start, task.plannedDurationHours) : start);
                    const dateOnlyTaskRange = !task.isMilestone
                        && isDateOnlyText(start)
                        && isDateOnlyText(finish)
                        && task.plannedDuration == null
                        && typeof task.plannedDurationHours !== "number";
                    if (dateOnlyTaskRange) {
                        start = withTimeOnDate(start, "09:00:00");
                        finish = withTimeOnDate(finish, "18:00:00");
                    }
                    const hasChildren = (byParent.get(task.uid) || []).length > 0;
                    orderedTasks.push({
                        uid: task.uid,
                        id: task.uid,
                        name: task.name,
                        outlineLevel: currentPath.length,
                        outlineNumber,
                        wbs: outlineNumber,
                        start,
                        finish,
                        duration: task.plannedDuration || (typeof task.plannedDurationHours === "number" ? `PT${task.plannedDurationHours}H` : "PT0H0M0S"),
                        milestone: task.isMilestone,
                        summary: task.isSummary || hasChildren,
                        percentComplete: task.percentComplete,
                        predecessors: task.predecessorUids.map((predecessorUid) => ({ predecessorUid })),
                        extendedAttributes: [],
                        baselines: [],
                        timephasedData: []
                    });
                    walk(task.uid, currentPath);
                });
            }
            walk(null, []);
            const taskFinishes = orderedTasks.map((task) => task.finish).filter(Boolean).sort();
            const orderedResources = inputResources.map((resource, index) => ({
                uid: draftResourceUidMap.get(String(resource.uid || "").trim()) || String(index + 1),
                id: draftResourceUidMap.get(String(resource.uid || "").trim()) || String(index + 1),
                name: String(resource.name || "").trim(),
                initials: String(resource.initials || "").trim() || undefined,
                group: String(resource.group || "").trim() || undefined,
                maxUnits: typeof resource.max_units === "number" && Number.isFinite(resource.max_units) ? resource.max_units : undefined,
                calendarUID: String(resource.calendar_uid || "").trim() || undefined,
                extendedAttributes: [],
                baselines: [],
                timephasedData: []
            }));
            const orderedAssignments = inputAssignments.map((assignment, index) => ({
                uid: String(index + 1),
                taskUid: draftUidMap.get(String(assignment.task_uid || "").trim()) || String(assignment.task_uid || "").trim(),
                resourceUid: draftResourceUidMap.get(String(assignment.resource_uid || "").trim()) || String(assignment.resource_uid || "").trim(),
                start: String(assignment.start || "").trim() || undefined,
                finish: String(assignment.finish || "").trim() || undefined,
                units: typeof assignment.units === "number" && Number.isFinite(assignment.units) ? assignment.units : undefined,
                work: String(assignment.work || "").trim() || undefined,
                percentWorkComplete: typeof assignment.percent_work_complete === "number" && Number.isFinite(assignment.percent_work_complete)
                    ? Math.max(0, Math.min(100, assignment.percent_work_complete))
                    : undefined,
                extendedAttributes: [],
                baselines: [],
                timephasedData: []
            }));
            return input.normalizeProjectModel(input.ensureDefaultProjectCalendar({
                project: {
                    name: data.project.name.trim(),
                    title: data.project.name.trim(),
                    startDate: projectStart,
                    finishDate: data.project.planned_finish || taskFinishes.at(-1) || projectStart,
                    scheduleFromStart: data.project.schedule_from_start !== undefined ? Boolean(data.project.schedule_from_start) : true,
                    minutesPerDay: typeof data.project.minutes_per_day === "number" && Number.isFinite(data.project.minutes_per_day)
                        ? data.project.minutes_per_day
                        : input.defaultProjectMinutesPerDay,
                    minutesPerWeek: typeof data.project.minutes_per_week === "number" && Number.isFinite(data.project.minutes_per_week)
                        ? data.project.minutes_per_week
                        : input.defaultProjectMinutesPerWeek,
                    daysPerMonth: typeof data.project.days_per_month === "number" && Number.isFinite(data.project.days_per_month)
                        ? data.project.days_per_month
                        : input.defaultProjectDaysPerMonth,
                    outlineCodes: [],
                    wbsMasks: [],
                    extendedAttributes: []
                },
                tasks: orderedTasks,
                resources: orderedResources,
                assignments: orderedAssignments,
                calendars: []
            }));
        },
        exportProjectOverviewView(model) {
            const parentMap = buildTaskParentMap(model.tasks);
            const phaseTasks = collectTopLevelPhases(model.tasks);
            const phaseUidSet = new Set(phaseTasks.map((task) => task.uid));
            const allMilestones = model.tasks.filter((task) => !isPlaceholderUid(task.uid) && task.milestone);
            const topLevelDependencyMap = new Map();
            for (const task of model.tasks) {
                const toPhaseUid = resolvePhaseUidForTask(task.uid, parentMap, phaseUidSet);
                if (!toPhaseUid) {
                    continue;
                }
                for (const predecessor of task.predecessors) {
                    const fromPhaseUid = resolvePhaseUidForTask(predecessor.predecessorUid, parentMap, phaseUidSet);
                    if (!fromPhaseUid || fromPhaseUid === toPhaseUid) {
                        continue;
                    }
                    const key = `${fromPhaseUid}->${toPhaseUid}:${formatDependencyType(predecessor.type)}`;
                    if (!topLevelDependencyMap.has(key)) {
                        topLevelDependencyMap.set(key, {
                            from_uid: fromPhaseUid,
                            to_uid: toPhaseUid,
                            type: formatDependencyType(predecessor.type)
                        });
                    }
                }
            }
            return {
                view_type: "project_overview_view",
                project: {
                    name: model.project.name,
                    planned_start: model.project.startDate,
                    planned_finish: model.project.finishDate,
                    status_date: model.project.statusDate
                },
                summary: {
                    task_count: model.tasks.filter((task) => !isPlaceholderUid(task.uid)).length,
                    summary_task_count: model.tasks.filter((task) => !isPlaceholderUid(task.uid) && task.summary).length,
                    milestone_count: allMilestones.length,
                    max_outline_level: model.tasks.reduce((max, task) => Math.max(max, task.outlineLevel || 0), 0)
                },
                phases: phaseTasks.map((phase) => {
                    const phaseTaskUids = collectPhaseTaskUids(model.tasks, phase.uid);
                    const descendantTasks = model.tasks.filter((task) => phaseTaskUids.has(task.uid));
                    return {
                        uid: phase.uid,
                        name: phase.name,
                        wbs: phase.wbs || phase.outlineNumber,
                        task_count: descendantTasks.length,
                        milestone_count: descendantTasks.filter((task) => task.milestone).length,
                        planned_start: phase.start,
                        planned_finish: phase.finish,
                        duration: phase.duration,
                        duration_hours: parseDurationHours(phase.duration),
                        percent_complete: phase.percentComplete,
                        sample_tasks: descendantTasks.slice(0, 3).map((task) => ({
                            uid: task.uid,
                            name: task.name
                        }))
                    };
                }),
                milestones: allMilestones.map((task) => ({
                    uid: task.uid,
                    name: task.name,
                    parent_uid: parentMap.get(task.uid),
                    date: task.finish || task.start
                })),
                top_level_dependencies: Array.from(topLevelDependencyMap.values()),
                rules: buildDefaultRules("project_overview_view")
            };
        },
        exportPhaseDetailView(model, requestedPhaseUid, options) {
            var _a;
            const phaseTasks = collectTopLevelPhases(model.tasks);
            if (phaseTasks.length === 0) {
                throw new Error("phase が見つかりません");
            }
            const phase = requestedPhaseUid ? phaseTasks.find((task) => task.uid === requestedPhaseUid) : phaseTasks[0];
            if (!phase) {
                throw new Error(`phase が見つかりません: ${requestedPhaseUid}`);
            }
            const parentMap = buildTaskParentMap(model.tasks);
            const positionMap = buildTaskPositionMap(model.tasks, parentMap);
            const phaseTaskUids = collectPhaseTaskUids(model.tasks, phase.uid);
            const phaseTasksOnly = model.tasks.filter((task) => phaseTaskUids.has(task.uid));
            const mode = (options === null || options === void 0 ? void 0 : options.mode) === "scoped" ? "scoped" : "full";
            const rootUid = mode === "scoped" ? ((_a = options === null || options === void 0 ? void 0 : options.rootUid) === null || _a === void 0 ? void 0 : _a.trim()) || undefined : undefined;
            const maxDepth = mode === "scoped" && typeof (options === null || options === void 0 ? void 0 : options.maxDepth) === "number" && Number.isFinite(options.maxDepth) && options.maxDepth >= 0
                ? Math.floor(options.maxDepth)
                : undefined;
            let scopedTaskUids = phaseTaskUids;
            if (rootUid) {
                const rootTask = phaseTasksOnly.find((task) => task.uid === rootUid);
                if (!rootTask) {
                    throw new Error(`phase 配下に root_uid が見つかりません: ${rootUid}`);
                }
                scopedTaskUids = collectTaskSubtreeUids(phaseTasksOnly, rootUid, maxDepth);
            }
            const descendantTasks = phaseTasksOnly.filter((task) => scopedTaskUids.has(task.uid));
            return {
                view_type: "phase_detail_view",
                project: {
                    name: model.project.name,
                    planned_start: model.project.startDate,
                    planned_finish: model.project.finishDate
                },
                phase: {
                    uid: phase.uid,
                    name: phase.name,
                    wbs: phase.wbs || phase.outlineNumber,
                    planned_start: phase.start,
                    planned_finish: phase.finish,
                    task_count: descendantTasks.length,
                    milestone_count: descendantTasks.filter((task) => task.milestone).length,
                    percent_complete: phase.percentComplete
                },
                scope: {
                    mode,
                    root_uid: rootUid || null,
                    max_depth: maxDepth !== null && maxDepth !== void 0 ? maxDepth : null
                },
                tasks: descendantTasks.map((task) => {
                    var _a;
                    return ({
                        uid: task.uid,
                        name: task.name,
                        parent_uid: parentMap.get(task.uid),
                        position: (_a = positionMap.get(task.uid)) !== null && _a !== void 0 ? _a : 0,
                        is_summary: task.summary,
                        is_milestone: task.milestone,
                        planned_duration: task.duration,
                        planned_duration_hours: parseDurationHours(task.duration),
                        planned_start: task.start,
                        planned_finish: task.finish,
                        percent_complete: task.percentComplete,
                        predecessor_uids: task.predecessors.map((item) => item.predecessorUid)
                    });
                }),
                milestones: descendantTasks.filter((task) => task.milestone).map((task) => ({
                    uid: task.uid,
                    name: task.name,
                    date: task.finish || task.start
                })),
                dependency_summary: descendantTasks.flatMap((task) => task.predecessors
                    .filter((predecessor) => scopedTaskUids.has(predecessor.predecessorUid))
                    .map((predecessor) => {
                    var _a;
                    return ({
                        from_uid: predecessor.predecessorUid,
                        to_uid: task.uid,
                        type: formatDependencyType(predecessor.type),
                        lag: predecessor.linkLag || "PT0H0M0S",
                        lag_hours: (_a = parseDurationHours(predecessor.linkLag || "PT0H0M0S")) !== null && _a !== void 0 ? _a : 0
                    });
                })),
                rules: buildDefaultRules("phase_detail_view")
            };
        },
        exportTaskEditView(model, requestedTaskUid) {
            var _a, _b;
            const tasks = model.tasks.filter((task) => !isPlaceholderUid(task.uid));
            if (tasks.length === 0) {
                throw new Error("task が見つかりません");
            }
            const targetTask = requestedTaskUid ? tasks.find((task) => task.uid === requestedTaskUid) : tasks.find((task) => !task.summary) || tasks[0];
            if (!targetTask) {
                throw new Error(`task が見つかりません: ${requestedTaskUid}`);
            }
            const parentMap = buildTaskParentMap(model.tasks);
            const positionMap = buildTaskPositionMap(model.tasks, parentMap);
            const parentUid = parentMap.get(targetTask.uid) || null;
            const parentTask = parentUid ? tasks.find((task) => task.uid === parentUid) : undefined;
            const siblingTasks = tasks
                .filter((task) => (parentMap.get(task.uid) || null) === parentUid && task.uid !== targetTask.uid)
                .map((task) => {
                var _a;
                return ({
                    uid: task.uid,
                    name: task.name,
                    position: (_a = positionMap.get(task.uid)) !== null && _a !== void 0 ? _a : 0,
                    is_summary: task.summary,
                    is_milestone: task.milestone
                });
            });
            const phaseTasks = collectTopLevelPhases(model.tasks);
            const phaseUidSet = new Set(phaseTasks.map((task) => task.uid));
            const phaseUid = resolvePhaseUidForTask(targetTask.uid, parentMap, phaseUidSet);
            const phaseTask = phaseUid ? tasks.find((task) => task.uid === phaseUid) : undefined;
            const taskByUid = new Map(tasks.map((task) => [task.uid, task]));
            const predecessors = targetTask.predecessors.map((item) => {
                var _a, _b;
                return ({
                    task_uid: item.predecessorUid,
                    name: ((_a = taskByUid.get(item.predecessorUid)) === null || _a === void 0 ? void 0 : _a.name) || item.predecessorUid,
                    type: formatDependencyType(item.type),
                    lag: item.linkLag || "PT0H0M0S",
                    lag_hours: (_b = parseDurationHours(item.linkLag || "PT0H0M0S")) !== null && _b !== void 0 ? _b : 0
                });
            });
            const successors = tasks.flatMap((task) => task.predecessors
                .filter((item) => item.predecessorUid === targetTask.uid)
                .map((item) => {
                var _a;
                return ({
                    task_uid: task.uid,
                    name: task.name,
                    type: formatDependencyType(item.type),
                    lag: item.linkLag || "PT0H0M0S",
                    lag_hours: (_a = parseDurationHours(item.linkLag || "PT0H0M0S")) !== null && _a !== void 0 ? _a : 0
                });
            }));
            const assignments = model.assignments
                .filter((assignment) => assignment.taskUid === targetTask.uid)
                .map((assignment) => {
                const resource = model.resources.find((item) => item.uid === assignment.resourceUid);
                return {
                    uid: assignment.uid,
                    resource_uid: assignment.resourceUid,
                    resource_name: (resource === null || resource === void 0 ? void 0 : resource.name) || assignment.resourceUid,
                    start: assignment.start,
                    finish: assignment.finish,
                    units: assignment.units,
                    work: assignment.work,
                    percent_work_complete: assignment.percentWorkComplete
                };
            });
            return {
                view_type: "task_edit_view",
                project: {
                    name: model.project.name,
                    planned_start: model.project.startDate,
                    planned_finish: model.project.finishDate
                },
                phase: phaseTask ? { uid: phaseTask.uid, name: phaseTask.name } : null,
                target_task: {
                    uid: targetTask.uid,
                    name: targetTask.name,
                    parent_uid: parentUid,
                    position: (_a = positionMap.get(targetTask.uid)) !== null && _a !== void 0 ? _a : 0,
                    is_summary: targetTask.summary,
                    is_milestone: targetTask.milestone,
                    planned_duration: targetTask.duration,
                    planned_duration_hours: parseDurationHours(targetTask.duration),
                    planned_start: targetTask.start,
                    planned_finish: targetTask.finish,
                    percent_complete: targetTask.percentComplete,
                    notes: targetTask.notes,
                    calendar_uid: targetTask.calendarUID || null,
                    critical: targetTask.critical
                },
                parent_task: parentTask ? {
                    uid: parentTask.uid,
                    name: parentTask.name,
                    position: (_b = positionMap.get(parentTask.uid)) !== null && _b !== void 0 ? _b : 0
                } : null,
                sibling_tasks: siblingTasks,
                predecessors,
                successors,
                assignments,
                rules: buildDefaultRules("task_edit_view")
            };
        }
    };
    globalThis.__mikuprojectMsprojectAiViews = aiViews;
})();
