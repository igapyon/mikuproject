/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    function buildTaskParentUidMap(tasks) {
        const parentMap = new Map();
        const stack = [];
        for (const task of tasks) {
            while (stack.length > 0 && task.outlineLevel <= stack[stack.length - 1].outlineLevel) {
                stack.pop();
            }
            if (stack.length > 0) {
                parentMap.set(task.uid, stack[stack.length - 1].uid);
            }
            stack.push(task);
        }
        return parentMap;
    }
    function escapeCsvCell(value) {
        const text = String(value !== null && value !== void 0 ? value : "");
        if (/[",\n]/.test(text)) {
            return `"${text.replace(/"/g, "\"\"")}"`;
        }
        return text;
    }
    function parseCsvRows(csvText) {
        const rows = [];
        let row = [];
        let cell = "";
        let inQuotes = false;
        for (let index = 0; index < csvText.length; index += 1) {
            const char = csvText[index];
            const next = csvText[index + 1];
            if (char === "\"") {
                if (inQuotes && next === "\"") {
                    cell += "\"";
                    index += 1;
                }
                else {
                    inQuotes = !inQuotes;
                }
                continue;
            }
            if (!inQuotes && char === ",") {
                row.push(cell);
                cell = "";
                continue;
            }
            if (!inQuotes && (char === "\n" || char === "\r")) {
                if (char === "\r" && next === "\n") {
                    index += 1;
                }
                row.push(cell);
                rows.push(row);
                row = [];
                cell = "";
                continue;
            }
            cell += char;
        }
        if (cell.length > 0 || row.length > 0) {
            row.push(cell);
            rows.push(row);
        }
        return rows.filter((item) => item.some((cellValue) => String(cellValue).trim() !== ""));
    }
    function parseCsvMultiValueCell(value) {
        const normalized = String(value || "").trim();
        if (!normalized) {
            return [];
        }
        const items = normalized
            .split(/[|;,、]/)
            .map((item) => item.trim())
            .filter(Boolean);
        return Array.from(new Set(items));
    }
    function parseCsvBooleanCell(value, fallback) {
        const normalized = String(value || "").trim().toLowerCase();
        if (!normalized) {
            return fallback;
        }
        if (["1", "true", "yes", "y", "on"].includes(normalized)) {
            return true;
        }
        if (["0", "false", "no", "n", "off"].includes(normalized)) {
            return false;
        }
        return fallback;
    }
    globalThis.__mikuprojectMsprojectCsv = {
        exportCsvParentId(model) {
            const header = ["ID", "ParentID", "WBS", "Name", "Start", "Finish", "PredecessorID", "Resource", "PercentComplete", "PercentWorkComplete", "Milestone", "Summary", "Critical", "Type", "Priority", "Work", "CalendarUID", "ConstraintType", "ConstraintDate", "Deadline", "Notes"];
            const parentMap = buildTaskParentUidMap(model.tasks);
            const resourceMap = new Map(model.resources.map((resource) => [resource.uid, resource.name]));
            const assignmentMap = new Map();
            for (const assignment of model.assignments) {
                const resourceName = resourceMap.get(assignment.resourceUid);
                if (!resourceName) {
                    continue;
                }
                const names = assignmentMap.get(assignment.taskUid) || [];
                if (!names.includes(resourceName)) {
                    names.push(resourceName);
                }
                assignmentMap.set(assignment.taskUid, names);
            }
            const rows = model.tasks.map((task) => {
                var _a, _b, _c, _d;
                return [
                    task.uid,
                    parentMap.get(task.uid) || "",
                    task.wbs || task.outlineNumber || "",
                    task.name,
                    task.start || "",
                    task.finish || "",
                    task.predecessors.map((item) => item.predecessorUid).join("|"),
                    (assignmentMap.get(task.uid) || []).join("|"),
                    task.percentComplete,
                    (_a = task.percentWorkComplete) !== null && _a !== void 0 ? _a : "",
                    task.milestone ? 1 : 0,
                    task.summary ? 1 : 0,
                    task.critical === undefined ? "" : (task.critical ? 1 : 0),
                    (_b = task.type) !== null && _b !== void 0 ? _b : "",
                    (_c = task.priority) !== null && _c !== void 0 ? _c : "",
                    task.work || "",
                    task.calendarUID || "",
                    (_d = task.constraintType) !== null && _d !== void 0 ? _d : "",
                    task.constraintDate || "",
                    task.deadline || "",
                    task.notes || ""
                ];
            });
            return [header, ...rows].map((row) => row.map((cell) => escapeCsvCell(cell)).join(",")).join("\n") + "\n";
        },
        importCsvParentId(input) {
            const rows = parseCsvRows(input.csvText.trim());
            if (rows.length === 0) {
                throw new Error("CSV が空です");
            }
            const header = rows[0].map((item) => item.trim());
            const requiredColumns = ["ID", "ParentID", "Name"];
            for (const requiredColumn of requiredColumns) {
                if (!header.includes(requiredColumn)) {
                    throw new Error(`CSV に必須列がありません: ${requiredColumn}`);
                }
            }
            const columnIndex = (name) => header.indexOf(name);
            const entries = rows.slice(1).map((row) => ({
                id: String(row[columnIndex("ID")] || "").trim(),
                parentId: String(row[columnIndex("ParentID")] || "").trim(),
                wbs: String((columnIndex("WBS") >= 0 ? row[columnIndex("WBS")] : "") || "").trim(),
                name: String(row[columnIndex("Name")] || "").trim(),
                start: String((columnIndex("Start") >= 0 ? row[columnIndex("Start")] : "") || "").trim(),
                finish: String((columnIndex("Finish") >= 0 ? row[columnIndex("Finish")] : "") || "").trim(),
                predecessorId: String((columnIndex("PredecessorID") >= 0 ? row[columnIndex("PredecessorID")] : "") || "").trim(),
                resource: String((columnIndex("Resource") >= 0 ? row[columnIndex("Resource")] : "") || "").trim(),
                percentComplete: input.parseNumber(String((columnIndex("PercentComplete") >= 0 ? row[columnIndex("PercentComplete")] : "0") || "0").trim(), 0),
                percentWorkComplete: columnIndex("PercentWorkComplete") >= 0 && String(row[columnIndex("PercentWorkComplete")] || "").trim()
                    ? input.parseNumber(String(row[columnIndex("PercentWorkComplete")] || "").trim(), 0)
                    : undefined,
                milestone: parseCsvBooleanCell(String((columnIndex("Milestone") >= 0 ? row[columnIndex("Milestone")] : "") || "").trim(), false),
                summary: columnIndex("Summary") >= 0 && String(row[columnIndex("Summary")] || "").trim()
                    ? parseCsvBooleanCell(String(row[columnIndex("Summary")] || "").trim(), false)
                    : undefined,
                critical: columnIndex("Critical") >= 0 && String(row[columnIndex("Critical")] || "").trim()
                    ? parseCsvBooleanCell(String(row[columnIndex("Critical")] || "").trim(), false)
                    : undefined,
                type: columnIndex("Type") >= 0 && String(row[columnIndex("Type")] || "").trim()
                    ? input.parseNumber(String(row[columnIndex("Type")] || "").trim(), 0)
                    : undefined,
                priority: columnIndex("Priority") >= 0 && String(row[columnIndex("Priority")] || "").trim()
                    ? input.parseNumber(String(row[columnIndex("Priority")] || "").trim(), 0)
                    : undefined,
                work: String((columnIndex("Work") >= 0 ? row[columnIndex("Work")] : "") || "").trim(),
                calendarUID: String((columnIndex("CalendarUID") >= 0 ? row[columnIndex("CalendarUID")] : "") || "").trim(),
                constraintType: columnIndex("ConstraintType") >= 0 && String(row[columnIndex("ConstraintType")] || "").trim()
                    ? input.parseNumber(String(row[columnIndex("ConstraintType")] || "").trim(), 0)
                    : undefined,
                constraintDate: String((columnIndex("ConstraintDate") >= 0 ? row[columnIndex("ConstraintDate")] : "") || "").trim(),
                deadline: String((columnIndex("Deadline") >= 0 ? row[columnIndex("Deadline")] : "") || "").trim(),
                notes: String((columnIndex("Notes") >= 0 ? row[columnIndex("Notes")] : "") || "").trim(),
                children: []
            })).filter((entry) => entry.id);
            const seenIds = new Set();
            for (const entry of entries) {
                if (seenIds.has(entry.id)) {
                    throw new Error(`CSV の ID が重複しています: ${entry.id}`);
                }
                seenIds.add(entry.id);
                if (!entry.name) {
                    throw new Error(`CSV の Name が空です: ID=${entry.id}`);
                }
                if (entry.parentId && entry.parentId === entry.id) {
                    throw new Error(`CSV の ParentID が自身を指しています: ID=${entry.id}`);
                }
            }
            const entryMap = new Map(entries.map((entry) => [entry.id, entry]));
            for (const entry of entries) {
                if (entry.parentId && !entryMap.has(entry.parentId)) {
                    throw new Error(`CSV の ParentID が既存 ID を指していません: ID=${entry.id}, ParentID=${entry.parentId}`);
                }
            }
            const visiting = new Set();
            const visited = new Set();
            function checkCycle(entry) {
                if (visited.has(entry.id)) {
                    return;
                }
                if (visiting.has(entry.id)) {
                    throw new Error(`CSV の ParentID が循環しています: ID=${entry.id}`);
                }
                visiting.add(entry.id);
                if (entry.parentId) {
                    const parent = entryMap.get(entry.parentId);
                    if (parent) {
                        checkCycle(parent);
                    }
                }
                visiting.delete(entry.id);
                visited.add(entry.id);
            }
            entries.forEach((entry) => checkCycle(entry));
            const roots = [];
            for (const entry of entries) {
                const parent = entry.parentId ? entryMap.get(entry.parentId) : undefined;
                if (parent) {
                    parent.children.push(entry);
                }
                else {
                    roots.push(entry);
                }
            }
            const tasks = [];
            function walk(entry, outlinePath) {
                var _a;
                const children = entry.children;
                let start = entry.start;
                let finish = entry.finish;
                if ((!start || !finish) && children.length > 0) {
                    const childStarts = children.map((child) => child.start).filter(Boolean).sort();
                    const childFinishes = children.map((child) => child.finish).filter(Boolean).sort();
                    start = start || childStarts[0] || "";
                    finish = finish || childFinishes.at(-1) || "";
                }
                const outlineNumber = outlinePath.join(".");
                tasks.push({
                    uid: entry.id,
                    id: entry.id,
                    name: entry.name,
                    outlineLevel: outlinePath.length,
                    outlineNumber,
                    wbs: entry.wbs || outlineNumber,
                    type: entry.type,
                    priority: entry.priority,
                    work: entry.work || undefined,
                    calendarUID: entry.calendarUID || undefined,
                    start,
                    finish,
                    duration: "PT0H0M0S",
                    milestone: entry.milestone || Boolean(start && finish && start === finish),
                    summary: (_a = entry.summary) !== null && _a !== void 0 ? _a : (children.length > 0),
                    critical: entry.critical,
                    percentComplete: Math.max(0, Math.min(100, entry.percentComplete)),
                    percentWorkComplete: entry.percentWorkComplete !== undefined ? Math.max(0, Math.min(100, entry.percentWorkComplete)) : undefined,
                    constraintType: entry.constraintType,
                    constraintDate: entry.constraintDate || undefined,
                    deadline: entry.deadline || undefined,
                    notes: entry.notes || undefined,
                    predecessors: parseCsvMultiValueCell(entry.predecessorId).map((item) => ({ predecessorUid: item })),
                    extendedAttributes: [],
                    baselines: [],
                    timephasedData: []
                });
                children.forEach((child, index) => walk(child, [...outlinePath, index + 1]));
            }
            roots.forEach((root, index) => walk(root, [index + 1]));
            const resourceNames = Array.from(new Set(entries.flatMap((entry) => parseCsvMultiValueCell(entry.resource))));
            const resources = resourceNames.map((name, index) => ({
                uid: String(index + 1),
                id: String(index + 1),
                name,
                extendedAttributes: [],
                baselines: [],
                timephasedData: []
            }));
            const resourceUidByName = new Map(resources.map((resource) => [resource.name, resource.uid]));
            let assignmentUid = 1;
            const taskByUid = new Map(tasks.map((task) => [task.uid, task]));
            const assignments = entries.flatMap((entry) => {
                const task = taskByUid.get(entry.id);
                if (!task) {
                    return [];
                }
                return parseCsvMultiValueCell(entry.resource).map((name) => ({
                    uid: String(assignmentUid++),
                    taskUid: entry.id,
                    resourceUid: resourceUidByName.get(name) || "",
                    start: task.start || undefined,
                    finish: task.finish || undefined,
                    percentWorkComplete: task.percentComplete,
                    extendedAttributes: [],
                    baselines: [],
                    timephasedData: []
                }));
            });
            const taskStarts = tasks.map((task) => task.start).filter(Boolean).sort();
            const taskFinishes = tasks.map((task) => task.finish).filter(Boolean).sort();
            return input.normalizeProjectModel(input.ensureDefaultProjectCalendar({
                project: {
                    name: "CSV Imported Project",
                    title: "CSV Imported Project",
                    startDate: taskStarts[0] || "",
                    finishDate: taskFinishes.at(-1) || "",
                    scheduleFromStart: true,
                    outlineCodes: [],
                    wbsMasks: [],
                    extendedAttributes: []
                },
                tasks,
                resources,
                assignments,
                calendars: []
            }));
        }
    };
})();
