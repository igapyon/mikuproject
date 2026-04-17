/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const projectPatchJsonUtil = globalThis.__mikuprojectProjectPatchJsonUtil;
    if (!projectPatchJsonUtil) {
        throw new Error("mikuproject Project Patch JSON util module is not loaded");
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
    function applyLinkTasksOperation(operation, taskByUid, changes, warnings, operationIndex) {
        const relation = resolvePatchTaskRelation(operation, "link_tasks", taskByUid, warnings, operationIndex);
        if (!relation) {
            return;
        }
        const { fromTask, toTask } = relation;
        const typeCode = projectPatchJsonUtil.resolveDependencyType(operation.type, warnings, toTask, operationIndex, "link_tasks");
        const linkLag = projectPatchJsonUtil.resolveDependencyLag(operation, warnings, toTask, operationIndex, "link_tasks");
        const existing = toTask.predecessors.find((item) => item.predecessorUid === fromTask.uid);
        if (existing) {
            const beforeText = projectPatchJsonUtil.formatPredecessorList(toTask.predecessors);
            const nextType = typeCode !== null && typeCode !== void 0 ? typeCode : existing.type;
            const nextLag = linkLag !== null && linkLag !== void 0 ? linkLag : existing.linkLag;
            const relationText = projectPatchJsonUtil.formatRequestedDependencyRelation(fromTask.uid, toTask.uid, nextType, nextLag);
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
                after: projectPatchJsonUtil.formatPredecessorList(toTask.predecessors)
            });
            return;
        }
        const beforeText = projectPatchJsonUtil.formatPredecessorList(toTask.predecessors);
        const relationText = projectPatchJsonUtil.formatRequestedDependencyRelation(fromTask.uid, toTask.uid, typeCode, linkLag);
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
            after: projectPatchJsonUtil.formatPredecessorList(toTask.predecessors)
        });
    }
    function applyUnlinkTasksOperation(operation, taskByUid, changes, warnings, operationIndex) {
        const relation = resolvePatchTaskRelation(operation, "unlink_tasks", taskByUid, warnings, operationIndex);
        if (!relation) {
            return;
        }
        const { fromTask, toTask } = relation;
        const requestedType = projectPatchJsonUtil.resolveDependencyType(operation.type, warnings, toTask, operationIndex, "unlink_tasks", { allowEmpty: true });
        const requestedLag = projectPatchJsonUtil.resolveDependencyLag(operation, warnings, toTask, operationIndex, "unlink_tasks", { allowEmpty: true });
        const beforeText = projectPatchJsonUtil.formatPredecessorList(toTask.predecessors);
        const matched = toTask.predecessors.filter((item) => {
            if (item.predecessorUid !== fromTask.uid) {
                return false;
            }
            if (requestedType !== undefined && item.type !== requestedType) {
                return false;
            }
            if (requestedLag !== undefined && projectPatchJsonUtil.normalizeDurationText(item.linkLag) !== projectPatchJsonUtil.normalizeDurationText(requestedLag)) {
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
            if (requestedLag !== undefined && projectPatchJsonUtil.normalizeDurationText(item.linkLag) !== projectPatchJsonUtil.normalizeDurationText(requestedLag)) {
                return true;
            }
            return false;
        });
        if (filtered.length === toTask.predecessors.length) {
            warnings.push({
                message: `unlink_tasks の対象依存関係が見つかりません: ${projectPatchJsonUtil.formatRequestedDependencyRelation(fromTask.uid, toTask.uid, requestedType, requestedLag)}`,
                scope: "tasks",
                uid: toTask.uid,
                label: toTask.name || toTask.uid
            });
            return;
        }
        toTask.predecessors = filtered;
        if (matched.length > 1) {
            warnings.push({
                message: `unlink_tasks は一致した依存関係 ${matched.length} 件をすべて解除しました: ${projectPatchJsonUtil.formatRequestedDependencyRelation(fromTask.uid, toTask.uid, requestedType, requestedLag)}`,
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
            after: projectPatchJsonUtil.formatPredecessorList(toTask.predecessors)
        });
    }
    globalThis.__mikuprojectProjectPatchJsonLinks = {
        applyLinkTasksOperation,
        applyUnlinkTasksOperation
    };
})();
