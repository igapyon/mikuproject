/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  type ImportChange = {
    scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
    uid: string;
    label: string;
    field: string;
    before: string | number | boolean | undefined;
    after: string | number | boolean;
  };

  type PatchWarning = {
    message: string;
    scope?: "project" | "tasks" | "resources" | "assignments" | "calendars";
    uid?: string;
    label?: string;
  };

  type PatchOperation = {
    from_uid?: string;
    to_uid?: string;
    type?: string;
    lag?: string;
    lag_hours?: number;
  };

  const projectPatchJsonUtil = (globalThis as typeof globalThis & {
    __mikuprojectProjectPatchJsonUtil?: {
      resolveDependencyType: (
        rawType: unknown,
        warnings: PatchWarning[],
        task: TaskModel,
        operationIndex: number,
        opName: "link_tasks" | "unlink_tasks",
        options?: { allowEmpty?: boolean }
      ) => number | undefined;
      resolveDependencyLag: (
        operation: PatchOperation,
        warnings: PatchWarning[],
        task: TaskModel,
        operationIndex: number,
        opName: "link_tasks" | "unlink_tasks",
        options?: { allowEmpty?: boolean }
      ) => string | undefined;
      normalizeDurationText: (duration: string | undefined) => string | undefined;
      formatRequestedDependencyRelation: (
        fromUid: string,
        toUid: string,
        type: number | undefined,
        linkLag: string | undefined
      ) => string;
      formatPredecessorList: (predecessors: PredecessorModel[]) => string;
    };
  }).__mikuprojectProjectPatchJsonUtil;

  if (!projectPatchJsonUtil) {
    throw new Error("mikuproject Project Patch JSON util module is not loaded");
  }

  function resolvePatchTaskRelation(
    operation: PatchOperation,
    opName: "link_tasks" | "unlink_tasks",
    taskByUid: Map<string, TaskModel>,
    warnings: PatchWarning[],
    operationIndex: number
  ): { fromTask: TaskModel; toTask: TaskModel } | undefined {
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

  function createsDependencyCycle(
    taskByUid: Map<string, TaskModel>,
    fromUid: string,
    toUid: string
  ): boolean {
    const queue = [toUid];
    const visited = new Set<string>();
    while (queue.length > 0) {
      const currentUid = queue.shift() as string;
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

  function applyLinkTasksOperation(
    operation: PatchOperation,
    taskByUid: Map<string, TaskModel>,
    changes: ImportChange[],
    warnings: PatchWarning[],
    operationIndex: number
  ): void {
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
      const nextType = typeCode ?? existing.type;
      const nextLag = linkLag ?? existing.linkLag;
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

  function applyUnlinkTasksOperation(
    operation: PatchOperation,
    taskByUid: Map<string, TaskModel>,
    changes: ImportChange[],
    warnings: PatchWarning[],
    operationIndex: number
  ): void {
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

  (globalThis as typeof globalThis & {
    __mikuprojectProjectPatchJsonLinks?: {
      applyLinkTasksOperation: typeof applyLinkTasksOperation;
      applyUnlinkTasksOperation: typeof applyUnlinkTasksOperation;
    };
  }).__mikuprojectProjectPatchJsonLinks = {
    applyLinkTasksOperation,
    applyUnlinkTasksOperation
  };
})();
