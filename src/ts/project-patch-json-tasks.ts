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
    op?: string;
    uid?: string;
    name?: string;
    new_parent_uid?: string | null;
    new_index?: number;
    is_summary?: boolean;
    is_milestone?: boolean;
    planned_start?: string;
    planned_finish?: string;
    planned_duration?: string;
    planned_duration_hours?: number;
  };

  const projectPatchJsonUtil = (globalThis as typeof globalThis & {
    __mikuprojectProjectPatchJsonUtil?: {
      isZeroDuration: (duration: string | undefined) => boolean;
      normalizePatchedTaskDate: (
        value: unknown,
        kind: "start" | "finish",
        task: TaskModel,
        project: ProjectInfo
      ) => string | undefined;
      formatDurationHours: (hours: number) => string;
    };
  }).__mikuprojectProjectPatchJsonUtil;

  if (!projectPatchJsonUtil) {
    throw new Error("mikuproject Project Patch JSON util module is not loaded");
  }

  function collectTaskSubtreeRange(tasks: TaskModel[], startIndex: number): { start: number; end: number } {
    const rootTask = tasks[startIndex];
    let end = startIndex + 1;
    while (end < tasks.length && tasks[end].outlineLevel > rootTask.outlineLevel) {
      end += 1;
    }
    return { start: startIndex, end };
  }

  function buildTaskParentMap(tasks: TaskModel[]): Map<string, string | null> {
    const parentMap = new Map<string, string | null>();
    const stack: TaskModel[] = [];
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

  function buildTaskPositionMap(tasks: TaskModel[], parentMap: Map<string, string | null>): Map<string, number> {
    const counters = new Map<string, number>();
    const positionMap = new Map<string, number>();
    for (const task of tasks) {
      const parentUid = parentMap.get(task.uid) || "__root__";
      const position = counters.get(parentUid) || 0;
      positionMap.set(task.uid, position);
      counters.set(parentUid, position + 1);
    }
    return positionMap;
  }

  function resolveMoveInsertionIndex(
    remainingTasks: TaskModel[],
    siblingStartIndices: number[],
    nextIndex: number,
    nextParentUid: string | null
  ): number {
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

  function resolveAddInsertionIndex(
    tasks: TaskModel[],
    siblingStartIndices: number[],
    nextIndex: number,
    nextParentUid: string | null
  ): number {
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

  function isNoOpMove(
    originalTasks: TaskModel[],
    remainingTasks: TaskModel[],
    subtree: TaskModel[],
    insertionIndex: number,
    levelDelta: number
  ): boolean {
    const currentSignature = originalTasks.map((task) => `${task.uid}@${task.outlineLevel}`).join("|");
    const candidateTasks = remainingTasks.map((task) => ({ uid: task.uid, outlineLevel: task.outlineLevel })).slice();
    const subtreeSignature = subtree.map((task) => ({
      uid: task.uid,
      outlineLevel: task.outlineLevel + levelDelta
    }));
    candidateTasks.splice(insertionIndex, 0, ...subtreeSignature);
    const nextSignature = candidateTasks.map((task) => `${task.uid}@${task.outlineLevel}`).join("|");
    return currentSignature === nextSignature;
  }

  function rebuildTaskHierarchyMetadata(tasks: TaskModel[]): void {
    const counters: number[] = [];
    tasks.forEach((task, index) => {
      counters.length = Math.max(1, task.outlineLevel);
      counters[task.outlineLevel - 1] = (counters[task.outlineLevel - 1] || 0) + 1;
      task.outlineNumber = counters.slice(0, task.outlineLevel).join(".");
      task.id = String(index + 1);
    });
  }

  function normalizeOptionalPatchedTaskDate(
    rawValue: unknown,
    kind: "start" | "finish",
    task: Pick<TaskModel, "uid" | "name">,
    project: ProjectInfo
  ): string | undefined {
    if (rawValue === undefined) {
      return undefined;
    }
    return projectPatchJsonUtil.normalizePatchedTaskDate(rawValue, kind, task as TaskModel, project);
  }

  function resolveAddedTaskDuration(
    operation: PatchOperation,
    uid: string,
    name: string,
    warnings: PatchWarning[]
  ): string | undefined {
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
      return projectPatchJsonUtil.formatDurationHours(operation.planned_duration_hours);
    }
    return undefined;
  }

  function applyAddTaskOperation(
    operation: PatchOperation,
    tasks: TaskModel[],
    project: ProjectInfo,
    changes: ImportChange[],
    warnings: PatchWarning[],
    operationIndex: number,
    warnUnsupportedPatchKeys: (
      operation: PatchOperation,
      allowedKeys: Set<string>,
      opName: string,
      uid: string,
      label: string,
      warnings: PatchWarning[]
    ) => void
  ): TaskModel | undefined {
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
    warnUnsupportedPatchKeys(
      operation,
      new Set(["op", "uid", "name", "is_summary", "new_parent_uid", "new_index", "is_milestone", "planned_start", "planned_finish", "planned_duration", "planned_duration_hours"]),
      "add_task",
      uid,
      name,
      warnings
    );
    if (operation.is_summary !== undefined && typeof operation.is_summary !== "boolean") {
      warnings.push({ message: `add_task.is_summary は boolean が必要です: ${uid}`, scope: "tasks", uid, label: name });
      return undefined;
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
    let nextParent: TaskModel | undefined;
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
    const normalizedStartCandidate = normalizeOptionalPatchedTaskDate(operation.planned_start, "start", { uid, name } as TaskModel, project);
    if (operation.planned_start !== undefined && normalizedStartCandidate === undefined) {
      warnings.push({ message: `add_task.planned_start の日付形式が解釈できません: ${uid}`, scope: "tasks", uid, label: name });
      return undefined;
    }
    const normalizedFinishCandidate = normalizeOptionalPatchedTaskDate(operation.planned_finish, "finish", { uid, name } as TaskModel, project);
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
      if (!projectPatchJsonUtil.isZeroDuration(duration)) {
        warnings.push({ message: `add_task.is_milestone=true のため planned_duration は 0 に揃えます: ${uid}`, scope: "tasks", uid, label: name });
        duration = "PT0H0M0S";
      }
    }
    const insertionIndex = resolveAddInsertionIndex(tasks, siblingStartIndices, nextIndex, nextParentUid);
    const addedTask: TaskModel = {
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
    changes.push({ scope: "tasks", uid, label: name, field: "name", before: undefined, after: name });
    changes.push({ scope: "tasks", uid, label: name, field: "parent_uid", before: undefined, after: nextParentUid || "(root)" });
    changes.push({ scope: "tasks", uid, label: name, field: "position", before: undefined, after: nextPositionMap.get(uid) ?? nextIndex });
    if (operation.planned_start !== undefined) {
      changes.push({ scope: "tasks", uid, label: name, field: "planned_start", before: undefined, after: normalizedStart });
    }
    if (operation.planned_finish !== undefined) {
      changes.push({ scope: "tasks", uid, label: name, field: "planned_finish", before: undefined, after: normalizedFinish });
    }
    return addedTask;
  }

  function applyMoveTaskOperation(
    operation: PatchOperation,
    tasks: TaskModel[],
    changes: ImportChange[],
    warnings: PatchWarning[],
    operationIndex: number
  ): void {
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
    const nextParentUid = operation.new_parent_uid == null || operation.new_parent_uid === "" ? null : String(operation.new_parent_uid).trim();
    if (typeof operation.new_index !== "number" || !Number.isInteger(operation.new_index) || operation.new_index < 0) {
      warnings.push({ message: `move_task.new_index は 0 以上の整数が必要です: ${uid}`, scope: "tasks", uid, label: task.name || task.uid });
      return;
    }
    const nextIndex = operation.new_index;
    const currentRange = collectTaskSubtreeRange(tasks, taskIndex);
    const subtree = tasks.slice(currentRange.start, currentRange.end);
    const subtreeUidSet = new Set(subtree.map((item) => item.uid));
    if (nextParentUid && subtreeUidSet.has(nextParentUid)) {
      warnings.push({ message: `move_task では task を自身または配下へ移動できません: ${uid} -> ${nextParentUid}`, scope: "tasks", uid, label: task.name || task.uid });
      return;
    }
    let nextParent: TaskModel | undefined;
    if (nextParentUid) {
      nextParent = tasks.find((item) => item.uid === nextParentUid);
      if (!nextParent) {
        warnings.push({ message: `move_task.new_parent_uid が既存 task を指していません: ${nextParentUid}`, scope: "tasks", uid, label: task.name || task.uid });
        return;
      }
      if (!nextParent.summary) {
        warnings.push({ message: `move_task.new_parent_uid は summary task を指す必要があります: ${nextParentUid}`, scope: "tasks", uid, label: task.name || task.uid });
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
      warnings.push({ message: `move_task.new_index が sibling 範囲外です: ${uid} -> ${nextIndex}`, scope: "tasks", uid, label: task.name || task.uid });
      return;
    }

    const insertionIndex = resolveMoveInsertionIndex(remainingTasks, siblingStartIndices, nextIndex, nextParentUid);
    const nextOutlineLevel = nextParent ? nextParent.outlineLevel + 1 : 1;
    const levelDelta = nextOutlineLevel - task.outlineLevel;
    if (isNoOpMove(tasks, remainingTasks, subtree, insertionIndex, levelDelta)) {
      warnings.push({ message: `move_task は結果が変わらないため無視します: ${uid} -> parent=${nextParentUid || "(root)"} index=${nextIndex}`, scope: "tasks", uid, label: task.name || task.uid });
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
    const nextPosition = nextPositionMap.get(uid) ?? nextIndex;
    const currentPosition = currentPositionMap.get(uid) ?? 0;
    if (currentParentText !== nextParentText) {
      changes.push({ scope: "tasks", uid, label: task.name || task.uid, field: "parent_uid", before: currentParentText, after: nextParentText });
    }
    if (currentPosition !== nextPosition) {
      changes.push({ scope: "tasks", uid, label: task.name || task.uid, field: "position", before: currentPosition, after: nextPosition });
    }
  }

  function applyDeleteTaskOperation(
    operation: PatchOperation,
    model: ProjectModel,
    changes: ImportChange[],
    warnings: PatchWarning[],
    operationIndex: number
  ): void {
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
      warnings.push({ message: `delete_task first cut では summary task や子を持つ task は削除できません: ${uid} (children=${childCount})`, scope: "tasks", uid, label: task.name || task.uid });
      return;
    }
    const referencingAssignments = model.assignments.filter((assignment) => assignment.taskUid === uid);
    if (referencingAssignments.length > 0) {
      const assignmentText = referencingAssignments.map((assignment) => assignment.uid).join(", ");
      warnings.push({ message: `delete_task first cut では assignment がある task は削除できません: ${uid} (assignments=${assignmentText})`, scope: "tasks", uid, label: task.name || task.uid });
      return;
    }
    const successorTasks = model.tasks.filter((item) => item.predecessors.some((predecessor) => predecessor.predecessorUid === uid));
    if (successorTasks.length > 0) {
      const successorText = successorTasks.map((item) => item.uid).join(", ");
      warnings.push({ message: `delete_task first cut では後続依存がある task は削除できません: ${uid} (successors=${successorText})`, scope: "tasks", uid, label: task.name || task.uid });
      return;
    }

    const parentMap = buildTaskParentMap(model.tasks);
    const positionMap = buildTaskPositionMap(model.tasks, parentMap);
    const parentUid = parentMap.get(uid) || "(root)";
    const position = positionMap.get(uid) ?? 0;
    model.tasks.splice(taskIndex, 1);
    rebuildTaskHierarchyMetadata(model.tasks);
    changes.push({ scope: "tasks", uid, label: task.name || task.uid, field: "name", before: task.name || task.uid, after: "(deleted)" });
    changes.push({ scope: "tasks", uid, label: task.name || task.uid, field: "parent_uid", before: parentUid, after: "(deleted)" });
    changes.push({ scope: "tasks", uid, label: task.name || task.uid, field: "position", before: position, after: "(deleted)" });
  }

  (globalThis as typeof globalThis & {
    __mikuprojectProjectPatchJsonTasks?: {
      applyAddTaskOperation: typeof applyAddTaskOperation;
      applyMoveTaskOperation: typeof applyMoveTaskOperation;
      applyDeleteTaskOperation: typeof applyDeleteTaskOperation;
    };
  }).__mikuprojectProjectPatchJsonTasks = {
    applyAddTaskOperation,
    applyMoveTaskOperation,
    applyDeleteTaskOperation
  };
})();
