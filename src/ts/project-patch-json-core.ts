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
    initials?: string;
    group?: string;
    calendar_uid?: string;
    max_units?: number;
    is_summary?: boolean;
    new_parent_uid?: string | null;
    new_index?: number;
    is_milestone?: boolean;
    task_uid?: string;
    resource_uid?: string;
    start?: string;
    finish?: string;
    units?: number;
    work?: string;
    percent_work_complete?: number;
    planned_start?: string;
    planned_finish?: string;
    planned_duration?: string;
    planned_duration_hours?: number;
    from_uid?: string;
    to_uid?: string;
    type?: string;
    lag?: string;
    lag_hours?: number;
    fields?: Record<string, unknown>;
  };

  type PatchDocument = {
    operations: PatchOperation[];
  };

  const mikuprojectXml = (globalThis as typeof globalThis & {
    __mikuprojectXml?: {
      normalizeProjectModel: (model: ProjectModel) => ProjectModel;
    };
  }).__mikuprojectXml;

  if (!mikuprojectXml) {
    throw new Error("mikuproject XML module is not loaded");
  }

  const mikuprojectProjectPatchJsonUtil = (globalThis as typeof globalThis & {
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
      formatDependencyType: (type: number | undefined) => string;
      isZeroDuration: (duration: string | undefined) => boolean;
      normalizeDurationText: (duration: string | undefined) => string | undefined;
      isValidDurationText: (duration: string) => boolean;
      formatPredecessor: (predecessor: PredecessorModel) => string;
      formatRequestedDependencyRelation: (
        fromUid: string,
        toUid: string,
        type: number | undefined,
        linkLag: string | undefined
      ) => string;
      formatPredecessorList: (predecessors: PredecessorModel[]) => string;
      normalizePatchedTaskDate: (
        value: unknown,
        kind: "start" | "finish",
        task: TaskModel,
        project: ProjectInfo
      ) => string | undefined;
      normalizePatchedPlainDateTime: (
        value: unknown,
        kind: "start" | "finish",
        project: ProjectInfo
      ) => string | undefined;
      formatDurationHours: (hours: number) => string;
      parseDurationHours: (duration: string | undefined) => number | undefined;
      cloneProjectModel: (model: ProjectModel) => ProjectModel;
    };
  }).__mikuprojectProjectPatchJsonUtil;

  if (!mikuprojectProjectPatchJsonUtil) {
    throw new Error("mikuproject Project Patch JSON util module is not loaded");
  }

  const mikuprojectProjectPatchJsonLinks = (globalThis as typeof globalThis & {
    __mikuprojectProjectPatchJsonLinks?: {
      applyLinkTasksOperation: (
        operation: PatchOperation,
        taskByUid: Map<string, TaskModel>,
        changes: ImportChange[],
        warnings: PatchWarning[],
        operationIndex: number
      ) => void;
      applyUnlinkTasksOperation: (
        operation: PatchOperation,
        taskByUid: Map<string, TaskModel>,
        changes: ImportChange[],
        warnings: PatchWarning[],
        operationIndex: number
      ) => void;
    };
  }).__mikuprojectProjectPatchJsonLinks;

  if (!mikuprojectProjectPatchJsonLinks) {
    throw new Error("mikuproject Project Patch JSON links module is not loaded");
  }

  const mikuprojectProjectPatchJsonTasks = (globalThis as typeof globalThis & {
    __mikuprojectProjectPatchJsonTasks?: {
      applyAddTaskOperation: (
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
      ) => TaskModel | undefined;
      applyMoveTaskOperation: (
        operation: PatchOperation,
        tasks: TaskModel[],
        changes: ImportChange[],
        warnings: PatchWarning[],
        operationIndex: number
      ) => void;
      applyDeleteTaskOperation: (
        operation: PatchOperation,
        model: ProjectModel,
        changes: ImportChange[],
        warnings: PatchWarning[],
        operationIndex: number
      ) => void;
    };
  }).__mikuprojectProjectPatchJsonTasks;

  if (!mikuprojectProjectPatchJsonTasks) {
    throw new Error("mikuproject Project Patch JSON tasks module is not loaded");
  }

  const mikuprojectProjectPatchJsonEntities = (globalThis as typeof globalThis & {
    __mikuprojectProjectPatchJsonEntities?: {
      warnUnsupportedPatchKeys: (
        operation: PatchOperation,
        allowedKeys: Set<string>,
        opName: string,
        uid: string,
        label: string,
        warnings: PatchWarning[]
      ) => void;
      warnUnsupportedPatchKeysForScope: (
        operation: PatchOperation,
        allowedKeys: Set<string>,
        opName: string,
        scope: "project" | "tasks" | "resources" | "assignments" | "calendars",
        uid: string,
        label: string,
        warnings: PatchWarning[]
      ) => void;
      applyAddAssignmentOperation: (
        operation: PatchOperation,
        model: ProjectModel,
        changes: ImportChange[],
        warnings: PatchWarning[],
        operationIndex: number
      ) => AssignmentModel | undefined;
      applyAddResourceOperation: (
        operation: PatchOperation,
        model: ProjectModel,
        changes: ImportChange[],
        warnings: PatchWarning[],
        operationIndex: number
      ) => ResourceModel | undefined;
      applyAddCalendarOperation: (
        operation: PatchOperation,
        model: ProjectModel,
        changes: ImportChange[],
        warnings: PatchWarning[],
        operationIndex: number
      ) => CalendarModel | undefined;
      applyDeleteResourceOperation: (
        operation: PatchOperation,
        model: ProjectModel,
        changes: ImportChange[],
        warnings: PatchWarning[],
        operationIndex: number
      ) => void;
      applyDeleteCalendarOperation: (
        operation: PatchOperation,
        model: ProjectModel,
        changes: ImportChange[],
        warnings: PatchWarning[],
        operationIndex: number
      ) => void;
      applyDeleteAssignmentOperation: (
        operation: PatchOperation,
        model: ProjectModel,
        changes: ImportChange[],
        warnings: PatchWarning[],
        operationIndex: number
      ) => void;
    };
  }).__mikuprojectProjectPatchJsonEntities;

  if (!mikuprojectProjectPatchJsonEntities) {
    throw new Error("mikuproject Project Patch JSON entities module is not loaded");
  }

  const mikuprojectProjectPatchJsonUpdates = (globalThis as typeof globalThis & {
    __mikuprojectProjectPatchJsonUpdates?: {
      applyUpdateTaskOperation: (
        task: TaskModel,
        rawFields: Record<string, unknown>,
        model: ProjectModel,
        changes: ImportChange[],
        warnings: PatchWarning[],
        operationIndex: number
      ) => void;
      applyUpdateProjectOperation: (
        project: ProjectInfo,
        rawFields: Record<string, unknown>,
        model: ProjectModel,
        changes: ImportChange[],
        warnings: PatchWarning[],
        operationIndex: number
      ) => void;
      applyUpdateAssignmentOperation: (
        assignment: AssignmentModel,
        rawFields: Record<string, unknown>,
        project: ProjectInfo,
        changes: ImportChange[],
        warnings: PatchWarning[],
        operationIndex: number
      ) => void;
      applyUpdateResourceOperation: (
        resource: ResourceModel,
        rawFields: Record<string, unknown>,
        model: ProjectModel,
        changes: ImportChange[],
        warnings: PatchWarning[],
        operationIndex: number
      ) => void;
      applyUpdateCalendarOperation: (
        calendar: CalendarModel,
        rawFields: Record<string, unknown>,
        model: ProjectModel,
        changes: ImportChange[],
        warnings: PatchWarning[],
        operationIndex: number
      ) => void;
    };
  }).__mikuprojectProjectPatchJsonUpdates;

  if (!mikuprojectProjectPatchJsonUpdates) {
    throw new Error("mikuproject Project Patch JSON updates module is not loaded");
  }

  function importProjectPatchJson(documentLike: unknown, baseModel: ProjectModel): {
    model: ProjectModel;
    changes: ImportChange[];
    warnings: PatchWarning[];
  } {
    const validation = validatePatchDocument(documentLike);
    const nextModel = mikuprojectProjectPatchJsonUtil.cloneProjectModel(baseModel);
    const changes: ImportChange[] = [];
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

        mikuprojectProjectPatchJsonUpdates.applyUpdateProjectOperation(nextModel.project, operation.fields, nextModel, changes, warnings, index);
        return;
      }

      if (op === "add_task") {
        const addedTask = mikuprojectProjectPatchJsonTasks.applyAddTaskOperation(
          operation,
          nextModel.tasks,
          nextModel.project,
          changes,
          warnings,
          index,
          mikuprojectProjectPatchJsonEntities.warnUnsupportedPatchKeys
        );
        if (addedTask) {
          taskByUid.set(addedTask.uid, addedTask);
        }
        return;
      }

      if (op === "add_resource") {
        const addedResource = mikuprojectProjectPatchJsonEntities.applyAddResourceOperation(operation, nextModel, changes, warnings, index);
        if (addedResource) {
          resourceByUid.set(addedResource.uid, addedResource);
        }
        return;
      }

      if (op === "add_calendar") {
        const addedCalendar = mikuprojectProjectPatchJsonEntities.applyAddCalendarOperation(operation, nextModel, changes, warnings, index);
        if (addedCalendar) {
          calendarByUid.set(addedCalendar.uid, addedCalendar);
        }
        return;
      }

      if (op === "add_assignment") {
        const addedAssignment = mikuprojectProjectPatchJsonEntities.applyAddAssignmentOperation(operation, nextModel, changes, warnings, index);
        if (addedAssignment) {
          assignmentByUid.set(addedAssignment.uid, addedAssignment);
        }
        return;
      }

      if (op === "delete_assignment") {
        mikuprojectProjectPatchJsonEntities.applyDeleteAssignmentOperation(operation, nextModel, changes, warnings, index);
        assignmentByUid.clear();
        nextModel.assignments.forEach((assignment) => assignmentByUid.set(assignment.uid, assignment));
        return;
      }

      if (op === "delete_resource") {
        mikuprojectProjectPatchJsonEntities.applyDeleteResourceOperation(operation, nextModel, changes, warnings, index);
        resourceByUid.clear();
        nextModel.resources.forEach((resource) => resourceByUid.set(resource.uid, resource));
        return;
      }

      if (op === "delete_calendar") {
        mikuprojectProjectPatchJsonEntities.applyDeleteCalendarOperation(operation, nextModel, changes, warnings, index);
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

        mikuprojectProjectPatchJsonUpdates.applyUpdateTaskOperation(task, operation.fields, nextModel, changes, warnings, index);
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

        mikuprojectProjectPatchJsonUpdates.applyUpdateAssignmentOperation(assignment, operation.fields, nextModel.project, changes, warnings, index);
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

        mikuprojectProjectPatchJsonUpdates.applyUpdateResourceOperation(resource, operation.fields, nextModel, changes, warnings, index);
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

        mikuprojectProjectPatchJsonUpdates.applyUpdateCalendarOperation(calendar, operation.fields, nextModel, changes, warnings, index);
        return;
      }

      if (op === "move_task") {
        mikuprojectProjectPatchJsonTasks.applyMoveTaskOperation(operation, nextModel.tasks, changes, warnings, index);
        return;
      }

      if (op === "delete_task") {
        mikuprojectProjectPatchJsonTasks.applyDeleteTaskOperation(operation, nextModel, changes, warnings, index);
        taskByUid.clear();
        nextModel.tasks.forEach((task) => taskByUid.set(task.uid, task));
        return;
      }

      if (op === "link_tasks") {
        mikuprojectProjectPatchJsonLinks.applyLinkTasksOperation(operation, taskByUid, changes, warnings, index);
        return;
      }

      if (op === "unlink_tasks") {
        mikuprojectProjectPatchJsonLinks.applyUnlinkTasksOperation(operation, taskByUid, changes, warnings, index);
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

  function validatePatchDocument(documentLike: unknown): {
    document: PatchDocument;
    warnings: PatchWarning[];
  } {
    if (!documentLike || typeof documentLike !== "object" || Array.isArray(documentLike)) {
      throw new Error("Patch JSON がオブジェクトではありません");
    }
    const document = documentLike as Partial<PatchDocument>;
    if (!Array.isArray(document.operations)) {
      throw new Error("Patch JSON には operations 配列が必要です");
    }
    const warnings: PatchWarning[] = [];
    document.operations.forEach((operation, index) => {
      if (!operation || typeof operation !== "object" || Array.isArray(operation)) {
        throw new Error(`operations[${index}] がオブジェクトではありません`);
      }
    });
    return {
      document: document as PatchDocument,
      warnings
    };
  }

  function buildUnsupportedOperationWarningMessage(op: string, index: number): string {
    const normalizedOp = op || "(empty)";
    if (normalizedOp === "link_tasks" || normalizedOp === "unlink_tasks") {
      return `未対応の op は無視します: operations[${index}].op = ${normalizedOp}。依存関係は ${normalizedOp} で扱う方針ですが、現時点では未実装です`;
    }
    if (normalizedOp === "move_task") {
      return `未対応の op は無視します: operations[${index}].op = ${normalizedOp}。親子や順序変更は move_task で扱う方針ですが、現時点では未実装です`;
    }
    return `未対応の op は無視します: operations[${index}].op = ${normalizedOp}`;
  }

  (globalThis as typeof globalThis & {
    __mikuprojectProjectPatchJsonCore?: {
      importProjectPatchJson: (documentLike: unknown, baseModel: ProjectModel) => {
        model: ProjectModel;
        changes: ImportChange[];
        warnings: PatchWarning[];
      };
      validatePatchDocument: (documentLike: unknown) => {
        document: PatchDocument;
        warnings: PatchWarning[];
      };
    };
  }).__mikuprojectProjectPatchJsonCore = {
    importProjectPatchJson,
    validatePatchDocument
  };
})();
