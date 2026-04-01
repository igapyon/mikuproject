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
  };

  type PatchOperation = {
    op?: string;
    uid?: string;
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

  function importProjectPatchJson(documentLike: unknown, baseModel: ProjectModel): {
    model: ProjectModel;
    changes: ImportChange[];
    warnings: PatchWarning[];
  } {
    const validation = validatePatchDocument(documentLike);
    const nextModel = cloneProjectModel(baseModel);
    const changes: ImportChange[] = [];
    const warnings = [...validation.warnings];
    const taskByUid = new Map(nextModel.tasks.map((task) => [task.uid, task]));

    validation.document.operations.forEach((operation, index) => {
      const op = String(operation.op || "").trim();
      if (op !== "update_task") {
        warnings.push({ message: `未対応の op は無視します: operations[${index}].op = ${op || "(empty)"}` });
        return;
      }

      const uid = String(operation.uid || "").trim();
      if (!uid) {
        warnings.push({ message: `update_task の uid がありません: operations[${index}]` });
        return;
      }
      const task = taskByUid.get(uid);
      if (!task) {
        warnings.push({ message: `update_task の uid が既存 task を指していません: ${uid}` });
        return;
      }
      if (!operation.fields || typeof operation.fields !== "object" || Array.isArray(operation.fields)) {
        warnings.push({ message: `update_task.fields がオブジェクトではありません: ${uid}` });
        return;
      }

      applyUpdateTaskOperation(task, operation.fields, nextModel.project, changes, warnings, index);
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

  function applyUpdateTaskOperation(
    task: TaskModel,
    rawFields: Record<string, unknown>,
    project: ProjectInfo,
    changes: ImportChange[],
    warnings: PatchWarning[],
    operationIndex: number
  ): void {
    const fields = { ...rawFields };
    if (Object.keys(fields).length === 0) {
      warnings.push({ message: `update_task に fields がありません: ${task.uid}` });
      return;
    }
    if (fields.planned_duration !== undefined && fields.planned_duration_hours !== undefined) {
      warnings.push({ message: `planned_duration と planned_duration_hours が同時指定されたため、planned_duration_hours は無視します: ${task.uid}` });
      delete fields.planned_duration_hours;
    }

    const handledFields = new Set<string>();

    if (fields.name !== undefined) {
      handledFields.add("name");
      if (typeof fields.name !== "string" || fields.name.trim() === "") {
        warnings.push({ message: `update_task.name は空でない文字列が必要です: ${task.uid}` });
      } else if (task.name !== fields.name.trim()) {
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

    if (fields.planned_start !== undefined) {
      handledFields.add("planned_start");
      const normalizedStart = normalizePatchedTaskDate(fields.planned_start, "start", task, project);
      if (normalizedStart === undefined) {
        warnings.push({ message: `update_task.planned_start の日付形式が解釈できません: ${task.uid}` });
      } else if (task.start !== normalizedStart) {
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
      const normalizedFinish = normalizePatchedTaskDate(fields.planned_finish, "finish", task, project);
      if (normalizedFinish === undefined) {
        warnings.push({ message: `update_task.planned_finish の日付形式が解釈できません: ${task.uid}` });
      } else if (task.finish !== normalizedFinish) {
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
        warnings.push({ message: `update_task.planned_duration は空でない文字列が必要です: ${task.uid}` });
      } else if (task.duration !== fields.planned_duration.trim()) {
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
        warnings.push({ message: `update_task.planned_duration_hours は 0 以上の数値が必要です: ${task.uid}` });
      } else {
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

    Object.keys(fields).forEach((fieldName) => {
      if (handledFields.has(fieldName)) {
        return;
      }
      warnings.push({ message: `未対応の field は無視します: operations[${operationIndex}].fields.${fieldName}` });
    });
  }

  function normalizePatchedTaskDate(
    value: unknown,
    kind: "start" | "finish",
    task: TaskModel,
    project: ProjectInfo
  ): string | undefined {
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

  function isDateOnlyText(value: string): boolean {
    return /^\d{4}-\d{2}-\d{2}$/.test(value);
  }

  function isDateText(value: string): boolean {
    if (isDateOnlyText(value)) {
      return true;
    }
    return !Number.isNaN(new Date(value).getTime());
  }

  function formatDurationHours(hours: number): string {
    const totalSeconds = Math.round(hours * 60 * 60);
    const normalizedSeconds = Math.max(0, totalSeconds);
    const durationHours = Math.floor(normalizedSeconds / 3600);
    const durationMinutes = Math.floor((normalizedSeconds % 3600) / 60);
    const durationSeconds = normalizedSeconds % 60;
    return `PT${durationHours}H${durationMinutes}M${durationSeconds}S`;
  }

  function parseDurationHours(duration: string | undefined): number | undefined {
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

  function cloneProjectModel(model: ProjectModel): ProjectModel {
    return JSON.parse(JSON.stringify(model)) as ProjectModel;
  }

  (globalThis as typeof globalThis & {
    __mikuprojectProjectPatchJson?: {
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
  }).__mikuprojectProjectPatchJson = {
    importProjectPatchJson,
    validatePatchDocument
  };
})();
