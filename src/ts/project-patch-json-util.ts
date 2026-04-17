/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  type PatchWarning = {
    message: string;
    scope?: "project" | "tasks" | "resources" | "assignments" | "calendars";
    uid?: string;
    label?: string;
  };

  type PatchOperation = {
    lag?: string;
    lag_hours?: number;
  };

  const projectPatchJsonUtil = {
    resolveDependencyType(
      rawType: unknown,
      warnings: PatchWarning[],
      task: TaskModel,
      operationIndex: number,
      opName: "link_tasks" | "unlink_tasks",
      options: { allowEmpty?: boolean } = {}
    ): number | undefined {
      const trimmed = String(rawType || "").trim();
      if (!trimmed) {
        return options.allowEmpty ? undefined : 1;
      }
      const typeMap: Record<string, number> = { FS: 1, SS: 2, FF: 3, SF: 4 };
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
    },

    resolveDependencyLag(
      operation: PatchOperation,
      warnings: PatchWarning[],
      task: TaskModel,
      operationIndex: number,
      opName: "link_tasks" | "unlink_tasks",
      options: { allowEmpty?: boolean } = {}
    ): string | undefined {
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
        if (!projectPatchJsonUtil.isValidDurationText(normalizedLag)) {
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
        return projectPatchJsonUtil.formatDurationHours(operation.lag_hours);
      }
      return options.allowEmpty ? undefined : undefined;
    },

    formatDependencyType(type: number | undefined): string {
      const typeMap: Record<number, string> = { 1: "FS", 2: "SS", 3: "FF", 4: "SF" };
      return typeMap[type ?? 1] || `type=${type}`;
    },

    isZeroDuration(duration: string | undefined): boolean {
      const text = String(duration || "").trim();
      return text === "" || text === "PT0H0M0S" || text === "PT0M0S" || text === "PT0S";
    },

    normalizeDurationText(duration: string | undefined): string | undefined {
      const text = String(duration || "").trim();
      return text || undefined;
    },

    isValidDurationText(duration: string): boolean {
      const text = String(duration || "").trim();
      if (!text) {
        return false;
      }
      const match = /^-?P(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)$/.exec(text);
      if (!match) {
        return false;
      }
      return Boolean(match[1] || match[2] || match[3]);
    },

    formatPredecessor(predecessor: PredecessorModel): string {
      const extras = [projectPatchJsonUtil.formatDependencyType(predecessor.type)];
      if (!projectPatchJsonUtil.isZeroDuration(predecessor.linkLag)) {
        extras.push(`lag=${String(predecessor.linkLag || "").trim()}`);
      }
      return `${predecessor.predecessorUid}(${extras.join(", ")})`;
    },

    formatRequestedDependencyRelation(
      fromUid: string,
      toUid: string,
      type: number | undefined,
      linkLag: string | undefined
    ): string {
      const extras = [projectPatchJsonUtil.formatDependencyType(type)];
      if (!projectPatchJsonUtil.isZeroDuration(linkLag)) {
        extras.push(`lag=${String(linkLag || "").trim()}`);
      }
      return `${fromUid} -> ${toUid} (${extras.join(", ")})`;
    },

    formatPredecessorList(predecessors: PredecessorModel[]): string {
      return predecessors.map((item) => projectPatchJsonUtil.formatPredecessor(item)).join(", ");
    },

    normalizePatchedTaskDate(
      value: unknown,
      kind: "start" | "finish",
      task: TaskModel,
      project: ProjectInfo
    ): string | undefined {
      if (typeof value !== "string") {
        return undefined;
      }
      const trimmed = value.trim();
      if (!trimmed || !projectPatchJsonUtil.isDateText(trimmed)) {
        return undefined;
      }
      if (projectPatchJsonUtil.isDateOnlyText(trimmed) && !task.milestone) {
        const timeText = kind === "start" ? (project.defaultStartTime || "09:00:00") : (project.defaultFinishTime || "18:00:00");
        return `${trimmed}T${timeText}`;
      }
      return trimmed;
    },

    normalizePatchedPlainDateTime(
      value: unknown,
      kind: "start" | "finish",
      project: ProjectInfo
    ): string | undefined {
      if (typeof value !== "string") {
        return undefined;
      }
      const trimmed = value.trim();
      if (!trimmed || !projectPatchJsonUtil.isDateText(trimmed)) {
        return undefined;
      }
      if (projectPatchJsonUtil.isDateOnlyText(trimmed)) {
        const timeText = kind === "start" ? (project.defaultStartTime || "09:00:00") : (project.defaultFinishTime || "18:00:00");
        return `${trimmed}T${timeText}`;
      }
      return trimmed;
    },

    isDateOnlyText(value: string): boolean {
      return /^\d{4}-\d{2}-\d{2}$/.test(value);
    },

    isDateText(value: string): boolean {
      if (projectPatchJsonUtil.isDateOnlyText(value)) {
        return true;
      }
      return !Number.isNaN(new Date(value).getTime());
    },

    formatDurationHours(hours: number): string {
      const totalSeconds = Math.round(hours * 60 * 60);
      const normalizedSeconds = Math.max(0, totalSeconds);
      const durationHours = Math.floor(normalizedSeconds / 3600);
      const durationMinutes = Math.floor((normalizedSeconds % 3600) / 60);
      const durationSeconds = normalizedSeconds % 60;
      return `PT${durationHours}H${durationMinutes}M${durationSeconds}S`;
    },

    parseDurationHours(duration: string | undefined): number | undefined {
      const text = String(duration || "").trim();
      const match = text.match(/^PT(?:(\d+(?:\.\d+)?)H)?(?:(\d+(?:\.\d+)?)M)?(?:(\d+(?:\.\d+)?)S)?$/);
      if (!match) {
        return undefined;
      }
      const hours = Number(match[1] || 0);
      const minutes = Number(match[2] || 0);
      const seconds = Number(match[3] || 0);
      return hours + (minutes / 60) + (seconds / 3600);
    },

    cloneProjectModel(model: ProjectModel): ProjectModel {
      return JSON.parse(JSON.stringify(model)) as ProjectModel;
    }
  };

  (globalThis as typeof globalThis & {
    __mikuprojectProjectPatchJsonUtil?: typeof projectPatchJsonUtil;
  }).__mikuprojectProjectPatchJsonUtil = projectPatchJsonUtil;
})();
