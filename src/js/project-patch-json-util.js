/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const projectPatchJsonUtil = {
        resolveDependencyType(rawType, warnings, task, operationIndex, opName, options = {}) {
            const trimmed = String(rawType || "").trim();
            if (!trimmed) {
                return options.allowEmpty ? undefined : 1;
            }
            const typeMap = { FS: 1, SS: 2, FF: 3, SF: 4 };
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
        resolveDependencyLag(operation, warnings, task, operationIndex, opName, options = {}) {
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
        formatDependencyType(type) {
            const typeMap = { 1: "FS", 2: "SS", 3: "FF", 4: "SF" };
            return typeMap[type !== null && type !== void 0 ? type : 1] || `type=${type}`;
        },
        isZeroDuration(duration) {
            const text = String(duration || "").trim();
            return text === "" || text === "PT0H0M0S" || text === "PT0M0S" || text === "PT0S";
        },
        normalizeDurationText(duration) {
            const text = String(duration || "").trim();
            return text || undefined;
        },
        isValidDurationText(duration) {
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
        formatPredecessor(predecessor) {
            const extras = [projectPatchJsonUtil.formatDependencyType(predecessor.type)];
            if (!projectPatchJsonUtil.isZeroDuration(predecessor.linkLag)) {
                extras.push(`lag=${String(predecessor.linkLag || "").trim()}`);
            }
            return `${predecessor.predecessorUid}(${extras.join(", ")})`;
        },
        formatRequestedDependencyRelation(fromUid, toUid, type, linkLag) {
            const extras = [projectPatchJsonUtil.formatDependencyType(type)];
            if (!projectPatchJsonUtil.isZeroDuration(linkLag)) {
                extras.push(`lag=${String(linkLag || "").trim()}`);
            }
            return `${fromUid} -> ${toUid} (${extras.join(", ")})`;
        },
        formatPredecessorList(predecessors) {
            return predecessors.map((item) => projectPatchJsonUtil.formatPredecessor(item)).join(", ");
        },
        normalizePatchedTaskDate(value, kind, task, project) {
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
        normalizePatchedPlainDateTime(value, kind, project) {
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
        isDateOnlyText(value) {
            return /^\d{4}-\d{2}-\d{2}$/.test(value);
        },
        isDateText(value) {
            if (projectPatchJsonUtil.isDateOnlyText(value)) {
                return true;
            }
            return !Number.isNaN(new Date(value).getTime());
        },
        formatDurationHours(hours) {
            const totalSeconds = Math.round(hours * 60 * 60);
            const normalizedSeconds = Math.max(0, totalSeconds);
            const durationHours = Math.floor(normalizedSeconds / 3600);
            const durationMinutes = Math.floor((normalizedSeconds % 3600) / 60);
            const durationSeconds = normalizedSeconds % 60;
            return `PT${durationHours}H${durationMinutes}M${durationSeconds}S`;
        },
        parseDurationHours(duration) {
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
        cloneProjectModel(model) {
            return JSON.parse(JSON.stringify(model));
        }
    };
    globalThis.__mikuprojectProjectPatchJsonUtil = projectPatchJsonUtil;
})();
