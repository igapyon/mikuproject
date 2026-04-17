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

  const projectPatchJsonUtil = (globalThis as typeof globalThis & {
    __mikuprojectProjectPatchJsonUtil?: {
      isZeroDuration: (duration: string | undefined) => boolean;
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
    };
  }).__mikuprojectProjectPatchJsonUtil;

  if (!projectPatchJsonUtil) {
    throw new Error("mikuproject Project Patch JSON util module is not loaded");
  }

  function applyUpdateTaskOperation(
    task: TaskModel,
    rawFields: Record<string, unknown>,
    model: ProjectModel,
    changes: ImportChange[],
    warnings: PatchWarning[],
    operationIndex: number
  ): void {
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

    const handledFields = new Set<string>();
    let nextMilestone = task.milestone;

    if (fields.is_milestone !== undefined) {
      handledFields.add("is_milestone");
      if (typeof fields.is_milestone !== "boolean") {
        warnings.push({ message: `update_task.is_milestone は boolean が必要です: ${task.uid}`, scope: "tasks", uid: task.uid, label: task.name || task.uid });
      } else if (task.summary && fields.is_milestone) {
        warnings.push({ message: `update_task では summary task を milestone にできません: ${task.uid}`, scope: "tasks", uid: task.uid, label: task.name || task.uid });
      } else if (task.milestone !== fields.is_milestone) {
        changes.push({ scope: "tasks", uid: task.uid, label: task.name || task.uid, field: "is_milestone", before: task.milestone, after: fields.is_milestone });
        task.milestone = fields.is_milestone;
        nextMilestone = fields.is_milestone;
      }
    }

    if (fields.name !== undefined) {
      handledFields.add("name");
      if (typeof fields.name !== "string" || fields.name.trim() === "") {
        warnings.push({ message: `update_task.name は空でない文字列が必要です: ${task.uid}`, scope: "tasks", uid: task.uid, label: task.name || task.uid });
      } else if (task.name !== fields.name.trim()) {
        changes.push({ scope: "tasks", uid: task.uid, label: task.name || task.uid, field: "name", before: task.name, after: fields.name.trim() });
        task.name = fields.name.trim();
      }
    }

    if (fields.notes !== undefined) {
      handledFields.add("notes");
      if (typeof fields.notes !== "string") {
        warnings.push({ message: `update_task.notes は文字列が必要です: ${task.uid}`, scope: "tasks", uid: task.uid, label: task.name || task.uid });
      } else if ((task.notes || "") !== fields.notes) {
        changes.push({ scope: "tasks", uid: task.uid, label: task.name || task.uid, field: "notes", before: task.notes, after: fields.notes });
        task.notes = fields.notes || undefined;
      }
    }

    if (fields.calendar_uid !== undefined) {
      handledFields.add("calendar_uid");
      if (typeof fields.calendar_uid !== "string") {
        warnings.push({ message: `update_task.calendar_uid は文字列が必要です: ${task.uid}`, scope: "tasks", uid: task.uid, label: task.name || task.uid });
      } else {
        const nextCalendarUid = fields.calendar_uid.trim() || undefined;
        if (nextCalendarUid && !model.calendars.some((calendar) => calendar.uid === nextCalendarUid)) {
          warnings.push({ message: `update_task.calendar_uid が既存 calendar を指していません: ${task.uid} -> ${nextCalendarUid}`, scope: "tasks", uid: task.uid, label: task.name || task.uid });
        } else if (task.calendarUID !== nextCalendarUid) {
          changes.push({ scope: "tasks", uid: task.uid, label: task.name || task.uid, field: "calendarUID", before: task.calendarUID, after: nextCalendarUid || "(cleared)" });
          task.calendarUID = nextCalendarUid;
        }
      }
    }

    if (fields.percent_complete !== undefined) {
      handledFields.add("percent_complete");
      if (typeof fields.percent_complete !== "number" || !Number.isFinite(fields.percent_complete) || fields.percent_complete < 0 || fields.percent_complete > 100) {
        warnings.push({ message: `update_task.percent_complete は 0 以上 100 以下の数値が必要です: ${task.uid}`, scope: "tasks", uid: task.uid, label: task.name || task.uid });
      } else if (task.percentComplete !== fields.percent_complete) {
        changes.push({ scope: "tasks", uid: task.uid, label: task.name || task.uid, field: "percentComplete", before: task.percentComplete, after: fields.percent_complete });
        task.percentComplete = fields.percent_complete;
      }
    }

    if (fields.percent_work_complete !== undefined) {
      handledFields.add("percent_work_complete");
      if (typeof fields.percent_work_complete !== "number" || !Number.isFinite(fields.percent_work_complete) || fields.percent_work_complete < 0 || fields.percent_work_complete > 100) {
        warnings.push({ message: `update_task.percent_work_complete は 0 以上 100 以下の数値が必要です: ${task.uid}`, scope: "tasks", uid: task.uid, label: task.name || task.uid });
      } else if (task.percentWorkComplete !== fields.percent_work_complete) {
        changes.push({ scope: "tasks", uid: task.uid, label: task.name || task.uid, field: "percentWorkComplete", before: task.percentWorkComplete, after: fields.percent_work_complete });
        task.percentWorkComplete = fields.percent_work_complete;
      }
    }

    if (fields.critical !== undefined) {
      handledFields.add("critical");
      if (typeof fields.critical !== "boolean") {
        warnings.push({ message: `update_task.critical は boolean が必要です: ${task.uid}`, scope: "tasks", uid: task.uid, label: task.name || task.uid });
      } else if (task.critical !== fields.critical) {
        changes.push({ scope: "tasks", uid: task.uid, label: task.name || task.uid, field: "critical", before: task.critical, after: fields.critical });
        task.critical = fields.critical;
      }
    }

    if (fields.planned_start !== undefined) {
      handledFields.add("planned_start");
      const normalizedStart = projectPatchJsonUtil.normalizePatchedTaskDate(fields.planned_start, "start", { ...task, milestone: nextMilestone }, project);
      if (normalizedStart === undefined) {
        warnings.push({ message: `update_task.planned_start の日付形式が解釈できません: ${task.uid}`, scope: "tasks", uid: task.uid, label: task.name || task.uid });
      } else if (task.start !== normalizedStart) {
        changes.push({ scope: "tasks", uid: task.uid, label: task.name || task.uid, field: "planned_start", before: task.start, after: normalizedStart });
        task.start = normalizedStart;
      }
    }

    if (fields.planned_finish !== undefined) {
      handledFields.add("planned_finish");
      const normalizedFinish = projectPatchJsonUtil.normalizePatchedTaskDate(fields.planned_finish, "finish", { ...task, milestone: nextMilestone }, project);
      if (normalizedFinish === undefined) {
        warnings.push({ message: `update_task.planned_finish の日付形式が解釈できません: ${task.uid}`, scope: "tasks", uid: task.uid, label: task.name || task.uid });
      } else if (task.finish !== normalizedFinish) {
        changes.push({ scope: "tasks", uid: task.uid, label: task.name || task.uid, field: "planned_finish", before: task.finish, after: normalizedFinish });
        task.finish = normalizedFinish;
      }
    }

    if (fields.planned_duration !== undefined) {
      handledFields.add("planned_duration");
      if (typeof fields.planned_duration !== "string" || fields.planned_duration.trim() === "") {
        warnings.push({ message: `update_task.planned_duration は空でない文字列が必要です: ${task.uid}`, scope: "tasks", uid: task.uid, label: task.name || task.uid });
      } else if (task.duration !== fields.planned_duration.trim()) {
        changes.push({ scope: "tasks", uid: task.uid, label: task.name || task.uid, field: "planned_duration", before: task.duration, after: fields.planned_duration.trim() });
        task.duration = fields.planned_duration.trim();
      }
    }

    if (fields.planned_duration_hours !== undefined) {
      handledFields.add("planned_duration_hours");
      if (typeof fields.planned_duration_hours !== "number" || !Number.isFinite(fields.planned_duration_hours) || fields.planned_duration_hours < 0) {
        warnings.push({ message: `update_task.planned_duration_hours は 0 以上の数値が必要です: ${task.uid}`, scope: "tasks", uid: task.uid, label: task.name || task.uid });
      } else {
        const beforeHours = projectPatchJsonUtil.parseDurationHours(task.duration);
        const nextDuration = projectPatchJsonUtil.formatDurationHours(fields.planned_duration_hours);
        if (task.duration !== nextDuration) {
          changes.push({ scope: "tasks", uid: task.uid, label: task.name || task.uid, field: "planned_duration_hours", before: beforeHours, after: fields.planned_duration_hours });
          task.duration = nextDuration;
        }
      }
    }

    if (nextMilestone) {
      if (task.finish !== task.start) {
        warnings.push({ message: `update_task.is_milestone=true のため planned_finish は planned_start に揃えます: ${task.uid}`, scope: "tasks", uid: task.uid, label: task.name || task.uid });
        if (task.finish !== task.start) {
          changes.push({ scope: "tasks", uid: task.uid, label: task.name || task.uid, field: "planned_finish", before: task.finish, after: task.start });
          task.finish = task.start;
        }
      }
      if (!projectPatchJsonUtil.isZeroDuration(task.duration)) {
        warnings.push({ message: `update_task.is_milestone=true のため planned_duration は 0 に揃えます: ${task.uid}`, scope: "tasks", uid: task.uid, label: task.name || task.uid });
        changes.push({ scope: "tasks", uid: task.uid, label: task.name || task.uid, field: "planned_duration", before: task.duration, after: "PT0H0M0S" });
        task.duration = "PT0H0M0S";
      }
    }

    Object.keys(fields).forEach((fieldName) => {
      if (!handledFields.has(fieldName)) {
        warnings.push({ message: `未対応の field は無視します: operations[${operationIndex}].fields.${fieldName}`, scope: "tasks", uid: task.uid, label: task.name || task.uid });
      }
    });
  }

  function applyUpdateProjectOperation(
    project: ProjectInfo,
    rawFields: Record<string, unknown>,
    model: ProjectModel,
    changes: ImportChange[],
    warnings: PatchWarning[],
    operationIndex: number
  ): void {
    const fields = { ...rawFields };
    const projectUid = "project";
    const projectLabel = project.name || "project";
    if (Object.keys(fields).length === 0) {
      warnings.push({ message: "update_project に fields がありません", scope: "project", uid: projectUid, label: projectLabel });
      return;
    }

    const handledFields = new Set<string>();

    if (fields.name !== undefined) {
      handledFields.add("name");
      if (typeof fields.name !== "string" || fields.name.trim() === "") {
        warnings.push({ message: "update_project.name は空でない文字列が必要です", scope: "project", uid: projectUid, label: projectLabel });
      } else if (project.name !== fields.name.trim()) {
        changes.push({ scope: "project", uid: projectUid, label: projectLabel, field: "name", before: project.name, after: fields.name.trim() });
        project.name = fields.name.trim();
      }
    }
    if (fields.title !== undefined) {
      handledFields.add("title");
      if (typeof fields.title !== "string") {
        warnings.push({ message: "update_project.title は文字列が必要です", scope: "project", uid: projectUid, label: projectLabel });
      } else {
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
      } else {
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
      } else {
        const nextCompany = fields.company.trim() || undefined;
        if (project.company !== nextCompany) {
          changes.push({ scope: "project", uid: projectUid, label: projectLabel, field: "company", before: project.company, after: nextCompany || "(cleared)" });
          project.company = nextCompany;
        }
      }
    }
    if (fields.current_date !== undefined) {
      handledFields.add("current_date");
      const normalizedCurrentDate = projectPatchJsonUtil.normalizePatchedPlainDateTime(fields.current_date, "start", project);
      if (normalizedCurrentDate === undefined) {
        warnings.push({ message: "update_project.current_date の日付形式が解釈できません", scope: "project", uid: projectUid, label: projectLabel });
      } else if (project.currentDate !== normalizedCurrentDate) {
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
      const normalizedStartDate = projectPatchJsonUtil.normalizePatchedPlainDateTime(fields.start_date, "start", project);
      if (normalizedStartDate === undefined) {
        warnings.push({ message: "update_project.start_date の日付形式が解釈できません", scope: "project", uid: projectUid, label: projectLabel });
      } else {
        nextStartDate = normalizedStartDate;
        startDateTouched = true;
      }
    }
    if (fields.finish_date !== undefined) {
      handledFields.add("finish_date");
      const normalizedFinishDate = projectPatchJsonUtil.normalizePatchedPlainDateTime(fields.finish_date, "finish", project);
      if (normalizedFinishDate === undefined) {
        warnings.push({ message: "update_project.finish_date の日付形式が解釈できません", scope: "project", uid: projectUid, label: projectLabel });
      } else {
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
      const normalizedStatusDate = projectPatchJsonUtil.normalizePatchedPlainDateTime(fields.status_date, "start", project);
      if (normalizedStatusDate === undefined) {
        warnings.push({ message: "update_project.status_date の日付形式が解釈できません", scope: "project", uid: projectUid, label: projectLabel });
      } else if (project.statusDate !== normalizedStatusDate) {
        changes.push({ scope: "project", uid: projectUid, label: projectLabel, field: "statusDate", before: project.statusDate, after: normalizedStatusDate });
        project.statusDate = normalizedStatusDate;
      }
    }
    if (fields.calendar_uid !== undefined) {
      handledFields.add("calendar_uid");
      if (typeof fields.calendar_uid !== "string") {
        warnings.push({ message: "update_project.calendar_uid は文字列が必要です", scope: "project", uid: projectUid, label: projectLabel });
      } else {
        const nextCalendarUid = fields.calendar_uid.trim() || undefined;
        if (nextCalendarUid && !model.calendars.some((calendar) => calendar.uid === nextCalendarUid)) {
          warnings.push({ message: `update_project.calendar_uid が既存 calendar を指していません: ${nextCalendarUid}`, scope: "project", uid: projectUid, label: projectLabel });
        } else if (project.calendarUID !== nextCalendarUid) {
          changes.push({ scope: "project", uid: projectUid, label: projectLabel, field: "calendarUID", before: project.calendarUID, after: nextCalendarUid || "(cleared)" });
          project.calendarUID = nextCalendarUid;
        }
      }
    }
    if (fields.minutes_per_day !== undefined) {
      handledFields.add("minutes_per_day");
      if (typeof fields.minutes_per_day !== "number" || !Number.isFinite(fields.minutes_per_day) || fields.minutes_per_day <= 0) {
        warnings.push({ message: "update_project.minutes_per_day は 0 より大きい数値が必要です", scope: "project", uid: projectUid, label: projectLabel });
      } else if (project.minutesPerDay !== fields.minutes_per_day) {
        changes.push({ scope: "project", uid: projectUid, label: projectLabel, field: "minutesPerDay", before: project.minutesPerDay, after: fields.minutes_per_day });
        project.minutesPerDay = fields.minutes_per_day;
      }
    }
    if (fields.minutes_per_week !== undefined) {
      handledFields.add("minutes_per_week");
      if (typeof fields.minutes_per_week !== "number" || !Number.isFinite(fields.minutes_per_week) || fields.minutes_per_week <= 0) {
        warnings.push({ message: "update_project.minutes_per_week は 0 より大きい数値が必要です", scope: "project", uid: projectUid, label: projectLabel });
      } else if (project.minutesPerWeek !== fields.minutes_per_week) {
        changes.push({ scope: "project", uid: projectUid, label: projectLabel, field: "minutesPerWeek", before: project.minutesPerWeek, after: fields.minutes_per_week });
        project.minutesPerWeek = fields.minutes_per_week;
      }
    }
    if (fields.days_per_month !== undefined) {
      handledFields.add("days_per_month");
      if (typeof fields.days_per_month !== "number" || !Number.isFinite(fields.days_per_month) || fields.days_per_month <= 0) {
        warnings.push({ message: "update_project.days_per_month は 0 より大きい数値が必要です", scope: "project", uid: projectUid, label: projectLabel });
      } else if (project.daysPerMonth !== fields.days_per_month) {
        changes.push({ scope: "project", uid: projectUid, label: projectLabel, field: "daysPerMonth", before: project.daysPerMonth, after: fields.days_per_month });
        project.daysPerMonth = fields.days_per_month;
      }
    }
    if (fields.schedule_from_start !== undefined) {
      handledFields.add("schedule_from_start");
      if (typeof fields.schedule_from_start !== "boolean") {
        warnings.push({ message: "update_project.schedule_from_start は boolean が必要です", scope: "project", uid: projectUid, label: projectLabel });
      } else if (project.scheduleFromStart !== fields.schedule_from_start) {
        changes.push({ scope: "project", uid: projectUid, label: projectLabel, field: "scheduleFromStart", before: project.scheduleFromStart, after: fields.schedule_from_start });
        project.scheduleFromStart = fields.schedule_from_start;
      }
    }
    Object.keys(fields).forEach((fieldName) => {
      if (!handledFields.has(fieldName)) {
        warnings.push({ message: `未対応の field は無視します: operations[${operationIndex}].fields.${fieldName}`, scope: "project", uid: projectUid, label: projectLabel });
      }
    });
  }

  function applyUpdateAssignmentOperation(
    assignment: AssignmentModel,
    rawFields: Record<string, unknown>,
    project: ProjectInfo,
    changes: ImportChange[],
    warnings: PatchWarning[],
    operationIndex: number
  ): void {
    const fields = { ...rawFields };
    if (Object.keys(fields).length === 0) {
      warnings.push({ message: `update_assignment に fields がありません: ${assignment.uid}`, scope: "assignments", uid: assignment.uid, label: assignment.uid });
      return;
    }

    const handledFields = new Set<string>();
    let nextStart = assignment.start;
    let nextFinish = assignment.finish;
    let startTouched = false;
    let finishTouched = false;
    if (fields.start !== undefined) {
      handledFields.add("start");
      const normalizedStart = projectPatchJsonUtil.normalizePatchedPlainDateTime(fields.start, "start", project);
      if (normalizedStart === undefined) {
        warnings.push({ message: `update_assignment.start の日付形式が解釈できません: ${assignment.uid}`, scope: "assignments", uid: assignment.uid, label: assignment.uid });
      } else {
        nextStart = normalizedStart;
        startTouched = true;
      }
    }
    if (fields.finish !== undefined) {
      handledFields.add("finish");
      const normalizedFinish = projectPatchJsonUtil.normalizePatchedPlainDateTime(fields.finish, "finish", project);
      if (normalizedFinish === undefined) {
        warnings.push({ message: `update_assignment.finish の日付形式が解釈できません: ${assignment.uid}`, scope: "assignments", uid: assignment.uid, label: assignment.uid });
      } else {
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
      changes.push({ scope: "assignments", uid: assignment.uid, label: assignment.uid, field: "start", before: assignment.start, after: nextStart as string });
      assignment.start = nextStart;
    }
    if (finishTouched && assignment.finish !== nextFinish) {
      changes.push({ scope: "assignments", uid: assignment.uid, label: assignment.uid, field: "finish", before: assignment.finish, after: nextFinish as string });
      assignment.finish = nextFinish;
    }
    if (fields.units !== undefined) {
      handledFields.add("units");
      if (typeof fields.units !== "number" || !Number.isFinite(fields.units) || fields.units < 0) {
        warnings.push({ message: `update_assignment.units は 0 以上の数値が必要です: ${assignment.uid}`, scope: "assignments", uid: assignment.uid, label: assignment.uid });
      } else if (assignment.units !== fields.units) {
        changes.push({ scope: "assignments", uid: assignment.uid, label: assignment.uid, field: "units", before: assignment.units, after: fields.units });
        assignment.units = fields.units;
      }
    }
    if (fields.work !== undefined) {
      handledFields.add("work");
      if (typeof fields.work !== "string" || fields.work.trim() === "") {
        warnings.push({ message: `update_assignment.work は空でない文字列が必要です: ${assignment.uid}`, scope: "assignments", uid: assignment.uid, label: assignment.uid });
      } else if (assignment.work !== fields.work.trim()) {
        changes.push({ scope: "assignments", uid: assignment.uid, label: assignment.uid, field: "work", before: assignment.work, after: fields.work.trim() });
        assignment.work = fields.work.trim();
      }
    }
    if (fields.percent_work_complete !== undefined) {
      handledFields.add("percent_work_complete");
      if (typeof fields.percent_work_complete !== "number" || !Number.isFinite(fields.percent_work_complete) || fields.percent_work_complete < 0 || fields.percent_work_complete > 100) {
        warnings.push({ message: `update_assignment.percent_work_complete は 0 以上 100 以下の数値が必要です: ${assignment.uid}`, scope: "assignments", uid: assignment.uid, label: assignment.uid });
      } else if (assignment.percentWorkComplete !== fields.percent_work_complete) {
        changes.push({ scope: "assignments", uid: assignment.uid, label: assignment.uid, field: "percentWorkComplete", before: assignment.percentWorkComplete, after: fields.percent_work_complete });
        assignment.percentWorkComplete = fields.percent_work_complete;
      }
    }
    Object.keys(fields).forEach((fieldName) => {
      if (!handledFields.has(fieldName)) {
        warnings.push({ message: `未対応の field は無視します: operations[${operationIndex}].fields.${fieldName}`, scope: "assignments", uid: assignment.uid, label: assignment.uid });
      }
    });
  }

  function applyUpdateResourceOperation(
    resource: ResourceModel,
    rawFields: Record<string, unknown>,
    model: ProjectModel,
    changes: ImportChange[],
    warnings: PatchWarning[],
    operationIndex: number
  ): void {
    const fields = { ...rawFields };
    if (Object.keys(fields).length === 0) {
      warnings.push({ message: `update_resource に fields がありません: ${resource.uid}`, scope: "resources", uid: resource.uid, label: resource.name || resource.uid });
      return;
    }
    const handledFields = new Set<string>();
    if (fields.name !== undefined) {
      handledFields.add("name");
      if (typeof fields.name !== "string" || fields.name.trim() === "") {
        warnings.push({ message: `update_resource.name は空でない文字列が必要です: ${resource.uid}`, scope: "resources", uid: resource.uid, label: resource.name || resource.uid });
      } else if (resource.name !== fields.name.trim()) {
        changes.push({ scope: "resources", uid: resource.uid, label: resource.name || resource.uid, field: "name", before: resource.name, after: fields.name.trim() });
        resource.name = fields.name.trim();
      }
    }
    if (fields.initials !== undefined) {
      handledFields.add("initials");
      if (typeof fields.initials !== "string") {
        warnings.push({ message: `update_resource.initials は文字列が必要です: ${resource.uid}`, scope: "resources", uid: resource.uid, label: resource.name || resource.uid });
      } else {
        const nextInitials = fields.initials.trim() || undefined;
        if (resource.initials !== nextInitials) {
          changes.push({ scope: "resources", uid: resource.uid, label: resource.name || resource.uid, field: "initials", before: resource.initials, after: nextInitials || "(cleared)" });
          resource.initials = nextInitials;
        }
      }
    }
    if (fields.group !== undefined) {
      handledFields.add("group");
      if (typeof fields.group !== "string") {
        warnings.push({ message: `update_resource.group は文字列が必要です: ${resource.uid}`, scope: "resources", uid: resource.uid, label: resource.name || resource.uid });
      } else {
        const nextGroup = fields.group.trim() || undefined;
        if (resource.group !== nextGroup) {
          changes.push({ scope: "resources", uid: resource.uid, label: resource.name || resource.uid, field: "group", before: resource.group, after: nextGroup || "(cleared)" });
          resource.group = nextGroup;
        }
      }
    }
    if (fields.calendar_uid !== undefined) {
      handledFields.add("calendar_uid");
      if (typeof fields.calendar_uid !== "string") {
        warnings.push({ message: `update_resource.calendar_uid は文字列が必要です: ${resource.uid}`, scope: "resources", uid: resource.uid, label: resource.name || resource.uid });
      } else {
        const nextCalendarUid = fields.calendar_uid.trim() || undefined;
        if (nextCalendarUid && !model.calendars.some((calendar) => calendar.uid === nextCalendarUid)) {
          warnings.push({ message: `update_resource.calendar_uid が既存 calendar を指していません: ${resource.uid} -> ${nextCalendarUid}`, scope: "resources", uid: resource.uid, label: resource.name || resource.uid });
        } else if (resource.calendarUID !== nextCalendarUid) {
          changes.push({ scope: "resources", uid: resource.uid, label: resource.name || resource.uid, field: "calendarUID", before: resource.calendarUID, after: nextCalendarUid || "(cleared)" });
          resource.calendarUID = nextCalendarUid;
        }
      }
    }
    if (fields.max_units !== undefined) {
      handledFields.add("max_units");
      if (typeof fields.max_units !== "number" || !Number.isFinite(fields.max_units) || fields.max_units < 0) {
        warnings.push({ message: `update_resource.max_units は 0 以上の数値が必要です: ${resource.uid}`, scope: "resources", uid: resource.uid, label: resource.name || resource.uid });
      } else if (resource.maxUnits !== fields.max_units) {
        changes.push({ scope: "resources", uid: resource.uid, label: resource.name || resource.uid, field: "maxUnits", before: resource.maxUnits, after: fields.max_units });
        resource.maxUnits = fields.max_units;
      }
    }
    if (fields.standard_rate !== undefined) {
      handledFields.add("standard_rate");
      if (typeof fields.standard_rate !== "string") {
        warnings.push({ message: `update_resource.standard_rate は文字列が必要です: ${resource.uid}`, scope: "resources", uid: resource.uid, label: resource.name || resource.uid });
      } else {
        const nextStandardRate = fields.standard_rate.trim() || undefined;
        if (resource.standardRate !== nextStandardRate) {
          changes.push({ scope: "resources", uid: resource.uid, label: resource.name || resource.uid, field: "standardRate", before: resource.standardRate, after: nextStandardRate || "(cleared)" });
          resource.standardRate = nextStandardRate;
        }
      }
    }
    if (fields.overtime_rate !== undefined) {
      handledFields.add("overtime_rate");
      if (typeof fields.overtime_rate !== "string") {
        warnings.push({ message: `update_resource.overtime_rate は文字列が必要です: ${resource.uid}`, scope: "resources", uid: resource.uid, label: resource.name || resource.uid });
      } else {
        const nextOvertimeRate = fields.overtime_rate.trim() || undefined;
        if (resource.overtimeRate !== nextOvertimeRate) {
          changes.push({ scope: "resources", uid: resource.uid, label: resource.name || resource.uid, field: "overtimeRate", before: resource.overtimeRate, after: nextOvertimeRate || "(cleared)" });
          resource.overtimeRate = nextOvertimeRate;
        }
      }
    }
    if (fields.cost_per_use !== undefined) {
      handledFields.add("cost_per_use");
      if (typeof fields.cost_per_use !== "number" || !Number.isFinite(fields.cost_per_use) || fields.cost_per_use < 0) {
        warnings.push({ message: `update_resource.cost_per_use は 0 以上の数値が必要です: ${resource.uid}`, scope: "resources", uid: resource.uid, label: resource.name || resource.uid });
      } else if (resource.costPerUse !== fields.cost_per_use) {
        changes.push({ scope: "resources", uid: resource.uid, label: resource.name || resource.uid, field: "costPerUse", before: resource.costPerUse, after: fields.cost_per_use });
        resource.costPerUse = fields.cost_per_use;
      }
    }
    if (fields.percent_work_complete !== undefined) {
      handledFields.add("percent_work_complete");
      if (typeof fields.percent_work_complete !== "number" || !Number.isFinite(fields.percent_work_complete) || fields.percent_work_complete < 0 || fields.percent_work_complete > 100) {
        warnings.push({ message: `update_resource.percent_work_complete は 0 以上 100 以下の数値が必要です: ${resource.uid}`, scope: "resources", uid: resource.uid, label: resource.name || resource.uid });
      } else if (resource.percentWorkComplete !== fields.percent_work_complete) {
        changes.push({ scope: "resources", uid: resource.uid, label: resource.name || resource.uid, field: "percentWorkComplete", before: resource.percentWorkComplete, after: fields.percent_work_complete });
        resource.percentWorkComplete = fields.percent_work_complete;
      }
    }
    Object.keys(fields).forEach((fieldName) => {
      if (!handledFields.has(fieldName)) {
        warnings.push({ message: `未対応の field は無視します: operations[${operationIndex}].fields.${fieldName}`, scope: "resources", uid: resource.uid, label: resource.name || resource.uid });
      }
    });
  }

  function applyUpdateCalendarOperation(
    calendar: CalendarModel,
    rawFields: Record<string, unknown>,
    model: ProjectModel,
    changes: ImportChange[],
    warnings: PatchWarning[],
    operationIndex: number
  ): void {
    const fields = { ...rawFields };
    if (Object.keys(fields).length === 0) {
      warnings.push({ message: `update_calendar に fields がありません: ${calendar.uid}`, scope: "calendars", uid: calendar.uid, label: calendar.name || calendar.uid });
      return;
    }
    const handledFields = new Set<string>();
    if (fields.name !== undefined) {
      handledFields.add("name");
      if (typeof fields.name !== "string" || fields.name.trim() === "") {
        warnings.push({ message: `update_calendar.name は空でない文字列が必要です: ${calendar.uid}`, scope: "calendars", uid: calendar.uid, label: calendar.name || calendar.uid });
      } else if (calendar.name !== fields.name.trim()) {
        changes.push({ scope: "calendars", uid: calendar.uid, label: calendar.name || calendar.uid, field: "name", before: calendar.name, after: fields.name.trim() });
        calendar.name = fields.name.trim();
      }
    }
    if (fields.is_base_calendar !== undefined) {
      handledFields.add("is_base_calendar");
      if (typeof fields.is_base_calendar !== "boolean") {
        warnings.push({ message: `update_calendar.is_base_calendar は boolean が必要です: ${calendar.uid}`, scope: "calendars", uid: calendar.uid, label: calendar.name || calendar.uid });
      } else if (calendar.isBaseCalendar !== fields.is_base_calendar) {
        changes.push({ scope: "calendars", uid: calendar.uid, label: calendar.name || calendar.uid, field: "isBaseCalendar", before: calendar.isBaseCalendar, after: fields.is_base_calendar });
        calendar.isBaseCalendar = fields.is_base_calendar;
      }
    }
    if (fields.base_calendar_uid !== undefined) {
      handledFields.add("base_calendar_uid");
      if (typeof fields.base_calendar_uid !== "string") {
        warnings.push({ message: `update_calendar.base_calendar_uid は文字列が必要です: ${calendar.uid}`, scope: "calendars", uid: calendar.uid, label: calendar.name || calendar.uid });
      } else {
        const nextBaseCalendarUid = fields.base_calendar_uid.trim() || undefined;
        if (nextBaseCalendarUid === calendar.uid) {
          warnings.push({ message: `update_calendar.base_calendar_uid は自身を指せません: ${calendar.uid}`, scope: "calendars", uid: calendar.uid, label: calendar.name || calendar.uid });
        } else if (nextBaseCalendarUid && !model.calendars.some((item) => item.uid === nextBaseCalendarUid)) {
          warnings.push({ message: `update_calendar.base_calendar_uid が既存 calendar を指していません: ${calendar.uid} -> ${nextBaseCalendarUid}`, scope: "calendars", uid: calendar.uid, label: calendar.name || calendar.uid });
        } else if (calendar.baseCalendarUID !== nextBaseCalendarUid) {
          changes.push({ scope: "calendars", uid: calendar.uid, label: calendar.name || calendar.uid, field: "baseCalendarUID", before: calendar.baseCalendarUID, after: nextBaseCalendarUid || "(cleared)" });
          calendar.baseCalendarUID = nextBaseCalendarUid;
        }
      }
    }
    Object.keys(fields).forEach((fieldName) => {
      if (!handledFields.has(fieldName)) {
        warnings.push({ message: `未対応の field は無視します: operations[${operationIndex}].fields.${fieldName}`, scope: "calendars", uid: calendar.uid, label: calendar.name || calendar.uid });
      }
    });
  }

  (globalThis as typeof globalThis & {
    __mikuprojectProjectPatchJsonUpdates?: {
      applyUpdateTaskOperation: typeof applyUpdateTaskOperation;
      applyUpdateProjectOperation: typeof applyUpdateProjectOperation;
      applyUpdateAssignmentOperation: typeof applyUpdateAssignmentOperation;
      applyUpdateResourceOperation: typeof applyUpdateResourceOperation;
      applyUpdateCalendarOperation: typeof applyUpdateCalendarOperation;
    };
  }).__mikuprojectProjectPatchJsonUpdates = {
    applyUpdateTaskOperation,
    applyUpdateProjectOperation,
    applyUpdateAssignmentOperation,
    applyUpdateResourceOperation,
    applyUpdateCalendarOperation
  };
})();
