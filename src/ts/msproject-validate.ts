/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  globalThis.__mikuprojectMsprojectValidate = {
    validateProjectModel(input: {
      model: ProjectModel;
      parseDateValue: (value: string | undefined) => number | null;
      isPlaceholderUid: (value: string | undefined) => boolean;
      isUnassignedResourceUid: (value: string | undefined) => boolean;
      describeTask: (task: TaskModel) => string;
      describeResource: (resource: ResourceModel) => string;
      describeCalendar: (calendar: CalendarModel) => string;
      describeAssignment: (assignment: AssignmentModel) => string;
      describeTaskRef: (model: ProjectModel, taskUid: string | undefined) => string;
      describeResourceRef: (model: ProjectModel, resourceUid: string | undefined) => string;
      detectTaskOrderIssue: (tasks: TaskModel[]) => { previous: TaskModel; current: TaskModel } | null;
    }): ValidationIssue[] {
      const { model } = input;
      const issues: ValidationIssue[] = [];
      const taskUidSet = new Set<string>();
      const taskIdSet = new Set<string>();
      const resourceUidSet = new Set<string>();
      const calendarUidSet = new Set<string>();

      if (!model.project.name) {
        issues.push({ level: "warning", scope: "project", message: "Project Name が空です" });
      }
      if (model.project.saveVersion !== undefined && model.project.saveVersion < 0) {
        issues.push({ level: "warning", scope: "project", message: "Project SaveVersion は 0 以上が望ましいです" });
      }
      if (!model.project.startDate) {
        issues.push({ level: "warning", scope: "project", message: "Project StartDate が空です" });
      }
      if (!model.project.finishDate) {
        issues.push({ level: "warning", scope: "project", message: "Project FinishDate が空です" });
      }
      if (model.project.minutesPerDay !== undefined && model.project.minutesPerDay <= 0) {
        issues.push({ level: "warning", scope: "project", message: "Project MinutesPerDay は正の値が望ましいです" });
      }
      if (model.project.minutesPerWeek !== undefined && model.project.minutesPerWeek <= 0) {
        issues.push({ level: "warning", scope: "project", message: "Project MinutesPerWeek は正の値が望ましいです" });
      }
      if (model.project.daysPerMonth !== undefined && model.project.daysPerMonth <= 0) {
        issues.push({ level: "warning", scope: "project", message: "Project DaysPerMonth は正の値が望ましいです" });
      }
      if (model.project.weekStartDay !== undefined && (model.project.weekStartDay < 1 || model.project.weekStartDay > 7)) {
        issues.push({ level: "warning", scope: "project", message: "Project WeekStartDay は 1..7 が望ましいです" });
      }
      if (model.project.workFormat !== undefined && model.project.workFormat < 0) {
        issues.push({ level: "warning", scope: "project", message: "Project WorkFormat は 0 以上が望ましいです" });
      }
      if (model.project.durationFormat !== undefined && model.project.durationFormat < 0) {
        issues.push({ level: "warning", scope: "project", message: "Project DurationFormat は 0 以上が望ましいです" });
      }
      if (model.project.currencyDigits !== undefined && model.project.currencyDigits < 0) {
        issues.push({ level: "warning", scope: "project", message: "Project CurrencyDigits は 0 以上が望ましいです" });
      }
      if (model.project.currencySymbolPosition !== undefined && model.project.currencySymbolPosition < 0) {
        issues.push({ level: "warning", scope: "project", message: "Project CurrencySymbolPosition は 0 以上が望ましいです" });
      }
      if (model.project.fyStartDate !== undefined && !input.parseDateValue(model.project.fyStartDate)) {
        issues.push({ level: "warning", scope: "project", message: "Project FYStartDate の日付形式が解釈できません" });
      }
      if (model.project.criticalSlackLimit !== undefined && model.project.criticalSlackLimit < 0) {
        issues.push({ level: "warning", scope: "project", message: "Project CriticalSlackLimit は 0 以上が望ましいです" });
      }
      if (model.project.defaultTaskType !== undefined && model.project.defaultTaskType < 0) {
        issues.push({ level: "warning", scope: "project", message: "Project DefaultTaskType は 0 以上が望ましいです" });
      }
      if (model.project.defaultFixedCostAccrual !== undefined && model.project.defaultFixedCostAccrual < 0) {
        issues.push({ level: "warning", scope: "project", message: "Project DefaultFixedCostAccrual は 0 以上が望ましいです" });
      }
      if (model.project.defaultTaskEVMethod !== undefined && model.project.defaultTaskEVMethod < 0) {
        issues.push({ level: "warning", scope: "project", message: "Project DefaultTaskEVMethod は 0 以上が望ましいです" });
      }
      if (model.project.newTaskStartDate !== undefined && model.project.newTaskStartDate < 0) {
        issues.push({ level: "warning", scope: "project", message: "Project NewTaskStartDate は 0 以上が望ましいです" });
      }
      for (const outlineCode of model.project.outlineCodes) {
        if (!outlineCode.fieldID && !outlineCode.fieldName) {
          issues.push({ level: "warning", scope: "project", message: "Project OutlineCode は FieldID または FieldName を持つことが望ましいです" });
        }
        for (const mask of outlineCode.masks) {
          if (mask.level < 1) {
            issues.push({ level: "warning", scope: "project", message: "Project OutlineCode Mask Level は 1 以上が望ましいです" });
          }
        }
      }
      for (const wbsMask of model.project.wbsMasks) {
        if (wbsMask.level < 1) {
          issues.push({ level: "warning", scope: "project", message: "Project WBSMask Level は 1 以上が望ましいです" });
        }
      }
      for (const attribute of model.project.extendedAttributes) {
        if (!attribute.fieldID && !attribute.fieldName) {
          issues.push({ level: "warning", scope: "project", message: "Project ExtendedAttribute は FieldID または FieldName を持つことが望ましいです" });
        }
        if (attribute.calculationType !== undefined && attribute.calculationType < 0) {
          issues.push({ level: "warning", scope: "project", message: "Project ExtendedAttribute CalculationType は 0 以上が望ましいです" });
        }
      }

      for (const calendar of model.calendars) {
        if (!calendar.uid) {
          issues.push({ level: "error", scope: "calendars", message: "Calendar UID が空です" });
        }
        if (calendar.isBaselineCalendar !== undefined && !calendar.isBaseCalendar && calendar.isBaselineCalendar) {
          issues.push({
            level: "warning",
            scope: "calendars",
            message: `Calendar IsBaselineCalendar は通常 BaseCalendar と整合していることが望ましいです: ${input.describeCalendar(calendar)}`
          });
        }
        if (calendarUidSet.has(calendar.uid)) {
          issues.push({ level: "error", scope: "calendars", message: `Calendar UID が重複しています: ${calendar.uid}` });
        }
        calendarUidSet.add(calendar.uid);
        for (const weekDay of calendar.weekDays) {
          if (weekDay.dayType < 1 || weekDay.dayType > 7) {
            issues.push({
              level: "warning",
              scope: "calendars",
              message: `Calendar WeekDay DayType が 1..7 の範囲外です: ${input.describeCalendar(calendar)}`
            });
          }
          for (const workingTime of weekDay.workingTimes) {
            if (!workingTime.fromTime || !workingTime.toTime) {
              issues.push({
                level: "warning",
                scope: "calendars",
                message: `Calendar WorkingTime の時刻が不足しています: ${input.describeCalendar(calendar)}`
              });
            }
          }
        }
        for (const exception of calendar.exceptions) {
          const exceptionFrom = input.parseDateValue(exception.fromDate);
          const exceptionTo = input.parseDateValue(exception.toDate);
          if (exceptionFrom !== null && exceptionTo !== null && exceptionFrom > exceptionTo) {
            issues.push({
              level: "warning",
              scope: "calendars",
              message: `Calendar Exception FromDate が ToDate より後です: ${input.describeCalendar(calendar)}`
            });
          }
          for (const workingTime of exception.workingTimes) {
            if (!workingTime.fromTime || !workingTime.toTime) {
              issues.push({
                level: "warning",
                scope: "calendars",
                message: `Calendar Exception WorkingTime の時刻が不足しています: ${input.describeCalendar(calendar)}`
              });
            }
          }
        }
        for (const workWeek of calendar.workWeeks) {
          const workWeekFrom = input.parseDateValue(workWeek.fromDate);
          const workWeekTo = input.parseDateValue(workWeek.toDate);
          if (workWeekFrom !== null && workWeekTo !== null && workWeekFrom > workWeekTo) {
            issues.push({
              level: "warning",
              scope: "calendars",
              message: `Calendar WorkWeek FromDate が ToDate より後です: ${input.describeCalendar(calendar)}`
            });
          }
          for (const weekDay of workWeek.weekDays) {
            if (weekDay.dayType < 1 || weekDay.dayType > 7) {
              issues.push({
                level: "warning",
                scope: "calendars",
                message: `Calendar WorkWeek DayType が 1..7 の範囲外です: ${input.describeCalendar(calendar)}`
              });
            }
          }
        }
      }

      if (model.project.calendarUID && !calendarUidSet.has(model.project.calendarUID)) {
        issues.push({
          level: "error",
          scope: "project",
          message: `Project CalendarUID が既存 Calendar を指していません: ${model.project.calendarUID}`
        });
      }

      for (const calendar of model.calendars) {
        if (calendar.baseCalendarUID && !calendarUidSet.has(calendar.baseCalendarUID)) {
          issues.push({
            level: "warning",
            scope: "calendars",
            message: `Calendar BaseCalendarUID が既存 Calendar を指していません: ${input.describeCalendar(calendar)}`
          });
        }
        if (calendar.baseCalendarUID && calendar.baseCalendarUID === calendar.uid) {
          issues.push({
            level: "warning",
            scope: "calendars",
            message: `Calendar BaseCalendarUID が自身を指しています: ${input.describeCalendar(calendar)}`
          });
        }
      }

      for (const task of model.tasks) {
        if (!task.uid) {
          issues.push({ level: "error", scope: "tasks", message: "Task UID が空です" });
        }
        if (!task.id) {
          issues.push({ level: "error", scope: "tasks", message: `Task ID が空です: ${task.name || "(無名)"}` });
        }
        if (!task.name) {
          if (!input.isPlaceholderUid(task.uid)) {
            issues.push({ level: "warning", scope: "tasks", message: `Task Name が空です: ${input.describeTask(task)}` });
          }
        }
        if (taskIdSet.has(task.id)) {
          issues.push({ level: "error", scope: "tasks", message: `Task ID が重複しています: ${task.id}` });
        }
        taskIdSet.add(task.id);
        if (!task.start) {
          issues.push({ level: "warning", scope: "tasks", message: `Task Start が空です: ${input.describeTask(task)}` });
        }
        if (!task.finish) {
          issues.push({ level: "warning", scope: "tasks", message: `Task Finish が空です: ${input.describeTask(task)}` });
        }
        if (task.outlineLevel < 1 && !input.isPlaceholderUid(task.uid)) {
          issues.push({ level: "error", scope: "tasks", message: `Task OutlineLevel が不正です: ${input.describeTask(task)}` });
        }
        if (task.calendarUID && !calendarUidSet.has(task.calendarUID)) {
          issues.push({
            level: "warning",
            scope: "tasks",
            message: `Task CalendarUID が既存 Calendar を指していません: ${input.describeTask(task)}`
          });
        }
        if (task.outlineNumber && !input.isPlaceholderUid(task.uid)) {
          const outlineParts = task.outlineNumber.split(".").filter(Boolean);
          if (outlineParts.length !== task.outlineLevel) {
            issues.push({
              level: "warning",
              scope: "tasks",
              message: `Task OutlineNumber と OutlineLevel の整合が取れていません: ${input.describeTask(task)}`
            });
          }
        }
        if (task.percentComplete < 0 || task.percentComplete > 100) {
          issues.push({
            level: "warning",
            scope: "tasks",
            message: `Task PercentComplete が 0..100 の範囲外です: ${input.describeTask(task)}`
          });
        }
        if (task.percentWorkComplete !== undefined && (task.percentWorkComplete < 0 || task.percentWorkComplete > 100)) {
          issues.push({
            level: "warning",
            scope: "tasks",
            message: `Task PercentWorkComplete が 0..100 の範囲外です: ${input.describeTask(task)}`
          });
        }
        const taskStart = input.parseDateValue(task.start);
        const taskFinish = input.parseDateValue(task.finish);
        const taskActualStart = input.parseDateValue(task.actualStart);
        const taskActualFinish = input.parseDateValue(task.actualFinish);
        const taskDeadline = input.parseDateValue(task.deadline);
        if (taskStart !== null && taskFinish !== null && taskStart > taskFinish) {
          issues.push({ level: "warning", scope: "tasks", message: `Task Start が Finish より後です: ${input.describeTask(task)}` });
        }
        if (taskFinish !== null && taskDeadline !== null && taskFinish > taskDeadline) {
          issues.push({ level: "warning", scope: "tasks", message: `Task Finish が Deadline より後です: ${input.describeTask(task)}` });
        }
        if (taskActualStart !== null && taskActualFinish !== null && taskActualStart > taskActualFinish) {
          issues.push({ level: "warning", scope: "tasks", message: `Task ActualStart が ActualFinish より後です: ${input.describeTask(task)}` });
        }
        if (taskUidSet.has(task.uid)) {
          issues.push({ level: "error", scope: "tasks", message: `Task UID が重複しています: ${task.uid}` });
        }
        for (const attribute of task.extendedAttributes) {
          if (!attribute.fieldID) {
            issues.push({ level: "warning", scope: "tasks", message: `Task ExtendedAttribute に FieldID がありません: ${input.describeTask(task)}` });
          }
        }
        for (const baseline of task.baselines) {
          if (baseline.number !== undefined && baseline.number < 0) {
            issues.push({ level: "warning", scope: "tasks", message: `Task Baseline Number は 0 以上が望ましいです: ${input.describeTask(task)}` });
          }
          const baselineStart = input.parseDateValue(baseline.start);
          const baselineFinish = input.parseDateValue(baseline.finish);
          if (baselineStart !== null && baselineFinish !== null && baselineStart > baselineFinish) {
            issues.push({ level: "warning", scope: "tasks", message: `Task Baseline Start が Finish より後です: ${input.describeTask(task)}` });
          }
        }
        for (const timephasedData of task.timephasedData) {
          if (timephasedData.type !== undefined && timephasedData.type < 0) {
            issues.push({ level: "warning", scope: "tasks", message: `Task TimephasedData Type は 0 以上が望ましいです: ${input.describeTask(task)}` });
          }
          const timephasedStart = input.parseDateValue(timephasedData.start);
          const timephasedFinish = input.parseDateValue(timephasedData.finish);
          if (timephasedStart !== null && timephasedFinish !== null && timephasedStart > timephasedFinish) {
            issues.push({ level: "warning", scope: "tasks", message: `Task TimephasedData Start が Finish より後です: ${input.describeTask(task)}` });
          }
        }
        taskUidSet.add(task.uid);
        if (task.priority !== undefined && (task.priority < 0 || task.priority > 1000)) {
          issues.push({ level: "warning", scope: "tasks", message: `Task Priority が 0..1000 の範囲外です: ${input.describeTask(task)}` });
        }
        if (task.cost !== undefined && task.cost < 0) {
          issues.push({ level: "warning", scope: "tasks", message: `Task Cost が負値です: ${input.describeTask(task)}` });
        }
        if (task.actualCost !== undefined && task.actualCost < 0) {
          issues.push({ level: "warning", scope: "tasks", message: `Task ActualCost が負値です: ${input.describeTask(task)}` });
        }
        if (task.remainingCost !== undefined && task.remainingCost < 0) {
          issues.push({ level: "warning", scope: "tasks", message: `Task RemainingCost が負値です: ${input.describeTask(task)}` });
        }
      }
      const taskOrderIssue = input.detectTaskOrderIssue(model.tasks);
      if (taskOrderIssue) {
        issues.push({
          level: "warning",
          scope: "tasks",
          message: `Task の並び順が OutlineNumber 順と一致していない可能性があります: ${input.describeTask(taskOrderIssue.current)} (直前: ${input.describeTask(taskOrderIssue.previous)})`
        });
      }

      for (const resource of model.resources) {
        if (!resource.uid) {
          issues.push({ level: "error", scope: "resources", message: "Resource UID が空です" });
        }
        if (!resource.name) {
          if (!input.isPlaceholderUid(resource.uid)) {
            issues.push({ level: "warning", scope: "resources", message: `Resource Name が空です: ${input.describeResource(resource)}` });
          }
        }
        if (resourceUidSet.has(resource.uid)) {
          issues.push({ level: "error", scope: "resources", message: `Resource UID が重複しています: ${resource.uid}` });
        }
        resourceUidSet.add(resource.uid);
        if (resource.calendarUID && !calendarUidSet.has(resource.calendarUID)) {
          issues.push({
            level: "warning",
            scope: "resources",
            message: `Resource CalendarUID が既存 Calendar を指していません: ${input.describeResource(resource)}`
          });
        }
        if (resource.workGroup !== undefined && resource.workGroup < 0) {
          issues.push({ level: "warning", scope: "resources", message: `Resource WorkGroup は 0 以上が望ましいです: ${input.describeResource(resource)}` });
        }
        if (resource.overtimeRateFormat !== undefined && resource.overtimeRateFormat < 0) {
          issues.push({ level: "warning", scope: "resources", message: `Resource OvertimeRateFormat は 0 以上が望ましいです: ${input.describeResource(resource)}` });
        }
        if (resource.cost !== undefined && resource.cost < 0) {
          issues.push({ level: "warning", scope: "resources", message: `Resource Cost が負値です: ${input.describeResource(resource)}` });
        }
        if (resource.actualCost !== undefined && resource.actualCost < 0) {
          issues.push({ level: "warning", scope: "resources", message: `Resource ActualCost が負値です: ${input.describeResource(resource)}` });
        }
        if (resource.remainingCost !== undefined && resource.remainingCost < 0) {
          issues.push({ level: "warning", scope: "resources", message: `Resource RemainingCost が負値です: ${input.describeResource(resource)}` });
        }
        if (resource.percentWorkComplete !== undefined && (resource.percentWorkComplete < 0 || resource.percentWorkComplete > 100)) {
          issues.push({ level: "warning", scope: "resources", message: `Resource PercentWorkComplete が 0..100 の範囲外です: ${input.describeResource(resource)}` });
        }
        for (const attribute of resource.extendedAttributes) {
          if (!attribute.fieldID) {
            issues.push({ level: "warning", scope: "resources", message: `Resource ExtendedAttribute に FieldID がありません: ${input.describeResource(resource)}` });
          }
        }
        for (const baseline of resource.baselines) {
          if (baseline.number !== undefined && baseline.number < 0) {
            issues.push({ level: "warning", scope: "resources", message: `Resource Baseline Number は 0 以上が望ましいです: ${input.describeResource(resource)}` });
          }
          const baselineStart = input.parseDateValue(baseline.start);
          const baselineFinish = input.parseDateValue(baseline.finish);
          if (baselineStart !== null && baselineFinish !== null && baselineStart > baselineFinish) {
            issues.push({ level: "warning", scope: "resources", message: `Resource Baseline Start が Finish より後です: ${input.describeResource(resource)}` });
          }
        }
        for (const timephasedData of resource.timephasedData) {
          if (timephasedData.type !== undefined && timephasedData.type < 0) {
            issues.push({ level: "warning", scope: "resources", message: `Resource TimephasedData Type は 0 以上が望ましいです: ${input.describeResource(resource)}` });
          }
          const timephasedStart = input.parseDateValue(timephasedData.start);
          const timephasedFinish = input.parseDateValue(timephasedData.finish);
          if (timephasedStart !== null && timephasedFinish !== null && timephasedStart > timephasedFinish) {
            issues.push({ level: "warning", scope: "resources", message: `Resource TimephasedData Start が Finish より後です: ${input.describeResource(resource)}` });
          }
        }
      }

      for (const task of model.tasks) {
        for (const predecessor of task.predecessors) {
          if (!taskUidSet.has(predecessor.predecessorUid)) {
            issues.push({
              level: "error",
              scope: "tasks",
              message: `PredecessorUID が既存 Task を指していません: ${input.describeTask(task)}, ${input.describeTaskRef(model, predecessor.predecessorUid)}`
            });
          }
        }
      }

      for (const assignment of model.assignments) {
        if (!assignment.uid) {
          issues.push({ level: "warning", scope: "assignments", message: "Assignment UID が空です" });
        }
        if (!taskUidSet.has(assignment.taskUid)) {
          issues.push({
            level: "error",
            scope: "assignments",
            message: `Assignment TaskUID が既存 Task を指していません: ${input.describeAssignment(assignment)}, ${input.describeTaskRef(model, assignment.taskUid)}`
          });
        }
        if (!resourceUidSet.has(assignment.resourceUid) && !input.isUnassignedResourceUid(assignment.resourceUid)) {
          issues.push({
            level: "error",
            scope: "assignments",
            message: `Assignment ResourceUID が既存 Resource を指していません: ${input.describeAssignment(assignment)}, ${input.describeTaskRef(model, assignment.taskUid)}, ${input.describeResourceRef(model, assignment.resourceUid)}`
          });
        }
        if (!assignment.start) {
          issues.push({ level: "warning", scope: "assignments", message: `Assignment Start が空です: ${input.describeAssignment(assignment)}` });
        }
        if (!assignment.finish) {
          issues.push({ level: "warning", scope: "assignments", message: `Assignment Finish が空です: ${input.describeAssignment(assignment)}` });
        }
        const assignmentStart = input.parseDateValue(assignment.start);
        const assignmentFinish = input.parseDateValue(assignment.finish);
        if (assignmentStart !== null && assignmentFinish !== null && assignmentStart > assignmentFinish) {
          issues.push({ level: "warning", scope: "assignments", message: `Assignment Start が Finish より後です: ${input.describeAssignment(assignment)}` });
        }
        if (assignment.units !== undefined && assignment.units < 0) {
          issues.push({ level: "warning", scope: "assignments", message: `Assignment Units が負値です: ${input.describeAssignment(assignment)}` });
        }
        if (assignment.cost !== undefined && assignment.cost < 0) {
          issues.push({ level: "warning", scope: "assignments", message: `Assignment Cost が負値です: ${input.describeAssignment(assignment)}` });
        }
        if (assignment.actualCost !== undefined && assignment.actualCost < 0) {
          issues.push({ level: "warning", scope: "assignments", message: `Assignment ActualCost が負値です: ${input.describeAssignment(assignment)}` });
        }
        if (assignment.remainingCost !== undefined && assignment.remainingCost < 0) {
          issues.push({ level: "warning", scope: "assignments", message: `Assignment RemainingCost が負値です: ${input.describeAssignment(assignment)}` });
        }
        if (assignment.percentWorkComplete !== undefined && (assignment.percentWorkComplete < 0 || assignment.percentWorkComplete > 100)) {
          issues.push({ level: "warning", scope: "assignments", message: `Assignment PercentWorkComplete が 0..100 の範囲外です: ${input.describeAssignment(assignment)}` });
        }
        if (assignment.overtimeWork !== undefined && !assignment.overtimeWork) {
          issues.push({ level: "warning", scope: "assignments", message: `Assignment OvertimeWork が空です: ${input.describeAssignment(assignment)}` });
        }
        if (assignment.actualOvertimeWork !== undefined && !assignment.actualOvertimeWork) {
          issues.push({ level: "warning", scope: "assignments", message: `Assignment ActualOvertimeWork が空です: ${input.describeAssignment(assignment)}` });
        }
        if (assignment.workContour !== undefined && assignment.workContour < 0) {
          issues.push({ level: "warning", scope: "assignments", message: `Assignment WorkContour は 0 以上が望ましいです: ${input.describeAssignment(assignment)}` });
        }
        if (assignment.startVariance !== undefined && !assignment.startVariance) {
          issues.push({ level: "warning", scope: "assignments", message: `Assignment StartVariance が空です: ${input.describeAssignment(assignment)}` });
        }
        for (const attribute of assignment.extendedAttributes) {
          if (!attribute.fieldID) {
            issues.push({ level: "warning", scope: "assignments", message: `Assignment ExtendedAttribute に FieldID がありません: ${input.describeAssignment(assignment)}` });
          }
        }
        for (const baseline of assignment.baselines) {
          if (baseline.number !== undefined && baseline.number < 0) {
            issues.push({ level: "warning", scope: "assignments", message: `Assignment Baseline Number は 0 以上が望ましいです: ${input.describeAssignment(assignment)}` });
          }
          const baselineStart = input.parseDateValue(baseline.start);
          const baselineFinish = input.parseDateValue(baseline.finish);
          if (baselineStart !== null && baselineFinish !== null && baselineStart > baselineFinish) {
            issues.push({ level: "warning", scope: "assignments", message: `Assignment Baseline Start が Finish より後です: ${input.describeAssignment(assignment)}` });
          }
        }
        for (const timephasedData of assignment.timephasedData) {
          if (timephasedData.type !== undefined && timephasedData.type < 0) {
            issues.push({ level: "warning", scope: "assignments", message: `Assignment TimephasedData Type は 0 以上が望ましいです: ${input.describeAssignment(assignment)}` });
          }
          const timephasedStart = input.parseDateValue(timephasedData.start);
          const timephasedFinish = input.parseDateValue(timephasedData.finish);
          if (timephasedStart !== null && timephasedFinish !== null && timephasedStart > timephasedFinish) {
            issues.push({ level: "warning", scope: "assignments", message: `Assignment TimephasedData Start が Finish より後です: ${input.describeAssignment(assignment)}` });
          }
        }
        if (assignment.finishVariance !== undefined && !assignment.finishVariance) {
          issues.push({ level: "warning", scope: "assignments", message: `Assignment FinishVariance が空です: ${input.describeAssignment(assignment)}` });
        }
      }

      return issues;
    }
  };
})();
