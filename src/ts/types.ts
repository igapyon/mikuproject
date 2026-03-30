/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  type ProjectInfo = {
    name: string;
    title?: string;
    author?: string;
    company?: string;
    creationDate?: string;
    lastSaved?: string;
    saveVersion?: number;
    startDate: string;
    finishDate: string;
    scheduleFromStart: boolean;
    currentDate?: string;
    defaultStartTime?: string;
    defaultFinishTime?: string;
    minutesPerDay?: number;
    minutesPerWeek?: number;
    daysPerMonth?: number;
    statusDate?: string;
    weekStartDay?: number;
    workFormat?: number;
    durationFormat?: number;
    currencyCode?: string;
    currencyDigits?: number;
    currencySymbol?: string;
    currencySymbolPosition?: number;
    fyStartDate?: string;
    fiscalYearStart?: boolean;
    criticalSlackLimit?: number;
    defaultTaskType?: number;
    defaultFixedCostAccrual?: number;
    defaultStandardRate?: string;
    defaultOvertimeRate?: string;
    defaultTaskEVMethod?: number;
    newTaskStartDate?: number;
    newTasksAreManual?: boolean;
    newTasksEffortDriven?: boolean;
    newTasksEstimated?: boolean;
    actualsInSync?: boolean;
    editableActualCosts?: boolean;
    honorConstraints?: boolean;
    insertedProjectsLikeSummary?: boolean;
    multipleCriticalPaths?: boolean;
    taskUpdatesResource?: boolean;
    updateManuallyScheduledTasksWhenEditingLinks?: boolean;
    calendarUID?: string;
    outlineCodes: OutlineCodeModel[];
    wbsMasks: WBSMaskModel[];
    extendedAttributes: ProjectExtendedAttributeModel[];
  };

  type OutlineCodeValueModel = {
    value: string;
    description?: string;
  };

  type OutlineCodeMaskModel = {
    level: number;
    mask?: string;
    length?: number;
    sequence?: number;
  };

  type OutlineCodeModel = {
    fieldID?: string;
    fieldName?: string;
    alias?: string;
    onlyTableValues?: boolean;
    enterprise?: boolean;
    resourceSubstitutionEnabled?: boolean;
    leafOnly?: boolean;
    allLevelsRequired?: boolean;
    masks: OutlineCodeMaskModel[];
    values: OutlineCodeValueModel[];
  };

  type WBSMaskModel = {
    level: number;
    mask?: string;
    length?: number;
    sequence?: number;
  };

  type ProjectExtendedAttributeModel = {
    fieldID?: string;
    fieldName?: string;
    alias?: string;
    calculationType?: number;
    restrictValues?: boolean;
    appendNewValues?: boolean;
  };

  type PredecessorModel = {
    predecessorUid: string;
    type?: number;
    linkLag?: string;
  };

  type TaskExtendedAttributeModel = {
    fieldID?: string;
    value?: string;
  };

  type TaskBaselineModel = {
    number?: number;
    start?: string;
    finish?: string;
    work?: string;
    cost?: number;
  };

  type TaskTimephasedDataModel = {
    type?: number;
    uid?: string;
    start?: string;
    finish?: string;
    unit?: number;
    value?: string;
  };

  type ResourceExtendedAttributeModel = {
    fieldID?: string;
    value?: string;
  };

  type ResourceBaselineModel = {
    number?: number;
    start?: string;
    finish?: string;
    work?: string;
    cost?: number;
  };

  type ResourceTimephasedDataModel = {
    type?: number;
    uid?: string;
    start?: string;
    finish?: string;
    unit?: number;
    value?: string;
  };

  type AssignmentExtendedAttributeModel = {
    fieldID?: string;
    value?: string;
  };

  type AssignmentBaselineModel = {
    number?: number;
    start?: string;
    finish?: string;
    work?: string;
    cost?: number;
  };

  type AssignmentTimephasedDataModel = {
    type?: number;
    uid?: string;
    start?: string;
    finish?: string;
    unit?: number;
    value?: string;
  };

  type TaskModel = {
    uid: string;
    id: string;
    name: string;
    outlineLevel: number;
    outlineNumber: string;
    wbs?: string;
    type?: number;
    calendarUID?: string;
    priority?: number;
    start: string;
    finish: string;
    duration: string;
    actualStart?: string;
    actualFinish?: string;
    deadline?: string;
    startVariance?: string;
    finishVariance?: string;
    work?: string;
    workVariance?: string;
    totalSlack?: string;
    freeSlack?: string;
    cost?: number;
    actualCost?: number;
    remainingCost?: number;
    remainingWork?: string;
    actualWork?: string;
    milestone: boolean;
    summary: boolean;
    critical?: boolean;
    percentComplete: number;
    percentWorkComplete?: number;
    notes?: string;
    constraintType?: number;
    constraintDate?: string;
    extendedAttributes: TaskExtendedAttributeModel[];
    baselines: TaskBaselineModel[];
    timephasedData: TaskTimephasedDataModel[];
    predecessors: PredecessorModel[];
  };

  type ResourceModel = {
    uid: string;
    id: string;
    name: string;
    type?: number;
    initials?: string;
    group?: string;
    workGroup?: number;
    maxUnits?: number;
    calendarUID?: string;
    standardRate?: string;
    standardRateFormat?: number;
    overtimeRate?: string;
    overtimeRateFormat?: number;
    costPerUse?: number;
    work?: string;
    actualWork?: string;
    remainingWork?: string;
    cost?: number;
    actualCost?: number;
    remainingCost?: number;
    percentWorkComplete?: number;
    extendedAttributes: ResourceExtendedAttributeModel[];
    baselines: ResourceBaselineModel[];
    timephasedData: ResourceTimephasedDataModel[];
  };

  type AssignmentModel = {
    uid: string;
    taskUid: string;
    resourceUid: string;
    start?: string;
    finish?: string;
    startVariance?: string;
    finishVariance?: string;
    delay?: string;
    milestone?: boolean;
    workContour?: number;
    units?: number;
    work?: string;
    cost?: number;
    actualCost?: number;
    remainingCost?: number;
    percentWorkComplete?: number;
    overtimeWork?: string;
    actualOvertimeWork?: string;
    actualWork?: string;
    remainingWork?: string;
    extendedAttributes: AssignmentExtendedAttributeModel[];
    baselines: AssignmentBaselineModel[];
    timephasedData: AssignmentTimephasedDataModel[];
  };

  type WorkingTimeModel = {
    fromTime: string;
    toTime: string;
  };

  type WeekDayModel = {
    dayType: number;
    dayWorking: boolean;
    workingTimes: WorkingTimeModel[];
  };

  type CalendarExceptionModel = {
    name?: string;
    fromDate?: string;
    toDate?: string;
    dayWorking?: boolean;
    workingTimes: WorkingTimeModel[];
  };

  type WorkWeekModel = {
    name?: string;
    fromDate?: string;
    toDate?: string;
    weekDays: WeekDayModel[];
  };

  type CalendarModel = {
    uid: string;
    name: string;
    isBaseCalendar: boolean;
    isBaselineCalendar?: boolean;
    baseCalendarUID?: string;
    weekDays: WeekDayModel[];
    exceptions: CalendarExceptionModel[];
    workWeeks: WorkWeekModel[];
  };

  type ProjectModel = {
    project: ProjectInfo;
    tasks: TaskModel[];
    resources: ResourceModel[];
    assignments: AssignmentModel[];
    calendars: CalendarModel[];
  };

  type ValidationIssue = {
    level: "error" | "warning";
    scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
    message: string;
  };

  (globalThis as typeof globalThis & {
    __mikuprojectTypes?: {
      __ready: true;
    };
  }).__mikuprojectTypes = {
    __ready: true
  };
})();
