/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  const SHEET_NAMES = [
    "Project",
    "Tasks",
    "Resources",
    "Assignments",
    "Calendars",
    "NonWorkingDays"
  ] as const;

  const HEADER_ROW_INDEX = 2;
  const DATA_ROW_START_INDEX = 3;

  const PROJECT_FIELD_ORDER = [
    "Name",
    "Title",
    "Author",
    "Company",
    "StartDate",
    "FinishDate",
    "CurrentDate",
    "StatusDate",
    "CalendarUID",
    "MinutesPerDay",
    "MinutesPerWeek",
    "DaysPerMonth",
    "ScheduleFromStart",
    "OutlineCodes",
    "WBSMasks",
    "ExtendedAttributes"
  ] as const;

  const PROJECT_EDITABLE_FIELDS = [
    "Name",
    "Title",
    "Author",
    "Company",
    "StartDate",
    "FinishDate",
    "CurrentDate",
    "StatusDate",
    "CalendarUID",
    "MinutesPerDay",
    "MinutesPerWeek",
    "DaysPerMonth",
    "ScheduleFromStart"
  ] as const;

  const SHEET_HEADERS = {
    Tasks: [
      "UID", "ID", "Name", "OutlineLevel", "OutlineNumber", "WBS",
      "Start", "Finish", "Duration", "PercentComplete", "PercentWorkComplete",
      "Milestone", "Summary", "Critical", "CalendarUID", "Predecessors", "Notes"
    ],
    Resources: [
      "UID", "ID", "Name", "Type", "Initials", "Group", "MaxUnits",
      "CalendarUID", "StandardRate", "OvertimeRate", "CostPerUse",
      "Work", "ActualWork", "RemainingWork"
    ],
    Assignments: [
      "UID", "TaskUID", "TaskName", "ResourceUID", "ResourceName", "Start",
      "Finish", "Units", "Work", "ActualWork", "RemainingWork", "PercentWorkComplete"
    ],
    Calendars: [
      "UID", "Name", "IsBaseCalendar", "BaseCalendarUID", "WeekDays", "Exceptions", "WorkWeeks"
    ],
    NonWorkingDays: [
      "CalendarUID", "Index", "CalendarName", "Name", "Date", "FromDate", "ToDate", "DayWorking"
    ]
  } as const;

  const IMPORTABLE_FIELDS = {
    Project: PROJECT_EDITABLE_FIELDS,
    Tasks: ["Name", "Start", "Finish", "Duration", "PercentComplete", "PercentWorkComplete", "Milestone", "Summary", "Critical", "CalendarUID", "Predecessors", "Notes"] as const,
    Resources: ["Name", "Group", "MaxUnits", "CalendarUID"] as const,
    Assignments: ["Units", "Work", "PercentWorkComplete"] as const,
    Calendars: ["Name", "IsBaseCalendar", "BaseCalendarUID"] as const,
    NonWorkingDays: ["Name", "Date", "FromDate", "ToDate", "DayWorking"] as const
  } as const;

  (globalThis as typeof globalThis & {
    __mikuprojectProjectWorkbookSchema?: {
      SHEET_NAMES: typeof SHEET_NAMES;
      HEADER_ROW_INDEX: typeof HEADER_ROW_INDEX;
      DATA_ROW_START_INDEX: typeof DATA_ROW_START_INDEX;
      PROJECT_FIELD_ORDER: typeof PROJECT_FIELD_ORDER;
      PROJECT_EDITABLE_FIELDS: typeof PROJECT_EDITABLE_FIELDS;
      SHEET_HEADERS: typeof SHEET_HEADERS;
      IMPORTABLE_FIELDS: typeof IMPORTABLE_FIELDS;
    };
  }).__mikuprojectProjectWorkbookSchema = {
    SHEET_NAMES,
    HEADER_ROW_INDEX,
    DATA_ROW_START_INDEX,
    PROJECT_FIELD_ORDER,
    PROJECT_EDITABLE_FIELDS,
    SHEET_HEADERS,
    IMPORTABLE_FIELDS
  };
})();
