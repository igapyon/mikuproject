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
    ];
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
    ];
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
    ];
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
    };
    const IMPORTABLE_FIELDS = {
        Project: PROJECT_EDITABLE_FIELDS,
        Tasks: ["Name", "Start", "Finish", "PercentComplete", "PercentWorkComplete", "Notes"],
        Resources: ["Name", "Group", "MaxUnits"],
        Assignments: ["Units", "Work", "PercentWorkComplete"],
        Calendars: ["Name", "IsBaseCalendar", "BaseCalendarUID"],
        NonWorkingDays: ["Name", "Date", "FromDate", "ToDate", "DayWorking"]
    };
    globalThis.__mikuprojectProjectWorkbookSchema = {
        SHEET_NAMES,
        HEADER_ROW_INDEX,
        DATA_ROW_START_INDEX,
        PROJECT_FIELD_ORDER,
        PROJECT_EDITABLE_FIELDS,
        SHEET_HEADERS,
        IMPORTABLE_FIELDS
    };
})();
