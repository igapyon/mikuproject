/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const HEADER_FILL = "#D9EAF7";
    const SECTION_FILL = "#BFD7EA";
    const LABEL_FILL = "#EDF5FB";
    const ALT_ROW_FILL = "#F9FBFD";
    const DATE_FILL = "#FFF4E8";
    const PERCENT_FILL = "#FCECF3";
    const REFERENCE_FILL = "#EEF7F4";
    const COUNT_FILL = "#F2F5F8";
    const EDITABLE_FILL = "#FDE7C7";
    const DURATION_FILL = "#FBF6ED";
    const NOTES_FILL = "#FFFBEA";
    const NAME_FILL = "#FAF6FF";
    const WORK_FILL = "#F1F8FD";
    const SUMMARY_FILL = "#E6F2E0";
    const MILESTONE_FILL = "#FFF0CF";
    const CRITICAL_FILL = "#F8DDE6";
    const SHEET_THEMES = {
        project: { section: "#BFD7EA", header: "#D9EAF7", label: "#EDF5FB" },
        tasks: { section: "#D4E0EC", header: "#E6EDF4", label: "#F2F6FA" },
        resources: { section: "#C8E3D8", header: "#DDF0E8", label: "#EFF8F4" },
        assignments: { section: "#D7D2EC", header: "#E7E3F5", label: "#F2F0FA" },
        calendars: { section: "#D7E3C4", header: "#E7F0DA", label: "#F2F7EA" },
        nonWorkingDays: { section: "#E9C7D5", header: "#F4DDE6", label: "#FBEEF3" }
    };
    function exportProjectWorkbook(model) {
        return {
            sheets: [
                buildProjectSheet(model),
                buildTasksSheet(model),
                buildResourcesSheet(model),
                buildAssignmentsSheet(model),
                buildCalendarsSheet(model),
                buildNonWorkingDaysSheet(model)
            ]
        };
    }
    function importProjectWorkbook(workbook, baseModel) {
        return importProjectWorkbookDetailed(workbook, baseModel).model;
    }
    function importProjectWorkbookDetailed(workbook, baseModel) {
        const nextModel = cloneProjectModel(baseModel);
        const changes = [];
        importProjectSheet(workbook, nextModel, changes);
        importTasksSheet(workbook, nextModel, changes);
        importResourcesSheet(workbook, nextModel, changes);
        importAssignmentsSheet(workbook, nextModel, changes);
        importCalendarsSheet(workbook, nextModel, changes);
        importNonWorkingDaysSheet(workbook, nextModel, changes);
        return {
            model: nextModel,
            changes
        };
    }
    function importProjectSheet(workbook, model, changes) {
        const projectSheet = workbook.sheets.find((sheet) => sheet.name === "Project");
        if (!projectSheet) {
            return;
        }
        const valueByField = new Map();
        for (const row of projectSheet.rows.slice(3)) {
            const field = readStringCell(row.cells[0]);
            if (!field) {
                continue;
            }
            valueByField.set(field, row.cells[1]);
        }
        const projectLabel = model.project.name;
        assignIfChanged(changes, "project", "project", projectLabel, model.project, "name", "Name", readStringCell(valueByField.get("Name")));
        assignIfChanged(changes, "project", "project", projectLabel, model.project, "title", "Title", readStringCell(valueByField.get("Title")));
        assignIfChanged(changes, "project", "project", projectLabel, model.project, "author", "Author", readStringCell(valueByField.get("Author")));
        assignIfChanged(changes, "project", "project", projectLabel, model.project, "company", "Company", readStringCell(valueByField.get("Company")));
        assignIfChanged(changes, "project", "project", projectLabel, model.project, "startDate", "StartDate", normalizeDateTimeInput(readStringCell(valueByField.get("StartDate"))));
        assignIfChanged(changes, "project", "project", projectLabel, model.project, "finishDate", "FinishDate", normalizeDateTimeInput(readStringCell(valueByField.get("FinishDate"))));
        assignIfChanged(changes, "project", "project", projectLabel, model.project, "currentDate", "CurrentDate", normalizeDateTimeInput(readStringCell(valueByField.get("CurrentDate"))));
        assignIfChanged(changes, "project", "project", projectLabel, model.project, "statusDate", "StatusDate", normalizeDateTimeInput(readStringCell(valueByField.get("StatusDate"))));
        assignIfChanged(changes, "project", "project", projectLabel, model.project, "calendarUID", "CalendarUID", readStringCell(valueByField.get("CalendarUID")));
        assignIfChanged(changes, "project", "project", projectLabel, model.project, "minutesPerDay", "MinutesPerDay", readNumberCell(valueByField.get("MinutesPerDay")));
        assignIfChanged(changes, "project", "project", projectLabel, model.project, "minutesPerWeek", "MinutesPerWeek", readNumberCell(valueByField.get("MinutesPerWeek")));
        assignIfChanged(changes, "project", "project", projectLabel, model.project, "daysPerMonth", "DaysPerMonth", readNumberCell(valueByField.get("DaysPerMonth")));
        assignIfChanged(changes, "project", "project", projectLabel, model.project, "scheduleFromStart", "ScheduleFromStart", readBooleanCell(valueByField.get("ScheduleFromStart")));
    }
    function buildProjectSheet(model) {
        const project = model.project;
        const rows = [
            titleRow("Project", SHEET_THEMES.project.section),
            titleRow("Basic Info", SHEET_THEMES.project.section),
            headerRow(["Field", "Value"], SHEET_THEMES.project.header),
            keyValueRow("Name", project.name, SHEET_THEMES.project.label),
            keyValueRow("Title", project.title, SHEET_THEMES.project.label),
            keyValueRow("Author", project.author, SHEET_THEMES.project.label),
            keyValueRow("Company", project.company, SHEET_THEMES.project.label),
            keyValueRow("StartDate", project.startDate, SHEET_THEMES.project.label),
            keyValueRow("FinishDate", project.finishDate, SHEET_THEMES.project.label),
            keyValueRow("CurrentDate", project.currentDate, SHEET_THEMES.project.label),
            keyValueRow("StatusDate", project.statusDate, SHEET_THEMES.project.label),
            titleRow("Settings", SHEET_THEMES.project.section),
            keyValueRow("CalendarUID", project.calendarUID, SHEET_THEMES.project.label),
            keyValueRow("MinutesPerDay", project.minutesPerDay, SHEET_THEMES.project.label),
            keyValueRow("MinutesPerWeek", project.minutesPerWeek, SHEET_THEMES.project.label),
            keyValueRow("DaysPerMonth", project.daysPerMonth, SHEET_THEMES.project.label),
            keyValueRow("ScheduleFromStart", project.scheduleFromStart, SHEET_THEMES.project.label),
            keyValueRow("OutlineCodes", project.outlineCodes.length, SHEET_THEMES.project.label),
            keyValueRow("WBSMasks", project.wbsMasks.length, SHEET_THEMES.project.label),
            keyValueRow("ExtendedAttributes", project.extendedAttributes.length, SHEET_THEMES.project.label)
        ];
        return {
            name: "Project",
            columns: [{ width: 26 }, { width: 42 }],
            mergedRanges: ["A11:B11"],
            rows
        };
    }
    function buildTasksSheet(model) {
        return {
            name: "Tasks",
            columns: [
                { width: 10 }, { width: 8 }, { width: 28 }, { width: 12 },
                { width: 14 }, { width: 12 }, { width: 20 }, { width: 20 },
                { width: 14 }, { width: 16 }, { width: 18 }, { width: 12 },
                { width: 12 }, { width: 12 }, { width: 12 }, { width: 20 },
                { width: 34 }
            ],
            mergedRanges: [],
            rows: [
                sectionTitleRow("Tasks", 17, SHEET_THEMES.tasks.section),
                sectionTitleRow("Task List", 17, SHEET_THEMES.tasks.section),
                headerRow([
                    "UID", "ID", "Name", "OutlineLevel", "OutlineNumber", "WBS",
                    "Start", "Finish", "Duration", "PercentComplete", "PercentWorkComplete",
                    "Milestone", "Summary", "Critical", "CalendarUID", "Predecessors", "Notes"
                ], SHEET_THEMES.tasks.header),
                ...model.tasks.map((task, index) => ({
                    cells: [
                        countCell(task.uid, index),
                        countCell(task.id, index),
                        editableCell(taskNameCell(task, index)),
                        countCell(task.outlineLevel, index),
                        textCell(task.outlineNumber, index),
                        textCell(task.wbs, index),
                        editableCell(dateTimeCell(task.start, index)),
                        editableCell(dateTimeCell(task.finish, index)),
                        durationCell(task.duration, index),
                        editableCell(percentCell(task.percentComplete, index)),
                        editableCell(percentCell(task.percentWorkComplete, index)),
                        taskFlagCell(task.milestone, index, MILESTONE_FILL),
                        taskFlagCell(task.summary, index, SUMMARY_FILL),
                        taskFlagCell(task.critical, index, CRITICAL_FILL),
                        referenceCell(task.calendarUID, index),
                        predecessorCell(task.predecessors.map((item) => item.predecessorUid).join(", "), index),
                        editableCell(notesCell(task.notes, index))
                    ]
                }))
            ]
        };
    }
    function buildResourcesSheet(model) {
        return {
            name: "Resources",
            columns: [
                { width: 10 }, { width: 8 }, { width: 24 }, { width: 10 },
                { width: 12 }, { width: 18 }, { width: 12 }, { width: 12 },
                { width: 14 }, { width: 14 }, { width: 12 }, { width: 14 },
                { width: 14 }, { width: 14 }
            ],
            mergedRanges: [],
            rows: [
                sectionTitleRow("Resources", 14, SHEET_THEMES.resources.section),
                sectionTitleRow("Resource List", 14, SHEET_THEMES.resources.section),
                headerRow([
                    "UID", "ID", "Name", "Type", "Initials", "Group", "MaxUnits",
                    "CalendarUID", "StandardRate", "OvertimeRate", "CostPerUse",
                    "Work", "ActualWork", "RemainingWork"
                ], SHEET_THEMES.resources.header),
                ...model.resources.map((resource, index) => ({
                    cells: [
                        countCell(resource.uid, index),
                        countCell(resource.id, index),
                        editableCell(entityNameCell(resource.name, index)),
                        countCell(resource.type, index),
                        textCell(resource.initials, index),
                        editableCell(textCell(resource.group, index)),
                        editableCell(countCell(resource.maxUnits, index)),
                        referenceCell(resource.calendarUID, index),
                        textCell(resource.standardRate, index),
                        textCell(resource.overtimeRate, index),
                        countCell(resource.costPerUse, index),
                        workCell(resource.work, index),
                        workCell(resource.actualWork, index),
                        workCell(resource.remainingWork, index)
                    ]
                }))
            ]
        };
    }
    function buildAssignmentsSheet(model) {
        const taskNameByUid = new Map(model.tasks.map((task) => [task.uid, task.name]));
        const resourceNameByUid = new Map(model.resources.map((resource) => [resource.uid, resource.name]));
        return {
            name: "Assignments",
            columns: [
                { width: 10 }, { width: 10 }, { width: 24 }, { width: 12 },
                { width: 24 }, { width: 20 }, { width: 20 }, { width: 10 },
                { width: 14 }, { width: 14 }, { width: 14 }, { width: 18 }
            ],
            mergedRanges: [],
            rows: [
                sectionTitleRow("Assignments", 12, SHEET_THEMES.assignments.section),
                sectionTitleRow("Assignment List", 12, SHEET_THEMES.assignments.section),
                headerRow([
                    "UID", "TaskUID", "TaskName", "ResourceUID", "ResourceName", "Start",
                    "Finish", "Units", "Work", "ActualWork", "RemainingWork", "PercentWorkComplete"
                ], SHEET_THEMES.assignments.header),
                ...model.assignments.map((assignment, index) => ({
                    cells: [
                        countCell(assignment.uid, index),
                        referenceCell(assignment.taskUid, index),
                        entityNameCell(taskNameByUid.get(assignment.taskUid), index),
                        referenceCell(assignment.resourceUid, index),
                        entityNameCell(resourceNameByUid.get(assignment.resourceUid), index),
                        dateTimeCell(assignment.start, index),
                        dateTimeCell(assignment.finish, index),
                        editableCell(countCell(assignment.units, index)),
                        editableCell(workCell(assignment.work, index)),
                        workCell(assignment.actualWork, index),
                        workCell(assignment.remainingWork, index),
                        editableCell(percentCell(assignment.percentWorkComplete, index))
                    ]
                }))
            ]
        };
    }
    function buildCalendarsSheet(model) {
        return {
            name: "Calendars",
            columns: [
                { width: 10 }, { width: 24 }, { width: 14 }, { width: 16 },
                { width: 10 }, { width: 12 }, { width: 10 }
            ],
            mergedRanges: [],
            rows: [
                sectionTitleRow("Calendars", 7, SHEET_THEMES.calendars.section),
                sectionTitleRow("Calendar List", 7, SHEET_THEMES.calendars.section),
                headerRow([
                    "UID", "Name", "IsBaseCalendar", "BaseCalendarUID", "WeekDays", "Exceptions", "WorkWeeks"
                ], SHEET_THEMES.calendars.header),
                ...model.calendars.map((calendar, index) => ({
                    cells: [
                        countCell(calendar.uid, index),
                        editableCell(entityNameCell(calendar.name, index)),
                        editableCell(countCell(calendar.isBaseCalendar, index)),
                        editableCell(referenceCell(calendar.baseCalendarUID, index)),
                        countCell(calendar.weekDays.length, index),
                        countCell(calendar.exceptions.length, index),
                        countCell(calendar.workWeeks.length, index)
                    ]
                }))
            ]
        };
    }
    function buildNonWorkingDaysSheet(model) {
        const rows = model.calendars.flatMap((calendar) => calendar.exceptions.map((exception, index) => ({
            cells: [
                countCell(calendar.uid, index),
                countCell(index + 1, index),
                textCell(calendar.name, index),
                editableCell(entityNameCell(exception.name, index)),
                editableCell(dateOnlyCell(formatExceptionDate(exception), index)),
                editableCell(dateOnlyCell(formatExceptionBoundaryDate(exception.fromDate), index)),
                editableCell(dateOnlyCell(formatExceptionBoundaryDate(exception.toDate), index)),
                editableCell(countCell(exception.dayWorking, index))
            ]
        })));
        return {
            name: "NonWorkingDays",
            columns: [
                { width: 12 }, { width: 10 }, { width: 22 }, { width: 24 },
                { width: 14 }, { width: 22 }, { width: 22 }, { width: 12 }
            ],
            mergedRanges: [],
            rows: [
                sectionTitleRow("NonWorkingDays", 8, SHEET_THEMES.nonWorkingDays.section),
                sectionTitleRow("Calendar Exceptions", 8, SHEET_THEMES.nonWorkingDays.section),
                headerRow([
                    "CalendarUID", "Index", "CalendarName", "Name", "Date", "FromDate", "ToDate", "DayWorking"
                ], SHEET_THEMES.nonWorkingDays.header),
                ...rows
            ]
        };
    }
    function headerRow(labels, fillColor = HEADER_FILL) {
        return {
            height: 24,
            cells: labels.map((label) => ({
                value: label,
                bold: true,
                fillColor,
                border: "thin",
                horizontalAlign: "center"
            }))
        };
    }
    function titleRow(title, fillColor = SECTION_FILL) {
        return {
            height: 28,
            cells: [
                {
                    value: title,
                    bold: true,
                    fontSize: 16,
                    fillColor,
                    horizontalAlign: "left"
                },
                {
                    fillColor
                }
            ]
        };
    }
    function sectionTitleRow(title, columnCount, fillColor = SECTION_FILL) {
        return {
            height: 26,
            cells: [
                {
                    value: title,
                    bold: true,
                    fontSize: 14,
                    fillColor,
                    horizontalAlign: "left"
                },
                ...Array.from({ length: Math.max(0, columnCount - 1) }, () => ({
                    fillColor
                }))
            ]
        };
    }
    function keyValueRow(label, value, labelFill = LABEL_FILL) {
        return {
            cells: [
                {
                    value: label,
                    bold: true,
                    fillColor: labelFill,
                    border: "thin"
                },
                keyValueCell(label, value)
            ]
        };
    }
    function cell(value) {
        if (value === undefined) {
            return {};
        }
        return {
            value,
            border: "thin"
        };
    }
    function keyValueCell(label, value) {
        if (isEditableProjectLabel(label)) {
            return editableCell(buildProjectValueCell(label, value));
        }
        return buildProjectValueCell(label, value);
    }
    function buildProjectValueCell(label, value) {
        if (isDateTimeLabel(label)) {
            return {
                ...cell(formatDateTimeDisplay(value)),
                fillColor: DATE_FILL
            };
        }
        if (label === "Name" || label === "Title") {
            return {
                ...cell(value),
                fillColor: NAME_FILL,
                bold: true
            };
        }
        if (label === "Author" || label === "Company") {
            return {
                ...cell(value),
                fillColor: NAME_FILL
            };
        }
        if (label === "CalendarUID") {
            return {
                ...cell(value),
                fillColor: REFERENCE_FILL,
                horizontalAlign: "center"
            };
        }
        if (label === "ScheduleFromStart") {
            return {
                ...cell(value),
                fillColor: COUNT_FILL,
                horizontalAlign: "center"
            };
        }
        return {
            ...cell(value),
            fillColor: isNumericSummaryLabel(label) ? COUNT_FILL : undefined
        };
    }
    function textCell(value, rowIndex) {
        return styledCell(value, rowIndex);
    }
    function taskNameCell(task, rowIndex) {
        const fillColor = task.summary ? SUMMARY_FILL : (task.critical ? CRITICAL_FILL : undefined);
        return {
            ...styledCell(task.name, rowIndex, { fillColor }),
            bold: task.summary || task.milestone
        };
    }
    function taskFlagCell(value, rowIndex, activeFillColor) {
        const isActive = value === true || value === 1 || value === "1";
        return styledCell(value, rowIndex, {
            fillColor: isActive ? activeFillColor : COUNT_FILL,
            horizontalAlign: "center"
        });
    }
    function referenceCell(value, rowIndex) {
        return styledCell(value, rowIndex, {
            fillColor: REFERENCE_FILL,
            horizontalAlign: "center"
        });
    }
    function countCell(value, rowIndex) {
        return styledCell(value, rowIndex, {
            fillColor: COUNT_FILL,
            horizontalAlign: "center"
        });
    }
    function percentCell(value, rowIndex) {
        return styledCell(value, rowIndex, {
            fillColor: PERCENT_FILL,
            horizontalAlign: "center"
        });
    }
    function durationCell(value, rowIndex) {
        return styledCell(value, rowIndex, {
            fillColor: DURATION_FILL,
            horizontalAlign: "center"
        });
    }
    function predecessorCell(value, rowIndex) {
        return styledCell(value, rowIndex, {
            fillColor: REFERENCE_FILL
        });
    }
    function notesCell(value, rowIndex) {
        return styledCell(value, rowIndex, {
            fillColor: NOTES_FILL
        });
    }
    function entityNameCell(value, rowIndex) {
        return styledCell(value, rowIndex, {
            fillColor: NAME_FILL
        });
    }
    function workCell(value, rowIndex) {
        return styledCell(value, rowIndex, {
            fillColor: WORK_FILL
        });
    }
    function editableCell(cellLike) {
        if (cellLike.value === undefined) {
            return cellLike;
        }
        return {
            ...cellLike,
            fillColor: EDITABLE_FILL
        };
    }
    function dateTimeCell(value, rowIndex) {
        return styledCell(formatDateTimeDisplay(value), rowIndex, {
            fillColor: DATE_FILL
        });
    }
    function dateOnlyCell(value, rowIndex) {
        return styledCell(value, rowIndex, {
            fillColor: DATE_FILL,
            horizontalAlign: "center"
        });
    }
    function styledCell(value, rowIndex, overrides = {}) {
        const base = cell(value);
        if (base.value === undefined) {
            return base;
        }
        return {
            ...base,
            fillColor: overrides.fillColor || (rowIndex % 2 === 0 ? ALT_ROW_FILL : undefined),
            horizontalAlign: overrides.horizontalAlign,
            numberFormat: overrides.numberFormat
        };
    }
    function formatDateTimeDisplay(value) {
        if (typeof value !== "string") {
            return value;
        }
        return value.replace("T", " ");
    }
    function isDateTimeLabel(label) {
        return ["StartDate", "FinishDate", "CurrentDate", "StatusDate"].includes(label);
    }
    function isNumericSummaryLabel(label) {
        return ["OutlineCodes", "WBSMasks", "ExtendedAttributes", "MinutesPerDay", "MinutesPerWeek", "DaysPerMonth"].includes(label);
    }
    function isEditableProjectLabel(label) {
        return [
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
        ].includes(label);
    }
    function cloneProjectModel(model) {
        return JSON.parse(JSON.stringify(model));
    }
    function importTasksSheet(workbook, model, changes) {
        const tasksSheet = workbook.sheets.find((sheet) => sheet.name === "Tasks");
        if (!tasksSheet) {
            return;
        }
        const columnIndexByLabel = readHeaderMap(tasksSheet, 2);
        const uidColumnIndex = columnIndexByLabel.get("UID");
        if (uidColumnIndex === undefined) {
            return;
        }
        const taskByUid = new Map(model.tasks.map((task) => [task.uid, task]));
        for (const row of tasksSheet.rows.slice(3)) {
            const uid = readStringCell(row.cells[uidColumnIndex]);
            if (!uid) {
                continue;
            }
            const task = taskByUid.get(uid);
            if (!task) {
                continue;
            }
            const taskLabel = task.name;
            assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "name", "Name", readStringCellAt(row.cells, columnIndexByLabel.get("Name")));
            assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "start", "Start", normalizeDateTimeInput(readStringCellAt(row.cells, columnIndexByLabel.get("Start"))));
            assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "finish", "Finish", normalizeDateTimeInput(readStringCellAt(row.cells, columnIndexByLabel.get("Finish"))));
            assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "percentComplete", "PercentComplete", readNumberCellAt(row.cells, columnIndexByLabel.get("PercentComplete")));
            assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "percentWorkComplete", "PercentWorkComplete", readNumberCellAt(row.cells, columnIndexByLabel.get("PercentWorkComplete")));
            assignIfChanged(changes, "tasks", task.uid, taskLabel, task, "notes", "Notes", readStringCellAt(row.cells, columnIndexByLabel.get("Notes")));
        }
    }
    function importResourcesSheet(workbook, model, changes) {
        const resourcesSheet = workbook.sheets.find((sheet) => sheet.name === "Resources");
        if (!resourcesSheet) {
            return;
        }
        const columnIndexByLabel = readHeaderMap(resourcesSheet, 2);
        const uidColumnIndex = columnIndexByLabel.get("UID");
        if (uidColumnIndex === undefined) {
            return;
        }
        const resourceByUid = new Map(model.resources.map((resource) => [resource.uid, resource]));
        for (const row of resourcesSheet.rows.slice(3)) {
            const uid = readStringCell(row.cells[uidColumnIndex]);
            if (!uid) {
                continue;
            }
            const resource = resourceByUid.get(uid);
            if (!resource) {
                continue;
            }
            const resourceLabel = resource.name;
            assignIfChanged(changes, "resources", resource.uid, resourceLabel, resource, "name", "Name", readStringCellAt(row.cells, columnIndexByLabel.get("Name")));
            assignIfChanged(changes, "resources", resource.uid, resourceLabel, resource, "group", "Group", readStringCellAt(row.cells, columnIndexByLabel.get("Group")));
            assignIfChanged(changes, "resources", resource.uid, resourceLabel, resource, "maxUnits", "MaxUnits", readNumberCellAt(row.cells, columnIndexByLabel.get("MaxUnits")));
        }
    }
    function importAssignmentsSheet(workbook, model, changes) {
        const assignmentsSheet = workbook.sheets.find((sheet) => sheet.name === "Assignments");
        if (!assignmentsSheet) {
            return;
        }
        const columnIndexByLabel = readHeaderMap(assignmentsSheet, 2);
        const uidColumnIndex = columnIndexByLabel.get("UID");
        if (uidColumnIndex === undefined) {
            return;
        }
        const assignmentByUid = new Map(model.assignments.map((assignment) => [assignment.uid, assignment]));
        for (const row of assignmentsSheet.rows.slice(3)) {
            const uid = readStringCell(row.cells[uidColumnIndex]);
            if (!uid) {
                continue;
            }
            const assignment = assignmentByUid.get(uid);
            if (!assignment) {
                continue;
            }
            const assignmentLabel = `TaskUID=${assignment.taskUid}`;
            assignIfChanged(changes, "assignments", assignment.uid, assignmentLabel, assignment, "units", "Units", readNumberCellAt(row.cells, columnIndexByLabel.get("Units")));
            assignIfChanged(changes, "assignments", assignment.uid, assignmentLabel, assignment, "work", "Work", readStringCellAt(row.cells, columnIndexByLabel.get("Work")));
            assignIfChanged(changes, "assignments", assignment.uid, assignmentLabel, assignment, "percentWorkComplete", "PercentWorkComplete", readNumberCellAt(row.cells, columnIndexByLabel.get("PercentWorkComplete")));
        }
    }
    function importCalendarsSheet(workbook, model, changes) {
        const calendarsSheet = workbook.sheets.find((sheet) => sheet.name === "Calendars");
        if (!calendarsSheet) {
            return;
        }
        const columnIndexByLabel = readHeaderMap(calendarsSheet, 2);
        const uidColumnIndex = columnIndexByLabel.get("UID");
        if (uidColumnIndex === undefined) {
            return;
        }
        const calendarByUid = new Map(model.calendars.map((calendar) => [calendar.uid, calendar]));
        for (const row of calendarsSheet.rows.slice(3)) {
            const uid = readStringCell(row.cells[uidColumnIndex]);
            if (!uid) {
                continue;
            }
            const calendar = calendarByUid.get(uid);
            if (!calendar) {
                continue;
            }
            const calendarLabel = calendar.name;
            assignIfChanged(changes, "calendars", calendar.uid, calendarLabel, calendar, "name", "Name", readStringCellAt(row.cells, columnIndexByLabel.get("Name")));
            assignIfChanged(changes, "calendars", calendar.uid, calendarLabel, calendar, "isBaseCalendar", "IsBaseCalendar", readBooleanCellAt(row.cells, columnIndexByLabel.get("IsBaseCalendar")));
            assignIfChanged(changes, "calendars", calendar.uid, calendarLabel, calendar, "baseCalendarUID", "BaseCalendarUID", readStringCellAt(row.cells, columnIndexByLabel.get("BaseCalendarUID")));
        }
    }
    function importNonWorkingDaysSheet(workbook, model, changes) {
        const nonWorkingDaysSheet = workbook.sheets.find((sheet) => sheet.name === "NonWorkingDays");
        if (!nonWorkingDaysSheet) {
            return;
        }
        const columnIndexByLabel = readHeaderMap(nonWorkingDaysSheet, 2);
        const calendarUidColumnIndex = columnIndexByLabel.get("CalendarUID");
        const indexColumnIndex = columnIndexByLabel.get("Index");
        if (calendarUidColumnIndex === undefined || indexColumnIndex === undefined) {
            return;
        }
        const calendarByUid = new Map(model.calendars.map((calendar) => [calendar.uid, calendar]));
        for (const row of nonWorkingDaysSheet.rows.slice(3)) {
            const calendarUid = readStringCell(row.cells[calendarUidColumnIndex]);
            const indexValue = readNumberCell(row.cells[indexColumnIndex]);
            if (!calendarUid || !indexValue) {
                continue;
            }
            const calendar = calendarByUid.get(calendarUid);
            if (!calendar) {
                continue;
            }
            const exception = calendar.exceptions[indexValue - 1];
            if (!exception) {
                continue;
            }
            const exceptionLabel = exception.name || `Exception ${indexValue}`;
            assignIfChanged(changes, "calendars", calendar.uid, calendar.name, exception, "name", `Exception${indexValue}.Name`, readStringCellAt(row.cells, columnIndexByLabel.get("Name")));
            const dateValue = readStringCellAt(row.cells, columnIndexByLabel.get("Date"));
            if (dateValue) {
                const normalizedDate = normalizeDateOnly(dateValue);
                if (normalizedDate) {
                    assignIfChanged(changes, "calendars", calendar.uid, calendar.name, exception, "fromDate", `Exception${indexValue}.FromDate`, `${normalizedDate}T00:00:00`);
                    assignIfChanged(changes, "calendars", calendar.uid, calendar.name, exception, "toDate", `Exception${indexValue}.ToDate`, `${normalizedDate}T23:59:59`);
                }
            }
            else {
                assignIfChanged(changes, "calendars", calendar.uid, calendar.name, exception, "fromDate", `Exception${indexValue}.FromDate`, normalizeExceptionBoundaryInput(readStringCellAt(row.cells, columnIndexByLabel.get("FromDate")), false));
                assignIfChanged(changes, "calendars", calendar.uid, calendar.name, exception, "toDate", `Exception${indexValue}.ToDate`, normalizeExceptionBoundaryInput(readStringCellAt(row.cells, columnIndexByLabel.get("ToDate")), true));
            }
            assignIfChanged(changes, "calendars", calendar.uid, calendar.name, exception, "dayWorking", `Exception${indexValue}.DayWorking`, readBooleanCellAt(row.cells, columnIndexByLabel.get("DayWorking")));
        }
    }
    function formatExceptionDate(exception) {
        var _a, _b;
        const fromDate = (_a = exception.fromDate) === null || _a === void 0 ? void 0 : _a.slice(0, 10);
        const toDate = (_b = exception.toDate) === null || _b === void 0 ? void 0 : _b.slice(0, 10);
        if (!fromDate || !toDate) {
            return undefined;
        }
        return fromDate === toDate ? fromDate : undefined;
    }
    function formatExceptionBoundaryDate(value) {
        return value === null || value === void 0 ? void 0 : value.slice(0, 10);
    }
    function normalizeDateOnly(value) {
        const trimmed = value.trim();
        const match = trimmed.match(/^(\d{4}-\d{2}-\d{2})/);
        return match ? match[1] : undefined;
    }
    function normalizeDateTimeInput(value) {
        if (!value) {
            return value;
        }
        const trimmed = value.trim();
        if (/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/.test(trimmed)) {
            return trimmed.replace(" ", "T");
        }
        return trimmed;
    }
    function normalizeExceptionBoundaryInput(value, isEndOfDay) {
        const normalizedDate = value ? normalizeDateOnly(value) : undefined;
        if (normalizedDate && normalizedDate === (value === null || value === void 0 ? void 0 : value.trim())) {
            return `${normalizedDate}${isEndOfDay ? "T23:59:59" : "T00:00:00"}`;
        }
        return normalizeDateTimeInput(value);
    }
    function readHeaderMap(sheet, headerRowIndex) {
        const headerRow = sheet.rows[headerRowIndex];
        const columnIndexByLabel = new Map();
        if (!headerRow) {
            return columnIndexByLabel;
        }
        headerRow.cells.forEach((cell, index) => {
            if (typeof cell.value === "string" && cell.value) {
                columnIndexByLabel.set(cell.value, index);
            }
        });
        return columnIndexByLabel;
    }
    function readStringCellAt(cells, index) {
        if (index === undefined) {
            return undefined;
        }
        return readStringCell(cells[index]);
    }
    function readNumberCellAt(cells, index) {
        if (index === undefined) {
            return undefined;
        }
        return readNumberCell(cells[index]);
    }
    function readBooleanCellAt(cells, index) {
        if (index === undefined) {
            return undefined;
        }
        return readBooleanCell(cells[index]);
    }
    function readStringCell(cell) {
        if (!cell || cell.value === undefined) {
            return undefined;
        }
        if (typeof cell.value === "string") {
            return cell.value;
        }
        if (typeof cell.value === "number" || typeof cell.value === "boolean") {
            return String(cell.value);
        }
        return undefined;
    }
    function readNumberCell(cell) {
        if (!cell || cell.value === undefined) {
            return undefined;
        }
        if (typeof cell.value === "number" && Number.isFinite(cell.value)) {
            return cell.value;
        }
        if (typeof cell.value === "string" && cell.value.trim() !== "") {
            const parsed = Number(cell.value);
            return Number.isFinite(parsed) ? parsed : undefined;
        }
        return undefined;
    }
    function readBooleanCell(cell) {
        if (!cell || cell.value === undefined) {
            return undefined;
        }
        if (typeof cell.value === "boolean") {
            return cell.value;
        }
        if (typeof cell.value === "number") {
            return cell.value !== 0;
        }
        if (typeof cell.value === "string") {
            if (cell.value === "true" || cell.value === "TRUE" || cell.value === "1") {
                return true;
            }
            if (cell.value === "false" || cell.value === "FALSE" || cell.value === "0") {
                return false;
            }
        }
        return undefined;
    }
    function assignIfChanged(changes, scope, uid, label, target, key, field, value) {
        if (value === undefined) {
            return;
        }
        const before = target[key];
        if (before === value) {
            return;
        }
        target[key] = value;
        changes.push({
            scope,
            uid,
            label,
            field,
            before: before,
            after: value
        });
    }
    globalThis.__mikuprojectProjectXlsx = {
        exportProjectWorkbook,
        importProjectWorkbook,
        importProjectWorkbookDetailed
    };
})();
