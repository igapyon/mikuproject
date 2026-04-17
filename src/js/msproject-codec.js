/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    globalThis.__mikuprojectMsprojectCodec = {
        importMsProjectXml(input) {
            var _a, _b, _c, _d, _e, _f, _g;
            const xml = input.parseXmlDocument(input.xmlText);
            const projectElement = xml.documentElement;
            const calendars = Array.from(((_a = projectElement.getElementsByTagName("Calendars")[0]) === null || _a === void 0 ? void 0 : _a.getElementsByTagName("Calendar")) || []);
            const tasks = Array.from(((_b = projectElement.getElementsByTagName("Tasks")[0]) === null || _b === void 0 ? void 0 : _b.getElementsByTagName("Task")) || []);
            const resources = Array.from(((_c = projectElement.getElementsByTagName("Resources")[0]) === null || _c === void 0 ? void 0 : _c.getElementsByTagName("Resource")) || []);
            const assignments = Array.from(((_d = projectElement.getElementsByTagName("Assignments")[0]) === null || _d === void 0 ? void 0 : _d.getElementsByTagName("Assignment")) || []);
            const textContent = input.textContent;
            const parseBoolean = input.parseBoolean;
            const parseNumber = input.parseNumber;
            return input.normalizeProjectModel(input.ensureDefaultProjectCalendar({
                project: {
                    name: textContent(projectElement, "Name"),
                    title: textContent(projectElement, "Title") || undefined,
                    author: textContent(projectElement, "Author") || undefined,
                    company: textContent(projectElement, "Company") || undefined,
                    creationDate: textContent(projectElement, "CreationDate") || undefined,
                    lastSaved: textContent(projectElement, "LastSaved") || undefined,
                    saveVersion: textContent(projectElement, "SaveVersion") ? parseNumber(textContent(projectElement, "SaveVersion"), 0) : undefined,
                    currentDate: textContent(projectElement, "CurrentDate") || undefined,
                    startDate: textContent(projectElement, "StartDate"),
                    finishDate: textContent(projectElement, "FinishDate"),
                    scheduleFromStart: parseBoolean(textContent(projectElement, "ScheduleFromStart")),
                    defaultStartTime: textContent(projectElement, "DefaultStartTime") || undefined,
                    defaultFinishTime: textContent(projectElement, "DefaultFinishTime") || undefined,
                    minutesPerDay: textContent(projectElement, "MinutesPerDay") ? parseNumber(textContent(projectElement, "MinutesPerDay"), 0) : undefined,
                    minutesPerWeek: textContent(projectElement, "MinutesPerWeek") ? parseNumber(textContent(projectElement, "MinutesPerWeek"), 0) : undefined,
                    daysPerMonth: textContent(projectElement, "DaysPerMonth") ? parseNumber(textContent(projectElement, "DaysPerMonth"), 0) : undefined,
                    statusDate: textContent(projectElement, "StatusDate") || undefined,
                    weekStartDay: textContent(projectElement, "WeekStartDay") ? parseNumber(textContent(projectElement, "WeekStartDay"), 0) : undefined,
                    workFormat: textContent(projectElement, "WorkFormat") ? parseNumber(textContent(projectElement, "WorkFormat"), 0) : undefined,
                    durationFormat: textContent(projectElement, "DurationFormat") ? parseNumber(textContent(projectElement, "DurationFormat"), 0) : undefined,
                    currencyCode: textContent(projectElement, "CurrencyCode") || undefined,
                    currencyDigits: textContent(projectElement, "CurrencyDigits") ? parseNumber(textContent(projectElement, "CurrencyDigits"), 0) : undefined,
                    currencySymbol: textContent(projectElement, "CurrencySymbol") || undefined,
                    currencySymbolPosition: textContent(projectElement, "CurrencySymbolPosition") ? parseNumber(textContent(projectElement, "CurrencySymbolPosition"), 0) : undefined,
                    fyStartDate: textContent(projectElement, "FYStartDate") || undefined,
                    fiscalYearStart: textContent(projectElement, "FiscalYearStart") ? parseBoolean(textContent(projectElement, "FiscalYearStart")) : undefined,
                    criticalSlackLimit: textContent(projectElement, "CriticalSlackLimit") ? parseNumber(textContent(projectElement, "CriticalSlackLimit"), 0) : undefined,
                    defaultTaskType: textContent(projectElement, "DefaultTaskType") ? parseNumber(textContent(projectElement, "DefaultTaskType"), 0) : undefined,
                    defaultFixedCostAccrual: textContent(projectElement, "DefaultFixedCostAccrual") ? parseNumber(textContent(projectElement, "DefaultFixedCostAccrual"), 0) : undefined,
                    defaultStandardRate: textContent(projectElement, "DefaultStandardRate") || undefined,
                    defaultOvertimeRate: textContent(projectElement, "DefaultOvertimeRate") || undefined,
                    defaultTaskEVMethod: textContent(projectElement, "DefaultTaskEVMethod") ? parseNumber(textContent(projectElement, "DefaultTaskEVMethod"), 0) : undefined,
                    newTaskStartDate: textContent(projectElement, "NewTaskStartDate") ? parseNumber(textContent(projectElement, "NewTaskStartDate"), 0) : undefined,
                    newTasksAreManual: textContent(projectElement, "NewTasksAreManual") ? parseBoolean(textContent(projectElement, "NewTasksAreManual")) : undefined,
                    newTasksEffortDriven: textContent(projectElement, "NewTasksEffortDriven") ? parseBoolean(textContent(projectElement, "NewTasksEffortDriven")) : undefined,
                    newTasksEstimated: textContent(projectElement, "NewTasksEstimated") ? parseBoolean(textContent(projectElement, "NewTasksEstimated")) : undefined,
                    actualsInSync: textContent(projectElement, "ActualsInSync") ? parseBoolean(textContent(projectElement, "ActualsInSync")) : undefined,
                    editableActualCosts: textContent(projectElement, "EditableActualCosts") ? parseBoolean(textContent(projectElement, "EditableActualCosts")) : undefined,
                    honorConstraints: textContent(projectElement, "HonorConstraints") ? parseBoolean(textContent(projectElement, "HonorConstraints")) : undefined,
                    insertedProjectsLikeSummary: textContent(projectElement, "InsertedProjectsLikeSummary") ? parseBoolean(textContent(projectElement, "InsertedProjectsLikeSummary")) : undefined,
                    multipleCriticalPaths: textContent(projectElement, "MultipleCriticalPaths") ? parseBoolean(textContent(projectElement, "MultipleCriticalPaths")) : undefined,
                    taskUpdatesResource: textContent(projectElement, "TaskUpdatesResource") ? parseBoolean(textContent(projectElement, "TaskUpdatesResource")) : undefined,
                    updateManuallyScheduledTasksWhenEditingLinks: textContent(projectElement, "UpdateManuallyScheduledTasksWhenEditingLinks") ? parseBoolean(textContent(projectElement, "UpdateManuallyScheduledTasksWhenEditingLinks")) : undefined,
                    calendarUID: textContent(projectElement, "CalendarUID") || undefined,
                    outlineCodes: Array.from(((_e = projectElement.getElementsByTagName("OutlineCodes")[0]) === null || _e === void 0 ? void 0 : _e.getElementsByTagName("OutlineCode")) || []).map((outlineCode) => ({
                        fieldID: textContent(outlineCode, "FieldID") || undefined,
                        fieldName: textContent(outlineCode, "FieldName") || undefined,
                        alias: textContent(outlineCode, "Alias") || undefined,
                        onlyTableValues: textContent(outlineCode, "OnlyTableValues") ? parseBoolean(textContent(outlineCode, "OnlyTableValues")) : undefined,
                        enterprise: textContent(outlineCode, "Enterprise") ? parseBoolean(textContent(outlineCode, "Enterprise")) : undefined,
                        resourceSubstitutionEnabled: textContent(outlineCode, "ResourceSubstitutionEnabled") ? parseBoolean(textContent(outlineCode, "ResourceSubstitutionEnabled")) : undefined,
                        leafOnly: textContent(outlineCode, "LeafOnly") ? parseBoolean(textContent(outlineCode, "LeafOnly")) : undefined,
                        allLevelsRequired: textContent(outlineCode, "AllLevelsRequired") ? parseBoolean(textContent(outlineCode, "AllLevelsRequired")) : undefined,
                        masks: input.parseOutlineCodeMasks(outlineCode),
                        values: input.parseOutlineCodeValues(outlineCode)
                    })),
                    wbsMasks: Array.from(((_f = projectElement.getElementsByTagName("WBSMasks")[0]) === null || _f === void 0 ? void 0 : _f.getElementsByTagName("WBSMask")) || []).map((wbsMask) => ({
                        level: parseNumber(textContent(wbsMask, "Level"), 0),
                        mask: textContent(wbsMask, "Mask") || undefined,
                        length: textContent(wbsMask, "Length") ? parseNumber(textContent(wbsMask, "Length"), 0) : undefined,
                        sequence: textContent(wbsMask, "Sequence") ? parseNumber(textContent(wbsMask, "Sequence"), 0) : undefined
                    })),
                    extendedAttributes: Array.from(((_g = projectElement.getElementsByTagName("ExtendedAttributes")[0]) === null || _g === void 0 ? void 0 : _g.getElementsByTagName("ExtendedAttribute")) || []).map((attribute) => ({
                        fieldID: textContent(attribute, "FieldID") || undefined,
                        fieldName: textContent(attribute, "FieldName") || undefined,
                        alias: textContent(attribute, "Alias") || undefined,
                        calculationType: textContent(attribute, "CalculationType") ? parseNumber(textContent(attribute, "CalculationType"), 0) : undefined,
                        restrictValues: textContent(attribute, "RestrictValues") ? parseBoolean(textContent(attribute, "RestrictValues")) : undefined,
                        appendNewValues: textContent(attribute, "AppendNewValues") ? parseBoolean(textContent(attribute, "AppendNewValues")) : undefined
                    }))
                },
                calendars: calendars.map((calendar) => {
                    var _a, _b;
                    return ({
                        uid: textContent(calendar, "UID"),
                        name: textContent(calendar, "Name"),
                        isBaseCalendar: parseBoolean(textContent(calendar, "IsBaseCalendar")),
                        isBaselineCalendar: textContent(calendar, "IsBaselineCalendar") ? parseBoolean(textContent(calendar, "IsBaselineCalendar")) : undefined,
                        baseCalendarUID: textContent(calendar, "BaseCalendarUID") || undefined,
                        weekDays: input.parseWeekDays(calendar),
                        exceptions: Array.from(((_a = calendar.getElementsByTagName("Exceptions")[0]) === null || _a === void 0 ? void 0 : _a.getElementsByTagName("Exception")) || []).map((exception) => ({
                            name: textContent(exception, "Name") || undefined,
                            fromDate: textContent(exception, "FromDate") || undefined,
                            toDate: textContent(exception, "ToDate") || undefined,
                            dayWorking: textContent(exception, "DayWorking") ? parseBoolean(textContent(exception, "DayWorking")) : undefined,
                            workingTimes: input.parseWorkingTimes(exception)
                        })),
                        workWeeks: Array.from(((_b = calendar.getElementsByTagName("WorkWeeks")[0]) === null || _b === void 0 ? void 0 : _b.getElementsByTagName("WorkWeek")) || []).map((workWeek) => ({
                            name: textContent(workWeek, "Name") || undefined,
                            fromDate: textContent(workWeek, "FromDate") || undefined,
                            toDate: textContent(workWeek, "ToDate") || undefined,
                            weekDays: input.parseWeekDays(workWeek)
                        }))
                    });
                }),
                tasks: tasks.map((task) => ({
                    uid: textContent(task, "UID"),
                    id: textContent(task, "ID"),
                    name: textContent(task, "Name"),
                    outlineLevel: parseNumber(textContent(task, "OutlineLevel"), 1),
                    outlineNumber: textContent(task, "OutlineNumber"),
                    wbs: textContent(task, "WBS") || undefined,
                    type: textContent(task, "Type") ? parseNumber(textContent(task, "Type"), 0) : undefined,
                    calendarUID: textContent(task, "CalendarUID") || undefined,
                    priority: textContent(task, "Priority") ? parseNumber(textContent(task, "Priority"), 0) : undefined,
                    start: textContent(task, "Start"),
                    finish: textContent(task, "Finish"),
                    duration: textContent(task, "Duration"),
                    actualStart: textContent(task, "ActualStart") || undefined,
                    actualFinish: textContent(task, "ActualFinish") || undefined,
                    deadline: textContent(task, "Deadline") || undefined,
                    startVariance: textContent(task, "StartVariance") || undefined,
                    finishVariance: textContent(task, "FinishVariance") || undefined,
                    work: textContent(task, "Work") || undefined,
                    workVariance: textContent(task, "WorkVariance") || undefined,
                    totalSlack: textContent(task, "TotalSlack") || undefined,
                    freeSlack: textContent(task, "FreeSlack") || undefined,
                    cost: textContent(task, "Cost") ? parseNumber(textContent(task, "Cost"), 0) : undefined,
                    actualCost: textContent(task, "ActualCost") ? parseNumber(textContent(task, "ActualCost"), 0) : undefined,
                    remainingCost: textContent(task, "RemainingCost") ? parseNumber(textContent(task, "RemainingCost"), 0) : undefined,
                    remainingWork: textContent(task, "RemainingWork") || undefined,
                    actualWork: textContent(task, "ActualWork") || undefined,
                    milestone: parseBoolean(textContent(task, "Milestone")),
                    summary: parseBoolean(textContent(task, "Summary")),
                    critical: textContent(task, "Critical") ? parseBoolean(textContent(task, "Critical")) : undefined,
                    percentComplete: parseNumber(textContent(task, "PercentComplete"), 0),
                    percentWorkComplete: textContent(task, "PercentWorkComplete") ? parseNumber(textContent(task, "PercentWorkComplete"), 0) : undefined,
                    notes: textContent(task, "Notes") || undefined,
                    constraintType: textContent(task, "ConstraintType") ? parseNumber(textContent(task, "ConstraintType"), 0) : undefined,
                    constraintDate: textContent(task, "ConstraintDate") || undefined,
                    extendedAttributes: Array.from(task.getElementsByTagName("ExtendedAttribute")).map((attribute) => ({
                        fieldID: textContent(attribute, "FieldID") || undefined,
                        value: textContent(attribute, "Value") || undefined
                    })),
                    baselines: Array.from(task.getElementsByTagName("Baseline")).map((baseline) => ({
                        number: textContent(baseline, "Number") ? parseNumber(textContent(baseline, "Number"), 0) : undefined,
                        start: textContent(baseline, "Start") || undefined,
                        finish: textContent(baseline, "Finish") || undefined,
                        work: textContent(baseline, "Work") || undefined,
                        cost: textContent(baseline, "Cost") ? parseNumber(textContent(baseline, "Cost"), 0) : undefined
                    })),
                    timephasedData: Array.from(task.getElementsByTagName("TimephasedData")).map((timephasedData) => ({
                        type: textContent(timephasedData, "Type") ? parseNumber(textContent(timephasedData, "Type"), 0) : undefined,
                        uid: textContent(timephasedData, "UID") || undefined,
                        start: textContent(timephasedData, "Start") || undefined,
                        finish: textContent(timephasedData, "Finish") || undefined,
                        unit: textContent(timephasedData, "Unit") ? parseNumber(textContent(timephasedData, "Unit"), 0) : undefined,
                        value: textContent(timephasedData, "Value") || undefined
                    })),
                    predecessors: Array.from(task.getElementsByTagName("PredecessorLink")).map((link) => ({
                        predecessorUid: textContent(link, "PredecessorUID"),
                        type: parseNumber(textContent(link, "Type"), 0),
                        linkLag: textContent(link, "LinkLag") || undefined
                    }))
                })),
                resources: resources.map((resource) => ({
                    uid: textContent(resource, "UID"),
                    id: textContent(resource, "ID"),
                    name: textContent(resource, "Name"),
                    type: parseNumber(textContent(resource, "Type"), 0),
                    initials: textContent(resource, "Initials") || undefined,
                    group: textContent(resource, "Group") || undefined,
                    workGroup: textContent(resource, "WorkGroup") ? parseNumber(textContent(resource, "WorkGroup"), 0) : undefined,
                    maxUnits: textContent(resource, "MaxUnits") ? parseNumber(textContent(resource, "MaxUnits"), 0) : undefined,
                    calendarUID: textContent(resource, "CalendarUID") || undefined,
                    standardRate: textContent(resource, "StandardRate") || undefined,
                    standardRateFormat: textContent(resource, "StandardRateFormat") ? parseNumber(textContent(resource, "StandardRateFormat"), 0) : undefined,
                    overtimeRate: textContent(resource, "OvertimeRate") || undefined,
                    overtimeRateFormat: textContent(resource, "OvertimeRateFormat") ? parseNumber(textContent(resource, "OvertimeRateFormat"), 0) : undefined,
                    costPerUse: textContent(resource, "CostPerUse") ? parseNumber(textContent(resource, "CostPerUse"), 0) : undefined,
                    work: textContent(resource, "Work") || undefined,
                    actualWork: textContent(resource, "ActualWork") || undefined,
                    remainingWork: textContent(resource, "RemainingWork") || undefined,
                    cost: textContent(resource, "Cost") ? parseNumber(textContent(resource, "Cost"), 0) : undefined,
                    actualCost: textContent(resource, "ActualCost") ? parseNumber(textContent(resource, "ActualCost"), 0) : undefined,
                    remainingCost: textContent(resource, "RemainingCost") ? parseNumber(textContent(resource, "RemainingCost"), 0) : undefined,
                    percentWorkComplete: textContent(resource, "PercentWorkComplete") ? parseNumber(textContent(resource, "PercentWorkComplete"), 0) : undefined,
                    extendedAttributes: Array.from(resource.getElementsByTagName("ExtendedAttribute")).map((attribute) => ({
                        fieldID: textContent(attribute, "FieldID") || undefined,
                        value: textContent(attribute, "Value") || undefined
                    })),
                    baselines: Array.from(resource.getElementsByTagName("Baseline")).map((baseline) => ({
                        number: textContent(baseline, "Number") ? parseNumber(textContent(baseline, "Number"), 0) : undefined,
                        start: textContent(baseline, "Start") || undefined,
                        finish: textContent(baseline, "Finish") || undefined,
                        work: textContent(baseline, "Work") || undefined,
                        cost: textContent(baseline, "Cost") ? parseNumber(textContent(baseline, "Cost"), 0) : undefined
                    })),
                    timephasedData: Array.from(resource.getElementsByTagName("TimephasedData")).map((timephasedData) => ({
                        type: textContent(timephasedData, "Type") ? parseNumber(textContent(timephasedData, "Type"), 0) : undefined,
                        uid: textContent(timephasedData, "UID") || undefined,
                        start: textContent(timephasedData, "Start") || undefined,
                        finish: textContent(timephasedData, "Finish") || undefined,
                        unit: textContent(timephasedData, "Unit") ? parseNumber(textContent(timephasedData, "Unit"), 0) : undefined,
                        value: textContent(timephasedData, "Value") || undefined
                    }))
                })),
                assignments: assignments.map((assignment) => ({
                    uid: textContent(assignment, "UID"),
                    taskUid: textContent(assignment, "TaskUID"),
                    resourceUid: textContent(assignment, "ResourceUID"),
                    start: textContent(assignment, "Start") || undefined,
                    finish: textContent(assignment, "Finish") || undefined,
                    startVariance: textContent(assignment, "StartVariance") || undefined,
                    finishVariance: textContent(assignment, "FinishVariance") || undefined,
                    delay: textContent(assignment, "Delay") || undefined,
                    milestone: textContent(assignment, "Milestone") ? parseBoolean(textContent(assignment, "Milestone")) : undefined,
                    workContour: textContent(assignment, "WorkContour") ? parseNumber(textContent(assignment, "WorkContour"), 0) : undefined,
                    units: parseNumber(textContent(assignment, "Units"), 0),
                    work: textContent(assignment, "Work") || undefined,
                    cost: textContent(assignment, "Cost") ? parseNumber(textContent(assignment, "Cost"), 0) : undefined,
                    actualCost: textContent(assignment, "ActualCost") ? parseNumber(textContent(assignment, "ActualCost"), 0) : undefined,
                    remainingCost: textContent(assignment, "RemainingCost") ? parseNumber(textContent(assignment, "RemainingCost"), 0) : undefined,
                    percentWorkComplete: textContent(assignment, "PercentWorkComplete") ? parseNumber(textContent(assignment, "PercentWorkComplete"), 0) : undefined,
                    overtimeWork: textContent(assignment, "OvertimeWork") || undefined,
                    actualOvertimeWork: textContent(assignment, "ActualOvertimeWork") || undefined,
                    actualWork: textContent(assignment, "ActualWork") || undefined,
                    remainingWork: textContent(assignment, "RemainingWork") || undefined,
                    extendedAttributes: Array.from(assignment.getElementsByTagName("ExtendedAttribute")).map((attribute) => ({
                        fieldID: textContent(attribute, "FieldID") || undefined,
                        value: textContent(attribute, "Value") || undefined
                    })),
                    baselines: Array.from(assignment.getElementsByTagName("Baseline")).map((baseline) => ({
                        number: textContent(baseline, "Number") ? parseNumber(textContent(baseline, "Number"), 0) : undefined,
                        start: textContent(baseline, "Start") || undefined,
                        finish: textContent(baseline, "Finish") || undefined,
                        work: textContent(baseline, "Work") || undefined,
                        cost: textContent(baseline, "Cost") ? parseNumber(textContent(baseline, "Cost"), 0) : undefined
                    })),
                    timephasedData: Array.from(assignment.getElementsByTagName("TimephasedData")).map((timephasedData) => ({
                        type: textContent(timephasedData, "Type") ? parseNumber(textContent(timephasedData, "Type"), 0) : undefined,
                        uid: textContent(timephasedData, "UID") || undefined,
                        start: textContent(timephasedData, "Start") || undefined,
                        finish: textContent(timephasedData, "Finish") || undefined,
                        unit: textContent(timephasedData, "Unit") ? parseNumber(textContent(timephasedData, "Unit"), 0) : undefined,
                        value: textContent(timephasedData, "Value") || undefined
                    }))
                }))
            }));
        },
        exportMsProjectXml(input) {
            const normalizedModel = input.ensureDefaultProjectCalendar(input.normalizeProjectModel(input.model));
            const doc = document.implementation.createDocument("", "Project", null);
            const project = doc.documentElement;
            project.setAttribute("xmlns", "http://schemas.microsoft.com/project");
            const appendTextElement = input.appendTextElement;
            appendTextElement(doc, project, "Name", normalizedModel.project.name);
            appendTextElement(doc, project, "Title", normalizedModel.project.title);
            appendTextElement(doc, project, "Company", normalizedModel.project.company);
            appendTextElement(doc, project, "Author", normalizedModel.project.author);
            appendTextElement(doc, project, "CreationDate", normalizedModel.project.creationDate);
            appendTextElement(doc, project, "LastSaved", normalizedModel.project.lastSaved);
            appendTextElement(doc, project, "SaveVersion", normalizedModel.project.saveVersion);
            appendTextElement(doc, project, "CurrentDate", normalizedModel.project.currentDate);
            appendTextElement(doc, project, "StartDate", normalizedModel.project.startDate);
            appendTextElement(doc, project, "FinishDate", normalizedModel.project.finishDate);
            appendTextElement(doc, project, "ScheduleFromStart", normalizedModel.project.scheduleFromStart);
            appendTextElement(doc, project, "DefaultStartTime", normalizedModel.project.defaultStartTime);
            appendTextElement(doc, project, "DefaultFinishTime", normalizedModel.project.defaultFinishTime);
            appendTextElement(doc, project, "MinutesPerDay", normalizedModel.project.minutesPerDay);
            appendTextElement(doc, project, "MinutesPerWeek", normalizedModel.project.minutesPerWeek);
            appendTextElement(doc, project, "DaysPerMonth", normalizedModel.project.daysPerMonth);
            appendTextElement(doc, project, "StatusDate", normalizedModel.project.statusDate);
            appendTextElement(doc, project, "WeekStartDay", normalizedModel.project.weekStartDay);
            appendTextElement(doc, project, "WorkFormat", normalizedModel.project.workFormat);
            appendTextElement(doc, project, "DurationFormat", normalizedModel.project.durationFormat);
            appendTextElement(doc, project, "CurrencyCode", normalizedModel.project.currencyCode);
            appendTextElement(doc, project, "CurrencyDigits", normalizedModel.project.currencyDigits);
            appendTextElement(doc, project, "CurrencySymbol", normalizedModel.project.currencySymbol);
            appendTextElement(doc, project, "CurrencySymbolPosition", normalizedModel.project.currencySymbolPosition);
            appendTextElement(doc, project, "FYStartDate", normalizedModel.project.fyStartDate);
            appendTextElement(doc, project, "FiscalYearStart", normalizedModel.project.fiscalYearStart);
            appendTextElement(doc, project, "CriticalSlackLimit", normalizedModel.project.criticalSlackLimit);
            appendTextElement(doc, project, "DefaultTaskType", normalizedModel.project.defaultTaskType);
            appendTextElement(doc, project, "DefaultFixedCostAccrual", normalizedModel.project.defaultFixedCostAccrual);
            appendTextElement(doc, project, "DefaultStandardRate", normalizedModel.project.defaultStandardRate);
            appendTextElement(doc, project, "DefaultOvertimeRate", normalizedModel.project.defaultOvertimeRate);
            appendTextElement(doc, project, "DefaultTaskEVMethod", normalizedModel.project.defaultTaskEVMethod);
            appendTextElement(doc, project, "NewTaskStartDate", normalizedModel.project.newTaskStartDate);
            appendTextElement(doc, project, "NewTasksAreManual", normalizedModel.project.newTasksAreManual);
            appendTextElement(doc, project, "NewTasksEffortDriven", normalizedModel.project.newTasksEffortDriven);
            appendTextElement(doc, project, "NewTasksEstimated", normalizedModel.project.newTasksEstimated);
            appendTextElement(doc, project, "ActualsInSync", normalizedModel.project.actualsInSync);
            appendTextElement(doc, project, "EditableActualCosts", normalizedModel.project.editableActualCosts);
            appendTextElement(doc, project, "HonorConstraints", normalizedModel.project.honorConstraints);
            appendTextElement(doc, project, "InsertedProjectsLikeSummary", normalizedModel.project.insertedProjectsLikeSummary);
            appendTextElement(doc, project, "MultipleCriticalPaths", normalizedModel.project.multipleCriticalPaths);
            appendTextElement(doc, project, "TaskUpdatesResource", normalizedModel.project.taskUpdatesResource);
            appendTextElement(doc, project, "UpdateManuallyScheduledTasksWhenEditingLinks", normalizedModel.project.updateManuallyScheduledTasksWhenEditingLinks);
            appendTextElement(doc, project, "CalendarUID", normalizedModel.project.calendarUID);
            if (normalizedModel.project.outlineCodes.length > 0) {
                const outlineCodesElement = doc.createElement("OutlineCodes");
                for (const outlineCode of normalizedModel.project.outlineCodes) {
                    const outlineCodeElement = doc.createElement("OutlineCode");
                    appendTextElement(doc, outlineCodeElement, "FieldID", outlineCode.fieldID);
                    appendTextElement(doc, outlineCodeElement, "FieldName", outlineCode.fieldName);
                    appendTextElement(doc, outlineCodeElement, "Alias", outlineCode.alias);
                    appendTextElement(doc, outlineCodeElement, "OnlyTableValues", outlineCode.onlyTableValues);
                    appendTextElement(doc, outlineCodeElement, "Enterprise", outlineCode.enterprise);
                    appendTextElement(doc, outlineCodeElement, "ResourceSubstitutionEnabled", outlineCode.resourceSubstitutionEnabled);
                    appendTextElement(doc, outlineCodeElement, "LeafOnly", outlineCode.leafOnly);
                    appendTextElement(doc, outlineCodeElement, "AllLevelsRequired", outlineCode.allLevelsRequired);
                    if (outlineCode.masks.length > 0) {
                        const masksElement = doc.createElement("Masks");
                        for (const mask of outlineCode.masks) {
                            const maskElement = doc.createElement("Mask");
                            appendTextElement(doc, maskElement, "Level", mask.level);
                            appendTextElement(doc, maskElement, "Mask", mask.mask);
                            appendTextElement(doc, maskElement, "Length", mask.length);
                            appendTextElement(doc, maskElement, "Sequence", mask.sequence);
                            masksElement.appendChild(maskElement);
                        }
                        outlineCodeElement.appendChild(masksElement);
                    }
                    if (outlineCode.values.length > 0) {
                        const valuesElement = doc.createElement("Values");
                        for (const value of outlineCode.values) {
                            const valueElement = doc.createElement("Value");
                            appendTextElement(doc, valueElement, "Value", value.value);
                            appendTextElement(doc, valueElement, "Description", value.description);
                            valuesElement.appendChild(valueElement);
                        }
                        outlineCodeElement.appendChild(valuesElement);
                    }
                    outlineCodesElement.appendChild(outlineCodeElement);
                }
                project.appendChild(outlineCodesElement);
            }
            if (normalizedModel.project.wbsMasks.length > 0) {
                const wbsMasksElement = doc.createElement("WBSMasks");
                for (const wbsMask of normalizedModel.project.wbsMasks) {
                    const wbsMaskElement = doc.createElement("WBSMask");
                    appendTextElement(doc, wbsMaskElement, "Level", wbsMask.level);
                    appendTextElement(doc, wbsMaskElement, "Mask", wbsMask.mask);
                    appendTextElement(doc, wbsMaskElement, "Length", wbsMask.length);
                    appendTextElement(doc, wbsMaskElement, "Sequence", wbsMask.sequence);
                    wbsMasksElement.appendChild(wbsMaskElement);
                }
                project.appendChild(wbsMasksElement);
            }
            if (normalizedModel.project.extendedAttributes.length > 0) {
                const extendedAttributesElement = doc.createElement("ExtendedAttributes");
                for (const attribute of normalizedModel.project.extendedAttributes) {
                    const extendedAttributeElement = doc.createElement("ExtendedAttribute");
                    appendTextElement(doc, extendedAttributeElement, "FieldID", attribute.fieldID);
                    appendTextElement(doc, extendedAttributeElement, "FieldName", attribute.fieldName);
                    appendTextElement(doc, extendedAttributeElement, "Alias", attribute.alias);
                    appendTextElement(doc, extendedAttributeElement, "CalculationType", attribute.calculationType);
                    appendTextElement(doc, extendedAttributeElement, "RestrictValues", attribute.restrictValues);
                    appendTextElement(doc, extendedAttributeElement, "AppendNewValues", attribute.appendNewValues);
                    extendedAttributesElement.appendChild(extendedAttributeElement);
                }
                project.appendChild(extendedAttributesElement);
            }
            const calendarsElement = doc.createElement("Calendars");
            for (const calendar of normalizedModel.calendars) {
                const calendarElement = doc.createElement("Calendar");
                appendTextElement(doc, calendarElement, "UID", calendar.uid);
                appendTextElement(doc, calendarElement, "Name", calendar.name);
                appendTextElement(doc, calendarElement, "IsBaseCalendar", calendar.isBaseCalendar);
                appendTextElement(doc, calendarElement, "IsBaselineCalendar", calendar.isBaselineCalendar);
                appendTextElement(doc, calendarElement, "BaseCalendarUID", calendar.baseCalendarUID);
                if (calendar.exceptions.length > 0) {
                    const exceptionsElement = doc.createElement("Exceptions");
                    for (const exception of calendar.exceptions) {
                        const exceptionElement = doc.createElement("Exception");
                        appendTextElement(doc, exceptionElement, "Name", exception.name);
                        appendTextElement(doc, exceptionElement, "FromDate", exception.fromDate);
                        appendTextElement(doc, exceptionElement, "ToDate", exception.toDate);
                        appendTextElement(doc, exceptionElement, "DayWorking", exception.dayWorking);
                        input.appendWorkingTimes(doc, exceptionElement, exception.workingTimes);
                        exceptionsElement.appendChild(exceptionElement);
                    }
                    calendarElement.appendChild(exceptionsElement);
                }
                if (calendar.workWeeks.length > 0) {
                    const workWeeksElement = doc.createElement("WorkWeeks");
                    for (const workWeek of calendar.workWeeks) {
                        const workWeekElement = doc.createElement("WorkWeek");
                        appendTextElement(doc, workWeekElement, "Name", workWeek.name);
                        appendTextElement(doc, workWeekElement, "FromDate", workWeek.fromDate);
                        appendTextElement(doc, workWeekElement, "ToDate", workWeek.toDate);
                        input.appendWeekDays(doc, workWeekElement, workWeek.weekDays);
                        workWeeksElement.appendChild(workWeekElement);
                    }
                    calendarElement.appendChild(workWeeksElement);
                }
                input.appendWeekDays(doc, calendarElement, calendar.weekDays);
                calendarsElement.appendChild(calendarElement);
            }
            project.appendChild(calendarsElement);
            const tasksElement = doc.createElement("Tasks");
            for (const task of normalizedModel.tasks) {
                const taskElement = doc.createElement("Task");
                appendTextElement(doc, taskElement, "UID", task.uid);
                appendTextElement(doc, taskElement, "ID", task.id);
                appendTextElement(doc, taskElement, "Name", task.name);
                appendTextElement(doc, taskElement, "OutlineLevel", task.outlineLevel);
                appendTextElement(doc, taskElement, "OutlineNumber", task.outlineNumber);
                appendTextElement(doc, taskElement, "WBS", task.wbs);
                appendTextElement(doc, taskElement, "Type", task.type);
                appendTextElement(doc, taskElement, "CalendarUID", task.calendarUID);
                appendTextElement(doc, taskElement, "Priority", task.priority);
                appendTextElement(doc, taskElement, "Start", task.start);
                appendTextElement(doc, taskElement, "Finish", task.finish);
                appendTextElement(doc, taskElement, "Duration", task.duration);
                appendTextElement(doc, taskElement, "ActualStart", task.actualStart);
                appendTextElement(doc, taskElement, "ActualFinish", task.actualFinish);
                appendTextElement(doc, taskElement, "Deadline", task.deadline);
                appendTextElement(doc, taskElement, "StartVariance", task.startVariance);
                appendTextElement(doc, taskElement, "FinishVariance", task.finishVariance);
                appendTextElement(doc, taskElement, "Work", task.work);
                appendTextElement(doc, taskElement, "WorkVariance", task.workVariance);
                appendTextElement(doc, taskElement, "TotalSlack", task.totalSlack);
                appendTextElement(doc, taskElement, "FreeSlack", task.freeSlack);
                appendTextElement(doc, taskElement, "Cost", task.cost);
                appendTextElement(doc, taskElement, "ActualCost", task.actualCost);
                appendTextElement(doc, taskElement, "RemainingCost", task.remainingCost);
                appendTextElement(doc, taskElement, "RemainingWork", task.remainingWork);
                appendTextElement(doc, taskElement, "ActualWork", task.actualWork);
                appendTextElement(doc, taskElement, "ConstraintType", task.constraintType);
                appendTextElement(doc, taskElement, "ConstraintDate", task.constraintDate);
                appendTextElement(doc, taskElement, "Milestone", task.milestone);
                appendTextElement(doc, taskElement, "Summary", task.summary);
                appendTextElement(doc, taskElement, "Critical", task.critical);
                appendTextElement(doc, taskElement, "PercentComplete", task.percentComplete);
                appendTextElement(doc, taskElement, "PercentWorkComplete", task.percentWorkComplete);
                appendTextElement(doc, taskElement, "Notes", task.notes);
                for (const attribute of task.extendedAttributes) {
                    const extendedAttributeElement = doc.createElement("ExtendedAttribute");
                    appendTextElement(doc, extendedAttributeElement, "FieldID", attribute.fieldID);
                    appendTextElement(doc, extendedAttributeElement, "Value", attribute.value);
                    taskElement.appendChild(extendedAttributeElement);
                }
                for (const baseline of task.baselines) {
                    const baselineElement = doc.createElement("Baseline");
                    appendTextElement(doc, baselineElement, "Number", baseline.number);
                    appendTextElement(doc, baselineElement, "Start", baseline.start);
                    appendTextElement(doc, baselineElement, "Finish", baseline.finish);
                    appendTextElement(doc, baselineElement, "Work", baseline.work);
                    appendTextElement(doc, baselineElement, "Cost", baseline.cost);
                    taskElement.appendChild(baselineElement);
                }
                for (const timephasedData of task.timephasedData) {
                    const timephasedDataElement = doc.createElement("TimephasedData");
                    appendTextElement(doc, timephasedDataElement, "Type", timephasedData.type);
                    appendTextElement(doc, timephasedDataElement, "UID", timephasedData.uid);
                    appendTextElement(doc, timephasedDataElement, "Start", timephasedData.start);
                    appendTextElement(doc, timephasedDataElement, "Finish", timephasedData.finish);
                    appendTextElement(doc, timephasedDataElement, "Unit", timephasedData.unit);
                    appendTextElement(doc, timephasedDataElement, "Value", timephasedData.value);
                    taskElement.appendChild(timephasedDataElement);
                }
                for (const predecessor of task.predecessors) {
                    const predecessorElement = doc.createElement("PredecessorLink");
                    appendTextElement(doc, predecessorElement, "PredecessorUID", predecessor.predecessorUid);
                    appendTextElement(doc, predecessorElement, "Type", predecessor.type);
                    appendTextElement(doc, predecessorElement, "LinkLag", predecessor.linkLag);
                    taskElement.appendChild(predecessorElement);
                }
                tasksElement.appendChild(taskElement);
            }
            project.appendChild(tasksElement);
            const resourcesElement = doc.createElement("Resources");
            for (const resource of normalizedModel.resources) {
                const resourceElement = doc.createElement("Resource");
                appendTextElement(doc, resourceElement, "UID", resource.uid);
                appendTextElement(doc, resourceElement, "ID", resource.id);
                appendTextElement(doc, resourceElement, "Name", resource.name);
                appendTextElement(doc, resourceElement, "Type", resource.type);
                appendTextElement(doc, resourceElement, "Initials", resource.initials);
                appendTextElement(doc, resourceElement, "Group", resource.group);
                appendTextElement(doc, resourceElement, "WorkGroup", resource.workGroup);
                appendTextElement(doc, resourceElement, "MaxUnits", resource.maxUnits);
                appendTextElement(doc, resourceElement, "CalendarUID", resource.calendarUID);
                appendTextElement(doc, resourceElement, "StandardRate", resource.standardRate);
                appendTextElement(doc, resourceElement, "StandardRateFormat", resource.standardRateFormat);
                appendTextElement(doc, resourceElement, "OvertimeRate", resource.overtimeRate);
                appendTextElement(doc, resourceElement, "OvertimeRateFormat", resource.overtimeRateFormat);
                appendTextElement(doc, resourceElement, "CostPerUse", resource.costPerUse);
                appendTextElement(doc, resourceElement, "Work", resource.work);
                appendTextElement(doc, resourceElement, "ActualWork", resource.actualWork);
                appendTextElement(doc, resourceElement, "RemainingWork", resource.remainingWork);
                appendTextElement(doc, resourceElement, "Cost", resource.cost);
                appendTextElement(doc, resourceElement, "ActualCost", resource.actualCost);
                appendTextElement(doc, resourceElement, "RemainingCost", resource.remainingCost);
                appendTextElement(doc, resourceElement, "PercentWorkComplete", resource.percentWorkComplete);
                for (const attribute of resource.extendedAttributes) {
                    const extendedAttributeElement = doc.createElement("ExtendedAttribute");
                    appendTextElement(doc, extendedAttributeElement, "FieldID", attribute.fieldID);
                    appendTextElement(doc, extendedAttributeElement, "Value", attribute.value);
                    resourceElement.appendChild(extendedAttributeElement);
                }
                for (const baseline of resource.baselines) {
                    const baselineElement = doc.createElement("Baseline");
                    appendTextElement(doc, baselineElement, "Number", baseline.number);
                    appendTextElement(doc, baselineElement, "Start", baseline.start);
                    appendTextElement(doc, baselineElement, "Finish", baseline.finish);
                    appendTextElement(doc, baselineElement, "Work", baseline.work);
                    appendTextElement(doc, baselineElement, "Cost", baseline.cost);
                    resourceElement.appendChild(baselineElement);
                }
                for (const timephasedData of resource.timephasedData) {
                    const timephasedDataElement = doc.createElement("TimephasedData");
                    appendTextElement(doc, timephasedDataElement, "Type", timephasedData.type);
                    appendTextElement(doc, timephasedDataElement, "UID", timephasedData.uid);
                    appendTextElement(doc, timephasedDataElement, "Start", timephasedData.start);
                    appendTextElement(doc, timephasedDataElement, "Finish", timephasedData.finish);
                    appendTextElement(doc, timephasedDataElement, "Unit", timephasedData.unit);
                    appendTextElement(doc, timephasedDataElement, "Value", timephasedData.value);
                    resourceElement.appendChild(timephasedDataElement);
                }
                resourcesElement.appendChild(resourceElement);
            }
            project.appendChild(resourcesElement);
            const assignmentsElement = doc.createElement("Assignments");
            for (const assignment of normalizedModel.assignments) {
                const assignmentElement = doc.createElement("Assignment");
                appendTextElement(doc, assignmentElement, "UID", assignment.uid);
                appendTextElement(doc, assignmentElement, "TaskUID", assignment.taskUid);
                appendTextElement(doc, assignmentElement, "ResourceUID", assignment.resourceUid);
                appendTextElement(doc, assignmentElement, "Start", assignment.start);
                appendTextElement(doc, assignmentElement, "Finish", assignment.finish);
                appendTextElement(doc, assignmentElement, "StartVariance", assignment.startVariance);
                appendTextElement(doc, assignmentElement, "FinishVariance", assignment.finishVariance);
                appendTextElement(doc, assignmentElement, "Delay", assignment.delay);
                appendTextElement(doc, assignmentElement, "Milestone", assignment.milestone);
                appendTextElement(doc, assignmentElement, "WorkContour", assignment.workContour);
                appendTextElement(doc, assignmentElement, "Units", assignment.units);
                appendTextElement(doc, assignmentElement, "Work", assignment.work);
                appendTextElement(doc, assignmentElement, "Cost", assignment.cost);
                appendTextElement(doc, assignmentElement, "ActualCost", assignment.actualCost);
                appendTextElement(doc, assignmentElement, "RemainingCost", assignment.remainingCost);
                appendTextElement(doc, assignmentElement, "PercentWorkComplete", assignment.percentWorkComplete);
                appendTextElement(doc, assignmentElement, "OvertimeWork", assignment.overtimeWork);
                appendTextElement(doc, assignmentElement, "ActualOvertimeWork", assignment.actualOvertimeWork);
                appendTextElement(doc, assignmentElement, "ActualWork", assignment.actualWork);
                appendTextElement(doc, assignmentElement, "RemainingWork", assignment.remainingWork);
                for (const attribute of assignment.extendedAttributes) {
                    const extendedAttributeElement = doc.createElement("ExtendedAttribute");
                    appendTextElement(doc, extendedAttributeElement, "FieldID", attribute.fieldID);
                    appendTextElement(doc, extendedAttributeElement, "Value", attribute.value);
                    assignmentElement.appendChild(extendedAttributeElement);
                }
                for (const baseline of assignment.baselines) {
                    const baselineElement = doc.createElement("Baseline");
                    appendTextElement(doc, baselineElement, "Number", baseline.number);
                    appendTextElement(doc, baselineElement, "Start", baseline.start);
                    appendTextElement(doc, baselineElement, "Finish", baseline.finish);
                    appendTextElement(doc, baselineElement, "Work", baseline.work);
                    appendTextElement(doc, baselineElement, "Cost", baseline.cost);
                    assignmentElement.appendChild(baselineElement);
                }
                for (const timephasedData of assignment.timephasedData) {
                    const timephasedDataElement = doc.createElement("TimephasedData");
                    appendTextElement(doc, timephasedDataElement, "Type", timephasedData.type);
                    appendTextElement(doc, timephasedDataElement, "UID", timephasedData.uid);
                    appendTextElement(doc, timephasedDataElement, "Start", timephasedData.start);
                    appendTextElement(doc, timephasedDataElement, "Finish", timephasedData.finish);
                    appendTextElement(doc, timephasedDataElement, "Unit", timephasedData.unit);
                    appendTextElement(doc, timephasedDataElement, "Value", timephasedData.value);
                    assignmentElement.appendChild(timephasedDataElement);
                }
                assignmentsElement.appendChild(assignmentElement);
            }
            project.appendChild(assignmentsElement);
            const serializer = new XMLSerializer();
            const serialized = serializer.serializeToString(doc);
            return `<?xml version="1.0" encoding="UTF-8"?>\n${input.formatXml(serialized)}\n`;
        }
    };
})();
