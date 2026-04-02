// @vitest-environment jsdom

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const typesCode = readFileSync(
  path.resolve(__dirname, "../src/js/types.js"),
  "utf8"
);
const msProjectXmlCode = readFileSync(
  path.resolve(__dirname, "../src/js/msproject-xml.js"),
  "utf8"
);
const minimalXml = readFileSync(
  path.resolve(__dirname, "../testdata/minimal.xml"),
  "utf8"
);
const hierarchyXml = readFileSync(
  path.resolve(__dirname, "../testdata/hierarchy.xml"),
  "utf8"
);
const dependencyXml = readFileSync(
  path.resolve(__dirname, "../testdata/dependency.xml"),
  "utf8"
);

function bootXmlModule() {
  new Function(`${typesCode}\n${msProjectXmlCode}`)();
  return globalThis.__mikuprojectXml;
}

describe("mikuproject msproject xml round-trip", () => {
  it("limits default calendar holiday exceptions to the project date range", () => {
    const xmlTools = bootXmlModule();

    const model = {
      project: {
        name: "祝日範囲確認",
        title: "祝日範囲確認",
        startDate: "2026-03-18T09:00:00",
        finishDate: "2026-03-22T18:00:00",
        scheduleFromStart: true,
        currentDate: "2026-03-20T09:00:00",
        calendarUID: undefined,
        outlineCodes: [],
        wbsMasks: [],
        extendedAttributes: []
      },
      tasks: [
        {
          uid: "1",
          id: "1",
          name: "確認タスク",
          outlineLevel: 1,
          outlineNumber: "1",
          start: "2026-03-18T09:00:00",
          finish: "2026-03-22T18:00:00",
          duration: "PT40H0M0S",
          milestone: false,
          summary: false,
          percentComplete: 0,
          extendedAttributes: [],
          baselines: [],
          timephasedData: [],
          predecessors: []
        }
      ],
      resources: [],
      assignments: [],
      calendars: []
    };

    const exportedXml = xmlTools.exportMsProjectXml(model);
    const reparsedModel = xmlTools.importMsProjectXml(exportedXml);

    expect(reparsedModel.calendars).toHaveLength(1);
    const holidayDates = reparsedModel.calendars[0].exceptions.map((exception) => exception.fromDate?.slice(0, 10));
    expect(holidayDates).toContain("2026-03-20");
    expect(holidayDates).not.toContain("2026-02-11");
    expect(holidayDates).not.toContain("2026-04-29");
  });

  it("round-trips the minimal xml sample", () => {
    const xmlTools = bootXmlModule();

    const model = xmlTools.importMsProjectXml(minimalXml);
    const exportedXml = xmlTools.exportMsProjectXml(model);
    const reparsedModel = xmlTools.importMsProjectXml(exportedXml);

    expect(model.project.name).toBe("Minimal Project");
    expect(reparsedModel.project.name).toBe("Minimal Project");
    expect(reparsedModel.project.title).toBeUndefined();
    expect(reparsedModel.tasks).toHaveLength(1);
    expect(reparsedModel.tasks[0].name).toBe("Single Task");
    expect(xmlTools.validateProjectModel(reparsedModel)).toHaveLength(0);
  });

  it("round-trips project metadata fields", () => {
    const xmlTools = bootXmlModule();
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<Project xmlns="http://schemas.microsoft.com/project">
  <Name>Metadata Project</Name>
  <Title>Metadata Title</Title>
  <Company>Example Company</Company>
  <Author>Example Author</Author>
  <CreationDate>2026-03-16T07:00:00</CreationDate>
  <LastSaved>2026-03-16T10:15:00</LastSaved>
  <SaveVersion>14</SaveVersion>
  <CurrencyCode>JPY</CurrencyCode>
  <CurrencyDigits>0</CurrencyDigits>
  <CurrencySymbol>¥</CurrencySymbol>
  <CurrencySymbolPosition>0</CurrencySymbolPosition>
  <FYStartDate>2026-04-01T00:00:00</FYStartDate>
  <FiscalYearStart>1</FiscalYearStart>
  <CriticalSlackLimit>0</CriticalSlackLimit>
  <DefaultTaskType>1</DefaultTaskType>
  <DefaultFixedCostAccrual>2</DefaultFixedCostAccrual>
  <DefaultStandardRate>5000/h</DefaultStandardRate>
  <DefaultOvertimeRate>7000/h</DefaultOvertimeRate>
  <DefaultTaskEVMethod>0</DefaultTaskEVMethod>
  <NewTaskStartDate>0</NewTaskStartDate>
  <NewTasksAreManual>0</NewTasksAreManual>
  <NewTasksEffortDriven>1</NewTasksEffortDriven>
  <NewTasksEstimated>1</NewTasksEstimated>
  <ActualsInSync>0</ActualsInSync>
  <EditableActualCosts>1</EditableActualCosts>
  <HonorConstraints>1</HonorConstraints>
  <InsertedProjectsLikeSummary>1</InsertedProjectsLikeSummary>
  <MultipleCriticalPaths>0</MultipleCriticalPaths>
  <TaskUpdatesResource>1</TaskUpdatesResource>
  <UpdateManuallyScheduledTasksWhenEditingLinks>0</UpdateManuallyScheduledTasksWhenEditingLinks>
  <OutlineCodes>
    <OutlineCode>
      <FieldID>188743731</FieldID>
      <FieldName>Outline Code1</FieldName>
      <Alias>Phase</Alias>
      <OnlyTableValues>1</OnlyTableValues>
      <Masks>
        <Mask>
          <Level>1</Level>
          <Mask>*</Mask>
          <Length>0</Length>
          <Sequence>0</Sequence>
        </Mask>
      </Masks>
      <Values>
        <Value>
          <Value>PLAN</Value>
          <Description>Planning</Description>
        </Value>
      </Values>
    </OutlineCode>
  </OutlineCodes>
  <WBSMasks>
    <WBSMask>
      <Level>1</Level>
      <Mask>A</Mask>
      <Length>1</Length>
      <Sequence>1</Sequence>
    </WBSMask>
  </WBSMasks>
  <ExtendedAttributes>
    <ExtendedAttribute>
      <FieldID>188743734</FieldID>
      <FieldName>Text1</FieldName>
      <Alias>Owner</Alias>
      <CalculationType>0</CalculationType>
      <RestrictValues>0</RestrictValues>
      <AppendNewValues>1</AppendNewValues>
    </ExtendedAttribute>
  </ExtendedAttributes>
  <StartDate>2026-03-16T09:00:00</StartDate>
  <FinishDate>2026-03-16T18:00:00</FinishDate>
  <ScheduleFromStart>1</ScheduleFromStart>
  <Tasks />
  <Resources />
  <Assignments />
</Project>`;

    const model = xmlTools.importMsProjectXml(xml);
    const exportedXml = xmlTools.exportMsProjectXml(model);
    const reparsedModel = xmlTools.importMsProjectXml(exportedXml);

    expect(reparsedModel.project.title).toBe("Metadata Title");
    expect(reparsedModel.project.company).toBe("Example Company");
    expect(reparsedModel.project.author).toBe("Example Author");
    expect(reparsedModel.project.creationDate).toBe("2026-03-16T07:00:00");
    expect(reparsedModel.project.lastSaved).toBe("2026-03-16T10:15:00");
    expect(reparsedModel.project.saveVersion).toBe(14);
    expect(reparsedModel.project.currencyCode).toBe("JPY");
    expect(reparsedModel.project.currencyDigits).toBe(0);
    expect(reparsedModel.project.currencySymbol).toBe("¥");
    expect(reparsedModel.project.currencySymbolPosition).toBe(0);
    expect(reparsedModel.project.fyStartDate).toBe("2026-04-01T00:00:00");
    expect(reparsedModel.project.fiscalYearStart).toBe(true);
    expect(reparsedModel.project.criticalSlackLimit).toBe(0);
    expect(reparsedModel.project.defaultTaskType).toBe(1);
    expect(reparsedModel.project.defaultFixedCostAccrual).toBe(2);
    expect(reparsedModel.project.defaultStandardRate).toBe("5000/h");
    expect(reparsedModel.project.defaultOvertimeRate).toBe("7000/h");
    expect(reparsedModel.project.defaultTaskEVMethod).toBe(0);
    expect(reparsedModel.project.newTaskStartDate).toBe(0);
    expect(reparsedModel.project.newTasksAreManual).toBe(false);
    expect(reparsedModel.project.newTasksEffortDriven).toBe(true);
    expect(reparsedModel.project.newTasksEstimated).toBe(true);
    expect(reparsedModel.project.actualsInSync).toBe(false);
    expect(reparsedModel.project.editableActualCosts).toBe(true);
    expect(reparsedModel.project.honorConstraints).toBe(true);
    expect(reparsedModel.project.insertedProjectsLikeSummary).toBe(true);
    expect(reparsedModel.project.multipleCriticalPaths).toBe(false);
    expect(reparsedModel.project.taskUpdatesResource).toBe(true);
    expect(reparsedModel.project.updateManuallyScheduledTasksWhenEditingLinks).toBe(false);
    expect(reparsedModel.project.outlineCodes).toHaveLength(1);
    expect(reparsedModel.project.outlineCodes[0].fieldID).toBe("188743731");
    expect(reparsedModel.project.outlineCodes[0].alias).toBe("Phase");
    expect(reparsedModel.project.outlineCodes[0].values[0].value).toBe("PLAN");
    expect(reparsedModel.project.wbsMasks).toHaveLength(1);
    expect(reparsedModel.project.wbsMasks[0].mask).toBe("A");
    expect(reparsedModel.project.extendedAttributes).toHaveLength(1);
    expect(reparsedModel.project.extendedAttributes[0].fieldName).toBe("Text1");
    expect(reparsedModel.project.extendedAttributes[0].alias).toBe("Owner");
    expect(reparsedModel.project.extendedAttributes[0].appendNewValues).toBe(true);
  });

  it("round-trips project scheduling metadata fields", () => {
    const xmlTools = bootXmlModule();
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<Project xmlns="http://schemas.microsoft.com/project">
  <Name>Schedule Metadata Project</Name>
  <StatusDate>2026-03-17T09:00:00</StatusDate>
  <WeekStartDay>2</WeekStartDay>
  <WorkFormat>2</WorkFormat>
  <DurationFormat>7</DurationFormat>
  <StartDate>2026-03-16T09:00:00</StartDate>
  <FinishDate>2026-03-18T18:00:00</FinishDate>
  <ScheduleFromStart>1</ScheduleFromStart>
  <Tasks />
  <Resources />
  <Assignments />
</Project>`;

    const model = xmlTools.importMsProjectXml(xml);
    const exportedXml = xmlTools.exportMsProjectXml(model);
    const reparsedModel = xmlTools.importMsProjectXml(exportedXml);

    expect(reparsedModel.project.statusDate).toBe("2026-03-17T09:00:00");
    expect(reparsedModel.project.weekStartDay).toBe(2);
    expect(reparsedModel.project.workFormat).toBe(2);
    expect(reparsedModel.project.durationFormat).toBe(7);
  });

  it("round-trips calendar base and weekday fields", () => {
    const xmlTools = bootXmlModule();
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<Project xmlns="http://schemas.microsoft.com/project">
  <Name>Calendar Detail Project</Name>
  <StartDate>2026-03-16T09:00:00</StartDate>
  <FinishDate>2026-03-16T18:00:00</FinishDate>
  <ScheduleFromStart>1</ScheduleFromStart>
  <Calendars>
    <Calendar>
      <UID>1</UID>
      <Name>Standard</Name>
      <IsBaseCalendar>1</IsBaseCalendar>
      <IsBaselineCalendar>1</IsBaselineCalendar>
    </Calendar>
    <Calendar>
      <UID>2</UID>
      <Name>Night Shift</Name>
      <IsBaseCalendar>0</IsBaseCalendar>
      <BaseCalendarUID>1</BaseCalendarUID>
      <WeekDays>
        <WeekDay>
          <DayType>7</DayType>
          <DayWorking>1</DayWorking>
          <WorkingTimes>
            <WorkingTime>
              <FromTime>18:00:00</FromTime>
              <ToTime>22:00:00</ToTime>
            </WorkingTime>
          </WorkingTimes>
        </WeekDay>
      </WeekDays>
    </Calendar>
  </Calendars>
  <Tasks />
  <Resources />
  <Assignments />
</Project>`;

    const model = xmlTools.importMsProjectXml(xml);
    const exportedXml = xmlTools.exportMsProjectXml(model);
    const reparsedModel = xmlTools.importMsProjectXml(exportedXml);

    expect(reparsedModel.calendars).toHaveLength(2);
    expect(reparsedModel.calendars[0].isBaselineCalendar).toBe(true);
    expect(reparsedModel.calendars[1].baseCalendarUID).toBe("1");
    expect(reparsedModel.calendars[1].weekDays[0].dayType).toBe(7);
    expect(reparsedModel.calendars[1].weekDays[0].dayWorking).toBe(true);
    expect(reparsedModel.calendars[1].weekDays[0].workingTimes[0].fromTime).toBe("18:00:00");
    expect(reparsedModel.calendars[1].weekDays[0].workingTimes[0].toTime).toBe("22:00:00");
  });

  it("round-trips calendar exceptions and workweeks", () => {
    const xmlTools = bootXmlModule();
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<Project xmlns="http://schemas.microsoft.com/project">
  <Name>Calendar Exception Project</Name>
  <StartDate>2026-03-16T09:00:00</StartDate>
  <FinishDate>2026-03-16T18:00:00</FinishDate>
  <ScheduleFromStart>1</ScheduleFromStart>
  <Calendars>
    <Calendar>
      <UID>1</UID>
      <Name>Standard</Name>
      <IsBaseCalendar>1</IsBaseCalendar>
      <Exceptions>
        <Exception>
          <Name>Holiday</Name>
          <FromDate>2026-03-20T00:00:00</FromDate>
          <ToDate>2026-03-20T23:59:59</ToDate>
          <DayWorking>0</DayWorking>
          <WorkingTimes>
            <WorkingTime>
              <FromTime>09:00:00</FromTime>
              <ToTime>12:00:00</ToTime>
            </WorkingTime>
          </WorkingTimes>
        </Exception>
      </Exceptions>
      <WorkWeeks>
        <WorkWeek>
          <Name>Sprint 1</Name>
          <FromDate>2026-03-16T00:00:00</FromDate>
          <ToDate>2026-03-31T23:59:59</ToDate>
          <WeekDays>
            <WeekDay>
              <DayType>2</DayType>
              <DayWorking>1</DayWorking>
              <WorkingTimes>
                <WorkingTime>
                  <FromTime>09:00:00</FromTime>
                  <ToTime>17:00:00</ToTime>
                </WorkingTime>
              </WorkingTimes>
            </WeekDay>
          </WeekDays>
        </WorkWeek>
      </WorkWeeks>
    </Calendar>
  </Calendars>
  <Tasks />
  <Resources />
  <Assignments />
</Project>`;

    const model = xmlTools.importMsProjectXml(xml);
    const exportedXml = xmlTools.exportMsProjectXml(model);
    const reparsedModel = xmlTools.importMsProjectXml(exportedXml);

    expect(reparsedModel.calendars[0].exceptions[0].name).toBe("Holiday");
    expect(reparsedModel.calendars[0].exceptions[0].dayWorking).toBe(false);
    expect(reparsedModel.calendars[0].exceptions[0].workingTimes[0].fromTime).toBe("09:00:00");
    expect(reparsedModel.calendars[0].exceptions[0].workingTimes[0].toTime).toBe("12:00:00");
    expect(reparsedModel.calendars[0].workWeeks[0].name).toBe("Sprint 1");
    expect(reparsedModel.calendars[0].workWeeks[0].weekDays[0].dayType).toBe(2);
    expect(reparsedModel.calendars[0].workWeeks[0].weekDays[0].workingTimes[0].toTime).toBe("17:00:00");
  });

  it("warns when calendar baseCalendarUID points to itself", () => {
    const xmlTools = bootXmlModule();
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<Project xmlns="http://schemas.microsoft.com/project">
  <Name>Self Base Calendar Project</Name>
  <StartDate>2026-03-16T09:00:00</StartDate>
  <FinishDate>2026-03-16T18:00:00</FinishDate>
  <ScheduleFromStart>1</ScheduleFromStart>
  <Calendars>
    <Calendar>
      <UID>1</UID>
      <Name>Loop Calendar</Name>
      <IsBaseCalendar>0</IsBaseCalendar>
      <BaseCalendarUID>1</BaseCalendarUID>
    </Calendar>
  </Calendars>
  <Tasks>
    <Task>
      <UID>1</UID>
      <ID>1</ID>
      <Name>Task</Name>
      <OutlineLevel>1</OutlineLevel>
      <OutlineNumber>1</OutlineNumber>
      <Start>2026-03-16T09:00:00</Start>
      <Finish>2026-03-16T18:00:00</Finish>
      <Duration>PT8H0M0S</Duration>
      <Milestone>0</Milestone>
      <Summary>0</Summary>
      <PercentComplete>0</PercentComplete>
    </Task>
  </Tasks>
</Project>`;

    const issues = xmlTools.validateProjectModel(xmlTools.importMsProjectXml(xml));

    expect(issues.some((issue) => issue.message.includes("BaseCalendarUID が自身を指しています"))).toBe(true);
  });

  it("round-trips resource and assignment practical fields", () => {
    const xmlTools = bootXmlModule();
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<Project xmlns="http://schemas.microsoft.com/project">
  <Name>Resource Assignment Project</Name>
  <StartDate>2026-03-16T09:00:00</StartDate>
  <FinishDate>2026-03-17T18:00:00</FinishDate>
  <ScheduleFromStart>1</ScheduleFromStart>
  <Calendars>
    <Calendar>
      <UID>1</UID>
      <Name>Standard</Name>
      <IsBaseCalendar>1</IsBaseCalendar>
    </Calendar>
  </Calendars>
  <Tasks>
    <Task>
      <UID>1</UID>
      <ID>1</ID>
      <Name>Assigned Task</Name>
      <OutlineLevel>1</OutlineLevel>
      <OutlineNumber>1</OutlineNumber>
      <Start>2026-03-16T09:00:00</Start>
      <Finish>2026-03-17T18:00:00</Finish>
      <Duration>PT16H0M0S</Duration>
      <Milestone>0</Milestone>
      <Summary>0</Summary>
      <PercentComplete>0</PercentComplete>
    </Task>
  </Tasks>
  <Resources>
    <Resource>
      <UID>1</UID>
      <ID>1</ID>
      <Name>Worker</Name>
      <Type>1</Type>
      <WorkGroup>0</WorkGroup>
      <CalendarUID>1</CalendarUID>
      <StandardRate>8000/h</StandardRate>
      <StandardRateFormat>2</StandardRateFormat>
      <OvertimeRate>12000/h</OvertimeRate>
      <OvertimeRateFormat>2</OvertimeRateFormat>
      <CostPerUse>1500</CostPerUse>
      <Work>PT24H0M0S</Work>
      <ActualWork>PT8H0M0S</ActualWork>
      <RemainingWork>PT16H0M0S</RemainingWork>
      <Cost>180000</Cost>
      <ActualCost>60000</ActualCost>
      <RemainingCost>120000</RemainingCost>
      <PercentWorkComplete>33</PercentWorkComplete>
    </Resource>
  </Resources>
  <Assignments>
    <Assignment>
      <UID>1</UID>
      <TaskUID>1</TaskUID>
      <ResourceUID>1</ResourceUID>
      <Start>2026-03-16T09:00:00</Start>
      <Finish>2026-03-17T18:00:00</Finish>
      <StartVariance>PT1H0M0S</StartVariance>
      <FinishVariance>PT2H0M0S</FinishVariance>
      <Delay>PT3H0M0S</Delay>
      <Milestone>0</Milestone>
      <WorkContour>1</WorkContour>
      <Units>1</Units>
      <Work>PT16H0M0S</Work>
      <Cost>100000</Cost>
      <ActualCost>30000</ActualCost>
      <RemainingCost>70000</RemainingCost>
      <PercentWorkComplete>50</PercentWorkComplete>
      <OvertimeWork>PT2H0M0S</OvertimeWork>
      <ActualOvertimeWork>PT1H0M0S</ActualOvertimeWork>
      <ActualWork>PT6H0M0S</ActualWork>
      <RemainingWork>PT10H0M0S</RemainingWork>
    </Assignment>
  </Assignments>
</Project>`;

    const model = xmlTools.importMsProjectXml(xml);
    const exportedXml = xmlTools.exportMsProjectXml(model);
    const reparsedModel = xmlTools.importMsProjectXml(exportedXml);

    expect(reparsedModel.resources[0].calendarUID).toBe("1");
    expect(reparsedModel.resources[0].workGroup).toBe(0);
    expect(reparsedModel.resources[0].standardRate).toBe("8000/h");
    expect(reparsedModel.resources[0].standardRateFormat).toBe(2);
    expect(reparsedModel.resources[0].overtimeRate).toBe("12000/h");
    expect(reparsedModel.resources[0].overtimeRateFormat).toBe(2);
    expect(reparsedModel.resources[0].costPerUse).toBe(1500);
    expect(reparsedModel.resources[0].work).toBe("PT24H0M0S");
    expect(reparsedModel.resources[0].actualWork).toBe("PT8H0M0S");
    expect(reparsedModel.resources[0].remainingWork).toBe("PT16H0M0S");
    expect(reparsedModel.resources[0].cost).toBe(180000);
    expect(reparsedModel.resources[0].actualCost).toBe(60000);
    expect(reparsedModel.resources[0].remainingCost).toBe(120000);
    expect(reparsedModel.resources[0].percentWorkComplete).toBe(33);
    expect(reparsedModel.assignments[0].startVariance).toBe("PT1H0M0S");
    expect(reparsedModel.assignments[0].finishVariance).toBe("PT2H0M0S");
    expect(reparsedModel.assignments[0].delay).toBe("PT3H0M0S");
    expect(reparsedModel.assignments[0].milestone).toBe(false);
    expect(reparsedModel.assignments[0].workContour).toBe(1);
    expect(reparsedModel.assignments[0].cost).toBe(100000);
    expect(reparsedModel.assignments[0].actualCost).toBe(30000);
    expect(reparsedModel.assignments[0].remainingCost).toBe(70000);
    expect(reparsedModel.assignments[0].percentWorkComplete).toBe(50);
    expect(reparsedModel.assignments[0].overtimeWork).toBe("PT2H0M0S");
    expect(reparsedModel.assignments[0].actualOvertimeWork).toBe("PT1H0M0S");
    expect(reparsedModel.assignments[0].actualWork).toBe("PT6H0M0S");
    expect(reparsedModel.assignments[0].remainingWork).toBe("PT10H0M0S");
  });

  it("round-trips task and assignment cost fields", () => {
    const xmlTools = bootXmlModule();
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<Project xmlns="http://schemas.microsoft.com/project">
  <Name>Cost Project</Name>
  <StartDate>2026-03-16T09:00:00</StartDate>
  <FinishDate>2026-03-18T18:00:00</FinishDate>
  <ScheduleFromStart>1</ScheduleFromStart>
  <Tasks>
    <Task>
      <UID>1</UID>
      <ID>1</ID>
      <Name>Cost Task</Name>
      <OutlineLevel>1</OutlineLevel>
      <OutlineNumber>1</OutlineNumber>
      <Start>2026-03-16T09:00:00</Start>
      <Finish>2026-03-18T18:00:00</Finish>
      <Duration>PT24H0M0S</Duration>
      <Work>PT24H0M0S</Work>
      <Cost>150000</Cost>
      <ActualCost>50000</ActualCost>
      <RemainingCost>100000</RemainingCost>
      <Milestone>0</Milestone>
      <Summary>0</Summary>
      <PercentComplete>0</PercentComplete>
    </Task>
  </Tasks>
  <Resources />
  <Assignments>
    <Assignment>
      <UID>1</UID>
      <TaskUID>1</TaskUID>
      <ResourceUID>-65535</ResourceUID>
      <Start>2026-03-16T09:00:00</Start>
      <Finish>2026-03-18T18:00:00</Finish>
      <Units>1</Units>
      <Work>PT24H0M0S</Work>
      <Cost>150000</Cost>
      <ActualCost>50000</ActualCost>
      <RemainingCost>100000</RemainingCost>
    </Assignment>
  </Assignments>
</Project>`;

    const model = xmlTools.importMsProjectXml(xml);
    const exportedXml = xmlTools.exportMsProjectXml(model);
    const reparsedModel = xmlTools.importMsProjectXml(exportedXml);

    expect(reparsedModel.tasks[0].cost).toBe(150000);
    expect(reparsedModel.tasks[0].actualCost).toBe(50000);
    expect(reparsedModel.tasks[0].remainingCost).toBe(100000);
    expect(reparsedModel.assignments[0].cost).toBe(150000);
    expect(reparsedModel.assignments[0].actualCost).toBe(50000);
    expect(reparsedModel.assignments[0].remainingCost).toBe(100000);
  });

  it("round-trips task deadline and variance fields", () => {
    const xmlTools = bootXmlModule();
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<Project xmlns="http://schemas.microsoft.com/project">
  <Name>Task Variance Project</Name>
  <StartDate>2026-03-16T09:00:00</StartDate>
  <FinishDate>2026-03-18T18:00:00</FinishDate>
  <ScheduleFromStart>1</ScheduleFromStart>
  <Tasks>
    <Task>
      <UID>1</UID>
      <ID>1</ID>
      <Name>Variance Task</Name>
      <OutlineLevel>1</OutlineLevel>
      <OutlineNumber>1</OutlineNumber>
      <Start>2026-03-16T09:00:00</Start>
      <Finish>2026-03-18T18:00:00</Finish>
      <Deadline>2026-03-19T18:00:00</Deadline>
      <Duration>PT24H0M0S</Duration>
      <StartVariance>PT1H0M0S</StartVariance>
      <FinishVariance>PT2H0M0S</FinishVariance>
      <Milestone>0</Milestone>
      <Summary>0</Summary>
      <PercentComplete>0</PercentComplete>
    </Task>
  </Tasks>
  <Resources />
  <Assignments />
</Project>`;

    const model = xmlTools.importMsProjectXml(xml);
    const exportedXml = xmlTools.exportMsProjectXml(model);
    const reparsedModel = xmlTools.importMsProjectXml(exportedXml);

    expect(reparsedModel.tasks[0].deadline).toBe("2026-03-19T18:00:00");
    expect(reparsedModel.tasks[0].startVariance).toBe("PT1H0M0S");
    expect(reparsedModel.tasks[0].finishVariance).toBe("PT2H0M0S");
  });

  it("round-trips extended task work fields", () => {
    const xmlTools = bootXmlModule();
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<Project xmlns="http://schemas.microsoft.com/project">
  <Name>Task Detail Project</Name>
  <StartDate>2026-03-16T09:00:00</StartDate>
  <FinishDate>2026-03-17T18:00:00</FinishDate>
  <ScheduleFromStart>1</ScheduleFromStart>
  <Tasks>
    <Task>
      <UID>1</UID>
      <ID>1</ID>
      <Name>Detailed Task</Name>
      <OutlineLevel>1</OutlineLevel>
      <OutlineNumber>1</OutlineNumber>
      <WBS>1</WBS>
      <Type>1</Type>
      <CalendarUID>1</CalendarUID>
      <Priority>700</Priority>
      <Start>2026-03-16T09:00:00</Start>
      <Finish>2026-03-17T18:00:00</Finish>
      <Duration>PT16H0M0S</Duration>
      <Work>PT16H0M0S</Work>
      <WorkVariance>PT1H0M0S</WorkVariance>
      <TotalSlack>PT4H0M0S</TotalSlack>
      <FreeSlack>PT2H0M0S</FreeSlack>
      <RemainingWork>PT8H0M0S</RemainingWork>
      <ActualWork>PT8H0M0S</ActualWork>
      <Milestone>0</Milestone>
      <Summary>0</Summary>
      <Critical>1</Critical>
      <PercentComplete>50</PercentComplete>
      <PercentWorkComplete>50</PercentWorkComplete>
    </Task>
  </Tasks>
  <Resources />
  <Assignments />
</Project>`;

    const model = xmlTools.importMsProjectXml(xml);
    const exportedXml = xmlTools.exportMsProjectXml(model);
    const reparsedModel = xmlTools.importMsProjectXml(exportedXml);

    expect(reparsedModel.tasks[0].wbs).toBe("1");
    expect(reparsedModel.tasks[0].type).toBe(1);
    expect(reparsedModel.tasks[0].calendarUID).toBe("1");
    expect(reparsedModel.tasks[0].priority).toBe(700);
    expect(reparsedModel.tasks[0].work).toBe("PT16H0M0S");
    expect(reparsedModel.tasks[0].workVariance).toBe("PT1H0M0S");
    expect(reparsedModel.tasks[0].totalSlack).toBe("PT4H0M0S");
    expect(reparsedModel.tasks[0].freeSlack).toBe("PT2H0M0S");
    expect(reparsedModel.tasks[0].remainingWork).toBe("PT8H0M0S");
    expect(reparsedModel.tasks[0].actualWork).toBe("PT8H0M0S");
    expect(reparsedModel.tasks[0].critical).toBe(true);
    expect(reparsedModel.tasks[0].percentWorkComplete).toBe(50);
  });

  it("round-trips the hierarchy xml sample", () => {
    const xmlTools = bootXmlModule();

    const model = xmlTools.importMsProjectXml(hierarchyXml);
    const exportedXml = xmlTools.exportMsProjectXml(model);
    const reparsedModel = xmlTools.importMsProjectXml(exportedXml);

    expect(reparsedModel.tasks).toHaveLength(3);
    expect(reparsedModel.tasks[0].summary).toBe(true);
    expect(reparsedModel.tasks[1].outlineNumber).toBe("1.1");
    expect(reparsedModel.tasks[2].notes).toBe("Second child task");
    expect(xmlTools.validateProjectModel(reparsedModel)).toHaveLength(0);
  });

  it("round-trips the dependency xml sample", () => {
    const xmlTools = bootXmlModule();

    const model = xmlTools.importMsProjectXml(dependencyXml);
    const exportedXml = xmlTools.exportMsProjectXml(model);
    const reparsedModel = xmlTools.importMsProjectXml(exportedXml);

    expect(reparsedModel.calendars).toHaveLength(1);
    expect(reparsedModel.tasks[1].predecessors).toHaveLength(1);
    expect(reparsedModel.tasks[1].predecessors[0].predecessorUid).toBe("1");
    expect(reparsedModel.assignments).toHaveLength(1);
    expect(reparsedModel.assignments[0].taskUid).toBe("2");
    expect(xmlTools.validateProjectModel(reparsedModel)).toHaveLength(0);
  });

  it("imports csv with parent id into a minimal project model", () => {
    const xmlTools = bootXmlModule();
    const csv = `ID,ParentID,WBS,Name,Start,Finish,PredecessorID,Resource,PercentComplete
1,,1,Project Summary,2026-03-16T09:00:00,2026-03-20T18:00:00,,,50
2,1,1.1,Design,2026-03-16T09:00:00,2026-03-17T18:00:00,,Miku,100
3,1,1.2,Implementation,2026-03-18T09:00:00,2026-03-20T18:00:00,2,Miku|Rin,0
4,3,1.2.1,Coding,2026-03-18T09:00:00,2026-03-19T18:00:00,2,Rin,20
`;

    const model = xmlTools.importCsvParentId(csv);

    expect(model.project.name).toBe("CSV Imported Project");
    expect(model.project.title).toBe("CSV Imported Project");
    expect(model.project.calendarUID).toBe("1");
    expect(model.tasks).toHaveLength(4);
    expect(model.tasks[0].summary).toBe(true);
    expect(model.tasks[1].outlineNumber).toBe("1.1");
    expect(model.tasks[2].predecessors[0].predecessorUid).toBe("2");
    expect(model.tasks[3].outlineLevel).toBe(3);
    expect(model.resources.map((item) => item.name)).toEqual(["Miku", "Rin"]);
    expect(model.assignments).toHaveLength(4);
    expect(model.assignments[1].resourceUid).toBe("1");
    expect(model.assignments[2].resourceUid).toBe("2");
    expect(model.calendars).toHaveLength(1);
    expect(model.calendars[0].name).toBe("Standard");
    expect(model.calendars[0].weekDays).toHaveLength(7);
    expect(model.calendars[0].exceptions.length).toBeGreaterThan(0);
  });

  it("imports extended task fields in csv with parent id", () => {
    const xmlTools = bootXmlModule();
    const csv = `ID,ParentID,WBS,Name,Start,Finish,PredecessorID,Resource,PercentComplete,PercentWorkComplete,Milestone,Summary,Critical,Type,Priority,Work,CalendarUID,ConstraintType,ConstraintDate,Deadline,Notes
1,,1,Project Summary,2026-03-16T09:00:00,2026-03-20T18:00:00,,,,,1,1,0,1,500,PT40H0M0S,1,,,,Root note
2,1,1.1,Design,2026-03-16T09:00:00,2026-03-17T18:00:00,,Miku,100,100,0,0,0,1,600,PT16H0M0S,1,,,,Design done
3,1,1.2,Release,2026-03-20T18:00:00,2026-03-20T18:00:00,2,Miku,100,100,1,0,1,1,700,PT0H0M0S,2,4,2026-03-20T09:00:00,2026-03-21T18:00:00,Release gate
`;

    const model = xmlTools.importCsvParentId(csv);

    expect(model.tasks[0].summary).toBe(true);
    expect(model.tasks[0].notes).toBe("Root note");
    expect(model.tasks[1].percentWorkComplete).toBe(100);
    expect(model.tasks[1].critical).toBe(false);
    expect(model.tasks[1].priority).toBe(600);
    expect(model.tasks[1].work).toBe("PT16H0M0S");
    expect(model.tasks[2].milestone).toBe(true);
    expect(model.tasks[2].critical).toBe(true);
    expect(model.tasks[2].type).toBe(1);
    expect(model.tasks[2].calendarUID).toBe("2");
    expect(model.tasks[2].constraintType).toBe(4);
    expect(model.tasks[2].constraintDate).toBe("2026-03-20T09:00:00");
    expect(model.tasks[2].deadline).toBe("2026-03-21T18:00:00");
    expect(model.tasks[2].work).toBe("PT0H0M0S");
    expect(model.tasks[2].notes).toBe("Release gate");
  });

  it("normalizes predecessor and resource separators in csv import", () => {
    const xmlTools = bootXmlModule();
    const csv = `ID,ParentID,WBS,Name,Start,Finish,PredecessorID,Resource,PercentComplete
1,,1,Project Summary,2026-03-16T09:00:00,2026-03-20T18:00:00,,,50
2,1,1.1,Design,2026-03-16T09:00:00,2026-03-17T18:00:00,,Miku,100
3,1,1.2,Implementation,2026-03-18T09:00:00,2026-03-20T18:00:00,"2, 4; 2","Miku; Rin、Luka| Rin",0
4,1,1.3,Review,2026-03-20T09:00:00,2026-03-20T18:00:00,,Luka,0
`;

    const model = xmlTools.importCsvParentId(csv);

    expect(model.tasks[2].predecessors.map((item) => item.predecessorUid)).toEqual(["2", "4"]);
    expect(model.resources.map((item) => item.name)).toEqual(["Miku", "Rin", "Luka"]);
    expect(model.assignments.filter((item) => item.taskUid === "3")).toHaveLength(3);
  });

  it("rejects duplicate id in csv import", () => {
    const xmlTools = bootXmlModule();
    const csv = `ID,ParentID,Name
1,,Root
1,,Duplicate
`;

    expect(() => xmlTools.importCsvParentId(csv)).toThrow("CSV の ID が重複しています");
  });

  it("rejects missing parent id in csv import", () => {
    const xmlTools = bootXmlModule();
    const csv = `ID,ParentID,Name
1,,Root
2,99,Child
`;

    expect(() => xmlTools.importCsvParentId(csv)).toThrow("CSV の ParentID が既存 ID を指していません");
  });

  it("rejects cyclic parent id in csv import", () => {
    const xmlTools = bootXmlModule();
    const csv = `ID,ParentID,Name
1,2,Root
2,1,Child
`;

    expect(() => xmlTools.importCsvParentId(csv)).toThrow("CSV の ParentID が循環しています");
  });

  it("allows placeholder UID=0 and unassigned ResourceUID=-65535", () => {
    const xmlTools = bootXmlModule();
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<Project xmlns="http://schemas.microsoft.com/project">
  <Name>Placeholder Project</Name>
  <StartDate>2026-03-16T09:00:00</StartDate>
  <FinishDate>2026-03-16T18:00:00</FinishDate>
  <ScheduleFromStart>1</ScheduleFromStart>
  <Tasks>
    <Task>
      <UID>0</UID>
      <ID>0</ID>
      <Name></Name>
      <OutlineLevel>0</OutlineLevel>
      <OutlineNumber></OutlineNumber>
      <Start>2026-03-16T09:00:00</Start>
      <Finish>2026-03-16T18:00:00</Finish>
      <Duration>PT8H0M0S</Duration>
      <Milestone>0</Milestone>
      <Summary>0</Summary>
      <PercentComplete>0</PercentComplete>
    </Task>
  </Tasks>
  <Resources>
    <Resource>
      <UID>0</UID>
      <ID>0</ID>
      <Name></Name>
      <Type>1</Type>
    </Resource>
  </Resources>
  <Assignments>
    <Assignment>
      <UID>1</UID>
      <TaskUID>0</TaskUID>
      <ResourceUID>-65535</ResourceUID>
      <Start>2026-03-16T09:00:00</Start>
      <Finish>2026-03-16T18:00:00</Finish>
      <Units>1</Units>
      <Work>PT8H0M0S</Work>
    </Assignment>
  </Assignments>
</Project>`;

    const model = xmlTools.importMsProjectXml(xml);
    const issues = xmlTools.validateProjectModel(model);

    expect(issues.some((issue) => issue.message.includes("OutlineLevel"))).toBe(false);
    expect(issues.some((issue) => issue.message.includes("ResourceUID が既存 Resource"))).toBe(false);
    expect(issues.some((issue) => issue.message.includes("Resource Name が空"))).toBe(false);
  });
});
