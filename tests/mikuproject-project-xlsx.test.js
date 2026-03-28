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
const excelIoCode = readFileSync(
  path.resolve(__dirname, "../src/js/excel-io.js"),
  "utf8"
);
const msProjectXmlCode = readFileSync(
  path.resolve(__dirname, "../src/js/msproject-xml.js"),
  "utf8"
);
const projectXlsxCode = readFileSync(
  path.resolve(__dirname, "../src/js/project-xlsx.js"),
  "utf8"
);

function bootModules() {
  new Function(`${typesCode}\n${excelIoCode}\n${msProjectXmlCode}\n${projectXlsxCode}`)();
  return {
    excelIo: globalThis.__mikuprojectExcelIo,
    xml: globalThis.__mikuprojectXml,
    projectXlsx: globalThis.__mikuprojectProjectXlsx
  };
}

function buildSampleWorkbook(mutator) {
  const { xml, projectXlsx } = bootModules();
  const model = xml.importMsProjectXml(xml.SAMPLE_XML);
  const workbook = projectXlsx.exportProjectWorkbook(model);
  if (mutator) {
    mutator(workbook);
  }
  return { xml, projectXlsx, model, workbook };
}

const SAMPLE_HOLIDAY_COUNT = 89;
const SAMPLE_FIRST_HOLIDAY_NAME = "元日（公式）";
const SAMPLE_FIRST_HOLIDAY_DATE = "2026-01-01";
const EDITABLE_FILL = "#FDE7C7";

describe("mikuproject project xlsx", () => {
  it("converts ProjectModel into workbook sheets", () => {
    const { xml, projectXlsx } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);

    const workbook = projectXlsx.exportProjectWorkbook(model);

    expect(workbook.sheets.map((sheet) => sheet.name)).toEqual([
      "Project",
      "Tasks",
      "Resources",
      "Assignments",
      "Calendars",
      "NonWorkingDays"
    ]);
    expect(workbook.sheets[0].mergedRanges).toEqual(["A11:B11"]);
    expect(workbook.sheets[0].rows[0].cells[0].value).toBe("Project");
    expect(workbook.sheets[0].rows[0].cells[0].bold).toBe(true);
    expect(workbook.sheets[0].rows[0].cells[0].fillColor).toBe("#BFD7EA");
    expect(workbook.sheets[0].rows[0].cells[1].fillColor).toBe("#BFD7EA");
    expect(workbook.sheets[0].rows[1].cells[0].value).toBe("Basic Info");
    expect(workbook.sheets[0].rows[1].cells[1].fillColor).toBe("#BFD7EA");
    expect(workbook.sheets[0].rows[2].cells[0].value).toBe("Field");
    expect(workbook.sheets[0].rows[3].cells[0].value).toBe("Name");
    expect(workbook.sheets[0].rows[3].cells[1].value).toBe("Sample Project");
    expect(workbook.sheets[0].rows[3].cells[1].fillColor).toBe(EDITABLE_FILL);
    expect(workbook.sheets[0].rows[4].cells[1].fillColor).toBe(EDITABLE_FILL);
    expect(workbook.sheets[0].rows[5].cells[1].fillColor).toBe(EDITABLE_FILL);
    expect(workbook.sheets[0].rows[6].cells[1].fillColor).toBe(EDITABLE_FILL);
    expect(workbook.sheets[0].rows[7].cells[1].value).toBe("2026-03-16 09:00:00");
    expect(workbook.sheets[0].rows[11].cells[0].value).toBe("Settings");
    expect(workbook.sheets[0].rows[12].cells[1].fillColor).toBe(EDITABLE_FILL);
  });

  it("maps tasks, resources, assignments, and calendars to tabular rows", () => {
    const { xml, projectXlsx } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);

    const workbook = projectXlsx.exportProjectWorkbook(model);
    const tasksSheet = workbook.sheets.find((sheet) => sheet.name === "Tasks");
    const resourcesSheet = workbook.sheets.find((sheet) => sheet.name === "Resources");
    const assignmentsSheet = workbook.sheets.find((sheet) => sheet.name === "Assignments");
    const calendarsSheet = workbook.sheets.find((sheet) => sheet.name === "Calendars");
    const nonWorkingDaysSheet = workbook.sheets.find((sheet) => sheet.name === "NonWorkingDays");

    expect(tasksSheet.mergedRanges).toEqual([]);
    expect(tasksSheet.rows[0].cells[0].value).toBe("Tasks");
    expect(tasksSheet.rows[0].cells[0].fillColor).toBe("#D4E0EC");
    expect(tasksSheet.rows[1].cells[0].value).toBe("Task List");
    expect(tasksSheet.rows[2].cells.map((cell) => cell.value)).toEqual([
      "UID",
      "ID",
      "Name",
      "OutlineLevel",
      "OutlineNumber",
      "WBS",
      "Start",
      "Finish",
      "Duration",
      "PercentComplete",
      "PercentWorkComplete",
      "Milestone",
      "Summary",
      "Critical",
      "CalendarUID",
      "Predecessors",
      "Notes"
    ]);
    expect(tasksSheet.rows[3].cells[2].value).toBe("Project Summary");
    expect(tasksSheet.rows[3].cells[6].value).toBe("2026-03-16 09:00:00");
    expect(tasksSheet.rows[3].cells[2].bold).toBe(true);
    expect(tasksSheet.rows[3].cells[2].fillColor).toBe(EDITABLE_FILL);
    expect(tasksSheet.rows[3].cells[12].fillColor).toBe("#E6F2E0");
    expect(tasksSheet.rows[5].cells[2].value).toBe("Implementation");
    expect(tasksSheet.rows[5].cells[8].fillColor).toBe("#FBF6ED");
    expect(tasksSheet.rows[5].cells[13].fillColor).toBe("#F8DDE6");
    expect(tasksSheet.rows[5].cells[15].value).toBe("2");
    expect(tasksSheet.rows[5].cells[15].fillColor).toBe("#EEF7F4");
    expect(tasksSheet.rows[5].cells[16].value).toBe("Implementation starts after design");
    expect(tasksSheet.rows[5].cells[16].fillColor).toBe(EDITABLE_FILL);

    expect(resourcesSheet.mergedRanges).toEqual([]);
    expect(resourcesSheet.rows[0].cells[0].value).toBe("Resources");
    expect(resourcesSheet.rows[0].cells[0].fillColor).toBe("#C8E3D8");
    expect(resourcesSheet.rows[1].cells[0].value).toBe("Resource List");
    expect(resourcesSheet.rows[2].cells.map((cell) => cell.value)).toEqual([
      "UID",
      "ID",
      "Name",
      "Type",
      "Initials",
      "Group",
      "MaxUnits",
      "CalendarUID",
      "StandardRate",
      "OvertimeRate",
      "CostPerUse",
      "Work",
      "ActualWork",
      "RemainingWork"
    ]);
    expect(resourcesSheet.rows[3].cells[2].value).toBe("Miku");
    expect(resourcesSheet.rows[3].cells[2].fillColor).toBe(EDITABLE_FILL);
    expect(resourcesSheet.rows[3].cells[5].fillColor).toBe(EDITABLE_FILL);
    expect(resourcesSheet.rows[3].cells[6].fillColor).toBe(EDITABLE_FILL);

    expect(assignmentsSheet.mergedRanges).toEqual([]);
    expect(assignmentsSheet.rows[0].cells[0].value).toBe("Assignments");
    expect(assignmentsSheet.rows[0].cells[0].fillColor).toBe("#D7D2EC");
    expect(assignmentsSheet.rows[1].cells[0].value).toBe("Assignment List");
    expect(assignmentsSheet.rows[2].cells.map((cell) => cell.value)).toEqual([
      "UID",
      "TaskUID",
      "TaskName",
      "ResourceUID",
      "ResourceName",
      "Start",
      "Finish",
      "Units",
      "Work",
      "ActualWork",
      "RemainingWork",
      "PercentWorkComplete"
    ]);
    expect(assignmentsSheet.rows[3].cells[2].value).toBe("Design");
    expect(assignmentsSheet.rows[3].cells[4].value).toBe("Miku");
    expect(assignmentsSheet.rows[3].cells[5].value).toBe("2026-03-16 09:00:00");
    expect(assignmentsSheet.rows[3].cells[7].fillColor).toBe(EDITABLE_FILL);
    expect(assignmentsSheet.rows[3].cells[8].fillColor).toBe(EDITABLE_FILL);
    expect(assignmentsSheet.rows[3].cells[11].fillColor).toBe(EDITABLE_FILL);

    expect(calendarsSheet.mergedRanges).toEqual([]);
    expect(calendarsSheet.rows[0].cells[0].value).toBe("Calendars");
    expect(calendarsSheet.rows[0].cells[0].fillColor).toBe("#D7E3C4");
    expect(calendarsSheet.rows[1].cells[0].value).toBe("Calendar List");
    expect(calendarsSheet.rows[2].cells.map((cell) => cell.value)).toEqual([
      "UID",
      "Name",
      "IsBaseCalendar",
      "BaseCalendarUID",
      "WeekDays",
      "Exceptions",
      "WorkWeeks"
    ]);
    expect(calendarsSheet.rows[3].cells[1].value).toBe("Standard");
    expect(calendarsSheet.rows[3].cells[1].fillColor).toBe(EDITABLE_FILL);
    expect(calendarsSheet.rows[3].cells[2].fillColor).toBe(EDITABLE_FILL);

    expect(nonWorkingDaysSheet.mergedRanges).toEqual([]);
    expect(nonWorkingDaysSheet.rows[0].cells[0].value).toBe("NonWorkingDays");
    expect(nonWorkingDaysSheet.rows[0].cells[0].fillColor).toBe("#E9C7D5");
    expect(nonWorkingDaysSheet.rows[1].cells[0].value).toBe("Calendar Exceptions");
    expect(nonWorkingDaysSheet.rows[2].cells.map((cell) => cell.value)).toEqual([
      "CalendarUID",
      "Index",
      "CalendarName",
      "Name",
      "Date",
      "FromDate",
      "ToDate",
      "DayWorking"
    ]);
    expect(nonWorkingDaysSheet.rows[3].cells[0].value).toBe("1");
    expect(nonWorkingDaysSheet.rows[3].cells[3].value).toBe(SAMPLE_FIRST_HOLIDAY_NAME);
    expect(nonWorkingDaysSheet.rows[3].cells[3].fillColor).toBe(EDITABLE_FILL);
    expect(nonWorkingDaysSheet.rows[3].cells[4].value).toBe(SAMPLE_FIRST_HOLIDAY_DATE);
    expect(nonWorkingDaysSheet.rows[3].cells[5].value).toBe(SAMPLE_FIRST_HOLIDAY_DATE);
    expect(nonWorkingDaysSheet.rows[3].cells[6].value).toBe(SAMPLE_FIRST_HOLIDAY_DATE);
    expect(nonWorkingDaysSheet.rows[3].cells[7].value).toBe(false);
    expect(nonWorkingDaysSheet.rows[3].cells[4].fillColor).toBe(EDITABLE_FILL);
    expect(nonWorkingDaysSheet.rows[3].cells[5].fillColor).toBe(EDITABLE_FILL);
    expect(nonWorkingDaysSheet.rows[3].cells[6].fillColor).toBe(EDITABLE_FILL);
    expect(nonWorkingDaysSheet.rows[3].cells[7].fillColor).toBe(EDITABLE_FILL);
    expect(nonWorkingDaysSheet.rows).toHaveLength(SAMPLE_HOLIDAY_COUNT + 3);
  });

  it("can generate a real xlsx from ProjectModel through the workbook adapter", () => {
    const { excelIo, xml, projectXlsx } = bootModules();
    const codec = new excelIo.XlsxWorkbookCodec();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);

    const workbook = projectXlsx.exportProjectWorkbook(model);
    const bytes = codec.exportWorkbook(workbook);
    const entries = codec.listEntries(bytes);
    const projectSheetXml = new TextDecoder().decode(codec.unpackEntries(bytes)["xl/worksheets/sheet1.xml"]);

    expect(entries).toContain("xl/workbook.xml");
    expect(entries).toContain("xl/worksheets/sheet1.xml");
    expect(entries).toContain("xl/styles.xml");
    expect(projectSheetXml).not.toContain('ref="A1:B1"');
    expect(projectSheetXml).toContain('ref="A11:B11"');
  });

  it("imports limited task fields from workbook rows back into ProjectModel", () => {
    const { xml, projectXlsx } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);
    const workbook = projectXlsx.exportProjectWorkbook(model);
    const tasksSheet = workbook.sheets.find((sheet) => sheet.name === "Tasks");

    tasksSheet.rows[4].cells[2].value = "Design Updated";
    tasksSheet.rows[4].cells[6].value = "2026-03-17 09:00:00";
    tasksSheet.rows[4].cells[7].value = "2026-03-18 18:00:00";
    tasksSheet.rows[4].cells[9].value = 80;
    tasksSheet.rows[4].cells[10].value = 90;
    tasksSheet.rows[4].cells[16].value = "Updated Notes";

    const importedModel = projectXlsx.importProjectWorkbook(workbook, model);
    const designTask = importedModel.tasks.find((task) => task.uid === "2");
    const implementationTask = importedModel.tasks.find((task) => task.uid === "3");

    expect(designTask.name).toBe("Design Updated");
    expect(designTask.start).toBe("2026-03-17T09:00:00");
    expect(designTask.finish).toBe("2026-03-18T18:00:00");
    expect(designTask.percentComplete).toBe(80);
    expect(designTask.percentWorkComplete).toBe(90);
    expect(designTask.notes).toBe("Updated Notes");
    expect(implementationTask.name).toBe("Implementation");
  });

  it("imports limited resource and assignment fields from workbook rows back into ProjectModel", () => {
    const { xml, projectXlsx } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);
    const workbook = projectXlsx.exportProjectWorkbook(model);
    const resourcesSheet = workbook.sheets.find((sheet) => sheet.name === "Resources");
    const assignmentsSheet = workbook.sheets.find((sheet) => sheet.name === "Assignments");

    resourcesSheet.rows[3].cells[2].value = "Miku Updated";
    resourcesSheet.rows[3].cells[5].value = "Platform";
    resourcesSheet.rows[3].cells[6].value = 0.8;

    assignmentsSheet.rows[3].cells[7].value = 0.75;
    assignmentsSheet.rows[3].cells[8].value = "PT12H0M0S";
    assignmentsSheet.rows[3].cells[11].value = 70;

    const importedModel = projectXlsx.importProjectWorkbook(workbook, model);
    const resource = importedModel.resources.find((item) => item.uid === "1");
    const assignment = importedModel.assignments.find((item) => item.uid === "1");

    expect(resource.name).toBe("Miku Updated");
    expect(resource.group).toBe("Platform");
    expect(resource.maxUnits).toBe(0.8);
    expect(assignment.units).toBe(0.75);
    expect(assignment.work).toBe("PT12H0M0S");
    expect(assignment.percentWorkComplete).toBe(70);
  });

  it("imports limited project fields from workbook rows back into ProjectModel", () => {
    const { xml, projectXlsx } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);
    const workbook = projectXlsx.exportProjectWorkbook(model);
    const projectSheet = workbook.sheets.find((sheet) => sheet.name === "Project");

    projectSheet.rows[3].cells[1].value = "Renamed Project";
    projectSheet.rows[7].cells[1].value = "2026-03-15 09:00:00";
    projectSheet.rows[12].cells[1].value = "2";
    projectSheet.rows[13].cells[1].value = 420;
    projectSheet.rows[16].cells[1].value = false;

    const importedModel = projectXlsx.importProjectWorkbook(workbook, model);

    expect(importedModel.project.name).toBe("Renamed Project");
    expect(importedModel.project.startDate).toBe("2026-03-15T09:00:00");
    expect(importedModel.project.calendarUID).toBe("2");
    expect(importedModel.project.minutesPerDay).toBe(420);
    expect(importedModel.project.scheduleFromStart).toBe(false);
  });

  it("imports project calendar and schedule mode fields from workbook rows back into ProjectModel", () => {
    const { xml, projectXlsx } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);
    const workbook = projectXlsx.exportProjectWorkbook(model);
    const projectSheet = workbook.sheets.find((sheet) => sheet.name === "Project");

    projectSheet.rows[12].cells[1].value = "2";
    projectSheet.rows[16].cells[1].value = false;

    const importedModel = projectXlsx.importProjectWorkbook(workbook, model);

    expect(importedModel.project.calendarUID).toBe("2");
    expect(importedModel.project.scheduleFromStart).toBe(false);
  });

  it("imports project metadata and date fields from workbook rows back into ProjectModel", () => {
    const { xml, projectXlsx } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);
    const workbook = projectXlsx.exportProjectWorkbook(model);
    const projectSheet = workbook.sheets.find((sheet) => sheet.name === "Project");

    projectSheet.rows[4].cells[1].value = "Updated Title";
    projectSheet.rows[6].cells[1].value = "Updated Company";
    projectSheet.rows[7].cells[1].value = "2026-03-15T09:00:00";
    projectSheet.rows[8].cells[1].value = "2026-03-28T18:00:00";

    const importedModel = projectXlsx.importProjectWorkbook(workbook, model);

    expect(importedModel.project.title).toBe("Updated Title");
    expect(importedModel.project.company).toBe("Updated Company");
    expect(importedModel.project.startDate).toBe("2026-03-15T09:00:00");
    expect(importedModel.project.finishDate).toBe("2026-03-28T18:00:00");
  });

  it("imports project author field from workbook rows back into ProjectModel", () => {
    const { xml, projectXlsx } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);
    const workbook = projectXlsx.exportProjectWorkbook(model);
    const projectSheet = workbook.sheets.find((sheet) => sheet.name === "Project");

    projectSheet.rows[5].cells[1].value = "Author From XLSX";

    const importedModel = projectXlsx.importProjectWorkbook(workbook, model);

    expect(importedModel.project.author).toBe("Author From XLSX");
  });

  it("imports project current and status dates plus weekly settings from workbook rows back into ProjectModel", () => {
    const { xml, projectXlsx } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);
    const workbook = projectXlsx.exportProjectWorkbook(model);
    const projectSheet = workbook.sheets.find((sheet) => sheet.name === "Project");

    projectSheet.rows[9].cells[1].value = "2026-03-18T09:00:00";
    projectSheet.rows[10].cells[1].value = "2026-03-22T09:00:00";
    projectSheet.rows[14].cells[1].value = 2100;
    projectSheet.rows[15].cells[1].value = 18;

    const importedModel = projectXlsx.importProjectWorkbook(workbook, model);

    expect(importedModel.project.currentDate).toBe("2026-03-18T09:00:00");
    expect(importedModel.project.statusDate).toBe("2026-03-22T09:00:00");
    expect(importedModel.project.minutesPerWeek).toBe(2100);
    expect(importedModel.project.daysPerMonth).toBe(18);
  });

  it("reports field-level import changes", () => {
    const { xml, projectXlsx } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);
    const workbook = projectXlsx.exportProjectWorkbook(model);
    const tasksSheet = workbook.sheets.find((sheet) => sheet.name === "Tasks");

    tasksSheet.rows[4].cells[2].value = "Design Updated";
    tasksSheet.rows[4].cells[9].value = 80;
    tasksSheet.rows[4].cells[16].value = "Updated Notes";

    const result = projectXlsx.importProjectWorkbookDetailed(workbook, model);

    expect(result.changes).toEqual([
      { scope: "tasks", uid: "2", label: "Design", field: "Name", before: "Design", after: "Design Updated" },
      { scope: "tasks", uid: "2", label: "Design", field: "PercentComplete", before: 100, after: 80 },
      { scope: "tasks", uid: "2", label: "Design", field: "Notes", before: "Design completed", after: "Updated Notes" }
    ]);
  });

  it("reports project-level import changes", () => {
    const { xml, projectXlsx } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);
    const workbook = projectXlsx.exportProjectWorkbook(model);
    const projectSheet = workbook.sheets.find((sheet) => sheet.name === "Project");

    projectSheet.rows[3].cells[1].value = "Renamed Project";
    projectSheet.rows[13].cells[1].value = 420;

    const result = projectXlsx.importProjectWorkbookDetailed(workbook, model);

    expect(result.changes).toEqual([
      { scope: "project", uid: "project", label: "Sample Project", field: "Name", before: "Sample Project", after: "Renamed Project" },
      { scope: "project", uid: "project", label: "Sample Project", field: "MinutesPerDay", before: 480, after: 420 }
    ]);
  });

  it("imports limited calendar fields from workbook rows back into ProjectModel", () => {
    const { xml, projectXlsx } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);
    const workbook = projectXlsx.exportProjectWorkbook(model);
    const calendarsSheet = workbook.sheets.find((sheet) => sheet.name === "Calendars");

    calendarsSheet.rows[3].cells[1].value = "Standard Updated";
    calendarsSheet.rows[3].cells[3].value = "2";

    const importedModel = projectXlsx.importProjectWorkbook(workbook, model);

    expect(importedModel.calendars.find((calendar) => calendar.uid === "1").name).toBe("Standard Updated");
    expect(importedModel.calendars.find((calendar) => calendar.uid === "1").baseCalendarUID).toBe("2");
    expect(importedModel.calendars.find((calendar) => calendar.uid === "2").name).toBe("Development");
  });

  it("reports calendar-level import changes", () => {
    const { xml, projectXlsx } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);
    const workbook = projectXlsx.exportProjectWorkbook(model);
    const calendarsSheet = workbook.sheets.find((sheet) => sheet.name === "Calendars");

    calendarsSheet.rows[3].cells[1].value = "Standard Updated";
    calendarsSheet.rows[3].cells[3].value = "2";

    const result = projectXlsx.importProjectWorkbookDetailed(workbook, model);

    expect(result.changes).toEqual([
      { scope: "calendars", uid: "1", label: "Standard", field: "Name", before: "Standard", after: "Standard Updated" },
      { scope: "calendars", uid: "1", label: "Standard", field: "BaseCalendarUID", before: undefined, after: "2" }
    ]);
  });

  it("imports calendar isBaseCalendar field from workbook rows back into ProjectModel", () => {
    const { xml, projectXlsx } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);
    const workbook = projectXlsx.exportProjectWorkbook(model);
    const calendarsSheet = workbook.sheets.find((sheet) => sheet.name === "Calendars");

    calendarsSheet.rows[4].cells[2].value = true;

    const importedModel = projectXlsx.importProjectWorkbook(workbook, model);

    expect(importedModel.calendars.find((calendar) => calendar.uid === "2").isBaseCalendar).toBe(true);
  });

  it("ignores calendar WeekDays, Exceptions, and WorkWeeks workbook edits", () => {
    const { xml, projectXlsx } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);
    const workbook = projectXlsx.exportProjectWorkbook(model);
    const calendarsSheet = workbook.sheets.find((sheet) => sheet.name === "Calendars");

    calendarsSheet.rows[3].cells[4].value = 77;
    calendarsSheet.rows[3].cells[5].value = 88;
    calendarsSheet.rows[3].cells[6].value = 99;

    const result = projectXlsx.importProjectWorkbookDetailed(workbook, model);

    expect(result.model.calendars.find((calendar) => calendar.uid === "1").weekDays).toHaveLength(1);
    expect(result.model.calendars.find((calendar) => calendar.uid === "1").exceptions).toHaveLength(SAMPLE_HOLIDAY_COUNT);
    expect(result.model.calendars.find((calendar) => calendar.uid === "1").workWeeks).toHaveLength(0);
    expect(result.changes).toEqual([]);
  });

  it("imports non-working day sheet fields from workbook rows back into ProjectModel", () => {
    const { xml, projectXlsx } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);
    const workbook = projectXlsx.exportProjectWorkbook(model);
    const nonWorkingDaysSheet = workbook.sheets.find((sheet) => sheet.name === "NonWorkingDays");

    nonWorkingDaysSheet.rows[3].cells[3].value = "Spring Holiday";
    nonWorkingDaysSheet.rows[3].cells[4].value = "2026-03-21";
    nonWorkingDaysSheet.rows[3].cells[7].value = false;

    const importedModel = projectXlsx.importProjectWorkbook(workbook, model);
    const exception = importedModel.calendars.find((calendar) => calendar.uid === "1").exceptions[0];

    expect(exception.name).toBe("Spring Holiday");
    expect(exception.fromDate).toBe("2026-03-21T00:00:00");
    expect(exception.toDate).toBe("2026-03-21T23:59:59");
    expect(exception.dayWorking).toBe(false);
  });

  it("reports non-working day import changes", () => {
    const { xml, projectXlsx } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);
    const workbook = projectXlsx.exportProjectWorkbook(model);
    const nonWorkingDaysSheet = workbook.sheets.find((sheet) => sheet.name === "NonWorkingDays");

    nonWorkingDaysSheet.rows[3].cells[3].value = "Spring Holiday";
    nonWorkingDaysSheet.rows[3].cells[4].value = "2026-03-21";

    const result = projectXlsx.importProjectWorkbookDetailed(workbook, model);

    expect(result.changes).toEqual([
      { scope: "calendars", uid: "1", label: "Standard", field: "Exception1.Name", before: SAMPLE_FIRST_HOLIDAY_NAME, after: "Spring Holiday" },
      { scope: "calendars", uid: "1", label: "Standard", field: "Exception1.FromDate", before: `${SAMPLE_FIRST_HOLIDAY_DATE}T00:00:00`, after: "2026-03-21T00:00:00" },
      { scope: "calendars", uid: "1", label: "Standard", field: "Exception1.ToDate", before: `${SAMPLE_FIRST_HOLIDAY_DATE}T23:59:59`, after: "2026-03-21T23:59:59" }
    ]);
  });

  it("reports assignment percent work complete import changes", () => {
    const { projectXlsx, model, workbook } = buildSampleWorkbook((nextWorkbook) => {
      const assignmentsSheet = nextWorkbook.sheets.find((sheet) => sheet.name === "Assignments");
      assignmentsSheet.rows[3].cells[11].value = 85;
    });

    const result = projectXlsx.importProjectWorkbookDetailed(workbook, model);

    expect(result.changes).toEqual([
      { scope: "assignments", uid: "1", label: "TaskUID=2", field: "PercentWorkComplete", before: 50, after: 85 }
    ]);
  });

  it("can validate imported task percent complete out of range without UI boot", () => {
    const { xml, projectXlsx, model, workbook } = buildSampleWorkbook((nextWorkbook) => {
      const tasksSheet = nextWorkbook.sheets.find((sheet) => sheet.name === "Tasks");
      tasksSheet.rows[4].cells[9].value = 120;
    });

    const importedModel = projectXlsx.importProjectWorkbook(workbook, model);
    const issues = xml.validateProjectModel(importedModel);

    expect(importedModel.tasks.find((task) => task.uid === "2").percentComplete).toBe(120);
    expect(issues).toHaveLength(1);
    expect(issues[0].level).toBe("warning");
    expect(issues[0].message).toContain("PercentComplete");
    expect(issues[0].message).toContain("0..100");
  });

  it("can validate imported task start later than finish without UI boot", () => {
    const { xml, projectXlsx, model, workbook } = buildSampleWorkbook((nextWorkbook) => {
      const tasksSheet = nextWorkbook.sheets.find((sheet) => sheet.name === "Tasks");
      tasksSheet.rows[4].cells[6].value = "2026-03-19T09:00:00";
      tasksSheet.rows[4].cells[7].value = "2026-03-18T18:00:00";
    });

    const importedModel = projectXlsx.importProjectWorkbook(workbook, model);
    const issues = xml.validateProjectModel(importedModel);

    expect(importedModel.tasks.find((task) => task.uid === "2").start).toBe("2026-03-19T09:00:00");
    expect(importedModel.tasks.find((task) => task.uid === "2").finish).toBe("2026-03-18T18:00:00");
    expect(issues).toHaveLength(1);
    expect(issues[0].level).toBe("warning");
    expect(issues[0].message).toContain("Start");
    expect(issues[0].message).toContain("Finish");
  });

  it("can validate imported missing calendar baseCalendarUID without UI boot", () => {
    const { xml, projectXlsx, model, workbook } = buildSampleWorkbook((nextWorkbook) => {
      const calendarsSheet = nextWorkbook.sheets.find((sheet) => sheet.name === "Calendars");
      calendarsSheet.rows[3].cells[3].value = "99";
    });

    const importedModel = projectXlsx.importProjectWorkbook(workbook, model);
    const issues = xml.validateProjectModel(importedModel);

    expect(importedModel.calendars.find((calendar) => calendar.uid === "1").baseCalendarUID).toBe("99");
    expect(issues).toHaveLength(1);
    expect(issues[0].level).toBe("warning");
    expect(issues[0].message).toContain("BaseCalendarUID");
    expect(issues[0].message).toContain("指していません");
  });

  it("can validate imported self-referencing calendar baseCalendarUID without UI boot", () => {
    const { xml, projectXlsx, model, workbook } = buildSampleWorkbook((nextWorkbook) => {
      const calendarsSheet = nextWorkbook.sheets.find((sheet) => sheet.name === "Calendars");
      calendarsSheet.rows[3].cells[3].value = "1";
    });

    const importedModel = projectXlsx.importProjectWorkbook(workbook, model);
    const issues = xml.validateProjectModel(importedModel);

    expect(importedModel.calendars.find((calendar) => calendar.uid === "1").baseCalendarUID).toBe("1");
    expect(issues).toHaveLength(1);
    expect(issues[0].level).toBe("warning");
    expect(issues[0].message).toContain("BaseCalendarUID");
    expect(issues[0].message).toContain("自身を指しています");
  });

  it("can export xml after imported xlsx edits without UI boot", () => {
    const { xml, projectXlsx, model, workbook } = buildSampleWorkbook((nextWorkbook) => {
      const tasksSheet = nextWorkbook.sheets.find((sheet) => sheet.name === "Tasks");
      tasksSheet.rows[4].cells[2].value = "Design Saved From XLSX";
    });

    const importedModel = projectXlsx.importProjectWorkbook(workbook, model);
    const exportedXml = xml.exportMsProjectXml(importedModel);

    expect(exportedXml).toContain("<Name>Design Saved From XLSX</Name>");
  });

  it("can export xml after imported invalid xlsx edits without UI boot", () => {
    const { xml, projectXlsx, model, workbook } = buildSampleWorkbook((nextWorkbook) => {
      const tasksSheet = nextWorkbook.sheets.find((sheet) => sheet.name === "Tasks");
      tasksSheet.rows[4].cells[6].value = "2026-03-19T09:00:00";
      tasksSheet.rows[4].cells[7].value = "2026-03-18T18:00:00";
    });

    const importedModel = projectXlsx.importProjectWorkbook(workbook, model);
    const issues = xml.validateProjectModel(importedModel);
    const exportedXml = xml.exportMsProjectXml(importedModel);

    expect(issues).toHaveLength(1);
    expect(issues[0].level).toBe("warning");
    expect(issues[0].message).toContain("Start");
    expect(issues[0].message).toContain("Finish");
    expect(exportedXml).toContain("<Start>2026-03-19T09:00:00</Start>");
    expect(exportedXml).toContain("<Finish>2026-03-18T18:00:00</Finish>");
  });

  it("round-trips editable fields through workbook and xlsx bytes", () => {
    const { excelIo, xml, projectXlsx } = bootModules();
    const codec = new excelIo.XlsxWorkbookCodec();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);
    const workbook = projectXlsx.exportProjectWorkbook(model);

    const tasksSheet = workbook.sheets.find((sheet) => sheet.name === "Tasks");
    const resourcesSheet = workbook.sheets.find((sheet) => sheet.name === "Resources");
    const assignmentsSheet = workbook.sheets.find((sheet) => sheet.name === "Assignments");
    const calendarsSheet = workbook.sheets.find((sheet) => sheet.name === "Calendars");
    const nonWorkingDaysSheet = workbook.sheets.find((sheet) => sheet.name === "NonWorkingDays");

    tasksSheet.rows[4].cells[2].value = "Design via XLSX";
    tasksSheet.rows[4].cells[9].value = 60;
    tasksSheet.rows[4].cells[16].value = "Design notes via XLSX";
    resourcesSheet.rows[3].cells[2].value = "Miku via XLSX";
    assignmentsSheet.rows[3].cells[11].value = 55;
    calendarsSheet.rows[4].cells[1].value = "Development via XLSX";
    calendarsSheet.rows[4].cells[2].value = true;
    calendarsSheet.rows[4].cells[3].value = "1";
    nonWorkingDaysSheet.rows[3].cells[3].value = "Holiday via XLSX";
    nonWorkingDaysSheet.rows[3].cells[4].value = "2026-03-22";

    const bytes = codec.exportWorkbook(workbook);
    const importedWorkbook = codec.importWorkbook(bytes);
    const importedModel = projectXlsx.importProjectWorkbook(importedWorkbook, model);

    expect(importedModel.tasks.find((task) => task.uid === "2").name).toBe("Design via XLSX");
    expect(importedModel.tasks.find((task) => task.uid === "2").percentComplete).toBe(60);
    expect(importedModel.tasks.find((task) => task.uid === "2").notes).toBe("Design notes via XLSX");
    expect(importedModel.resources.find((resource) => resource.uid === "1").name).toBe("Miku via XLSX");
    expect(importedModel.assignments.find((assignment) => assignment.uid === "1").percentWorkComplete).toBe(55);
    expect(importedModel.calendars.find((calendar) => calendar.uid === "2").name).toBe("Development via XLSX");
    expect(importedModel.calendars.find((calendar) => calendar.uid === "2").isBaseCalendar).toBe(true);
    expect(importedModel.calendars.find((calendar) => calendar.uid === "2").baseCalendarUID).toBe("1");
    expect(importedModel.calendars.find((calendar) => calendar.uid === "1").exceptions[0].name).toBe("Holiday via XLSX");
    expect(importedModel.calendars.find((calendar) => calendar.uid === "1").exceptions[0].fromDate).toBe("2026-03-22T00:00:00");
  });
});
