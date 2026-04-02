// @vitest-environment jsdom

import { beforeEach, describe, expect, it } from "vitest";

import {
  bootPage,
  dependencyXml,
  flushAsyncWork,
  parseXmlViaHook,
  setupMainAiJsonTestDom
} from "./helpers/main-ai-json-harness.js";

describe("mikuproject main ai json patch models", () => {
  beforeEach(() => {
    setupMainAiJsonTestDom();
  });

  it("imports patch json update_project into the current model and xml", async () => {
    bootPage();
    const xmlWithExtraCalendar = dependencyXml.replace(
      "</Calendars>",
      `  <Calendar>\n      <UID>2</UID>\n      <Name>Alt</Name>\n      <IsBaseCalendar>0</IsBaseCalendar>\n      <BaseCalendarUID>1</BaseCalendarUID>\n    </Calendar>\n  </Calendars>`
    );
    document.getElementById("xmlInput").value = xmlWithExtraCalendar;
    parseXmlViaHook();
    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [{
        op: "update_project",
        fields: {
          name: "Dependency Project Prime",
          title: "Prime Title",
          author: "Prime Author",
          company: "Prime Company",
          start_date: "2026-03-15",
          finish_date: "2026-03-20",
          current_date: "2026-03-17",
          status_date: "2026-03-18",
          calendar_uid: "2",
          minutes_per_day: 420,
          minutes_per_week: 2100,
          days_per_month: 18,
          schedule_from_start: false
        }
      }]
    }, null, 2);
    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();
    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON を読み込んで 13 件の変更を反映しました");
    expect(document.getElementById("xmlInput").value).toContain("<Name>Dependency Project Prime</Name>");
    expect(document.getElementById("xmlInput").value).toContain("<Title>Prime Title</Title>");
    expect(document.getElementById("xmlInput").value).toContain("<Author>Prime Author</Author>");
    expect(document.getElementById("xmlInput").value).toContain("<Company>Prime Company</Company>");
    expect(document.getElementById("xmlInput").value).toContain("<StartDate>2026-03-15T09:00:00</StartDate>");
    expect(document.getElementById("xmlInput").value).toContain("<FinishDate>2026-03-20T18:00:00</FinishDate>");
    expect(document.getElementById("xmlInput").value).toContain("<CurrentDate>2026-03-17T09:00:00</CurrentDate>");
    expect(document.getElementById("xmlInput").value).toContain("<StatusDate>2026-03-18T09:00:00</StatusDate>");
    expect(document.getElementById("xmlInput").value).toContain("<CalendarUID>2</CalendarUID>");
    expect(document.getElementById("xmlInput").value).toContain("<MinutesPerDay>420</MinutesPerDay>");
    expect(document.getElementById("xmlInput").value).toContain("<MinutesPerWeek>2100</MinutesPerWeek>");
    expect(document.getElementById("xmlInput").value).toContain("<DaysPerMonth>18</DaysPerMonth>");
    expect(document.getElementById("xmlInput").value).toContain("<ScheduleFromStart>0</ScheduleFromStart>");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Project");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("Title");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("StartDate");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("MinutesPerDay");
  });

  it("imports patch json update_resource into the current model and xml", async () => {
    bootPage();
    document.getElementById("xmlInput").value = dependencyXml;
    parseXmlViaHook();
    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [{
        op: "update_resource",
        uid: "1",
        fields: {
          name: "Miku Prime",
          initials: "MP",
          group: "Platform",
          calendar_uid: "1",
          max_units: 0.75,
          standard_rate: "1000",
          overtime_rate: "1500",
          cost_per_use: 250,
          percent_work_complete: 60
        }
      }]
    }, null, 2);
    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();
    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON を読み込んで 9 件の変更を反映しました");
    expect(document.getElementById("xmlInput").value).toContain("<Name>Miku Prime</Name>");
    expect(document.getElementById("xmlInput").value).toContain("<Initials>MP</Initials>");
    expect(document.getElementById("xmlInput").value).toContain("<Group>Platform</Group>");
    expect(document.getElementById("xmlInput").value).toContain("<CalendarUID>1</CalendarUID>");
    expect(document.getElementById("xmlInput").value).toContain("<MaxUnits>0.75</MaxUnits>");
    expect(document.getElementById("xmlInput").value).toContain("<StandardRate>1000</StandardRate>");
    expect(document.getElementById("xmlInput").value).toContain("<OvertimeRate>1500</OvertimeRate>");
    expect(document.getElementById("xmlInput").value).toContain("<CostPerUse>250</CostPerUse>");
    expect(document.getElementById("xmlInput").value).toContain("<PercentWorkComplete>60</PercentWorkComplete>");
    expect(document.getElementById("modelOutput").value).toContain("\"initials\": \"MP\"");
    expect(document.getElementById("modelOutput").value).toContain("\"maxUnits\": 0.75");
    expect(document.getElementById("modelOutput").value).toContain("\"costPerUse\": 250");
    expect(document.getElementById("modelOutput").value).toContain("\"percentWorkComplete\": 60");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Resources");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("Initials");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("MaxUnits");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("StandardRate");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("CostPerUse");
  });

  it("imports patch json update_calendar into the current model and xml", async () => {
    bootPage();
    const xmlWithExtraCalendar = dependencyXml.replace("</Calendars>", `  <Calendar>\n      <UID>2</UID>\n      <Name>Alt</Name>\n      <IsBaseCalendar>0</IsBaseCalendar>\n      <BaseCalendarUID>1</BaseCalendarUID>\n      <WeekDays />\n    </Calendar>\n  </Calendars>`);
    document.getElementById("xmlInput").value = xmlWithExtraCalendar;
    parseXmlViaHook();
    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [{ op: "update_calendar", uid: "2", fields: { name: "Alt Prime", is_base_calendar: true, base_calendar_uid: "" } }]
    }, null, 2);
    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();
    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON を読み込んで 3 件の変更を反映しました");
    expect(document.getElementById("xmlInput").value).toContain("<UID>2</UID>");
    expect(document.getElementById("xmlInput").value).toContain("<Name>Alt Prime</Name>");
    expect(document.getElementById("xmlInput").value).toContain("<IsBaseCalendar>1</IsBaseCalendar>");
    expect(document.getElementById("modelOutput").value).toContain("\"isBaseCalendar\": true");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Calendars");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("IsBaseCalendar");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("BaseCalendarUID");
  });

  it("imports patch json add_calendar into the current model and xml", async () => {
    bootPage();
    document.getElementById("xmlInput").value = dependencyXml;
    parseXmlViaHook();
    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [{ op: "add_calendar", uid: "2", name: "Night Shift", is_base_calendar: false, base_calendar_uid: "1" }]
    }, null, 2);
    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();
    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON を読み込んで 2 件の変更を反映しました");
    expect(document.getElementById("xmlInput").value).toContain("<UID>2</UID>");
    expect(document.getElementById("xmlInput").value).toContain("<Name>Night Shift</Name>");
    expect(document.getElementById("xmlInput").value).toContain("<BaseCalendarUID>1</BaseCalendarUID>");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Calendars");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("BaseCalendarUID");
  });

  it("imports patch json delete_calendar into the current model and xml", async () => {
    bootPage();
    const xmlWithExtraCalendar = dependencyXml.replace("</Calendars>", `  <Calendar>\n      <UID>2</UID>\n      <Name>Alt</Name>\n      <IsBaseCalendar>0</IsBaseCalendar>\n      <WeekDays />\n    </Calendar>\n  </Calendars>`);
    document.getElementById("xmlInput").value = xmlWithExtraCalendar;
    parseXmlViaHook();
    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [{ op: "delete_calendar", uid: "2" }]
    }, null, 2);
    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();
    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON を読み込んで 1 件の変更を反映しました");
    expect(document.getElementById("xmlInput").value).not.toContain("<Name>Alt</Name>");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("(deleted)");
  });

  it("imports patch json add_resource into the current model and xml", async () => {
    bootPage();
    document.getElementById("xmlInput").value = dependencyXml;
    parseXmlViaHook();
    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [{
        op: "add_resource",
        uid: "2",
        name: "Mikuku",
        initials: "M",
        group: "PMO",
        calendar_uid: "1",
        max_units: 1,
        standard_rate: "1200",
        overtime_rate: "1800",
        cost_per_use: 300,
        percent_work_complete: 25
      }]
    }, null, 2);
    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();
    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON を読み込んで 9 件の変更を反映しました");
    expect(document.getElementById("xmlInput").value).toContain("<UID>2</UID>");
    expect(document.getElementById("xmlInput").value).toContain("<Name>Mikuku</Name>");
    expect(document.getElementById("xmlInput").value).toContain("<Initials>M</Initials>");
    expect(document.getElementById("xmlInput").value).toContain("<Group>PMO</Group>");
    expect(document.getElementById("xmlInput").value).toContain("<CalendarUID>1</CalendarUID>");
    expect(document.getElementById("xmlInput").value).toContain("<MaxUnits>1</MaxUnits>");
    expect(document.getElementById("xmlInput").value).toContain("<StandardRate>1200</StandardRate>");
    expect(document.getElementById("xmlInput").value).toContain("<OvertimeRate>1800</OvertimeRate>");
    expect(document.getElementById("xmlInput").value).toContain("<CostPerUse>300</CostPerUse>");
    expect(document.getElementById("xmlInput").value).toContain("<PercentWorkComplete>25</PercentWorkComplete>");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Resources");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("Initials");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("StandardRate");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("CostPerUse");
  });

  it("imports patch json delete_resource into the current model and xml", async () => {
    bootPage();
    const xmlWithoutAssignments = dependencyXml.replace(/<Assignments>[\s\S]*<\/Assignments>/, "<Assignments />");
    document.getElementById("xmlInput").value = xmlWithoutAssignments;
    parseXmlViaHook();
    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [{ op: "delete_resource", uid: "1" }]
    }, null, 2);
    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();
    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON を読み込んで 1 件の変更を反映しました");
    expect(document.getElementById("xmlInput").value).not.toContain("<Name>Miku</Name>");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("(deleted)");
  });

});
