// @vitest-environment jsdom

import { beforeEach, describe, expect, it } from "vitest";

import {
  bootPage,
  dependencyXml,
  flushAsyncWork,
  getMainHooks,
  parseXmlViaHook,
  setImportFiles,
  setupMainXlsxImportDom,
  workbookImportSampleJson
} from "./helpers/main-xlsx-import-harness.js";

describe("mikuproject main xlsx import replace", () => {
  beforeEach(() => {
    setupMainXlsxImportDom();
  });

  it("imports workbook json edits back into the current model and xml", async () => {
    bootPage();
    parseXmlViaHook();

    const workbookJson = globalThis.__mikuprojectProjectWorkbookJson.exportProjectWorkbookJson(
      globalThis.__mikuprojectXml.importMsProjectXml(document.getElementById("xmlInput").value)
    );
    workbookJson.sheets.Tasks[2].Name = "初期実装 Imported From JSON";
    workbookJson.sheets.Tasks[2].PercentComplete = 66;

    const file = new File([JSON.stringify(workbookJson, null, 2)], "workbook-inline.json", {
      type: "application/json"
    });
    Object.defineProperty(file, "text", {
      configurable: true,
      value: () => Promise.resolve(JSON.stringify(workbookJson, null, 2))
    });

    setImportFiles(file);
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("modelOutput").value).toContain("\"name\": \"初期実装 Imported From JSON\"");
    expect(document.getElementById("statusMessage").textContent).toContain("JSON を読み込んで project 全体を置き換えました");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("JSON Replace 反映結果");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("workbook JSON による全置換結果です");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("PercentComplete");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("100");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("66");
  });

  it("imports workbook json from a file into the current model and xml", async () => {
    bootPage();
    parseXmlViaHook();

    const file = new File([workbookImportSampleJson], "workbook-import-sample.json", {
      type: "application/json"
    });
    Object.defineProperty(file, "text", {
      configurable: true,
      value: () => Promise.resolve(workbookImportSampleJson)
    });

    setImportFiles(file);
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("modelOutput").value).toContain("\"name\": \"初期実装 Imported From JSON File\"");
    expect(document.getElementById("statusMessage").textContent).toContain("JSON を読み込んで project 全体を置き換えました");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("PercentComplete");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("100");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("55");
  });

  it("reports ignored workbook json warnings in status message", async () => {
    bootPage();
    parseXmlViaHook();

    const workbookJson = globalThis.__mikuprojectProjectWorkbookJson.exportProjectWorkbookJson(
      globalThis.__mikuprojectXml.importMsProjectXml(document.getElementById("xmlInput").value)
    );
    workbookJson.sheets.Tasks[2].UnknownColumn = "ignored";
    workbookJson.sheets.UnknownSheet = [];

    const file = new File([JSON.stringify(workbookJson, null, 2)], "workbook-warning.json", {
      type: "application/json"
    });
    Object.defineProperty(file, "text", {
      configurable: true,
      value: () => Promise.resolve(JSON.stringify(workbookJson, null, 2))
    });

    setImportFiles(file);
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("JSON 取込で 2 件の warning を無視しました");
    expect(document.getElementById("importWarnings").textContent).toContain("未知の列は無視します: Tasks[2].UnknownColumn");
    expect(document.getElementById("importWarnings").textContent).toContain("未知の sheet は無視します: UnknownSheet");
  });

  it("imports project sheet edits back into the current model and xml", async () => {
    bootPage();
    parseXmlViaHook();

    const codec = new globalThis.__mikuprojectExcelIo.XlsxWorkbookCodec();
    const workbook = globalThis.__mikuprojectProjectXlsx.exportProjectWorkbook(
      globalThis.__mikuprojectXml.importMsProjectXml(document.getElementById("xmlInput").value)
    );
    const projectSheet = workbook.sheets.find((sheet) => sheet.name === "Project");
    projectSheet.rows[3].cells[1].value = "Project From XLSX";
    projectSheet.rows[13].cells[1].value = 420;
    const bytes = codec.exportWorkbook(workbook);

    const file = new File([bytes], "project-sheet.xlsx", {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    Object.defineProperty(file, "arrayBuffer", {
      configurable: true,
      value: () => Promise.resolve(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength))
    });

    setImportFiles(file);
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("modelOutput").value).toContain("\"name\": \"Project From XLSX\"");
    expect(document.getElementById("modelOutput").value).toContain("\"minutesPerDay\": 420");
    expect(document.getElementById("statusMessage").textContent).toContain("XLSX を読み込んで project 全体を置き換えました");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("XLSX Replace 反映結果");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("XLSX による全置換結果です");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Project 1");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("MinutesPerDay");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("480");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("420");
  });

  it("renders project import summary content without xlsx import wiring", () => {
    bootPage();

    getMainHooks().renderXlsxImportSummary([
      { scope: "project", uid: "project", label: "mikuproject開発", field: "CalendarUID", before: "1", after: "2" },
      { scope: "project", uid: "project", label: "mikuproject開発", field: "ScheduleFromStart", before: true, after: false },
      { scope: "project", uid: "project", label: "mikuproject開発", field: "Author", before: undefined, after: "Author From XLSX" }
    ]);

    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Project 1");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("3 fields");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Before");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("After");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("CalendarUID");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("1");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("2");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Author");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("(empty)");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Author From XLSX");
  });

  it("renders grouped xlsx import summary content without xlsx import wiring", () => {
    bootPage();

    getMainHooks().renderXlsxImportSummary([
      { scope: "calendars", uid: "1", label: "Standard", field: "Name", before: "Standard", after: "Standard Updated" },
      { scope: "tasks", uid: "3", label: "初期実装（MS Project XML 調査・基軸フォーマット選定・内部モデルの概要確定）", field: "Start", before: "2026-03-16", after: "2026-03-17" },
      { scope: "tasks", uid: "3", label: "初期実装（MS Project XML 調査・基軸フォーマット選定・内部モデルの概要確定）", field: "Finish", before: "2026-03-16", after: "2026-03-18" },
      { scope: "resources", uid: "1", label: "Miku", field: "Name", before: "Miku", after: "Miku Renamed" },
      { scope: "assignments", uid: "1", label: "TaskUID=2", field: "Work", before: "PT16H0M0S", after: "PT12H0M0S" }
    ]);

    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Tasks 1");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Resources 1");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Assignments 1");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Calendars 1");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("2 fields");
  });
});
