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

describe("mikuproject main xlsx import", () => {
  beforeEach(() => {
    setupMainXlsxImportDom();
  });

  it("imports xlsx edits back into the current model and xml", async () => {
    bootPage();
    parseXmlViaHook();
    document.getElementById("downloadXmlBtn").click();
    expect(document.getElementById("xmlSaveState").textContent).toContain("XML 保存状態: 保存済み (2026-03-16 23:12)");

    const codec = new globalThis.__mikuprojectExcelIo.XlsxWorkbookCodec();
    const workbook = globalThis.__mikuprojectProjectXlsx.exportProjectWorkbook(
      globalThis.__mikuprojectXml.importMsProjectXml(document.getElementById("xmlInput").value)
    );
    const tasksSheet = workbook.sheets.find((sheet) => sheet.name === "Tasks");
    tasksSheet.rows[5].cells[2].value = "初期実装 Imported From XLSX";
    tasksSheet.rows[5].cells[8].value = "PT24H0M0S";
    tasksSheet.rows[5].cells[9].value = 77;
    tasksSheet.rows[5].cells[14].value = "2";
    tasksSheet.rows[5].cells[15].value = "2";
    const bytes = codec.exportWorkbook(workbook);

    const file = new File([bytes], "sample.xlsx", {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    Object.defineProperty(file, "arrayBuffer", {
      configurable: true,
      value: () => Promise.resolve(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength))
    });

    setImportFiles(file);
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("modelOutput").value).toContain("\"name\": \"初期実装 Imported From XLSX\"");
    expect(document.getElementById("modelOutput").value).toContain("\"duration\": \"PT24H0M0S\"");
    expect(document.getElementById("modelOutput").value).toContain("\"percentComplete\": 77");
    expect(document.getElementById("modelOutput").value).toContain("\"calendarUID\": \"2\"");
    expect(document.getElementById("modelOutput").value).toContain("\"predecessorUid\": \"2\"");
    expect(document.getElementById("xmlInput").value).toContain("<Name>初期実装 Imported From XLSX</Name>");
    expect(document.getElementById("xmlInput").value).toContain("<PredecessorUID>2</PredecessorUID>");
    expect(document.getElementById("statusMessage").textContent).toContain("XLSX を読み込んで project 全体を置き換えました");
    expect(document.getElementById("statusMessage").textContent).toContain("XML Export で保存できます");
    expect(document.getElementById("xmlSaveState").textContent).toContain("XML 保存状態: 未保存");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Tasks 1");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Duration");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("PT0H0M0S");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("PT24H0M0S");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("PercentComplete");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("77");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("CalendarUID");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("(empty)");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("2");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Predecessors");
  });

  it("imports resource sheet edits back into the current model and xml", async () => {
    bootPage();
    document.getElementById("xmlInput").value = dependencyXml;
    parseXmlViaHook();

    const codec = new globalThis.__mikuprojectExcelIo.XlsxWorkbookCodec();
    const workbook = globalThis.__mikuprojectProjectXlsx.exportProjectWorkbook(
      globalThis.__mikuprojectXml.importMsProjectXml(document.getElementById("xmlInput").value)
    );
    const resourcesSheet = workbook.sheets.find((sheet) => sheet.name === "Resources");
    resourcesSheet.rows[3].cells[2].value = "Miku Updated";
    resourcesSheet.rows[3].cells[5].value = "Dev";
    resourcesSheet.rows[3].cells[6].value = 1;
    resourcesSheet.rows[3].cells[7].value = "1";
    const bytes = codec.exportWorkbook(workbook);

    const file = new File([bytes], "resource-sheet.xlsx", {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    Object.defineProperty(file, "arrayBuffer", {
      configurable: true,
      value: () => Promise.resolve(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength))
    });

    setImportFiles(file);
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("modelOutput").value).toContain("\"name\": \"Miku Updated\"");
    expect(document.getElementById("modelOutput").value).toContain("\"group\": \"Dev\"");
    expect(document.getElementById("modelOutput").value).toContain("\"maxUnits\": 1");
    expect(document.getElementById("modelOutput").value).toContain("\"calendarUID\": \"1\"");
    expect(document.getElementById("statusMessage").textContent).toContain("XLSX を読み込んで project 全体を置き換えました");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Resources 1");
  });

  it("reports when xlsx import has no applicable changes", async () => {
    bootPage();
    parseXmlViaHook();

    const codec = new globalThis.__mikuprojectExcelIo.XlsxWorkbookCodec();
    const originalXml = document.getElementById("xmlInput").value;
    const workbook = globalThis.__mikuprojectProjectXlsx.exportProjectWorkbook(
      globalThis.__mikuprojectXml.importMsProjectXml(originalXml)
    );
    const bytes = codec.exportWorkbook(workbook);

    const file = new File([bytes], "no-change.xlsx", {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    Object.defineProperty(file, "arrayBuffer", {
      configurable: true,
      value: () => Promise.resolve(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength))
    });

    setImportFiles(file);
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("XLSX を読み込んで project 全体を置き換えました");
    expect(document.getElementById("xlsxImportSummary").classList.contains("md-hidden")).toBe(true);
  });

  it("imports duration edits and ignores unsupported xlsx columns and sheets", async () => {
    bootPage();
    parseXmlViaHook();

    const codec = new globalThis.__mikuprojectExcelIo.XlsxWorkbookCodec();
    const originalXml = document.getElementById("xmlInput").value;
    const workbook = globalThis.__mikuprojectProjectXlsx.exportProjectWorkbook(
      globalThis.__mikuprojectXml.importMsProjectXml(originalXml)
    );
    const tasksSheet = workbook.sheets.find((sheet) => sheet.name === "Tasks");
    const calendarsSheet = workbook.sheets.find((sheet) => sheet.name === "Calendars");

    tasksSheet.rows[4].cells[8].value = "PT99H0M0S";
    calendarsSheet.rows[3].cells[4].value = 99;

    const bytes = codec.exportWorkbook(workbook);
    const file = new File([bytes], "unsupported-columns.xlsx", {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    Object.defineProperty(file, "arrayBuffer", {
      configurable: true,
      value: () => Promise.resolve(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength))
    });

    setImportFiles(file);
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("XLSX を読み込んで project 全体を置き換えました");
    expect(document.getElementById("modelOutput").value).toContain("\"duration\": \"PT99H0M0S\"");
    expect(document.getElementById("modelOutput").value).not.toContain("\"weekDays\": 99");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Duration");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("PT0H0M0S");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("PT99H0M0S");
  });

  it("ignores calendar WeekDays, Exceptions, and WorkWeeks edits in xlsx import", async () => {
    bootPage();
    parseXmlViaHook();

    const codec = new globalThis.__mikuprojectExcelIo.XlsxWorkbookCodec();
    const originalXml = document.getElementById("xmlInput").value;
    const workbook = globalThis.__mikuprojectProjectXlsx.exportProjectWorkbook(
      globalThis.__mikuprojectXml.importMsProjectXml(originalXml)
    );
    const calendarsSheet = workbook.sheets.find((sheet) => sheet.name === "Calendars");

    calendarsSheet.rows[3].cells[4].value = 77;
    calendarsSheet.rows[3].cells[5].value = 88;
    calendarsSheet.rows[3].cells[6].value = 99;

    const bytes = codec.exportWorkbook(workbook);
    const file = new File([bytes], "ignored-calendar-structure.xlsx", {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    Object.defineProperty(file, "arrayBuffer", {
      configurable: true,
      value: () => Promise.resolve(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength))
    });

    setImportFiles(file);
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("XLSX を読み込んで project 全体を置き換えました");
    expect(document.getElementById("modelOutput").value).not.toContain("\"weekDays\": 77");
    expect(document.getElementById("modelOutput").value).not.toContain("\"exceptions\": 88");
    expect(document.getElementById("modelOutput").value).not.toContain("\"workWeeks\": 99");
    expect(document.getElementById("xlsxImportSummary").classList.contains("md-hidden")).toBe(true);
  });
});
