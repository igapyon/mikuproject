// @vitest-environment jsdom

import { beforeEach, describe, expect, it, vi } from "vitest";

import {
  SAMPLE_HOLIDAY_COUNT,
  bootPage,
  flushAsyncWork,
  getDefaultSampleHolidayDates,
  hierarchyXml,
  parseXmlViaHook,
  setupMainPreviewExportDom
} from "./helpers/main-preview-export-harness.js";

describe("mikuproject main preview export", () => {
  beforeEach(() => {
    setupMainPreviewExportDom();
  });

  it("exports xml from the current model", () => {
    bootPage();
    parseXmlViaHook();
    document.getElementById("downloadXmlBtn").click();
    const xmlText = document.getElementById("xmlInput").value;
    expect(xmlText).toContain("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
    expect(xmlText).toContain("<Name>mikuproject開発</Name>");
    expect(xmlText).toContain("<StartDate>2026-03-16</StartDate>");
    expect(xmlText).toContain("<FinishDate>2026-04-01</FinishDate>");
    expect(xmlText).toContain("<CalendarUID>1</CalendarUID>");
  });

  it("exports csv with parent id from the current model", async () => {
    bootPage();
    parseXmlViaHook();
    document.getElementById("downloadXmlBtn").click();
    const xmlInput = document.getElementById("xmlInput");
    xmlInput.value = `${xmlInput.value}\n<!-- edited -->`;
    xmlInput.dispatchEvent(new Event("input"));
    document.getElementById("exportCsvBtn").click();
    const csvBlob = URL.createObjectURL.mock.calls.at(-1)?.[0];
    expect(csvBlob).toBeTruthy();
    expect(csvBlob.type).toBe("text/csv;charset=utf-8");
    expect(document.getElementById("xmlInput").value).not.toContain("<!-- edited -->");
    expect(document.getElementById("statusMessage").textContent).toContain("CSV + ParentID を生成して保存しました");
  });

  it("smoke-tests lightweight download/export actions", async () => {
    bootPage();
    parseXmlViaHook();
    document.getElementById("downloadXmlBtn").click();
    expect(HTMLAnchorElement.prototype.click.mock.instances.at(-1).download).toBe("mikuproject-export-202603162312.xml");
    const xmlInput = document.getElementById("xmlInput");
    xmlInput.value = `${xmlInput.value}\n<!-- edited -->`;
    xmlInput.dispatchEvent(new Event("input"));
    document.getElementById("exportXlsxBtn").click();
    expect(HTMLAnchorElement.prototype.click.mock.instances.at(-1).download).toBe("mikuproject-export-202603162312.xlsx");
    document.getElementById("exportWorkbookJsonBtn").click();
    expect(HTMLAnchorElement.prototype.click.mock.instances.at(-1).download).toBe("mikuproject-workbook-202603162312.json");
    expect(JSON.parse(document.getElementById("workbookJsonOutput").value).format).toBe("mikuproject_workbook_json");
    document.getElementById("downloadWeeklySvgBtn").click();
    expect(HTMLAnchorElement.prototype.click.mock.instances.at(-1).download).toBe("mikuproject-wbs-weekly-202603162312.svg");
    document.getElementById("exportMermaidMdBtn").click();
    expect(HTMLAnchorElement.prototype.click.mock.instances.at(-1).download).toBe("mikuproject-wbs-mermaid-202603162312.md");
    document.getElementById("exportWbsMdBtn").click();
    expect(HTMLAnchorElement.prototype.click.mock.instances.at(-1).download).toBe("mikuproject-wbs-20260316.md");
    const OriginalBlob = Blob;
    class InspectableBlob extends OriginalBlob {
      constructor(parts = [], options = {}) {
        super(parts, options);
        this._parts = parts;
      }
    }
    globalThis.Blob = InspectableBlob;
    try {
      document.getElementById("downloadSvgBtn").click();
      await flushAsyncWork();
      await flushAsyncWork();
      expect(HTMLAnchorElement.prototype.click.mock.instances.at(-1).download).toBe("mikuproject-wbs-daily-202603162312.svg");
      const svgBlob = URL.createObjectURL.mock.calls.at(-1)?.[0];
      const svgText = String(svgBlob._parts?.[0] || "");
      expect(svgText).not.toContain("data-chart-origin-x");
    } finally {
      globalThis.Blob = OriginalBlob;
    }
  });

  it("downloads current wbs xlsx", () => {
    bootPage();
    const exportSpy = vi.spyOn(globalThis.__mikuprojectWbsXlsx, "exportWbsWorkbook");
    parseXmlViaHook();
    const defaultHolidayDates = getDefaultSampleHolidayDates();
    document.getElementById("downloadXmlBtn").click();
    document.getElementById("exportWbsXlsxBtn").click();
    expect(HTMLAnchorElement.prototype.click.mock.instances.at(-1).download).toBe("mikuproject-wbs-202603162312.xlsx");
    expect(exportSpy.mock.calls.at(-1)?.[1]).toEqual({
      holidayDates: defaultHolidayDates,
      displayDaysBeforeBaseDate: undefined,
      displayDaysAfterBaseDate: undefined,
      useBusinessDaysForDisplayRange: true,
      useBusinessDaysForProgressBand: true
    });
    expect(document.getElementById("statusMessage").textContent).toContain(`祝日 ${SAMPLE_HOLIDAY_COUNT} 件`);
    document.getElementById("wbsDisplayDaysBeforeInput").value = "1";
    document.getElementById("wbsDisplayDaysAfterInput").value = "2";
    document.getElementById("exportWbsXlsxBtn").click();
    expect(exportSpy.mock.calls.at(-1)?.[1]).toEqual({
      holidayDates: defaultHolidayDates,
      displayDaysBeforeBaseDate: 1,
      displayDaysAfterBaseDate: 2,
      useBusinessDaysForDisplayRange: true,
      useBusinessDaysForProgressBand: true
    });
    document.getElementById("exportWbsXlsxBtn").click();
    expect(exportSpy.mock.calls.at(-1)?.[1]).toEqual({
      holidayDates: defaultHolidayDates,
      displayDaysBeforeBaseDate: 1,
      displayDaysAfterBaseDate: 2,
      useBusinessDaysForDisplayRange: true,
      useBusinessDaysForProgressBand: true
    });
  });

  it("returns xml save state to unsaved after manual xml edit", async () => {
    bootPage();
    document.getElementById("downloadXmlBtn").click();
    const xmlInput = document.getElementById("xmlInput");
    xmlInput.value = `${xmlInput.value}\n<!-- edited -->`;
    xmlInput.dispatchEvent(new Event("input"));
    await flushAsyncWork();
    expect(document.getElementById("xmlSaveState").textContent).toContain("未保存");
  });

  it("exports regenerated xml instead of manual textarea edits when a model exists", async () => {
    bootPage();
    parseXmlViaHook();
    const xmlInput = document.getElementById("xmlInput");
    xmlInput.value = `${xmlInput.value}\n<!-- edited -->`;
    xmlInput.dispatchEvent(new Event("input"));
    document.getElementById("downloadXmlBtn").click();
    expect(document.getElementById("xmlInput").value).not.toContain("<!-- edited -->");
  });

  it("downloads monthly wbs calendar svg zip", async () => {
    bootPage();
    parseXmlViaHook();
    document.getElementById("downloadMonthlyCalendarSvgBtn").click();
    await flushAsyncWork();
    expect(HTMLAnchorElement.prototype.click.mock.instances.at(-1).download).toBe("mikuproject-monthly-wbs-calendar-202603162312.zip");
  });

  it("downloads all outputs as zip", () => {
    bootPage();
    parseXmlViaHook();
    document.getElementById("downloadAllOutputsBtn").click();
    expect(HTMLAnchorElement.prototype.click.mock.instances.at(-1).download).toBe("mikuproject-all-202603162312.zip");
  });

  it("imports xml from a file into the textarea", async () => {
    bootPage();
    const importInput = document.getElementById("importFileInput");
    const file = new File(["<Project><Name>Imported</Name></Project>"], "sample.xml", { type: "application/xml" });
    Object.defineProperty(file, "text", {
      configurable: true,
      value: () => Promise.resolve("<Project><Name>Imported</Name></Project>")
    });
    Object.defineProperty(importInput, "files", {
      configurable: true,
      value: [file]
    });
    importInput.dispatchEvent(new Event("change"));
    await Promise.resolve();
    await Promise.resolve();
    expect(document.getElementById("xmlInput").value).toContain("<Name>Imported</Name>");
    expect(document.getElementById("summaryProjectName").textContent).toBe("Imported");
    expect(document.getElementById("nativeSvgPreview").innerHTML).toContain("<svg");
  });

  it("parses csv with parent id into internal model summary", async () => {
    bootPage();
    const file = new File([[
      "ID,ParentID,WBS,Name,Start,Finish,PredecessorID,Resource,PercentComplete",
      "1,,1,Project Summary,2026-03-16T09:00:00,2026-03-20T18:00:00,,,50",
      "2,1,1.1,Design,2026-03-16T09:00:00,2026-03-17T18:00:00,,Miku,100",
      "3,1,1.2,Implementation,2026-03-18T09:00:00,2026-03-20T18:00:00,2,Miku,0"
    ].join("\n")], "sample.csv", { type: "text/csv" });
    Object.defineProperty(file, "text", {
      configurable: true,
      value: () => Promise.resolve([
        "ID,ParentID,WBS,Name,Start,Finish,PredecessorID,Resource,PercentComplete",
        "1,,1,Project Summary,2026-03-16T09:00:00,2026-03-20T18:00:00,,,50",
        "2,1,1.1,Design,2026-03-16T09:00:00,2026-03-17T18:00:00,,Miku,100",
        "3,1,1.2,Implementation,2026-03-18T09:00:00,2026-03-20T18:00:00,2,Miku,0"
      ].join("\n"))
    });
    const importInput = document.getElementById("importFileInput");
    Object.defineProperty(importInput, "files", {
      value: [file],
      configurable: true
    });
    importInput.dispatchEvent(new Event("change"));
    await flushAsyncWork();
    await flushAsyncWork();
    expect(document.getElementById("summaryProjectName").textContent).toBe("CSV Imported Project");
    expect(document.getElementById("nativeSvgPreview").innerHTML).toContain("<svg");
  });
});
