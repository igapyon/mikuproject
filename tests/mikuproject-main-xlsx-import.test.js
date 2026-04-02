// @vitest-environment jsdom

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { beforeEach, describe, expect, it, vi } from "vitest";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const typesCode = readFileSync(
  path.resolve(__dirname, "../src/js/types.js"),
  "utf8"
);
const markdownEscapeCode = readFileSync(
  path.resolve(__dirname, "../src/js/markdown-escape.js"),
  "utf8"
);
const aiJsonUtilCode = readFileSync(
  path.resolve(__dirname, "../src/js/ai-json-util.js"),
  "utf8"
);
const mainUtilCode = readFileSync(
  path.resolve(__dirname, "../src/js/main-util.js"),
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
const projectWorkbookSchemaCode = readFileSync(
  path.resolve(__dirname, "../src/js/project-workbook-schema.js"),
  "utf8"
);
const projectXlsxCode = readFileSync(
  path.resolve(__dirname, "../src/js/project-xlsx.js"),
  "utf8"
);
const projectWorkbookJsonCode = readFileSync(
  path.resolve(__dirname, "../src/js/project-workbook-json.js"),
  "utf8"
);
const mainRenderCode = readFileSync(
  path.resolve(__dirname, "../src/js/main-render.js"),
  "utf8"
);
const mainCode = readFileSync(
  path.resolve(__dirname, "../src/js/main.js"),
  "utf8"
);
const dependencyXml = readFileSync(
  path.resolve(__dirname, "../testdata/dependency.xml"),
  "utf8"
);
const workbookImportSampleJson = readFileSync(
  path.resolve(__dirname, "../testdata/workbook-import-sample.json"),
  "utf8"
);

const projectPatchJsonStubCode = `
globalThis.__mikuprojectProjectPatchJson = {
  importProjectPatchJson: () => ({ model: null, changes: [], warnings: [] })
};
`;

const wbsXlsxStubCode = `
globalThis.__mikuprojectWbsXlsx = {
  collectWbsHolidayDates: () => [],
  exportWbsWorkbook: () => ({ sheets: [] })
};
`;

const wbsMarkdownStubCode = `
globalThis.__mikuprojectWbsMarkdown = {
  exportWbsMarkdown: () => "# stub"
};
`;

const nativeSvgStubCode = `
globalThis.__mikuprojectNativeSvg = {
  exportNativeSvg: () => "<svg data-stub=\\"daily\\"></svg>",
  exportWeeklyNativeSvg: () => "<svg data-stub=\\"weekly\\"></svg>",
  exportMonthlyWbsCalendarSvgArchive: () => ({
    entries: [{ fileName: "2026-03.svg", svg: "<svg data-stub=\\"monthly\\"></svg>" }],
    zipBytes: new Uint8Array()
  })
};
`;

function mountDom() {
  document.body.innerHTML = `
    <button id="importFileBtn" type="button"></button>
    <button id="loadSampleBtn" type="button"></button>
    <button id="downloadAllOutputsBtn" type="button"></button>
    <button id="exportXlsxBtn" type="button"></button>
    <button id="exportWorkbookJsonBtn" type="button"></button>
    <button id="exportCsvBtn" type="button"></button>
    <button id="exportWbsXlsxBtn" type="button"></button>
    <button id="exportWbsMdBtn" type="button"></button>
    <button id="downloadWeeklySvgBtn" type="button"></button>
    <button id="downloadMonthlyCalendarSvgBtn" type="button"></button>
    <button id="exportMermaidMdBtn" type="button"></button>
    <button id="downloadSvgBtn" type="button"></button>
    <button id="previewDailySvgBtn" type="button"></button>
    <button id="previewWeeklySvgBtn" type="button"></button>
    <button id="previewMonthlySvgBtn" type="button"></button>
    <button id="exportAiBundleBtn" type="button"></button>
    <button id="exportProjectOverviewBtn" type="button"></button>
    <button id="exportTaskEditBtn" type="button"></button>
    <button id="exportPhaseDetailFullBtn" type="button"></button>
    <button id="exportPhaseDetailBtn" type="button"></button>
    <button id="loadProjectDraftSampleBtn" type="button"></button>
    <button id="importProjectDraftBtn" type="button"></button>
    <button id="downloadXmlBtn" type="button"></button>
    <button id="roundTripBtn" type="button"></button>
    <button id="copyAiPromptBtn" type="button"></button>
    <input id="importFileInput" type="file" />
    <input id="phaseDetailUidInput" type="text" />
    <input id="taskEditUidInput" type="text" />
    <input id="phaseDetailRootUidInput" type="text" />
    <input id="phaseDetailMaxDepthInput" type="text" />
    <input id="wbsDisplayDaysBeforeInput" />
    <input id="wbsDisplayDaysAfterInput" />
    <input id="wbsBusinessDayRangeInput" type="checkbox" />
    <input id="wbsBusinessDayProgressInput" type="checkbox" />
    <div id="statusMessage"></div>
    <div class="md-top-tabs">
      <button type="button" class="md-top-tab is-active" data-tab="input"></button>
      <button type="button" class="md-top-tab" data-tab="transform"></button>
      <button type="button" class="md-top-tab" data-tab="output"></button>
    </div>
    <section id="tabPanelInput" class="md-tab-panel" data-tab-panel="input">
      <textarea id="xmlInput"></textarea>
      <template id="aiPromptTemplate"># mikuproject AI JSON Spec</template>
      <textarea id="projectDraftImportInput"></textarea>
      <div id="xmlSaveState"></div>
    </section>
    <section id="tabPanelTransform" class="md-tab-panel" data-tab-panel="transform" hidden>
      <div id="summaryProjectName"></div>
      <div id="summaryTaskCount"></div>
      <div id="summaryResourceCount"></div>
      <div id="summaryAssignmentCount"></div>
      <div id="summaryCalendarCount"></div>
      <div class="md-feedback-stack">
        <div class="md-label md-hidden">検証結果</div>
        <div id="validationIssues" class="md-hidden"></div>
        <div class="md-label md-hidden">import warnings</div>
        <div id="importWarnings" class="md-hidden"></div>
        <div class="md-label md-hidden">import summary</div>
        <div id="xlsxImportSummary" class="md-hidden"></div>
      </div>
      <div id="projectPreview"></div>
      <div id="taskPreview"></div>
      <div id="resourcePreview"></div>
      <div id="assignmentPreview"></div>
      <div id="calendarPreview"></div>
      <textarea id="modelOutput"></textarea>
      <textarea id="mermaidOutput"></textarea>
      <div id="wbsPreviewTitle"></div>
      <div id="wbsPreviewRange"></div>
      <div id="nativeSvgPreview"></div>
    </section>
    <section id="tabPanelOutput" class="md-tab-panel" data-tab-panel="output" hidden>
      <textarea id="workbookJsonOutput"></textarea>
      <textarea id="aiBundleOutput"></textarea>
      <textarea id="projectOverviewOutput"></textarea>
      <textarea id="phaseDetailOutput"></textarea>
    </section>
    <div id="toast"></div>
  `;
  const toast = document.getElementById("toast");
  toast.show = vi.fn();
}

function bootPage() {
  mountDom();
  new Function(`${typesCode}\n${markdownEscapeCode}\n${aiJsonUtilCode}\n${mainUtilCode}\n${excelIoCode}\n${msProjectXmlCode}\n${projectWorkbookSchemaCode}\n${projectXlsxCode}\n${projectWorkbookJsonCode}\n${projectPatchJsonStubCode}\n${wbsXlsxStubCode}\n${wbsMarkdownStubCode}\n${nativeSvgStubCode}\n${mainRenderCode}\n${mainCode}`)();
  document.dispatchEvent(new Event("DOMContentLoaded"));
}

function getMainHooks() {
  return globalThis.__mikuprojectMainTestHooks;
}

function parseXmlViaHook() {
  getMainHooks().parseCurrentXml();
}

async function flushAsyncWork() {
  await Promise.resolve();
  await Promise.resolve();
}

function setImportFiles(file) {
  const importInput = document.getElementById("importFileInput");
  Object.defineProperty(importInput, "files", {
    configurable: true,
    value: [file]
  });
  importInput.dispatchEvent(new Event("change"));
}

describe("mikuproject main xlsx import", () => {
  beforeEach(() => {
    document.body.innerHTML = "";
    Object.defineProperty(URL, "createObjectURL", {
      value: vi.fn(() => "blob:mock"),
      configurable: true
    });
    Object.defineProperty(URL, "revokeObjectURL", {
      value: vi.fn(),
      configurable: true
    });
    HTMLAnchorElement.prototype.click = vi.fn();
    const clipboard = {
      writeText: vi.fn(async () => {})
    };
    Object.defineProperty(globalThis.navigator, "clipboard", {
      value: clipboard,
      configurable: true
    });
    Object.defineProperty(window.navigator, "clipboard", {
      value: clipboard,
      configurable: true
    });
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2026-03-16T23:12:00+09:00"));
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
