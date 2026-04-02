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
const projectPatchJsonCode = readFileSync(
  path.resolve(__dirname, "../src/js/project-patch-json.js"),
  "utf8"
);
const wbsXlsxCode = readFileSync(
  path.resolve(__dirname, "../src/js/wbs-xlsx.js"),
  "utf8"
);
const wbsMarkdownCode = readFileSync(
  path.resolve(__dirname, "../src/js/wbs-markdown.js"),
  "utf8"
);
const nativeSvgCode = readFileSync(
  path.resolve(__dirname, "../src/js/wbs-svg.js"),
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

const hierarchyXml = readFileSync(
  path.resolve(__dirname, "../testdata/hierarchy.xml"),
  "utf8"
);

function mountDom() {
  document.body.innerHTML = `
    <button id="importFileBtn" type="button">Load from file</button>
    <button id="loadSampleBtn" type="button">サンプル</button>
    <button id="downloadAllOutputsBtn" type="button">All</button>
    <button id="exportXlsxBtn" type="button">XLSX</button>
    <button id="exportWorkbookJsonBtn" type="button">JSON</button>
    <button id="exportCsvBtn" type="button">CSV</button>
    <button id="exportWbsXlsxBtn" type="button">WBS XLSX</button>
    <button id="exportWbsMdBtn" type="button">WBS Markdown</button>
    <button id="downloadWeeklySvgBtn" type="button" disabled>Weekly SVG</button>
    <button id="downloadMonthlyCalendarSvgBtn" type="button" disabled>Monthly Calendar SVG</button>
    <button id="exportMermaidMdBtn" type="button">Mermaid</button>
    <button id="downloadSvgBtn" type="button" disabled>Daily SVG</button>
    <button id="previewDailySvgBtn" type="button">Daily SVG</button>
    <button id="previewWeeklySvgBtn" type="button">Weekly SVG</button>
    <button id="previewMonthlySvgBtn" type="button">Monthly Calendar SVG</button>
    <button id="exportAiBundleBtn" type="button">project_overview + all phase_detail full</button>
    <button id="exportProjectOverviewBtn" type="button">project_overview_view</button>
    <button id="exportTaskEditBtn" type="button">task_edit_view</button>
    <button id="exportPhaseDetailFullBtn" type="button">phase_detail_view full</button>
    <button id="exportPhaseDetailBtn" type="button">phase_detail_view</button>
    <button id="loadProjectDraftSampleBtn" type="button">サンプル draft</button>
    <button id="importProjectDraftBtn" type="button">project_draft_view を取り込む</button>
    <button id="downloadXmlBtn" type="button">MS Project XML</button>
    <button id="roundTripBtn" type="button">Round Trip</button>
    <button id="copyAiPromptBtn" type="button">Copy AI Prompt</button>
    <input id="importFileInput" type="file" />
    <input id="phaseDetailUidInput" type="text" />
    <input id="taskEditUidInput" type="text" />
    <input id="phaseDetailRootUidInput" type="text" />
    <input id="phaseDetailMaxDepthInput" type="text" />
    <div id="statusMessage"></div>
    <div class="md-top-tabs" role="tablist" aria-label="mikuproject sections">
      <button type="button" class="md-top-tab is-active" data-tab="input" role="tab" aria-selected="true" aria-controls="tabPanelInput">
        <span class="md-top-tab-no">1</span>
        <span class="md-top-tab-label">Input</span>
      </button>
      <button type="button" class="md-top-tab" data-tab="transform" role="tab" aria-selected="false" aria-controls="tabPanelTransform">
        <span class="md-top-tab-no">2</span>
        <span class="md-top-tab-label">Transform</span>
      </button>
      <button type="button" class="md-top-tab" data-tab="output" role="tab" aria-selected="false" aria-controls="tabPanelOutput">
        <span class="md-top-tab-no">3</span>
        <span class="md-top-tab-label">Output</span>
      </button>
    </div>
    <section id="tabPanelInput" class="md-flow-section md-tab-panel" data-tab-panel="input">
      <textarea id="xmlInput"></textarea>
      <template id="aiPromptTemplate"># mikuproject AI JSON Spec</template>
      <textarea id="projectDraftImportInput"></textarea>
      <div id="xmlSaveState"></div>
    </section>
    <section id="tabPanelTransform" class="md-flow-section md-tab-panel" data-tab-panel="transform" hidden>
      <div id="summaryProjectName"></div>
      <div id="summaryTaskCount"></div>
      <div id="summaryResourceCount"></div>
      <div id="summaryAssignmentCount"></div>
      <div id="summaryCalendarCount"></div>
      <div class="md-feedback-stack md-hidden">
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
      <input id="wbsDisplayDaysBeforeInput" />
      <input id="wbsDisplayDaysAfterInput" />
      <input id="wbsBusinessDayRangeInput" type="checkbox" />
      <input id="wbsBusinessDayProgressInput" type="checkbox" />
    </section>
    <section id="tabPanelOutput" class="md-flow-section md-tab-panel" data-tab-panel="output" hidden>
      <textarea id="workbookJsonOutput"></textarea>
      <textarea id="aiBundleOutput"></textarea>
      <textarea id="projectOverviewOutput"></textarea>
      <textarea id="taskEditOutput"></textarea>
      <textarea id="phaseDetailOutput"></textarea>
    </section>
    <div id="toast"></div>
  `;
  const toast = document.getElementById("toast");
  toast.show = vi.fn();
}

function bootPage() {
  mountDom();
  new Function(`${typesCode}\n${markdownEscapeCode}\n${aiJsonUtilCode}\n${mainUtilCode}\n${excelIoCode}\n${msProjectXmlCode}\n${projectWorkbookSchemaCode}\n${projectXlsxCode}\n${projectWorkbookJsonCode}\n${projectPatchJsonCode}\n${wbsXlsxCode}\n${wbsMarkdownCode}\n${nativeSvgCode}\n${mainRenderCode}\n${mainCode}`)();
  document.dispatchEvent(new Event("DOMContentLoaded"));
}

function bootXmlModule() {
  new Function(`${typesCode}\n${msProjectXmlCode}`)();
  return globalThis.__mikuprojectXml;
}

function getMainHooks() {
  return globalThis.__mikuprojectMainTestHooks;
}

function parseXmlViaHook() {
  getMainHooks().parseCurrentXml();
}

async function exportMermaidViaHook() {
  await getMainHooks().exportCurrentMermaid();
}

const SAMPLE_HOLIDAY_COUNT = 1;

function getDefaultSampleHolidayDates() {
  return globalThis.__mikuprojectWbsXlsx.collectWbsHolidayDates(
    globalThis.__mikuprojectXml.importMsProjectXml(globalThis.__mikuprojectXml.SAMPLE_XML)
  );
}

async function flushAsyncWork() {
  await Promise.resolve();
  await Promise.resolve();
}

describe("mikuproject main preview export", () => {
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

  it("exports mermaid gantt from the current model", async () => {
    bootPage();

    parseXmlViaHook();
    await exportMermaidViaHook();

    const mermaidText = document.getElementById("mermaidOutput").value;
    expect(mermaidText).toContain("gantt");
    expect(mermaidText).toContain("title mikuproject開発");
    expect(mermaidText).toContain("section 基盤整備");
    expect(document.getElementById("nativeSvgPreview").innerHTML).toContain("<svg");
    expect(document.getElementById("downloadSvgBtn").disabled).toBe(false);
    expect(document.getElementById("downloadWeeklySvgBtn").disabled).toBe(false);
  });

  it("switches overview svg preview between daily, weekly, and monthly", async () => {
    bootPage();

    parseXmlViaHook();
    await exportMermaidViaHook();

    document.getElementById("previewWeeklySvgBtn").click();
    expect(document.getElementById("nativeSvgPreview").innerHTML).toContain("weekly overview");

    document.getElementById("previewMonthlySvgBtn").click();
    expect(document.getElementById("nativeSvgPreview").innerHTML).toContain("2026-03");
    expect(document.getElementById("nativeSvgPreview").innerHTML).toContain("2026-04");

    document.getElementById("previewDailySvgBtn").click();
    expect(document.getElementById("nativeSvgPreview").innerHTML).toContain("<svg");
    expect(document.getElementById("nativeSvgPreview").innerHTML).not.toContain("weekly overview");
  });

  it("keeps long daily task labels outside narrow on-bar labels", async () => {
    bootPage();

    parseXmlViaHook();
    await exportMermaidViaHook();

    const dailySvg = document.getElementById("nativeSvgPreview").innerHTML;
    expect(dailySvg).not.toContain('text-anchor="middle">round-trip拡張（MS Project XML → 内部JSON形式 → MS Project XML の往復対応）</text>');
    expect(dailySvg).not.toContain('text-anchor="middle">MS Project XML と XLSX の相互変換・round-trip実装</text>');
  });

  it("places daily labels on the left after they pass the chart midpoint", async () => {
    bootPage();

    parseXmlViaHook();
    await exportMermaidViaHook();

    const dailySvg = document.getElementById("nativeSvgPreview").innerHTML;
    expect(dailySvg).toContain('text-anchor="start">架空検討フェーズ【架空】</text>');
    expect(dailySvg).toContain('text-anchor="end">MS Project XML と XLSX の相互変換・round-trip実装</text>');
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
    document.getElementById("downloadXmlBtn").click();
    const xmlInput = document.getElementById("xmlInput");
    xmlInput.value = `${xmlInput.value}\n<!-- edited -->`;
    xmlInput.dispatchEvent(new Event("input"));
    document.getElementById("exportWbsXlsxBtn").click();
    const defaultHolidayDates = getDefaultSampleHolidayDates();

    expect(HTMLAnchorElement.prototype.click.mock.instances.at(-1).download).toBe("mikuproject-wbs-202603162312.xlsx");
    expect(exportSpy.mock.calls.at(-1)?.[1]).toEqual({
      holidayDates: defaultHolidayDates,
      displayDaysBeforeBaseDate: undefined,
      displayDaysAfterBaseDate: undefined,
      useBusinessDaysForDisplayRange: true,
      useBusinessDaysForProgressBand: true
    });
    expect(document.getElementById("statusMessage").textContent).toContain(`祝日 ${SAMPLE_HOLIDAY_COUNT} 件`);
  });

  it("downloads current wbs xlsx with configured display range", async () => {
    bootPage();
    const exportSpy = vi.spyOn(globalThis.__mikuprojectWbsXlsx, "exportWbsWorkbook");

    parseXmlViaHook();
    document.getElementById("wbsDisplayDaysBeforeInput").value = "1";
    document.getElementById("wbsDisplayDaysAfterInput").value = "2";
    document.getElementById("exportWbsXlsxBtn").click();
    const defaultHolidayDates = getDefaultSampleHolidayDates();

    expect(exportSpy.mock.calls.at(-1)?.[1]).toEqual({
      holidayDates: defaultHolidayDates,
      displayDaysBeforeBaseDate: 1,
      displayDaysAfterBaseDate: 2,
      useBusinessDaysForDisplayRange: true,
      useBusinessDaysForProgressBand: true
    });
  });

  it("downloads current wbs xlsx with business-day display range", async () => {
    bootPage();
    const exportSpy = vi.spyOn(globalThis.__mikuprojectWbsXlsx, "exportWbsWorkbook");

    parseXmlViaHook();
    document.getElementById("wbsDisplayDaysBeforeInput").value = "1";
    document.getElementById("wbsDisplayDaysAfterInput").value = "2";
    document.getElementById("exportWbsXlsxBtn").click();
    const defaultHolidayDates = getDefaultSampleHolidayDates();

    expect(exportSpy.mock.calls.at(-1)?.[1]).toEqual({
      holidayDates: defaultHolidayDates,
      displayDaysBeforeBaseDate: 1,
      displayDaysAfterBaseDate: 2,
      useBusinessDaysForDisplayRange: true,
      useBusinessDaysForProgressBand: true
    });
  });

  it("downloads current wbs xlsx with business-day progress band", async () => {
    bootPage();
    const exportSpy = vi.spyOn(globalThis.__mikuprojectWbsXlsx, "exportWbsWorkbook");

    parseXmlViaHook();
    document.getElementById("exportWbsXlsxBtn").click();
    const defaultHolidayDates = getDefaultSampleHolidayDates();

    expect(exportSpy.mock.calls.at(-1)?.[1]).toEqual({
      holidayDates: defaultHolidayDates,
      displayDaysBeforeBaseDate: undefined,
      displayDaysAfterBaseDate: undefined,
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

    expect(document.getElementById("xmlSaveState").textContent).toContain("XML 保存状態: 未保存");
  });

  it("exports regenerated xml instead of manual textarea edits when a model exists", async () => {
    bootPage();

    parseXmlViaHook();
    const xmlInput = document.getElementById("xmlInput");
    xmlInput.value = `${xmlInput.value}\n<!-- edited -->`;
    xmlInput.dispatchEvent(new Event("input"));
    await flushAsyncWork();

    const OriginalBlob = Blob;
    class InspectableBlob extends OriginalBlob {
      constructor(parts, options) {
        super(parts, options);
        this.__parts = parts;
      }

      text() {
        return Promise.resolve(this.__parts.join(""));
      }
    }
    globalThis.Blob = InspectableBlob;
    URL.createObjectURL.mockClear();
    try {
      document.getElementById("downloadXmlBtn").click();
      const exportedBlob = URL.createObjectURL.mock.calls.at(-1)?.[0];
      await expect(exportedBlob.text()).resolves.not.toContain("<!-- edited -->");
      expect(document.getElementById("xmlInput").value).not.toContain("<!-- edited -->");
    } finally {
      globalThis.Blob = OriginalBlob;
    }
  });

  it("downloads monthly wbs calendar svg zip", async () => {
    bootPage();

    parseXmlViaHook();
    const OriginalBlob = Blob;
    class InspectableBlob extends OriginalBlob {
      constructor(parts = [], options = {}) {
        super(parts, options);
        this._parts = parts;
      }
    }
    globalThis.Blob = InspectableBlob;
    URL.createObjectURL.mockClear();
    try {
      document.getElementById("downloadMonthlyCalendarSvgBtn").click();
      const zipBlob = URL.createObjectURL.mock.calls.at(-1)?.[0];
      const rawPart = zipBlob._parts?.[0];
      const codec = new globalThis.__mikuprojectExcelIo.XlsxWorkbookCodec();
      const entries = codec.unpackEntries(rawPart);
      expect(Object.keys(entries).sort()).toContain("monthly-calendar/2026-03.svg");
      const marchSvg = new TextDecoder().decode(entries["monthly-calendar/2026-03.svg"]);
      expect(marchSvg).toContain("mikuproject開発");
    } finally {
      globalThis.Blob = OriginalBlob;
    }
  });

  it("downloads all outputs as zip", () => {
    bootPage();

    parseXmlViaHook();
    const OriginalBlob = Blob;
    class InspectableBlob extends OriginalBlob {
      constructor(parts = [], options = {}) {
        super(parts, options);
        this._parts = parts;
      }
    }
    globalThis.Blob = InspectableBlob;
    URL.createObjectURL.mockClear();
    try {
      document.getElementById("downloadAllOutputsBtn").click();
      const zipBlob = URL.createObjectURL.mock.calls.at(-1)?.[0];
      const rawPart = zipBlob._parts?.[0];
      const codec = new globalThis.__mikuprojectExcelIo.XlsxWorkbookCodec();
      const entries = codec.unpackEntries(rawPart);
      const sortedEntryNames = Object.keys(entries).sort();
      expect(sortedEntryNames).toContain("README.txt");
      expect(sortedEntryNames).toContain("mikuproject-export-202603162312.xml");
      expect(sortedEntryNames).toContain("mikuproject-export-202603162312.xlsx");
      expect(sortedEntryNames).toContain("mikuproject-workbook-202603162312.json");
      expect(sortedEntryNames).toContain("mikuproject-export-202603162312.csv");
      expect(sortedEntryNames).toContain("mikuproject-wbs-202603162312.xlsx");
      expect(sortedEntryNames).toContain("mikuproject-wbs-20260316.md");
      expect(sortedEntryNames).toContain("mikuproject-wbs-daily-202603162312.svg");
      expect(sortedEntryNames).toContain("mikuproject-wbs-weekly-202603162312.svg");
      expect(sortedEntryNames).toContain("monthly-calendar/2026-03.svg");
      expect(sortedEntryNames).toContain("monthly-calendar/2026-04.svg");
      expect(sortedEntryNames).toContain("mikuproject-wbs-mermaid-202603162312.md");
      expect(sortedEntryNames).toContain("mikuproject-project-overview-view.editjson");
      expect(sortedEntryNames).toContain("mikuproject-full-bundle.editjson");
      expect(sortedEntryNames).toContain("mikuproject-phase-detail-view-full.editjson");
      const monthlySvg = new TextDecoder().decode(entries["monthly-calendar/2026-03.svg"]);
      expect(monthlySvg).toContain("mikuproject開発");
    } finally {
      globalThis.Blob = OriginalBlob;
    }
  });

  it("keeps complex mermaid dependencies as comments", () => {
    const xmlTools = bootXmlModule();
    const model = {
      project: {
        name: "Mermaid Complex",
        startDate: "2026-03-16T09:00:00",
        finishDate: "2026-03-20T18:00:00",
        scheduleFromStart: true,
        outlineCodes: [],
        wbsMasks: [],
        extendedAttributes: []
      },
      tasks: [
        {
          uid: "1",
          id: "1",
          name: "Prep",
          outlineLevel: 1,
          outlineNumber: "1",
          start: "2026-03-16T09:00:00",
          finish: "2026-03-16T18:00:00",
          duration: "PT8H0M0S",
          milestone: false,
          summary: false,
          percentComplete: 100,
          predecessors: [],
          extendedAttributes: [],
          baselines: [],
          timephasedData: []
        },
        {
          uid: "2",
          id: "2",
          name: "Review",
          outlineLevel: 1,
          outlineNumber: "2",
          start: "2026-03-17T09:00:00",
          finish: "2026-03-17T18:00:00",
          duration: "PT8H0M0S",
          milestone: false,
          summary: false,
          percentComplete: 0,
          predecessors: [],
          extendedAttributes: [],
          baselines: [],
          timephasedData: []
        },
        {
          uid: "3",
          id: "3",
          name: "Ship",
          outlineLevel: 1,
          outlineNumber: "3",
          start: "2026-03-18T09:00:00",
          finish: "2026-03-18T18:00:00",
          duration: "PT8H0M0S",
          milestone: false,
          summary: false,
          percentComplete: 0,
          predecessors: [
            { predecessorUid: "1", type: 1, linkLag: "PT2H0M0S" },
            { predecessorUid: "2", type: 4 }
          ],
          extendedAttributes: [],
          baselines: [],
          timephasedData: []
        }
      ],
      resources: [],
      assignments: [],
      calendars: []
    };

    const mermaidText = xmlTools.exportMermaidGantt(model);
    expect(mermaidText).toContain("Ship :task_3, 2026-03-18T09:00:00, 2026-03-18T18:00:00");
    expect(mermaidText).toContain("%% dependency: Ship after Prep (type=FS, lag=2h) [task_3 after task_1]");
    expect(mermaidText).toContain("%% dependency(note): Ship has multiple predecessors");
  });

  it("sanitizes date-leading mermaid gantt labels", () => {
    const xmlTools = bootXmlModule();
    const model = {
      project: {
        name: "2026-03 mikuproject開発",
        startDate: "2026-03-16T09:00:00",
        finishDate: "2026-03-16T18:00:00",
        scheduleFromStart: true,
        outlineCodes: [],
        wbsMasks: [],
        extendedAttributes: []
      },
      tasks: [
        {
          uid: "1",
          id: "1",
          name: "2026-03-16 初期実装（42513dd：XML import/export）",
          outlineLevel: 1,
          outlineNumber: "1",
          start: "2026-03-16T09:00:00",
          finish: "2026-03-16T18:00:00",
          duration: "PT8H0M0S",
          milestone: false,
          summary: false,
          percentComplete: 0,
          predecessors: [],
          extendedAttributes: [],
          baselines: [],
          timephasedData: []
        }
      ],
      resources: [],
      assignments: [],
      calendars: []
    };

    const mermaidText = xmlTools.exportMermaidGantt(model);
    expect(mermaidText).toContain("title Project 2026-03 mikuproject開発");
    expect(mermaidText).toContain("Task 2026-03-16 初期実装（42513dd XML import/export） :task_1, 2026-03-16T09:00:00, 2026-03-16T18:00:00");
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
