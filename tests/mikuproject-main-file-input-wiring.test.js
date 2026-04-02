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
const mainRenderCode = readFileSync(
  path.resolve(__dirname, "../src/js/main-render.js"),
  "utf8"
);
const mainCode = readFileSync(
  path.resolve(__dirname, "../src/js/main.js"),
  "utf8"
);

const msProjectXmlStubCode = `
globalThis.__mikuprojectXml = {
  SAMPLE_XML: "<Project><Name>Stub</Name><StartDate>2026-03-16T09:00:00</StartDate><FinishDate>2026-03-16T18:00:00</FinishDate><ScheduleFromStart>1</ScheduleFromStart><Tasks /><Resources /><Assignments /></Project>",
  SAMPLE_PROJECT_DRAFT_VIEW: {},
  importMsProjectXml: () => ({
    project: { name: "Stub", title: "Stub", startDate: "2026-03-16T09:00:00", finishDate: "2026-03-16T18:00:00", scheduleFromStart: true, outlineCodes: [], wbsMasks: [], extendedAttributes: [] },
    tasks: [],
    resources: [],
    assignments: [],
    calendars: []
  }),
  importCsvParentId: () => ({
    project: { name: "Stub", title: "Stub", startDate: "2026-03-16T09:00:00", finishDate: "2026-03-16T18:00:00", scheduleFromStart: true, outlineCodes: [], wbsMasks: [], extendedAttributes: [] },
    tasks: [],
    resources: [],
    assignments: [],
    calendars: []
  }),
  exportMsProjectXml: () => "<Project />",
  exportMermaidGantt: () => "gantt",
  buildProjectDraftRequest: () => ({}),
  importProjectDraftView: () => ({
    project: { name: "Stub", title: "Stub", startDate: "2026-03-16T09:00:00", finishDate: "2026-03-16T18:00:00", scheduleFromStart: true, outlineCodes: [], wbsMasks: [], extendedAttributes: [] },
    tasks: [],
    resources: [],
    assignments: [],
    calendars: []
  }),
  exportProjectOverviewView: () => ({}),
  exportPhaseDetailView: () => ({}),
  exportCsvParentId: () => "ID,Name",
  normalizeProjectModel: (model) => model,
  validateProjectModel: () => []
};
`;

const excelIoStubCode = `
globalThis.__mikuprojectExcelIo = {
  XlsxWorkbookCodec: function () {
    this.exportWorkbook = () => new Uint8Array();
    this.importWorkbook = () => ({ sheets: [] });
  }
};
`;

const projectXlsxStubCode = `
globalThis.__mikuprojectProjectXlsx = {
  exportProjectWorkbook: () => ({ sheets: [] }),
  importProjectWorkbook: (_workbook, baseModel) => baseModel,
  importProjectWorkbookDetailed: (_workbook, baseModel) => ({ model: baseModel, changes: [] })
};
`;

const projectWorkbookJsonStubCode = `
globalThis.__mikuprojectProjectWorkbookJson = {
  exportProjectWorkbookJson: () => ({}),
  importProjectWorkbookJson: (_documentLike, baseModel) => ({ model: baseModel, changes: [], warnings: [] }),
  validateWorkbookJsonDocument: (documentLike) => ({ document: documentLike, warnings: [] })
};
`;

const projectPatchJsonStubCode = `
globalThis.__mikuprojectProjectPatchJson = {
  importProjectPatchJson: (_documentLike, baseModel) => ({ model: baseModel, changes: [], warnings: [] }),
  validatePatchDocument: (documentLike) => ({ document: documentLike, warnings: [] })
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
  exportNativeSvg: () => "<svg></svg>",
  exportWeeklyNativeSvg: () => "<svg></svg>",
  exportMonthlyWbsCalendarSvgArchive: () => ({ entries: [], zipBytes: new Uint8Array() })
};
`;

const bootPageCode = `${typesCode}\n${markdownEscapeCode}\n${aiJsonUtilCode}\n${mainUtilCode}\n${msProjectXmlStubCode}\n${excelIoStubCode}\n${projectXlsxStubCode}\n${projectWorkbookJsonStubCode}\n${projectPatchJsonStubCode}\n${wbsXlsxStubCode}\n${wbsMarkdownStubCode}\n${nativeSvgStubCode}\n${mainRenderCode}\n${mainCode}`;
const bootPageRunner = new Function(bootPageCode);

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
      <div id="validationIssues" class="md-hidden"></div>
      <div id="importWarnings" class="md-hidden"></div>
      <div id="xlsxImportSummary" class="md-hidden"></div>
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
  bootPageRunner();
  document.dispatchEvent(new Event("DOMContentLoaded"));
}

describe("mikuproject main file input wiring", () => {
  beforeEach(() => {
    document.body.innerHTML = "";
  });

  it("clears file input value before reselecting the same file", () => {
    bootPage();

    const importInput = document.getElementById("importFileInput");
    Object.defineProperty(importInput, "value", {
      configurable: true,
      writable: true,
      value: "same-file.xlsx"
    });

    importInput.dispatchEvent(new Event("click"));

    expect(importInput.value).toBe("");
  });

  it("opens file chooser from the visible import button after clearing the input value", () => {
    bootPage();

    const importInput = document.getElementById("importFileInput");
    Object.defineProperty(importInput, "value", {
      configurable: true,
      writable: true,
      value: "stale-value.xlsx"
    });
    importInput.click = vi.fn();

    document.getElementById("importFileBtn").click();

    expect(importInput.value).toBe("");
    expect(importInput.click).toHaveBeenCalled();
  });
});
