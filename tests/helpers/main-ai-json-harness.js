import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { vi } from "vitest";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const typesCode = readFileSync(
  path.resolve(__dirname, "../../src/js/types.js"),
  "utf8"
);
const markdownEscapeCode = readFileSync(
  path.resolve(__dirname, "../../src/js/markdown-escape.js"),
  "utf8"
);
const aiJsonUtilCode = readFileSync(
  path.resolve(__dirname, "../../src/js/ai-json-util.js"),
  "utf8"
);
const mainUtilCode = readFileSync(
  path.resolve(__dirname, "../../src/js/main-util.js"),
  "utf8"
);
const msProjectAiViewsCode = readFileSync(
  path.resolve(__dirname, "../../src/js/msproject-ai-views.js"),
  "utf8"
);
const msProjectCalendarCode = readFileSync(
  path.resolve(__dirname, "../../src/js/msproject-calendar.js"),
  "utf8"
);
const msProjectSamplesCode = readFileSync(
  path.resolve(__dirname, "../../src/js/msproject-samples.js"),
  "utf8"
);
const msProjectCsvCode = readFileSync(
  path.resolve(__dirname, "../../src/js/msproject-csv.js"),
  "utf8"
);
const msProjectValidateHelpersCode = readFileSync(
  path.resolve(__dirname, "../../src/js/msproject-validate-helpers.js"),
  "utf8"
);
const msProjectValidateCode = readFileSync(
  path.resolve(__dirname, "../../src/js/msproject-validate.js"),
  "utf8"
);
const msProjectXmlDomCode = readFileSync(
  path.resolve(__dirname, "../../src/js/msproject-xml-dom.js"),
  "utf8"
);
const msProjectCodecCode = readFileSync(
  path.resolve(__dirname, "../../src/js/msproject-codec.js"),
  "utf8"
);
const msProjectMermaidCode = readFileSync(
  path.resolve(__dirname, "../../src/js/msproject-mermaid.js"),
  "utf8"
);
const msProjectXmlCode = readFileSync(
  path.resolve(__dirname, "../../src/js/msproject-xml.js"),
  "utf8"
);
const projectPatchJsonCode = readFileSync(
  path.resolve(__dirname, "../../src/js/project-patch-json.js"),
  "utf8"
);
const projectPatchJsonUtilCode = readFileSync(
  path.resolve(__dirname, "../../src/js/project-patch-json-util.js"),
  "utf8"
);
const projectPatchJsonLinksCode = readFileSync(
  path.resolve(__dirname, "../../src/js/project-patch-json-links.js"),
  "utf8"
);
const projectPatchJsonTasksCode = readFileSync(
  path.resolve(__dirname, "../../src/js/project-patch-json-tasks.js"),
  "utf8"
);
const projectPatchJsonEntitiesCode = readFileSync(
  path.resolve(__dirname, "../../src/js/project-patch-json-entities.js"),
  "utf8"
);
const projectPatchJsonUpdatesCode = readFileSync(
  path.resolve(__dirname, "../../src/js/project-patch-json-updates.js"),
  "utf8"
);
const projectPatchJsonCoreCode = readFileSync(
  path.resolve(__dirname, "../../src/js/project-patch-json-core.js"),
  "utf8"
);
const mainIoCode = readFileSync(
  path.resolve(__dirname, "../../src/js/main-io.js"),
  "utf8"
);
const mainImportCode = readFileSync(
  path.resolve(__dirname, "../../src/js/main-import.js"),
  "utf8"
);
const mainExportCode = readFileSync(
  path.resolve(__dirname, "../../src/js/main-export.js"),
  "utf8"
);
const mainInputEventsCode = readFileSync(
  path.resolve(__dirname, "../../src/js/main-input-events.js"),
  "utf8"
);
const mainButtonEventsCode = readFileSync(
  path.resolve(__dirname, "../../src/js/main-button-events.js"),
  "utf8"
);
const mainEventsCode = readFileSync(
  path.resolve(__dirname, "../../src/js/main-events.js"),
  "utf8"
);
const mainPreviewCode = readFileSync(
  path.resolve(__dirname, "../../src/js/main-preview.js"),
  "utf8"
);
const mainUiCode = readFileSync(
  path.resolve(__dirname, "../../src/js/main-ui.js"),
  "utf8"
);
const mainSupportCode = readFileSync(
  path.resolve(__dirname, "../../src/js/main-support.js"),
  "utf8"
);
const mainSamplesCode = readFileSync(
  path.resolve(__dirname, "../../src/js/main-samples.js"),
  "utf8"
);
const mainDownloadsCode = readFileSync(
  path.resolve(__dirname, "../../src/js/main-downloads.js"),
  "utf8"
);
const mainOutputActionsCode = readFileSync(
  path.resolve(__dirname, "../../src/js/main-output-actions.js"),
  "utf8"
);
const mainImportActionsCode = readFileSync(
  path.resolve(__dirname, "../../src/js/main-import-actions.js"),
  "utf8"
);
const mainXmlActionsCode = readFileSync(
  path.resolve(__dirname, "../../src/js/main-xml-actions.js"),
  "utf8"
);
const mainArchiveActionsCode = readFileSync(
  path.resolve(__dirname, "../../src/js/main-archive-actions.js"),
  "utf8"
);
const mainSaveStateCode = readFileSync(
  path.resolve(__dirname, "../../src/js/main-save-state.js"),
  "utf8"
);
const mainTabActionsCode = readFileSync(
  path.resolve(__dirname, "../../src/js/main-tab-actions.js"),
  "utf8"
);
const mainPreviewActionsCode = readFileSync(
  path.resolve(__dirname, "../../src/js/main-preview-actions.js"),
  "utf8"
);
const mainTransformCode = readFileSync(
  path.resolve(__dirname, "../../src/js/main-transform.js"),
  "utf8"
);
const mainFlowCode = readFileSync(
  path.resolve(__dirname, "../../src/js/main-flow.js"),
  "utf8"
);
const mainModelCode = readFileSync(
  path.resolve(__dirname, "../../src/js/main-model.js"),
  "utf8"
);
const mainRenderCode = readFileSync(
  path.resolve(__dirname, "../../src/js/main-render.js"),
  "utf8"
);
const mainCode = readFileSync(
  path.resolve(__dirname, "../../src/js/main.js"),
  "utf8"
);

export const hierarchyXml = readFileSync(
  path.resolve(__dirname, "../../testdata/hierarchy.xml"),
  "utf8"
);
export const dependencyXml = readFileSync(
  path.resolve(__dirname, "../../testdata/dependency.xml"),
  "utf8"
);

const excelIoStubCode = `
globalThis.__mikuprojectExcelIo = {
  XlsxWorkbookCodec: function () {
    this.exportWorkbook = () => new Uint8Array();
    this.importWorkbook = () => ({ sheets: [] });
    this.importWorkbookAsync = async () => ({ sheets: [] });
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
  exportProjectWorkbookJson: () => ({ format: "mikuproject_workbook_json", sheets: [] }),
  importProjectWorkbookJson: (_documentLike, baseModel) => ({ model: baseModel, changes: [], warnings: [] }),
  validateWorkbookJsonDocument: (documentLike) => ({ document: documentLike, warnings: [] })
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
  exportWeeklyNativeSvg: () => "<svg data-stub=\\"weekly\\">weekly overview</svg>",
  exportMonthlyWbsCalendarSvgArchive: () => ({
    entries: [{ fileName: "2026-03.svg", svg: "<svg data-stub=\\"monthly\\">2026-03</svg>" }],
    zipBytes: new Uint8Array()
  })
};
`;

const bootPageCode = `${typesCode}\n${markdownEscapeCode}\n${aiJsonUtilCode}\n${mainUtilCode}\n${msProjectAiViewsCode}\n${msProjectCalendarCode}\n${msProjectSamplesCode}\n${msProjectCsvCode}\n${msProjectValidateHelpersCode}\n${msProjectValidateCode}\n${msProjectXmlDomCode}\n${msProjectCodecCode}\n${msProjectMermaidCode}\n${msProjectXmlCode}\n${projectPatchJsonUtilCode}\n${projectPatchJsonLinksCode}\n${projectPatchJsonTasksCode}\n${projectPatchJsonEntitiesCode}\n${projectPatchJsonUpdatesCode}\n${projectPatchJsonCoreCode}\n${projectPatchJsonCode}\n${excelIoStubCode}\n${projectXlsxStubCode}\n${projectWorkbookJsonStubCode}\n${wbsXlsxStubCode}\n${wbsMarkdownStubCode}\n${nativeSvgStubCode}\n${mainIoCode}\n${mainImportCode}\n${mainExportCode}\n${mainInputEventsCode}\n${mainButtonEventsCode}\n${mainEventsCode}\n${mainPreviewCode}\n${mainUiCode}\n${mainSupportCode}\n${mainSamplesCode}\n${mainDownloadsCode}\n${mainOutputActionsCode}\n${mainImportActionsCode}\n${mainXmlActionsCode}\n${mainArchiveActionsCode}\n${mainSaveStateCode}\n${mainTabActionsCode}\n${mainPreviewActionsCode}\n${mainTransformCode}\n${mainFlowCode}\n${mainModelCode}\n${mainRenderCode}\n${mainCode}`;
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
    <button id="copyAiPromptBtnPane" type="button"></button>
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
      <template id="aiPromptTemplate"># mikuproject AI JSON Spec

あなたはこれから mikuproject とやりとりします。</template>
      <textarea id="projectDraftImportInput"></textarea>
      <div id="xmlSaveState"></div>
    </section>
    <section id="tabPanelTransform" class="md-tab-panel" data-tab-panel="transform" hidden>
      <div id="summaryProjectName"></div>
      <div id="summaryTaskCount"></div>
      <div id="summaryResourceCount"></div>
      <div id="summaryAssignmentCount"></div>
      <div id="summaryCalendarCount"></div>
      <div id="nativeSvgPreview"></div>
      <textarea id="modelOutput"></textarea>
      <div id="projectPreview"></div>
      <div id="taskPreview"></div>
      <div id="resourcePreview"></div>
      <div id="assignmentPreview"></div>
      <div id="calendarPreview"></div>
      <textarea id="mermaidOutput"></textarea>
      <div id="validationIssues" class="md-hidden"></div>
      <div id="importWarnings" class="md-hidden"></div>
      <div id="xlsxImportSummary" class="md-hidden"></div>
    </section>
    <section id="tabPanelOutput" class="md-tab-panel" data-tab-panel="output" hidden>
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
  const copyButton = document.getElementById("copyAiPromptBtnPane");
  if (copyButton) {
    copyButton.id = "copyAiPromptBtn";
  }
}

export function bootPage() {
  mountDom();
  bootPageRunner();
  document.dispatchEvent(new Event("DOMContentLoaded"));
}

export function setupMainAiJsonTestDom() {
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
}

export function getMainHooks() {
  return globalThis.__mikuprojectMainTestHooks;
}

export function parseXmlViaHook() {
  getMainHooks().parseCurrentXml();
}

export async function flushAsyncWork() {
  await Promise.resolve();
  await Promise.resolve();
}
