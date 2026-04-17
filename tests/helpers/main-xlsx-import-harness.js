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
const excelIoUtilCode = readFileSync(
  path.resolve(__dirname, "../../src/js/excel-io-util.js"),
  "utf8"
);
const excelIoCode = readFileSync(
  path.resolve(__dirname, "../../src/js/excel-io.js"),
  "utf8"
);
const excelIoZipCode = readFileSync(
  path.resolve(__dirname, "../../src/js/excel-io-zip.js"),
  "utf8"
);
const excelIoNormalizeCode = readFileSync(
  path.resolve(__dirname, "../../src/js/excel-io-normalize.js"),
  "utf8"
);
const excelIoPackageXmlCode = readFileSync(
  path.resolve(__dirname, "../../src/js/excel-io-package-xml.js"),
  "utf8"
);
const excelIoWorksheetBuildCode = readFileSync(
  path.resolve(__dirname, "../../src/js/excel-io-worksheet-build.js"),
  "utf8"
);
const excelIoWorksheetParseCode = readFileSync(
  path.resolve(__dirname, "../../src/js/excel-io-worksheet-parse.js"),
  "utf8"
);
const excelIoWorkbookParseCode = readFileSync(
  path.resolve(__dirname, "../../src/js/excel-io-workbook-parse.js"),
  "utf8"
);
const excelIoWorkbookBuildCode = readFileSync(
  path.resolve(__dirname, "../../src/js/excel-io-workbook-build.js"),
  "utf8"
);
const excelIoStylesBuildCode = readFileSync(
  path.resolve(__dirname, "../../src/js/excel-io-styles-build.js"),
  "utf8"
);
const excelIoStylesParseCode = readFileSync(
  path.resolve(__dirname, "../../src/js/excel-io-styles-parse.js"),
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
const projectWorkbookSchemaCode = readFileSync(
  path.resolve(__dirname, "../../src/js/project-workbook-schema.js"),
  "utf8"
);
const projectXlsxImportUtilCode = readFileSync(
  path.resolve(__dirname, "../../src/js/project-xlsx-import-util.js"),
  "utf8"
);
const projectXlsxImportProjectCode = readFileSync(
  path.resolve(__dirname, "../../src/js/project-xlsx-import-project.js"),
  "utf8"
);
const projectXlsxImportCalendarsCode = readFileSync(
  path.resolve(__dirname, "../../src/js/project-xlsx-import-calendars.js"),
  "utf8"
);
const projectXlsxImportEntitiesCode = readFileSync(
  path.resolve(__dirname, "../../src/js/project-xlsx-import-entities.js"),
  "utf8"
);
const projectXlsxImportCode = readFileSync(
  path.resolve(__dirname, "../../src/js/project-xlsx-import.js"),
  "utf8"
);
const projectXlsxExportUtilCode = readFileSync(
  path.resolve(__dirname, "../../src/js/project-xlsx-export-util.js"),
  "utf8"
);
const projectXlsxExportProjectCode = readFileSync(
  path.resolve(__dirname, "../../src/js/project-xlsx-export-project.js"),
  "utf8"
);
const projectXlsxExportEntitiesCode = readFileSync(
  path.resolve(__dirname, "../../src/js/project-xlsx-export-entities.js"),
  "utf8"
);
const projectXlsxExportCalendarsCode = readFileSync(
  path.resolve(__dirname, "../../src/js/project-xlsx-export-calendars.js"),
  "utf8"
);
const projectXlsxExportCode = readFileSync(
  path.resolve(__dirname, "../../src/js/project-xlsx-export.js"),
  "utf8"
);
const projectXlsxCode = readFileSync(
  path.resolve(__dirname, "../../src/js/project-xlsx.js"),
  "utf8"
);
const projectWorkbookJsonValidateCode = readFileSync(
  path.resolve(__dirname, "../../src/js/project-workbook-json-validate.js"),
  "utf8"
);
const projectWorkbookJsonImportCode = readFileSync(
  path.resolve(__dirname, "../../src/js/project-workbook-json-import.js"),
  "utf8"
);
const projectWorkbookJsonExportCode = readFileSync(
  path.resolve(__dirname, "../../src/js/project-workbook-json-export.js"),
  "utf8"
);
const projectWorkbookJsonCode = readFileSync(
  path.resolve(__dirname, "../../src/js/project-workbook-json.js"),
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

export const dependencyXml = readFileSync(
  path.resolve(__dirname, "../../testdata/dependency.xml"),
  "utf8"
);
export const workbookImportSampleJson = readFileSync(
  path.resolve(__dirname, "../../testdata/workbook-import-sample.json"),
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

const bootPageCode = `${typesCode}\n${markdownEscapeCode}\n${aiJsonUtilCode}\n${mainUtilCode}\n${excelIoUtilCode}\n${excelIoZipCode}\n${excelIoNormalizeCode}\n${excelIoPackageXmlCode}\n${excelIoWorksheetBuildCode}\n${excelIoWorksheetParseCode}\n${excelIoWorkbookParseCode}\n${excelIoWorkbookBuildCode}\n${excelIoStylesBuildCode}\n${excelIoStylesParseCode}\n${excelIoCode}\n${msProjectAiViewsCode}\n${msProjectCalendarCode}\n${msProjectSamplesCode}\n${msProjectCsvCode}\n${msProjectValidateHelpersCode}\n${msProjectValidateCode}\n${msProjectXmlDomCode}\n${msProjectCodecCode}\n${msProjectMermaidCode}\n${msProjectXmlCode}\n${projectWorkbookSchemaCode}\n${projectXlsxImportUtilCode}\n${projectXlsxImportProjectCode}\n${projectXlsxImportCalendarsCode}\n${projectXlsxImportEntitiesCode}\n${projectXlsxImportCode}\n${projectXlsxExportUtilCode}\n${projectXlsxExportProjectCode}\n${projectXlsxExportEntitiesCode}\n${projectXlsxExportCalendarsCode}\n${projectXlsxExportCode}\n${projectXlsxCode}\n${projectWorkbookJsonValidateCode}\n${projectWorkbookJsonImportCode}\n${projectWorkbookJsonExportCode}\n${projectWorkbookJsonCode}\n${projectPatchJsonStubCode}\n${wbsXlsxStubCode}\n${wbsMarkdownStubCode}\n${nativeSvgStubCode}\n${mainIoCode}\n${mainImportCode}\n${mainExportCode}\n${mainInputEventsCode}\n${mainButtonEventsCode}\n${mainEventsCode}\n${mainPreviewCode}\n${mainUiCode}\n${mainSupportCode}\n${mainSamplesCode}\n${mainDownloadsCode}\n${mainOutputActionsCode}\n${mainImportActionsCode}\n${mainXmlActionsCode}\n${mainArchiveActionsCode}\n${mainSaveStateCode}\n${mainTabActionsCode}\n${mainPreviewActionsCode}\n${mainTransformCode}\n${mainFlowCode}\n${mainModelCode}\n${mainRenderCode}\n${mainCode}`;
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

export function bootPage() {
  mountDom();
  bootPageRunner();
  document.dispatchEvent(new Event("DOMContentLoaded"));
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

export function setImportFiles(file) {
  const importInput = document.getElementById("importFileInput");
  Object.defineProperty(importInput, "files", {
    configurable: true,
    value: [file]
  });
  importInput.dispatchEvent(new Event("change"));
}

export function setupMainXlsxImportDom() {
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
