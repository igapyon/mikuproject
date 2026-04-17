import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { vi } from "vitest";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const typesCode = readFileSync(path.resolve(__dirname, "../../src/js/types.js"), "utf8");
const markdownEscapeCode = readFileSync(path.resolve(__dirname, "../../src/js/markdown-escape.js"), "utf8");
const aiJsonUtilCode = readFileSync(path.resolve(__dirname, "../../src/js/ai-json-util.js"), "utf8");
const mainUtilCode = readFileSync(path.resolve(__dirname, "../../src/js/main-util.js"), "utf8");
const excelIoUtilCode = readFileSync(path.resolve(__dirname, "../../src/js/excel-io-util.js"), "utf8");
const excelIoZipCode = readFileSync(path.resolve(__dirname, "../../src/js/excel-io-zip.js"), "utf8");
const excelIoNormalizeCode = readFileSync(path.resolve(__dirname, "../../src/js/excel-io-normalize.js"), "utf8");
const excelIoPackageXmlCode = readFileSync(path.resolve(__dirname, "../../src/js/excel-io-package-xml.js"), "utf8");
const excelIoWorksheetBuildCode = readFileSync(path.resolve(__dirname, "../../src/js/excel-io-worksheet-build.js"), "utf8");
const excelIoWorksheetParseCode = readFileSync(path.resolve(__dirname, "../../src/js/excel-io-worksheet-parse.js"), "utf8");
const excelIoWorkbookParseCode = readFileSync(path.resolve(__dirname, "../../src/js/excel-io-workbook-parse.js"), "utf8");
const excelIoWorkbookBuildCode = readFileSync(path.resolve(__dirname, "../../src/js/excel-io-workbook-build.js"), "utf8");
const excelIoStylesBuildCode = readFileSync(path.resolve(__dirname, "../../src/js/excel-io-styles-build.js"), "utf8");
const excelIoStylesParseCode = readFileSync(path.resolve(__dirname, "../../src/js/excel-io-styles-parse.js"), "utf8");
const excelIoCode = readFileSync(path.resolve(__dirname, "../../src/js/excel-io.js"), "utf8");
const msProjectAiViewsCode = readFileSync(path.resolve(__dirname, "../../src/js/msproject-ai-views.js"), "utf8");
const msProjectCalendarCode = readFileSync(path.resolve(__dirname, "../../src/js/msproject-calendar.js"), "utf8");
const msProjectSamplesCode = readFileSync(path.resolve(__dirname, "../../src/js/msproject-samples.js"), "utf8");
const msProjectCsvCode = readFileSync(path.resolve(__dirname, "../../src/js/msproject-csv.js"), "utf8");
const msProjectValidateHelpersCode = readFileSync(path.resolve(__dirname, "../../src/js/msproject-validate-helpers.js"), "utf8");
const msProjectValidateCode = readFileSync(path.resolve(__dirname, "../../src/js/msproject-validate.js"), "utf8");
const msProjectXmlDomCode = readFileSync(path.resolve(__dirname, "../../src/js/msproject-xml-dom.js"), "utf8");
const msProjectCodecCode = readFileSync(path.resolve(__dirname, "../../src/js/msproject-codec.js"), "utf8");
const msProjectMermaidCode = readFileSync(path.resolve(__dirname, "../../src/js/msproject-mermaid.js"), "utf8");
const msProjectXmlCode = readFileSync(path.resolve(__dirname, "../../src/js/msproject-xml.js"), "utf8");
const projectWorkbookSchemaCode = readFileSync(path.resolve(__dirname, "../../src/js/project-workbook-schema.js"), "utf8");
const projectXlsxImportUtilCode = readFileSync(path.resolve(__dirname, "../../src/js/project-xlsx-import-util.js"), "utf8");
const projectXlsxImportProjectCode = readFileSync(path.resolve(__dirname, "../../src/js/project-xlsx-import-project.js"), "utf8");
const projectXlsxImportCalendarsCode = readFileSync(path.resolve(__dirname, "../../src/js/project-xlsx-import-calendars.js"), "utf8");
const projectXlsxImportEntitiesCode = readFileSync(path.resolve(__dirname, "../../src/js/project-xlsx-import-entities.js"), "utf8");
const projectXlsxImportCode = readFileSync(path.resolve(__dirname, "../../src/js/project-xlsx-import.js"), "utf8");
const projectXlsxExportUtilCode = readFileSync(path.resolve(__dirname, "../../src/js/project-xlsx-export-util.js"), "utf8");
const projectXlsxExportProjectCode = readFileSync(path.resolve(__dirname, "../../src/js/project-xlsx-export-project.js"), "utf8");
const projectXlsxExportEntitiesCode = readFileSync(path.resolve(__dirname, "../../src/js/project-xlsx-export-entities.js"), "utf8");
const projectXlsxExportCalendarsCode = readFileSync(path.resolve(__dirname, "../../src/js/project-xlsx-export-calendars.js"), "utf8");
const projectXlsxExportCode = readFileSync(path.resolve(__dirname, "../../src/js/project-xlsx-export.js"), "utf8");
const projectXlsxCode = readFileSync(path.resolve(__dirname, "../../src/js/project-xlsx.js"), "utf8");
const projectWorkbookJsonValidateCode = readFileSync(path.resolve(__dirname, "../../src/js/project-workbook-json-validate.js"), "utf8");
const projectWorkbookJsonImportCode = readFileSync(path.resolve(__dirname, "../../src/js/project-workbook-json-import.js"), "utf8");
const projectWorkbookJsonExportCode = readFileSync(path.resolve(__dirname, "../../src/js/project-workbook-json-export.js"), "utf8");
const projectWorkbookJsonCode = readFileSync(path.resolve(__dirname, "../../src/js/project-workbook-json.js"), "utf8");
const projectPatchJsonUtilCode = readFileSync(path.resolve(__dirname, "../../src/js/project-patch-json-util.js"), "utf8");
const projectPatchJsonLinksCode = readFileSync(path.resolve(__dirname, "../../src/js/project-patch-json-links.js"), "utf8");
const projectPatchJsonTasksCode = readFileSync(path.resolve(__dirname, "../../src/js/project-patch-json-tasks.js"), "utf8");
const projectPatchJsonEntitiesCode = readFileSync(path.resolve(__dirname, "../../src/js/project-patch-json-entities.js"), "utf8");
const projectPatchJsonUpdatesCode = readFileSync(path.resolve(__dirname, "../../src/js/project-patch-json-updates.js"), "utf8");
const projectPatchJsonCoreCode = readFileSync(path.resolve(__dirname, "../../src/js/project-patch-json-core.js"), "utf8");
const projectPatchJsonCode = readFileSync(path.resolve(__dirname, "../../src/js/project-patch-json.js"), "utf8");
const wbsDatebandCode = readFileSync(path.resolve(__dirname, "../../src/js/wbs-dateband.js"), "utf8");
const wbsXlsxBaseCode = readFileSync(path.resolve(__dirname, "../../src/js/wbs-xlsx-base.js"), "utf8");
const wbsXlsxTaskmetaCode = readFileSync(path.resolve(__dirname, "../../src/js/wbs-xlsx-taskmeta.js"), "utf8");
const wbsXlsxLayoutCode = readFileSync(path.resolve(__dirname, "../../src/js/wbs-xlsx-layout.js"), "utf8");
const wbsXlsxSectionsCode = readFileSync(path.resolve(__dirname, "../../src/js/wbs-xlsx-sections.js"), "utf8");
const wbsXlsxCellsCode = readFileSync(path.resolve(__dirname, "../../src/js/wbs-xlsx-cells.js"), "utf8");
const wbsXlsxExportCode = readFileSync(path.resolve(__dirname, "../../src/js/wbs-xlsx-export.js"), "utf8");
const wbsXlsxPublicCode = readFileSync(path.resolve(__dirname, "../../src/js/wbs-xlsx-public.js"), "utf8");
const wbsXlsxCode = readFileSync(path.resolve(__dirname, "../../src/js/wbs-xlsx.js"), "utf8");
const wbsMarkdownCode = readFileSync(path.resolve(__dirname, "../../src/js/wbs-markdown.js"), "utf8");
const wbsSvgZipCode = readFileSync(path.resolve(__dirname, "../../src/js/wbs-svg-zip.js"), "utf8");
const wbsSvgBarsCode = readFileSync(path.resolve(__dirname, "../../src/js/wbs-svg-bars.js"), "utf8");
const wbsSvgViewportCode = readFileSync(path.resolve(__dirname, "../../src/js/wbs-svg-viewport.js"), "utf8");
const wbsSvgAxisCode = readFileSync(path.resolve(__dirname, "../../src/js/wbs-svg-axis.js"), "utf8");
const wbsSvgScaffoldCode = readFileSync(path.resolve(__dirname, "../../src/js/wbs-svg-scaffold.js"), "utf8");
const wbsSvgLabelsCode = readFileSync(path.resolve(__dirname, "../../src/js/wbs-svg-labels.js"), "utf8");
const wbsSvgCalendarCode = readFileSync(path.resolve(__dirname, "../../src/js/wbs-svg-calendar.js"), "utf8");
const wbsSvgTimelineCode = readFileSync(path.resolve(__dirname, "../../src/js/wbs-svg-timeline.js"), "utf8");
const wbsSvgRenderCode = readFileSync(path.resolve(__dirname, "../../src/js/wbs-svg-render.js"), "utf8");
const wbsSvgPublicCode = readFileSync(path.resolve(__dirname, "../../src/js/wbs-svg-public.js"), "utf8");
const nativeSvgCode = readFileSync(path.resolve(__dirname, "../../src/js/wbs-svg.js"), "utf8");
const mainIoCode = readFileSync(path.resolve(__dirname, "../../src/js/main-io.js"), "utf8");
const mainImportCode = readFileSync(path.resolve(__dirname, "../../src/js/main-import.js"), "utf8");
const mainExportCode = readFileSync(path.resolve(__dirname, "../../src/js/main-export.js"), "utf8");
const mainInputEventsCode = readFileSync(path.resolve(__dirname, "../../src/js/main-input-events.js"), "utf8");
const mainButtonEventsCode = readFileSync(path.resolve(__dirname, "../../src/js/main-button-events.js"), "utf8");
const mainEventsCode = readFileSync(path.resolve(__dirname, "../../src/js/main-events.js"), "utf8");
const mainPreviewCode = readFileSync(path.resolve(__dirname, "../../src/js/main-preview.js"), "utf8");
const mainUiCode = readFileSync(path.resolve(__dirname, "../../src/js/main-ui.js"), "utf8");
const mainSupportCode = readFileSync(path.resolve(__dirname, "../../src/js/main-support.js"), "utf8");
const mainSamplesCode = readFileSync(path.resolve(__dirname, "../../src/js/main-samples.js"), "utf8");
const mainDownloadsCode = readFileSync(path.resolve(__dirname, "../../src/js/main-downloads.js"), "utf8");
const mainOutputActionsCode = readFileSync(path.resolve(__dirname, "../../src/js/main-output-actions.js"), "utf8");
const mainImportActionsCode = readFileSync(path.resolve(__dirname, "../../src/js/main-import-actions.js"), "utf8");
const mainXmlActionsCode = readFileSync(path.resolve(__dirname, "../../src/js/main-xml-actions.js"), "utf8");
const mainArchiveActionsCode = readFileSync(path.resolve(__dirname, "../../src/js/main-archive-actions.js"), "utf8");
const mainSaveStateCode = readFileSync(path.resolve(__dirname, "../../src/js/main-save-state.js"), "utf8");
const mainTabActionsCode = readFileSync(path.resolve(__dirname, "../../src/js/main-tab-actions.js"), "utf8");
const mainPreviewActionsCode = readFileSync(path.resolve(__dirname, "../../src/js/main-preview-actions.js"), "utf8");
const mainTransformCode = readFileSync(path.resolve(__dirname, "../../src/js/main-transform.js"), "utf8");
const mainFlowCode = readFileSync(path.resolve(__dirname, "../../src/js/main-flow.js"), "utf8");
const mainModelCode = readFileSync(path.resolve(__dirname, "../../src/js/main-model.js"), "utf8");
const mainRenderCode = readFileSync(path.resolve(__dirname, "../../src/js/main-render.js"), "utf8");
const mainCode = readFileSync(path.resolve(__dirname, "../../src/js/main.js"), "utf8");
const bootPageCode = `${typesCode}\n${markdownEscapeCode}\n${aiJsonUtilCode}\n${mainUtilCode}\n${excelIoUtilCode}\n${excelIoZipCode}\n${excelIoNormalizeCode}\n${excelIoPackageXmlCode}\n${excelIoWorksheetBuildCode}\n${excelIoWorksheetParseCode}\n${excelIoWorkbookParseCode}\n${excelIoWorkbookBuildCode}\n${excelIoStylesBuildCode}\n${excelIoStylesParseCode}\n${excelIoCode}\n${msProjectAiViewsCode}\n${msProjectCalendarCode}\n${msProjectSamplesCode}\n${msProjectCsvCode}\n${msProjectValidateHelpersCode}\n${msProjectValidateCode}\n${msProjectXmlDomCode}\n${msProjectCodecCode}\n${msProjectMermaidCode}\n${msProjectXmlCode}\n${projectWorkbookSchemaCode}\n${projectXlsxImportUtilCode}\n${projectXlsxImportProjectCode}\n${projectXlsxImportCalendarsCode}\n${projectXlsxImportEntitiesCode}\n${projectXlsxImportCode}\n${projectXlsxExportUtilCode}\n${projectXlsxExportProjectCode}\n${projectXlsxExportEntitiesCode}\n${projectXlsxExportCalendarsCode}\n${projectXlsxExportCode}\n${projectXlsxCode}\n${projectWorkbookJsonValidateCode}\n${projectWorkbookJsonImportCode}\n${projectWorkbookJsonExportCode}\n${projectWorkbookJsonCode}\n${projectPatchJsonUtilCode}\n${projectPatchJsonLinksCode}\n${projectPatchJsonTasksCode}\n${projectPatchJsonEntitiesCode}\n${projectPatchJsonUpdatesCode}\n${projectPatchJsonCoreCode}\n${projectPatchJsonCode}\n${wbsDatebandCode}\n${wbsXlsxBaseCode}\n${wbsXlsxTaskmetaCode}\n${wbsXlsxLayoutCode}\n${wbsXlsxSectionsCode}\n${wbsXlsxCellsCode}\n${wbsXlsxExportCode}\n${wbsXlsxPublicCode}\n${wbsXlsxCode}\n${wbsMarkdownCode}\n${wbsSvgZipCode}\n${wbsSvgBarsCode}\n${wbsSvgViewportCode}\n${wbsSvgAxisCode}\n${wbsSvgScaffoldCode}\n${wbsSvgLabelsCode}\n${wbsSvgCalendarCode}\n${wbsSvgTimelineCode}\n${wbsSvgRenderCode}\n${wbsSvgPublicCode}\n${nativeSvgCode}\n${mainIoCode}\n${mainImportCode}\n${mainExportCode}\n${mainInputEventsCode}\n${mainButtonEventsCode}\n${mainEventsCode}\n${mainPreviewCode}\n${mainUiCode}\n${mainSupportCode}\n${mainSamplesCode}\n${mainDownloadsCode}\n${mainOutputActionsCode}\n${mainImportActionsCode}\n${mainXmlActionsCode}\n${mainArchiveActionsCode}\n${mainSaveStateCode}\n${mainTabActionsCode}\n${mainPreviewActionsCode}\n${mainTransformCode}\n${mainFlowCode}\n${mainModelCode}\n${mainRenderCode}\n${mainCode}`;
const bootPageRunner = new Function(bootPageCode);
const bootXmlModuleCode = `${typesCode}\n${msProjectAiViewsCode}\n${msProjectCalendarCode}\n${msProjectSamplesCode}\n${msProjectCsvCode}\n${msProjectValidateHelpersCode}\n${msProjectValidateCode}\n${msProjectXmlDomCode}\n${msProjectCodecCode}\n${msProjectMermaidCode}\n${msProjectXmlCode}\n${wbsMarkdownCode}`;
const bootXmlModuleRunner = new Function(bootXmlModuleCode);

export const hierarchyXml = readFileSync(path.resolve(__dirname, "../../testdata/hierarchy.xml"), "utf8");
export const SAMPLE_HOLIDAY_COUNT = 1;

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
      <button type="button" class="md-top-tab is-active" data-tab="input" role="tab" aria-selected="true" aria-controls="tabPanelInput"></button>
      <button type="button" class="md-top-tab" data-tab="transform" role="tab" aria-selected="false" aria-controls="tabPanelTransform"></button>
      <button type="button" class="md-top-tab" data-tab="output" role="tab" aria-selected="false" aria-controls="tabPanelOutput"></button>
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

export function bootPage() {
  mountDom();
  bootPageRunner();
  document.dispatchEvent(new Event("DOMContentLoaded"));
}

export function bootXmlModule() {
  bootXmlModuleRunner();
  return globalThis.__mikuprojectXml;
}

export function getMainHooks() {
  return globalThis.__mikuprojectMainTestHooks;
}

export function parseXmlViaHook() {
  getMainHooks().parseCurrentXml();
}

export async function exportMermaidViaHook() {
  await getMainHooks().exportCurrentMermaid();
}

export function getDefaultSampleHolidayDates() {
  return globalThis.__mikuprojectWbsXlsx.collectWbsHolidayDates(
    globalThis.__mikuprojectXml.importMsProjectXml(globalThis.__mikuprojectXml.SAMPLE_XML)
  );
}

export async function flushAsyncWork() {
  await Promise.resolve();
  await Promise.resolve();
}

export function setupMainPreviewExportDom() {
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
  const clipboard = { writeText: vi.fn(async () => {}) };
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
