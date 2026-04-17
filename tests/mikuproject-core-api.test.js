// @vitest-environment jsdom

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { beforeEach, describe, expect, it } from "vitest";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const typesCode = readFileSync(
  path.resolve(__dirname, "../src/js/types.js"),
  "utf8"
);
const aiJsonUtilCode = readFileSync(
  path.resolve(__dirname, "../src/js/ai-json-util.js"),
  "utf8"
);
const aiJsonSpecCode = readFileSync(
  path.resolve(__dirname, "../src/js/ai-json-spec.js"),
  "utf8"
);
const mainUtilCode = readFileSync(
  path.resolve(__dirname, "../src/js/main-util.js"),
  "utf8"
);
const excelIoUtilCode = readFileSync(
  path.resolve(__dirname, "../src/js/excel-io-util.js"),
  "utf8"
);
const msProjectAiViewsCode = readFileSync(
  path.resolve(__dirname, "../src/js/msproject-ai-views.js"),
  "utf8"
);
const msProjectCalendarCode = readFileSync(
  path.resolve(__dirname, "../src/js/msproject-calendar.js"),
  "utf8"
);
const msProjectSamplesCode = readFileSync(
  path.resolve(__dirname, "../src/js/msproject-samples.js"),
  "utf8"
);
const msProjectCsvCode = readFileSync(
  path.resolve(__dirname, "../src/js/msproject-csv.js"),
  "utf8"
);
const msProjectValidateHelpersCode = readFileSync(
  path.resolve(__dirname, "../src/js/msproject-validate-helpers.js"),
  "utf8"
);
const msProjectValidateCode = readFileSync(
  path.resolve(__dirname, "../src/js/msproject-validate.js"),
  "utf8"
);
const msProjectXmlDomCode = readFileSync(
  path.resolve(__dirname, "../src/js/msproject-xml-dom.js"),
  "utf8"
);
const msProjectCodecCode = readFileSync(
  path.resolve(__dirname, "../src/js/msproject-codec.js"),
  "utf8"
);
const msProjectMermaidCode = readFileSync(
  path.resolve(__dirname, "../src/js/msproject-mermaid.js"),
  "utf8"
);
const msProjectXmlCode = readFileSync(
  path.resolve(__dirname, "../src/js/msproject-xml.js"),
  "utf8"
);
const markdownEscapeCode = readFileSync(
  path.resolve(__dirname, "../src/js/markdown-escape.js"),
  "utf8"
);
const projectWorkbookSchemaCode = readFileSync(
  path.resolve(__dirname, "../src/js/project-workbook-schema.js"),
  "utf8"
);
const projectXlsxImportUtilCode = readFileSync(
  path.resolve(__dirname, "../src/js/project-xlsx-import-util.js"),
  "utf8"
);
const projectXlsxImportProjectCode = readFileSync(
  path.resolve(__dirname, "../src/js/project-xlsx-import-project.js"),
  "utf8"
);
const projectXlsxImportCalendarsCode = readFileSync(
  path.resolve(__dirname, "../src/js/project-xlsx-import-calendars.js"),
  "utf8"
);
const projectXlsxImportEntitiesCode = readFileSync(
  path.resolve(__dirname, "../src/js/project-xlsx-import-entities.js"),
  "utf8"
);
const projectXlsxImportCode = readFileSync(
  path.resolve(__dirname, "../src/js/project-xlsx-import.js"),
  "utf8"
);
const projectXlsxExportUtilCode = readFileSync(
  path.resolve(__dirname, "../src/js/project-xlsx-export-util.js"),
  "utf8"
);
const projectXlsxExportProjectCode = readFileSync(
  path.resolve(__dirname, "../src/js/project-xlsx-export-project.js"),
  "utf8"
);
const projectXlsxExportEntitiesCode = readFileSync(
  path.resolve(__dirname, "../src/js/project-xlsx-export-entities.js"),
  "utf8"
);
const projectXlsxExportCalendarsCode = readFileSync(
  path.resolve(__dirname, "../src/js/project-xlsx-export-calendars.js"),
  "utf8"
);
const projectXlsxExportCode = readFileSync(
  path.resolve(__dirname, "../src/js/project-xlsx-export.js"),
  "utf8"
);
const projectXlsxCode = readFileSync(
  path.resolve(__dirname, "../src/js/project-xlsx.js"),
  "utf8"
);
const projectWorkbookJsonValidateCode = readFileSync(
  path.resolve(__dirname, "../src/js/project-workbook-json-validate.js"),
  "utf8"
);
const projectWorkbookJsonImportCode = readFileSync(
  path.resolve(__dirname, "../src/js/project-workbook-json-import.js"),
  "utf8"
);
const projectWorkbookJsonExportCode = readFileSync(
  path.resolve(__dirname, "../src/js/project-workbook-json-export.js"),
  "utf8"
);
const excelIoCode = readFileSync(
  path.resolve(__dirname, "../src/js/excel-io.js"),
  "utf8"
);
const excelIoZipCode = readFileSync(
  path.resolve(__dirname, "../src/js/excel-io-zip.js"),
  "utf8"
);
const excelIoNormalizeCode = readFileSync(
  path.resolve(__dirname, "../src/js/excel-io-normalize.js"),
  "utf8"
);
const excelIoPackageXmlCode = readFileSync(
  path.resolve(__dirname, "../src/js/excel-io-package-xml.js"),
  "utf8"
);
const excelIoWorksheetBuildCode = readFileSync(
  path.resolve(__dirname, "../src/js/excel-io-worksheet-build.js"),
  "utf8"
);
const excelIoWorksheetParseCode = readFileSync(
  path.resolve(__dirname, "../src/js/excel-io-worksheet-parse.js"),
  "utf8"
);
const excelIoWorkbookParseCode = readFileSync(
  path.resolve(__dirname, "../src/js/excel-io-workbook-parse.js"),
  "utf8"
);
const excelIoWorkbookBuildCode = readFileSync(
  path.resolve(__dirname, "../src/js/excel-io-workbook-build.js"),
  "utf8"
);
const excelIoStylesBuildCode = readFileSync(
  path.resolve(__dirname, "../src/js/excel-io-styles-build.js"),
  "utf8"
);
const excelIoStylesParseCode = readFileSync(
  path.resolve(__dirname, "../src/js/excel-io-styles-parse.js"),
  "utf8"
);
const projectWorkbookJsonCode = readFileSync(
  path.resolve(__dirname, "../src/js/project-workbook-json.js"),
  "utf8"
);
const projectPatchJsonUtilCode = readFileSync(
  path.resolve(__dirname, "../src/js/project-patch-json-util.js"),
  "utf8"
);
const projectPatchJsonLinksCode = readFileSync(
  path.resolve(__dirname, "../src/js/project-patch-json-links.js"),
  "utf8"
);
const projectPatchJsonTasksCode = readFileSync(
  path.resolve(__dirname, "../src/js/project-patch-json-tasks.js"),
  "utf8"
);
const projectPatchJsonEntitiesCode = readFileSync(
  path.resolve(__dirname, "../src/js/project-patch-json-entities.js"),
  "utf8"
);
const projectPatchJsonUpdatesCode = readFileSync(
  path.resolve(__dirname, "../src/js/project-patch-json-updates.js"),
  "utf8"
);
const projectPatchJsonCoreCode = readFileSync(
  path.resolve(__dirname, "../src/js/project-patch-json-core.js"),
  "utf8"
);
const projectPatchJsonCode = readFileSync(
  path.resolve(__dirname, "../src/js/project-patch-json.js"),
  "utf8"
);
const coreApiMsprojectAiCode = readFileSync(
  path.resolve(__dirname, "../src/js/core-api-msproject-ai.js"),
  "utf8"
);
const coreApiMsprojectCode = readFileSync(
  path.resolve(__dirname, "../src/js/core-api-msproject.js"),
  "utf8"
);
const coreApiWorkbookXlsxCode = readFileSync(
  path.resolve(__dirname, "../src/js/core-api-workbook-xlsx.js"),
  "utf8"
);
const coreApiWorkbookCode = readFileSync(
  path.resolve(__dirname, "../src/js/core-api-workbook.js"),
  "utf8"
);
const coreApiAiJsonImportCode = readFileSync(
  path.resolve(__dirname, "../src/js/core-api-ai-json-import.js"),
  "utf8"
);
const coreApiAiJsonCode = readFileSync(
  path.resolve(__dirname, "../src/js/core-api-ai-json.js"),
  "utf8"
);
const coreApiExternalBinaryCode = readFileSync(
  path.resolve(__dirname, "../src/js/core-api-external-binary.js"),
  "utf8"
);
const coreApiExternalDocumentCode = readFileSync(
  path.resolve(__dirname, "../src/js/core-api-external-document.js"),
  "utf8"
);
const coreApiExternalImportCode = readFileSync(
  path.resolve(__dirname, "../src/js/core-api-external-import.js"),
  "utf8"
);
const coreApiImportCode = readFileSync(
  path.resolve(__dirname, "../src/js/core-api-import.js"),
  "utf8"
);
const coreApiReportCode = readFileSync(
  path.resolve(__dirname, "../src/js/core-api-report.js"),
  "utf8"
);
const coreApiReportAdaptersCode = readFileSync(
  path.resolve(__dirname, "../src/js/core-api-report-adapters.js"),
  "utf8"
);
const coreApiReportPublicCode = readFileSync(
  path.resolve(__dirname, "../src/js/core-api-report-public.js"),
  "utf8"
);
const coreApiRegistryCode = readFileSync(
  path.resolve(__dirname, "../src/js/core-api-registry.js"),
  "utf8"
);
const coreApiPublicCode = readFileSync(
  path.resolve(__dirname, "../src/js/core-api-public.js"),
  "utf8"
);
const wbsDatebandCode = readFileSync(
  path.resolve(__dirname, "../src/js/wbs-dateband.js"),
  "utf8"
);
const wbsXlsxBaseCode = readFileSync(
  path.resolve(__dirname, "../src/js/wbs-xlsx-base.js"),
  "utf8"
);
const wbsXlsxTaskmetaCode = readFileSync(
  path.resolve(__dirname, "../src/js/wbs-xlsx-taskmeta.js"),
  "utf8"
);
const wbsXlsxLayoutCode = readFileSync(
  path.resolve(__dirname, "../src/js/wbs-xlsx-layout.js"),
  "utf8"
);
const wbsXlsxSectionsCode = readFileSync(
  path.resolve(__dirname, "../src/js/wbs-xlsx-sections.js"),
  "utf8"
);
const wbsXlsxCellsCode = readFileSync(
  path.resolve(__dirname, "../src/js/wbs-xlsx-cells.js"),
  "utf8"
);
const wbsXlsxExportCode = readFileSync(
  path.resolve(__dirname, "../src/js/wbs-xlsx-export.js"),
  "utf8"
);
const wbsXlsxPublicCode = readFileSync(
  path.resolve(__dirname, "../src/js/wbs-xlsx-public.js"),
  "utf8"
);
const wbsXlsxCode = readFileSync(
  path.resolve(__dirname, "../src/js/wbs-xlsx.js"),
  "utf8"
);
const wbsSvgZipCode = readFileSync(
  path.resolve(__dirname, "../src/js/wbs-svg-zip.js"),
  "utf8"
);
const wbsSvgBarsCode = readFileSync(
  path.resolve(__dirname, "../src/js/wbs-svg-bars.js"),
  "utf8"
);
const wbsSvgViewportCode = readFileSync(
  path.resolve(__dirname, "../src/js/wbs-svg-viewport.js"),
  "utf8"
);
const wbsSvgAxisCode = readFileSync(
  path.resolve(__dirname, "../src/js/wbs-svg-axis.js"),
  "utf8"
);
const wbsSvgScaffoldCode = readFileSync(
  path.resolve(__dirname, "../src/js/wbs-svg-scaffold.js"),
  "utf8"
);
const wbsSvgLabelsCode = readFileSync(
  path.resolve(__dirname, "../src/js/wbs-svg-labels.js"),
  "utf8"
);
const wbsSvgCalendarCode = readFileSync(
  path.resolve(__dirname, "../src/js/wbs-svg-calendar.js"),
  "utf8"
);
const wbsSvgTimelineCode = readFileSync(
  path.resolve(__dirname, "../src/js/wbs-svg-timeline.js"),
  "utf8"
);
const wbsSvgRenderCode = readFileSync(
  path.resolve(__dirname, "../src/js/wbs-svg-render.js"),
  "utf8"
);
const wbsSvgPublicCode = readFileSync(
  path.resolve(__dirname, "../src/js/wbs-svg-public.js"),
  "utf8"
);
const wbsSvgCode = readFileSync(
  path.resolve(__dirname, "../src/js/wbs-svg.js"),
  "utf8"
);
const wbsMarkdownCode = readFileSync(
  path.resolve(__dirname, "../src/js/wbs-markdown.js"),
  "utf8"
);
const coreApiCode = readFileSync(
  path.resolve(__dirname, "../src/js/core-api.js"),
  "utf8"
);
const dependencyXml = readFileSync(
  path.resolve(__dirname, "../testdata/dependency.xml"),
  "utf8"
);

function bootModules() {
  new Function([
    typesCode,
    aiJsonUtilCode,
    aiJsonSpecCode,
    mainUtilCode,
    excelIoUtilCode,
    msProjectAiViewsCode,
    msProjectCalendarCode,
    msProjectSamplesCode,
    msProjectCsvCode,
    msProjectValidateHelpersCode,
    msProjectValidateCode,
    msProjectXmlDomCode,
    msProjectCodecCode,
    msProjectMermaidCode,
    msProjectXmlCode,
    markdownEscapeCode,
    projectWorkbookSchemaCode,
    excelIoZipCode,
    excelIoNormalizeCode,
    excelIoPackageXmlCode,
    excelIoWorksheetBuildCode,
    excelIoWorksheetParseCode,
    excelIoWorkbookParseCode,
    excelIoWorkbookBuildCode,
    excelIoStylesBuildCode,
    excelIoStylesParseCode,
    excelIoCode,
    projectXlsxImportUtilCode,
    projectXlsxImportProjectCode,
    projectXlsxImportCalendarsCode,
    projectXlsxImportEntitiesCode,
    projectXlsxImportCode,
    projectXlsxExportUtilCode,
    projectXlsxExportProjectCode,
    projectXlsxExportEntitiesCode,
    projectXlsxExportCalendarsCode,
    projectXlsxExportCode,
    projectXlsxCode,
    projectWorkbookJsonValidateCode,
    projectWorkbookJsonImportCode,
    projectWorkbookJsonExportCode,
    projectWorkbookJsonCode,
    projectPatchJsonUtilCode,
    projectPatchJsonLinksCode,
    projectPatchJsonTasksCode,
    projectPatchJsonEntitiesCode,
    projectPatchJsonUpdatesCode,
    projectPatchJsonCoreCode,
    projectPatchJsonCode,
    coreApiMsprojectAiCode,
    coreApiMsprojectCode,
    coreApiWorkbookXlsxCode,
    coreApiWorkbookCode,
    coreApiAiJsonImportCode,
    coreApiAiJsonCode,
    coreApiExternalBinaryCode,
    coreApiExternalDocumentCode,
    coreApiExternalImportCode,
    coreApiImportCode,
    wbsDatebandCode,
    wbsXlsxBaseCode,
    wbsXlsxTaskmetaCode,
    wbsXlsxLayoutCode,
    wbsXlsxSectionsCode,
    wbsXlsxCellsCode,
    wbsXlsxExportCode,
    wbsXlsxPublicCode,
    wbsXlsxCode,
    wbsSvgZipCode,
    wbsSvgBarsCode,
    wbsSvgViewportCode,
    wbsSvgAxisCode,
    wbsSvgScaffoldCode,
    wbsSvgLabelsCode,
    wbsSvgCalendarCode,
    wbsSvgTimelineCode,
    wbsSvgRenderCode,
    wbsSvgPublicCode,
    wbsSvgCode,
    wbsMarkdownCode,
    coreApiReportCode,
    coreApiReportAdaptersCode,
    coreApiReportPublicCode,
    coreApiRegistryCode,
    coreApiPublicCode,
    coreApiCode
  ].join("\n"))();
  return globalThis.__mikuprojectCoreApi;
}

describe("mikuproject core api", () => {
  beforeEach(() => {
    delete globalThis.__mikuprojectAiJsonUtil;
    delete globalThis.__mikuprojectAiJsonSpec;
    delete globalThis.__mikuprojectMainUtil;
    delete globalThis.__mikuprojectXml;
    delete globalThis.__mikuprojectMarkdownEscape;
    delete globalThis.__mikuprojectProjectWorkbookSchema;
    delete globalThis.__mikuprojectExcelIo;
    delete globalThis.__mikuprojectProjectXlsxImportUtil;
    delete globalThis.__mikuprojectProjectXlsxImportProject;
    delete globalThis.__mikuprojectProjectXlsxImportCalendars;
    delete globalThis.__mikuprojectProjectXlsxImportEntities;
    delete globalThis.__mikuprojectProjectXlsxImport;
    delete globalThis.__mikuprojectProjectXlsxExportUtil;
    delete globalThis.__mikuprojectProjectXlsxExportProject;
    delete globalThis.__mikuprojectProjectXlsxExportEntities;
    delete globalThis.__mikuprojectProjectXlsxExportCalendars;
    delete globalThis.__mikuprojectProjectXlsxExport;
    delete globalThis.__mikuprojectProjectXlsx;
    delete globalThis.__mikuprojectProjectWorkbookJsonValidate;
    delete globalThis.__mikuprojectProjectWorkbookJsonImport;
    delete globalThis.__mikuprojectProjectWorkbookJsonExport;
    delete globalThis.__mikuprojectProjectWorkbookJson;
    delete globalThis.__mikuprojectProjectPatchJson;
    delete globalThis.__mikuprojectCoreApiMsprojectAi;
    delete globalThis.__mikuprojectCoreApiMsproject;
    delete globalThis.__mikuprojectCoreApiWorkbookXlsx;
    delete globalThis.__mikuprojectCoreApiWorkbook;
    delete globalThis.__mikuprojectCoreApiAiJsonImport;
    delete globalThis.__mikuprojectCoreApiAiJson;
    delete globalThis.__mikuprojectCoreApiExternalBinary;
    delete globalThis.__mikuprojectCoreApiExternalDocument;
    delete globalThis.__mikuprojectCoreApiExternalImport;
    delete globalThis.__mikuprojectCoreApiImport;
    delete globalThis.__mikuprojectCoreApiReport;
    delete globalThis.__mikuprojectCoreApiReportAdapters;
    delete globalThis.__mikuprojectCoreApiReportPublic;
    delete globalThis.__mikuprojectCoreApiRegistry;
    delete globalThis.__mikuprojectCoreApiPublic;
    delete globalThis.__mikuprojectWbsXlsx;
    delete globalThis.__mikuprojectNativeSvg;
    delete globalThis.__mikuprojectWbsMarkdown;
    delete globalThis.__mikuprojectCoreApi;
  });

  it("exposes the ai json spec through a stable function", () => {
    const api = bootModules();

    const spec = api.getAiJsonSpec();

    expect(spec.id).toBe("mikuproject-ai-json-spec");
    expect(spec.version).toBe("v20260410");
    expect(spec.text).toContain("# mikuproject AI JSON Prompt / Spec");
    expect(api.getAiJsonSpecText()).toBe(spec.text);
  });

  it("exposes stable sample getters", () => {
    const api = bootModules();

    const sampleXml = api.samples.getSampleXml();
    const sampleDraft = api.samples.getSampleProjectDraftView();

    expect(sampleXml).toContain("<Project");
    expect(sampleXml).toContain("<Name>mikuproject開発</Name>");
    expect(sampleDraft.view_type).toBe("project_draft_view");
    expect(sampleDraft.project.name).toBe("mikuproject開発");
    expect(Array.isArray(sampleDraft.tasks)).toBe(true);
    expect(sampleDraft.tasks.length).toBeGreaterThan(0);
  });

  it("parses fenced ai json text and detects the kind", () => {
    const api = bootModules();

    const parsed = api.parseAiJsonText([
      "説明文",
      "```json",
      JSON.stringify({ view_type: "project_draft_view", project: { name: "Test" }, tasks: [] }),
      "```"
    ].join("\n"));

    expect(parsed.kind).toBe("project_draft_view");
    expect(parsed.document.view_type).toBe("project_draft_view");
  });

  it("imports project_draft_view without UI dependencies", () => {
    const api = bootModules();

    const result = api.importAiJsonDocument({
      view_type: "project_draft_view",
      project: {
        name: "API draft import",
        planned_start: "2026-04-01"
      },
      tasks: [
        { uid: "draft-1", name: "開始", parent_uid: null, position: 0, is_milestone: true, planned_start: "2026-04-01", planned_finish: "2026-04-01" }
      ],
      resources: [],
      assignments: []
    });

    expect(result.kind).toBe("project_draft_view");
    expect(result.mode).toBe("replace");
    expect(result.model.project.name).toBe("API draft import");
    expect(result.model.tasks).toHaveLength(1);
  });

  it("exposes ai view helpers through the unified entrypoint", () => {
    const api = bootModules();
    const hierarchyModel = api.msProject.importFromXml(readFileSync(
      path.resolve(__dirname, "../testdata/hierarchy.xml"),
      "utf8"
    ));

    const request = api.aiViews.buildProjectDraftRequest({
      name: "Core API draft request",
      plannedStart: "2026-04-01",
      goal: "Stabilize wrappers",
      teamCount: 2,
      mustHavePhases: ["Plan"],
      mustHaveMilestones: ["Release"]
    });
    const overview = api.aiViews.exportProjectOverviewView(hierarchyModel);
    const phaseDetail = api.aiViews.exportPhaseDetailView(hierarchyModel, "1", { mode: "scoped", rootUid: "2", maxDepth: 1 });
    const taskEdit = api.aiViews.exportTaskEditView(hierarchyModel, "3");

    expect(request.view_type).toBe("project_draft_request");
    expect(request.project.name).toBe("Core API draft request");
    expect(request.requirements.team_count).toBe(2);
    expect(overview.view_type).toBe("project_overview_view");
    expect(overview.phases[0].uid).toBe("1");
    expect(phaseDetail.view_type).toBe("phase_detail_view");
    expect(phaseDetail.scope).toEqual({ mode: "scoped", root_uid: "2", max_depth: 1 });
    expect(phaseDetail.tasks.map((task) => task.uid)).toEqual(["2"]);
    expect(taskEdit.view_type).toBe("task_edit_view");
    expect(taskEdit.target_task.uid).toBe("3");
    expect(Array.isArray(taskEdit.predecessors)).toBe(true);
  });

  it("imports workbook json with and without a base model", () => {
    const api = bootModules();
    const baseModel = api.msProject.importFromXml(dependencyXml);
    const documentLike = api.workbookJson.exportDocument(baseModel);
    documentLike.sheets.Project.find((row) => row.Field === "Name").Value = "Core API workbook import";

    const replaceResult = api.importAiJsonDocument(documentLike);
    const mergeResult = api.importAiJsonDocument(documentLike, { baseModel });

    expect(replaceResult.kind).toBe("workbook_json");
    expect(replaceResult.mode).toBe("replace");
    expect(replaceResult.model.project.name).toBe("Core API workbook import");
    expect(mergeResult.kind).toBe("workbook_json");
    expect(mergeResult.mode).toBe("merge");
    expect(mergeResult.model.project.name).toBe("Core API workbook import");
    expect(Array.isArray(mergeResult.changes)).toBe(true);
  });

  it("imports ai json text and validates workbook json through the unified entrypoint", () => {
    const api = bootModules();
    const baseModel = api.msProject.importFromXml(dependencyXml);
    const documentLike = api.workbookJson.exportDocument(baseModel);
    documentLike.sheets.Project.find((row) => row.Field === "Name").Value = "Core API text import";
    documentLike.sheets.UnknownSheet = [];

    const parsedImport = api.importAiJsonText([
      "説明",
      "```json",
      JSON.stringify(documentLike, null, 2),
      "```"
    ].join("\n"), { baseModel });
    const validation = api.workbookJson.validateDocument(documentLike);

    expect(parsedImport.kind).toBe("workbook_json");
    expect(parsedImport.result.kind).toBe("workbook_json");
    expect(parsedImport.result.mode).toBe("merge");
    expect(parsedImport.result.model.project.name).toBe("Core API text import");
    expect(validation.document.format).toBe("mikuproject_workbook_json");
    expect(validation.warnings.map((item) => item.message)).toContain("未知の sheet は無視します: UnknownSheet");
  });

  it("exposes xlsx encode/decode and workbook import/export through the unified entrypoint", () => {
    const api = bootModules();
    const baseModel = api.msProject.importFromXml(dependencyXml);
    const workbook = api.xlsx.exportWorkbook(baseModel);
    const projectNameRow = workbook.sheets
      .find((sheet) => sheet.name === "Project")
      .rows.find((row) => row.cells[0]?.value === "Name");

    projectNameRow.cells[1].value = "Core API xlsx import";

    const bytes = api.xlsx.encodeWorkbook(workbook);
    const decodedWorkbook = api.xlsx.decodeWorkbook(bytes);
    const replaceModel = api.xlsx.importAsProjectModel(decodedWorkbook);
    const mergedModel = api.xlsx.importIntoProjectModel(decodedWorkbook, baseModel);

    expect(bytes).toBeInstanceOf(Uint8Array);
    expect(decodedWorkbook.sheets.some((sheet) => sheet.name === "Tasks")).toBe(true);
    expect(replaceModel.project.name).toBe("Core API xlsx import");
    expect(mergedModel.project.name).toBe("Core API xlsx import");
  });

  it("imports external formats through importExternal", () => {
    const api = bootModules();
    const baseModel = api.msProject.importFromXml(dependencyXml);
    const workbook = api.xlsx.exportWorkbook(baseModel);
    workbook.sheets.find((sheet) => sheet.name === "Project").rows.find((row) => row.cells[0]?.value === "Name").cells[1].value = "Core API external xlsx";
    const xlsxBytes = api.xlsx.encodeWorkbook(workbook);
    const workbookJson = api.workbookJson.exportDocument(baseModel);
    workbookJson.sheets.Project.find((row) => row.Field === "Name").Value = "Core API external workbook json";

    const xmlResult = api.importExternal({
      source: { format: "ms_project_xml", text: dependencyXml },
      mode: "replace"
    });
    const xlsxReplaceResult = api.importExternal({
      source: { format: "xlsx", bytes: xlsxBytes },
      mode: "replace"
    });
    const xlsxMergeResult = api.importExternal({
      source: { format: "xlsx", bytes: xlsxBytes },
      mode: "merge",
      baseModel
    });
    const workbookJsonMergeResult = api.importExternal({
      source: { format: "workbook_json", document: workbookJson },
      mode: "merge",
      baseModel
    });
    const patchResult = api.importExternal({
      source: {
        format: "patch_json",
        document: {
          operations: [
            {
              op: "update_project",
              fields: {
                name: "Core API external patch"
              }
            }
          ]
        }
      },
      mode: "patch",
      baseModel
    });

    expect(xmlResult.kind).toBe("ms_project_xml");
    expect(xmlResult.mode).toBe("replace");
    expect(xlsxReplaceResult.kind).toBe("xlsx");
    expect(xlsxReplaceResult.model.project.name).toBe("Core API external xlsx");
    expect(xlsxMergeResult.kind).toBe("xlsx");
    expect(xlsxMergeResult.mode).toBe("merge");
    expect(Array.isArray(xlsxMergeResult.changes)).toBe(true);
    expect(workbookJsonMergeResult.kind).toBe("workbook_json");
    expect(workbookJsonMergeResult.mode).toBe("merge");
    expect(workbookJsonMergeResult.model.project.name).toBe("Core API external workbook json");
    expect(patchResult.kind).toBe("patch_json");
    expect(patchResult.model.project.name).toBe("Core API external patch");
  });

  it("applies patch json through the unified entrypoint", () => {
    const api = bootModules();
    const baseModel = api.msProject.importFromXml(dependencyXml);

    const result = api.importAiJsonDocument({
      operations: [
        {
          op: "update_project",
          fields: {
            name: "Core API patch import"
          }
        }
      ]
    }, { baseModel });

    expect(result.kind).toBe("patch_json");
    expect(result.mode).toBe("patch");
    expect(result.model.project.name).toBe("Core API patch import");
    expect(result.changes.some((change) => String(change.field).toLowerCase() === "name")).toBe(true);
  });

  it("exposes report exports through the unified entrypoint", () => {
    const api = bootModules();
    const model = api.msProject.importFromXml(dependencyXml);

    const wbsWorkbook = api.report.wbsXlsx.exportWorkbook(model, {
      displayDaysBeforeBaseDate: 1,
      displayDaysAfterBaseDate: 2
    });
    const wbsBytes = api.report.wbsXlsx.exportBytes(model);
    const dailySvg = api.report.svg.exportDaily(model, { labelMode: "list" });
    const weeklySvg = api.report.svg.exportWeekly(model);
    const monthlyCalendar = api.report.svg.exportMonthlyCalendar(model);
    const reportBundle = api.report.all.export(model);
    const wbsMarkdown = api.report.wbsMarkdown.export(model);
    const mermaidText = api.report.mermaid.exportGantt(model);

    expect(wbsWorkbook.sheets.length).toBeGreaterThan(0);
    expect(wbsBytes).toBeInstanceOf(Uint8Array);
    expect(dailySvg).toContain("<svg");
    expect(weeklySvg).toContain("<svg");
    expect(monthlyCalendar.entries.length).toBeGreaterThan(0);
    expect(monthlyCalendar.zipBytes).toBeInstanceOf(Uint8Array);
    expect(reportBundle.zipBytes).toBeInstanceOf(Uint8Array);
    expect(reportBundle.entries.map((entry) => entry.name)).toContain("wbs.xlsx");
    expect(reportBundle.entries.map((entry) => entry.name)).toContain("wbs.md");
    expect(reportBundle.entries.map((entry) => entry.name)).toContain("mermaid.mmd");
    expect(reportBundle.entries.map((entry) => entry.name)).toContain("daily.svg");
    expect(reportBundle.entries.map((entry) => entry.name)).toContain("weekly.svg");
    expect(reportBundle.entries.some((entry) => entry.name.startsWith("monthly-calendar/"))).toBe(true);
    expect(reportBundle.entries.some((entry) => entry.name === "monthly-calendar.zip")).toBe(false);
    expect(wbsMarkdown).toContain("# WBS テーブル");
    expect(mermaidText).toContain("gantt");
  });

  it("rejects patch json when baseModel is missing", () => {
    const api = bootModules();

    expect(() => api.importAiJsonDocument({ operations: [] })).toThrow("baseModel");
  });

  it("rejects unsupported format and mode combinations in importExternal", () => {
    const api = bootModules();
    const baseModel = api.msProject.importFromXml(dependencyXml);

    expect(() => api.importExternal({
      source: { format: "ms_project_xml", text: dependencyXml },
      mode: "merge",
      baseModel
    })).toThrow("replace");
    expect(() => api.importExternal({
      source: { format: "patch_json", document: { operations: [] } },
      mode: "replace",
      baseModel
    })).toThrow("patch");
    expect(() => api.importExternal({
      source: { format: "xlsx", bytes: new Uint8Array() },
      mode: "patch",
      baseModel
    })).toThrow("replace または merge");
    expect(() => api.importExternal({
      source: { format: "workbook_json", document: {} },
      mode: "patch",
      baseModel
    })).toThrow("patch import");
  });

  it("rejects merge imports when baseModel is missing", () => {
    const api = bootModules();
    const baseModel = api.msProject.importFromXml(dependencyXml);
    const workbook = api.xlsx.exportWorkbook(baseModel);
    const xlsxBytes = api.xlsx.encodeWorkbook(workbook);
    const workbookJson = api.workbookJson.exportDocument(baseModel);

    expect(() => api.importExternal({
      source: { format: "xlsx", bytes: xlsxBytes },
      mode: "merge"
    })).toThrow("baseModel");
    expect(() => api.importExternal({
      source: { format: "workbook_json", document: workbookJson },
      mode: "merge"
    })).toThrow("baseModel");
  });

  it("validates patch documents through the unified entrypoint", () => {
    const api = bootModules();

    const valid = api.patchJson.validateDocument({
      operations: [
        { op: "update_project", fields: { name: "Validated Patch" } }
      ]
    });

    expect(valid.document.operations).toHaveLength(1);
    expect(valid.warnings).toEqual([]);
    expect(() => api.patchJson.validateDocument(null)).toThrow("オブジェクト");
    expect(() => api.patchJson.validateDocument({})).toThrow("operations");
    expect(() => api.patchJson.validateDocument({ operations: [null] })).toThrow("operations[0]");
  });
});
