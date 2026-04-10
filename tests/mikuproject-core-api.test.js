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
const projectXlsxCode = readFileSync(
  path.resolve(__dirname, "../src/js/project-xlsx.js"),
  "utf8"
);
const excelIoCode = readFileSync(
  path.resolve(__dirname, "../src/js/excel-io.js"),
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
    msProjectXmlCode,
    markdownEscapeCode,
    projectWorkbookSchemaCode,
    excelIoCode,
    projectXlsxCode,
    projectWorkbookJsonCode,
    projectPatchJsonCode,
    wbsXlsxCode,
    wbsSvgCode,
    wbsMarkdownCode,
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
    delete globalThis.__mikuprojectProjectXlsx;
    delete globalThis.__mikuprojectProjectWorkbookJson;
    delete globalThis.__mikuprojectProjectPatchJson;
    delete globalThis.__mikuprojectWbsXlsx;
    delete globalThis.__mikuprojectNativeSvg;
    delete globalThis.__mikuprojectWbsMarkdown;
    delete globalThis.__mikuprojectCoreApi;
  });

  it("exposes the ai json spec through a stable function", () => {
    const api = bootModules();

    const spec = api.getAiJsonSpec();

    expect(spec.id).toBe("mikuproject-ai-json-spec");
    expect(spec.version).toBe("v20260403");
    expect(spec.text).toContain("# mikuproject AI JSON Prompt / Spec");
    expect(api.getAiJsonSpecText()).toBe(spec.text);
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
});
