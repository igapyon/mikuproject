// @vitest-environment jsdom

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const typesCode = readFileSync(
  path.resolve(__dirname, "../src/js/types.js"),
  "utf8"
);
const excelIoUtilCode = readFileSync(
  path.resolve(__dirname, "../src/js/excel-io-util.js"),
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
const projectWorkbookJsonCode = readFileSync(
  path.resolve(__dirname, "../src/js/project-workbook-json.js"),
  "utf8"
);
const dependencyXml = readFileSync(
  path.resolve(__dirname, "../testdata/dependency.xml"),
  "utf8"
);

function bootModules() {
  new Function(`${typesCode}\n${excelIoUtilCode}\n${excelIoZipCode}\n${excelIoNormalizeCode}\n${excelIoPackageXmlCode}\n${excelIoWorksheetBuildCode}\n${excelIoWorksheetParseCode}\n${excelIoWorkbookParseCode}\n${excelIoWorkbookBuildCode}\n${excelIoStylesBuildCode}\n${excelIoStylesParseCode}\n${excelIoCode}\n${msProjectAiViewsCode}\n${msProjectCalendarCode}\n${msProjectSamplesCode}\n${msProjectCsvCode}\n${msProjectValidateHelpersCode}\n${msProjectValidateCode}\n${msProjectXmlDomCode}\n${msProjectCodecCode}\n${msProjectMermaidCode}\n${msProjectXmlCode}\n${projectWorkbookSchemaCode}\n${projectXlsxImportUtilCode}\n${projectXlsxImportProjectCode}\n${projectXlsxImportCalendarsCode}\n${projectXlsxImportEntitiesCode}\n${projectXlsxImportCode}\n${projectXlsxExportUtilCode}\n${projectXlsxExportProjectCode}\n${projectXlsxExportEntitiesCode}\n${projectXlsxExportCalendarsCode}\n${projectXlsxExportCode}\n${projectXlsxCode}\n${projectWorkbookJsonValidateCode}\n${projectWorkbookJsonImportCode}\n${projectWorkbookJsonExportCode}\n${projectWorkbookJsonCode}`)();
  return {
    xml: globalThis.__mikuprojectXml,
    projectWorkbookJson: globalThis.__mikuprojectProjectWorkbookJson
  };
}

describe("mikuproject project workbook json", () => {
  it("exports workbook json with fixed format and sheets", () => {
    const { xml, projectWorkbookJson } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);

    const documentLike = projectWorkbookJson.exportProjectWorkbookJson(model);

    expect(documentLike.format).toBe("mikuproject_workbook_json");
    expect(documentLike.version).toBe(1);
    expect(Object.keys(documentLike.sheets)).toEqual([
      "Project",
      "Tasks",
      "Resources",
      "Assignments",
      "Calendars",
      "NonWorkingDays"
    ]);
    expect(documentLike.sheets.Project[0]).toEqual({ Field: "Name", Value: "mikuproject開発" });
    expect(documentLike.sheets.Tasks[0].UID).toBe("1");
    expect(documentLike.sheets.Tasks[0].Name).toBe("基盤整備");
    expect(documentLike.sheets.Resources[0].Name).toBe("Mikuku");
    expect(documentLike.sheets.Assignments[0].ResourceName).toBe("Mikuku");
  });

  it("imports limited editable fields through workbook json", () => {
    const { xml, projectWorkbookJson } = bootModules();
    const baseModel = xml.importMsProjectXml(dependencyXml);
    const documentLike = projectWorkbookJson.exportProjectWorkbookJson(baseModel);
    const executeTaskRow = documentLike.sheets.Tasks.find((row) => row.UID === "2");
    const resourceRow = documentLike.sheets.Resources.find((row) => row.UID === "1");

    documentLike.sheets.Project.find((row) => row.Field === "Name").Value = "JSON import project";
    executeTaskRow.Name = "JSON import task";
    executeTaskRow.Start = "2026-03-16 10:00:00";
    executeTaskRow.Duration = "PT24H0M0S";
    executeTaskRow.Milestone = "○";
    executeTaskRow.Summary = "○";
    executeTaskRow.Critical = "ー";
    executeTaskRow.CalendarUID = "2";
    executeTaskRow.Predecessors = "2";
    executeTaskRow.OutlineNumber = "999";
    resourceRow.Name = "Miku Updated";
    resourceRow.Group = "Dev";
    resourceRow.MaxUnits = 1;
    resourceRow.CalendarUID = "1";

    const result = projectWorkbookJson.importProjectWorkbookJson(documentLike, baseModel);

    expect(result.model.project.name).toBe("JSON import project");
    expect(result.model.tasks[1].name).toBe("JSON import task");
    expect(result.model.tasks[1].start).toBe("2026-03-16T10:00:00");
    expect(result.model.tasks[1].duration).toBe("PT24H0M0S");
    expect(result.model.tasks[1].milestone).toBe(true);
    expect(result.model.tasks[1].summary).toBe(true);
    expect(result.model.tasks[1].critical).toBe(false);
    expect(result.model.tasks[1].calendarUID).toBe("2");
    expect(result.model.tasks[1].predecessors).toEqual([{ predecessorUid: "2" }]);
    expect(result.model.resources[0].name).toBe("Miku Updated");
    expect(result.model.resources[0].group).toBe("Dev");
    expect(result.model.resources[0].maxUnits).toBe(1);
    expect(result.model.resources[0].calendarUID).toBe("1");
    expect(result.model.tasks[1].outlineNumber).toBe(baseModel.tasks[1].outlineNumber);
    expect(result.changes.some((change) => change.field === "Name")).toBe(true);
  });

  it("rejects invalid workbook json format", () => {
    const { projectWorkbookJson } = bootModules();

    expect(() => projectWorkbookJson.validateWorkbookJsonDocument({
      format: "other",
      version: 1,
      sheets: {}
    })).toThrow("format が mikuproject_workbook_json ではありません");
  });

  it("reports warnings for unknown sheet and unknown columns", () => {
    const { projectWorkbookJson } = bootModules();

    const result = projectWorkbookJson.validateWorkbookJsonDocument({
      format: "mikuproject_workbook_json",
      version: 1,
      sheets: {
        Project: [{ Field: "Name", Value: "x", Extra: "ignored" }],
        UnknownSheet: []
      }
    });

    expect(result.warnings.map((item) => item.message)).toEqual([
      "未知の列は無視します: Project[0].Extra",
      "未知の sheet は無視します: UnknownSheet"
    ]);
  });

  it("rejects non-array sheets and non-object rows", () => {
    const { projectWorkbookJson } = bootModules();

    expect(() => projectWorkbookJson.validateWorkbookJsonDocument({
      format: "mikuproject_workbook_json",
      version: 1,
      sheets: {
        Project: {}
      }
    })).toThrow("sheets.Project は配列である必要があります");

    expect(() => projectWorkbookJson.validateWorkbookJsonDocument({
      format: "mikuproject_workbook_json",
      version: 1,
      sheets: {
        Tasks: [null]
      }
    })).toThrow("sheets.Tasks にオブジェクトではない行があります");
  });

  it("keeps non-editable task columns unchanged through workbook json import", () => {
    const { xml, projectWorkbookJson } = bootModules();
    const baseModel = xml.importMsProjectXml(dependencyXml);
    const documentLike = projectWorkbookJson.exportProjectWorkbookJson(baseModel);
    const executeTaskRow = documentLike.sheets.Tasks.find((row) => row.UID === "2");

    executeTaskRow.ID = "999";
    executeTaskRow.OutlineLevel = 9;
    executeTaskRow.OutlineNumber = "999";
    executeTaskRow.WBS = "WBS-999";

    const result = projectWorkbookJson.importProjectWorkbookJson(documentLike, baseModel);
    const importedTask = result.model.tasks.find((task) => task.uid === "2");
    const baseTask = baseModel.tasks.find((task) => task.uid === "2");

    expect(importedTask.id).toBe(baseTask.id);
    expect(importedTask.outlineLevel).toBe(baseTask.outlineLevel);
    expect(importedTask.outlineNumber).toBe(baseTask.outlineNumber);
    expect(importedTask.wbs).toBe(baseTask.wbs);
    expect(result.changes).toEqual([]);
  });
});
