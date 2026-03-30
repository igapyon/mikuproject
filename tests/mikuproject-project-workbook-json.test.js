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

function bootModules() {
  new Function(`${typesCode}\n${excelIoCode}\n${msProjectXmlCode}\n${projectWorkbookSchemaCode}\n${projectXlsxCode}\n${projectWorkbookJsonCode}`)();
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
  });

  it("imports limited editable fields through workbook json", () => {
    const { xml, projectWorkbookJson } = bootModules();
    const baseModel = xml.importMsProjectXml(xml.SAMPLE_XML);
    const documentLike = projectWorkbookJson.exportProjectWorkbookJson(baseModel);

    documentLike.sheets.Project.find((row) => row.Field === "Name").Value = "JSON import project";
    documentLike.sheets.Tasks[2].Name = "JSON import task";
    documentLike.sheets.Tasks[2].Start = "2026-03-16 10:00:00";
    documentLike.sheets.Tasks[2].OutlineNumber = "999";

    const result = projectWorkbookJson.importProjectWorkbookJson(documentLike, baseModel);

    expect(result.model.project.name).toBe("JSON import project");
    expect(result.model.tasks[2].name).toBe("JSON import task");
    expect(result.model.tasks[2].start).toBe("2026-03-16T10:00:00");
    expect(result.model.tasks[2].outlineNumber).toBe(baseModel.tasks[2].outlineNumber);
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
});
