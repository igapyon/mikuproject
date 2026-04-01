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
const dependencyXml = readFileSync(
  path.resolve(__dirname, "../testdata/dependency.xml"),
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
});
