// @vitest-environment jsdom

import { beforeEach, describe, expect, it, vi } from "vitest";

import {
  bootPage,
  flushAsyncWork,
  getMainHooks,
  setupMainPreviewExportDom
} from "./helpers/main-preview-export-harness.js";

function createTextFile(contents, name, type = "application/json") {
  const file = new File([contents], name, { type });
  Object.defineProperty(file, "text", {
    configurable: true,
    value: () => Promise.resolve(contents)
  });
  return file;
}

describe("mikuproject main refactor safety", () => {
  beforeEach(() => {
    setupMainPreviewExportDom();
  });

  it("builds archive entries with the expected core outputs", () => {
    bootPage();

    const entries = getMainHooks().buildCurrentOutputArchiveEntries();
    const names = entries.map((entry) => entry.name);
    const readmeText = new TextDecoder().decode(entries.find((entry) => entry.name === "README.txt").data);
    const bundleText = new TextDecoder().decode(entries.find((entry) => entry.name === "mikuproject-full-bundle.editjson").data);

    expect(names).toContain("README.txt");
    expect(names.some((name) => /^mikuproject-export-\d{12}\.xml$/.test(name))).toBe(true);
    expect(names.some((name) => /^mikuproject-export-\d{12}\.xlsx$/.test(name))).toBe(true);
    expect(names.some((name) => /^mikuproject-workbook-\d{12}\.json$/.test(name))).toBe(true);
    expect(names.some((name) => /^mikuproject-wbs-\d{8}\.md$/.test(name))).toBe(true);
    expect(names).toContain("mikuproject-project-overview-view.editjson");
    expect(names).toContain("mikuproject-full-bundle.editjson");
    expect(names).toContain("mikuproject-phase-detail-view-full.editjson");
    expect(names.some((name) => name.startsWith("monthly-calendar/"))).toBe(true);
    expect(readmeText).toContain("mikuproject ALL ZIP");
    expect(readmeText).toContain("monthly-calendar/YYYY-MM.svg");

    const bundle = JSON.parse(bundleText);
    expect(bundle.view_type).toBe("ai_projection_bundle");
    expect(bundle.project_overview_view.view_type).toBe("project_overview_view");
    expect(Array.isArray(bundle.phase_detail_views_full)).toBe(true);
    expect(bundle.phase_detail_views_full.length).toBeGreaterThan(0);
    expect(Array.isArray(bundle.task_edit_views_full)).toBe(true);
    expect(bundle.task_edit_views_full.length).toBeGreaterThan(0);
  });

  it("dispatches project_draft_view json files through importFromFile", async () => {
    bootPage();

    const file = createTextFile(JSON.stringify({
      view_type: "project_draft_view",
      project: {
        name: "File Draft Import",
        planned_start: "2026-04-01"
      },
      tasks: [
        {
          uid: "draft-1",
          name: "着手",
          parent_uid: null,
          position: 0,
          is_milestone: true,
          planned_start: "2026-04-01",
          planned_finish: "2026-04-01"
        }
      ],
      resources: [],
      assignments: []
    }, null, 2), "draft.json");

    await getMainHooks().importFromFile(file);
    await flushAsyncWork();

    expect(document.getElementById("projectDraftImportInput").value).toContain("\"view_type\": \"project_draft_view\"");
    expect(document.getElementById("summaryProjectName").textContent).toBe("File Draft Import");
    expect(document.getElementById("xmlInput").value).toContain("<Name>File Draft Import</Name>");
    expect(document.getElementById("statusMessage").textContent).toContain("project_draft_view を取り込みました");
  });

  it("dispatches patch editjson files through importFromFile", async () => {
    bootPage();

    const file = createTextFile(JSON.stringify({
      operations: [
        {
          op: "update_project",
          fields: {
            name: "Patched From File"
          }
        }
      ]
    }, null, 2), "update.editjson");

    await getMainHooks().importFromFile(file);
    await flushAsyncWork();

    expect(document.getElementById("projectDraftImportInput").value).toContain("\"update_project\"");
    expect(document.getElementById("summaryProjectName").textContent).toBe("Patched From File");
    expect(document.getElementById("modelOutput").value).toContain("\"name\": \"Patched From File\"");
    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON");
  });

  it("fails round-trip check when normalized models diverge", () => {
    bootPage();

    const originalNormalize = globalThis.__mikuprojectXml.normalizeProjectModel;
    globalThis.__mikuprojectXml.normalizeProjectModel = vi.fn()
      .mockImplementationOnce((model) => model)
      .mockImplementationOnce((model) => ({
        ...model,
        project: {
          ...model.project,
          name: `${model.project.name} mismatch`
        }
      }));

    expect(() => getMainHooks().runRoundTripCheck()).toThrow("再読込後の内部モデルが一致しません");

    globalThis.__mikuprojectXml.normalizeProjectModel = originalNormalize;
  });
});
