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
const mainRenderCode = readFileSync(
  path.resolve(__dirname, "../src/js/main-render.js"),
  "utf8"
);

function bootRenderModule() {
  new Function(`${typesCode}\n${mainRenderCode}`)();
  return globalThis.__mikuprojectMainRender;
}

function mountFeedbackDom() {
  document.body.innerHTML = `
    <div class="md-feedback-stack md-hidden">
      <div class="md-label md-hidden">検証結果</div>
      <div id="validationIssues" class="md-hidden"></div>
      <div class="md-label md-hidden">import warnings</div>
      <div id="importWarnings" class="md-hidden"></div>
      <div class="md-label md-hidden">import summary</div>
      <div id="xlsxImportSummary" class="md-hidden"></div>
    </div>
  `;
}

function mountSummaryDom() {
  document.body.innerHTML = `
    <div id="summaryProjectName"></div>
    <div id="summaryTaskCount"></div>
    <div id="summaryResourceCount"></div>
    <div id="summaryAssignmentCount"></div>
    <div id="summaryCalendarCount"></div>
    <textarea id="modelOutput"></textarea>
    <div id="projectPreview"></div>
    <div id="taskPreview"></div>
    <div id="resourcePreview"></div>
    <div id="assignmentPreview"></div>
    <div id="calendarPreview"></div>
  `;
}

describe("mikuproject main render", () => {
  beforeEach(() => {
    document.body.innerHTML = "";
  });

  it("groups import warnings by uid and shows generic warnings", () => {
    const render = bootRenderModule();
    mountFeedbackDom();

    render.renderImportWarnings(document, [
      { message: "generic warning" },
      { message: "task warning 1", scope: "tasks", uid: "3", label: "Task C" },
      { message: "task warning 2", scope: "tasks", uid: "3", label: "Task C" },
      { message: "calendar warning", scope: "calendars", uid: "1", label: "Standard" }
    ], { sourceLabel: "Patch JSON" });

    expect(document.getElementById("importWarnings").classList.contains("md-hidden")).toBe(false);
    expect(document.querySelector(".md-feedback-stack")?.classList.contains("md-hidden")).toBe(false);
    expect(document.getElementById("importWarnings").textContent).toContain("Patch JSON warning");
    expect(document.getElementById("importWarnings").textContent).toContain("generic warning");
    expect(document.getElementById("importWarnings").textContent).toContain("UID=3 Task C");
    expect(document.getElementById("importWarnings").textContent).toContain("task warning 1");
    expect(document.getElementById("importWarnings").textContent).toContain("task warning 2");
    expect(document.getElementById("importWarnings").textContent).toContain("UID=1 Standard");
  });

  it("renders grouped import summary with warning digest and unchanged scopes", () => {
    const render = bootRenderModule();
    mountFeedbackDom();

    render.renderXlsxImportSummary(document, [
      { scope: "tasks", uid: "3", label: "Task C", field: "Name", before: "Before", after: "After" },
      { scope: "tasks", uid: "3", label: "Task C", field: "Notes", before: undefined, after: "memo" },
      { scope: "project", uid: "project", label: "Project A", field: "Name", before: "Project A", after: "Project B" }
    ], {
      sourceLabel: "Patch JSON",
      warnings: [
        { message: "warning a", scope: "tasks", uid: "3", label: "Task C" },
        { message: "warning b", scope: "tasks", uid: "3", label: "Task C" }
      ]
    });

    expect(document.getElementById("xlsxImportSummary").classList.contains("md-hidden")).toBe(false);
    expect(document.querySelector(".md-feedback-stack")?.classList.contains("md-hidden")).toBe(false);
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Patch JSON 反映結果");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Project 1 / 1 fields");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Tasks 1 / 2 fields");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Patch JSON warning 2 件 / UID=3 Task C");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("変更なし: Resources, Assignments, Calendars");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("UID=3 Task C");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Name");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Notes");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("(empty)");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Patch JSON の部分適用結果です");
  });

  it("hides feedback sections when validation issues are cleared", () => {
    const render = bootRenderModule();
    mountFeedbackDom();

    render.renderValidationIssues(document, [
      { level: "warning", scope: "tasks", message: "Task warning" }
    ]);
    expect(document.querySelector(".md-feedback-stack")?.classList.contains("md-hidden")).toBe(false);

    render.renderValidationIssues(document, []);
    render.renderImportWarnings(document, []);
    render.renderXlsxImportSummary(document, []);

    expect(document.getElementById("validationIssues").classList.contains("md-hidden")).toBe(true);
    expect(document.getElementById("importWarnings").classList.contains("md-hidden")).toBe(true);
    expect(document.getElementById("xlsxImportSummary").classList.contains("md-hidden")).toBe(true);
    expect(document.querySelector(".md-feedback-stack")?.classList.contains("md-hidden")).toBe(true);
  });

  it("escapes validation issue html-sensitive text", () => {
    const render = bootRenderModule();
    mountFeedbackDom();

    render.renderValidationIssues(document, [
      { level: "warning", scope: "tasks", message: "<script>alert(1)</script>" }
    ]);

    expect(document.getElementById("validationIssues").innerHTML).not.toContain("<script>");
    expect(document.getElementById("validationIssues").innerHTML).toContain("&lt;script&gt;alert(1)&lt;/script&gt;");
  });

  it("renders empty previews when summary model is null", () => {
    const render = bootRenderModule();
    mountSummaryDom();
    const updateSvgButton = { fn: 0 };

    render.updateSummary(document, null, () => {
      updateSvgButton.fn += 1;
    });

    expect(updateSvgButton.fn).toBe(1);
    expect(document.getElementById("summaryProjectName").textContent).toBe("-");
    expect(document.getElementById("summaryTaskCount").textContent).toBe("0");
    expect(document.getElementById("projectPreview").textContent).toContain("まだ表示できる項目がありません");
    expect(document.getElementById("taskPreview").textContent).toContain("まだ表示できる項目がありません");
  });

  it("renders summary previews with missing links and calendar references", () => {
    const render = bootRenderModule();
    mountSummaryDom();

    render.updateSummary(document, {
      project: {
        name: "Render Project",
        title: "Render Title",
        author: "Render Author",
        company: "Render Co",
        startDate: "2026-03-16T09:00:00",
        finishDate: "2026-03-18T18:00:00",
        scheduleFromStart: true,
        calendarUID: "1",
        outlineCodes: [],
        wbsMasks: [],
        extendedAttributes: []
      },
      tasks: [
        {
          uid: "1",
          id: "1",
          name: "Task A",
          outlineLevel: 1,
          outlineNumber: "1",
          start: "2026-03-16T09:00:00",
          finish: "2026-03-16T18:00:00",
          duration: "PT8H0M0S",
          milestone: false,
          summary: false,
          percentComplete: 0,
          calendarUID: "99",
          extendedAttributes: [],
          baselines: [],
          timephasedData: [],
          predecessors: [{ predecessorUid: "88" }]
        }
      ],
      resources: [
        {
          uid: "1",
          id: "1",
          name: "Miku",
          calendarUID: "77",
          extendedAttributes: [],
          baselines: [],
          timephasedData: []
        }
      ],
      assignments: [
        {
          uid: "1",
          taskUid: "1",
          resourceUid: "9",
          start: "2026-03-16T09:00:00",
          finish: "2026-03-16T18:00:00",
          extendedAttributes: [],
          baselines: [],
          timephasedData: []
        }
      ],
      calendars: [
        {
          uid: "1",
          name: "Standard",
          isBaseCalendar: true,
          baseCalendarUID: undefined,
          weekDays: [],
          exceptions: [],
          workWeeks: []
        }
      ]
    }, () => {});

    expect(document.getElementById("summaryProjectName").textContent).toBe("Render Project");
    expect(document.getElementById("summaryTaskCount").textContent).toBe("1");
    expect(document.getElementById("summaryResourceCount").textContent).toBe("1");
    expect(document.getElementById("summaryAssignmentCount").textContent).toBe("1");
    expect(document.getElementById("summaryCalendarCount").textContent).toBe("1");
    expect(document.getElementById("modelOutput").value).toContain("\"Render Project\"");
    expect(document.getElementById("projectPreview").textContent).toContain("Calendar=1 (Standard)");
    expect(document.getElementById("taskPreview").textContent).toContain("Calendar=99 (missing)");
    expect(document.getElementById("taskPreview").textContent).toContain("Predecessors=88");
    expect(document.getElementById("resourcePreview").textContent).toContain("Calendar=77 (missing)");
    expect(document.getElementById("assignmentPreview").textContent).toContain("Resource=9 (missing)");
    expect(document.getElementById("calendarPreview").textContent).toContain("Refs=Project=1 / Tasks=0 / Resources=0 / BaseOf=0");
  });

  it("escapes warning and summary html-sensitive text", () => {
    const render = bootRenderModule();
    mountFeedbackDom();

    render.renderImportWarnings(document, [
      { message: "<script>alert(1)</script>", scope: "tasks", uid: "3", label: "<Task>" }
    ], { sourceLabel: "Patch JSON" });
    render.renderXlsxImportSummary(document, [
      {
        scope: "tasks",
        uid: "3",
        label: "<Task>",
        field: "Name",
        before: "<before>",
        after: "<after>"
      }
    ]);

    expect(document.getElementById("importWarnings").innerHTML).not.toContain("<script>");
    expect(document.getElementById("importWarnings").innerHTML).toContain("&lt;script&gt;alert(1)&lt;/script&gt;");
    expect(document.getElementById("xlsxImportSummary").innerHTML).not.toContain("<before>");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("&lt;before&gt;");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("&lt;after&gt;");
  });
});
