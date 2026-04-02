// @vitest-environment jsdom

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { beforeEach, describe, expect, it, vi } from "vitest";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const typesCode = readFileSync(
  path.resolve(__dirname, "../src/js/types.js"),
  "utf8"
);
const markdownEscapeCode = readFileSync(
  path.resolve(__dirname, "../src/js/markdown-escape.js"),
  "utf8"
);
const aiJsonUtilCode = readFileSync(
  path.resolve(__dirname, "../src/js/ai-json-util.js"),
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
const projectPatchJsonCode = readFileSync(
  path.resolve(__dirname, "../src/js/project-patch-json.js"),
  "utf8"
);
const mainRenderCode = readFileSync(
  path.resolve(__dirname, "../src/js/main-render.js"),
  "utf8"
);
const mainCode = readFileSync(
  path.resolve(__dirname, "../src/js/main.js"),
  "utf8"
);
const dependencyXml = readFileSync(
  path.resolve(__dirname, "../testdata/dependency.xml"),
  "utf8"
);
const hierarchyXml = readFileSync(
  path.resolve(__dirname, "../testdata/hierarchy.xml"),
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

const bootPageCode = `${typesCode}\n${markdownEscapeCode}\n${aiJsonUtilCode}\n${mainUtilCode}\n${excelIoStubCode}\n${msProjectXmlCode}\n${projectXlsxStubCode}\n${projectWorkbookJsonStubCode}\n${projectPatchJsonStubCode}\n${wbsXlsxStubCode}\n${wbsMarkdownStubCode}\n${nativeSvgStubCode}\n${mainRenderCode}\n${mainCode}`;
const bootPageRunner = new Function(bootPageCode);
const bootPatchModuleRunner = new Function(`${typesCode}\n${msProjectXmlCode}\n${projectPatchJsonCode}`);

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

function bootPage() {
  mountDom();
  bootPageRunner();
  document.dispatchEvent(new Event("DOMContentLoaded"));
}

function getMainHooks() {
  return globalThis.__mikuprojectMainTestHooks;
}

function parseXmlViaHook() {
  getMainHooks().parseCurrentXml();
}

function applyPatchWithoutMain(xmlText, documentLike) {
  bootPatchModuleRunner();
  const xml = globalThis.__mikuprojectXml;
  const patch = globalThis.__mikuprojectProjectPatchJson;
  const model = xml.importMsProjectXml(xmlText);
  const result = patch.importProjectPatchJson(documentLike, model);
  return {
    ...result,
    xml: xml.exportMsProjectXml(result.model)
  };
}

describe("mikuproject main validation", () => {
  beforeEach(() => {
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
  });

  it("renders validation issues without xlsx import wiring", () => {
    bootPage();

    getMainHooks().renderValidationIssues([
      { level: "warning", scope: "tasks", message: "Task UID=2 (Design) PercentComplete must be within 0..100" },
      { level: "warning", scope: "tasks", message: "Task UID=2 (Design) Start must be earlier than or equal to Finish" },
      { level: "warning", scope: "calendars", message: "Calendar BaseCalendarUID が自身を指しています: UID=1 Name=Standard" }
    ]);

    expect(document.getElementById("validationIssues").classList.contains("md-hidden")).toBe(false);
    expect(document.querySelector(".md-feedback-stack")?.classList.contains("md-hidden")).toBe(false);
    expect(document.getElementById("validationIssues").textContent).toContain("Tasks");
    expect(document.getElementById("validationIssues").textContent).toContain("Calendars");
    expect(document.getElementById("validationIssues").textContent).toContain("PercentComplete");
    expect(document.getElementById("validationIssues").textContent).toContain("Start");
    expect(document.getElementById("validationIssues").textContent).toContain("BaseCalendarUID");
  });

  it("reports validation error when assignment references a missing resource", () => {
    bootPage();

    document.getElementById("xmlInput").value = dependencyXml.replace(
      "<ResourceUID>1</ResourceUID>",
      "<ResourceUID>99</ResourceUID>"
    );
    parseXmlViaHook();
    document.getElementById("roundTripBtn").click();

    expect(document.getElementById("validationIssues").textContent).toContain("Assignment ResourceUID");
    expect(document.getElementById("validationIssues").textContent).toContain("UID=1");
    expect(document.getElementById("validationIssues").textContent).toContain("TaskUID=2");
    expect(document.getElementById("validationIssues").textContent).toContain("Execute");
    expect(document.getElementById("validationIssues").textContent).toContain("ResourceUID=99");
  });

  it("reports validation error when project calendar does not exist", () => {
    bootPage();

    document.getElementById("xmlInput").value = dependencyXml.replace(
      "<CalendarUID>1</CalendarUID>",
      "<CalendarUID>99</CalendarUID>"
    );
    parseXmlViaHook();

    expect(document.getElementById("statusMessage").textContent).toContain("検証で");
    expect(document.getElementById("validationIssues").textContent).toContain("Project");
    expect(document.getElementById("validationIssues").textContent).toContain("Project CalendarUID");
  });

  it("reports validation warning when task calendar does not exist", () => {
    bootPage();

    document.getElementById("xmlInput").value = dependencyXml.replace(
      "<PercentComplete>0</PercentComplete>",
      "<PercentComplete>0</PercentComplete>\n      <CalendarUID>99</CalendarUID>"
    );
    parseXmlViaHook();

    expect(document.getElementById("validationIssues").textContent).toContain("Task CalendarUID");
    expect(document.getElementById("validationIssues").textContent).toContain("UID=2");
    expect(document.getElementById("validationIssues").textContent).toContain("Execute");
  });

  it("reports validation warning when percent complete is out of range", () => {
    bootPage();

    document.getElementById("xmlInput").value = dependencyXml.replace(
      "<PercentComplete>0</PercentComplete>",
      "<PercentComplete>120</PercentComplete>"
    );
    parseXmlViaHook();

    expect(document.getElementById("validationIssues").textContent).toContain("PercentComplete");
  });

  it("reports validation warning when task start is after finish", () => {
    bootPage();

    document.getElementById("xmlInput").value = dependencyXml.replace(
      "<Start>2026-03-18T09:00:00</Start>\n      <Finish>2026-03-19T18:00:00</Finish>",
      "<Start>2026-03-21T09:00:00</Start>\n      <Finish>2026-03-20T18:00:00</Finish>"
    );
    parseXmlViaHook();

    expect(document.getElementById("validationIssues").textContent).toContain("Task Start が Finish より後");
  });

  it("reports validation warning when task order does not match outline order", () => {
    bootPage();

    document.getElementById("xmlInput").value = dependencyXml.replace(
      "<OutlineNumber>2</OutlineNumber>",
      "<OutlineNumber>1</OutlineNumber>"
    );
    parseXmlViaHook();

    expect(document.getElementById("validationIssues").textContent).toContain("Task の並び順");
    expect(document.getElementById("validationIssues").textContent).toContain("UID=2");
    expect(document.getElementById("validationIssues").textContent).toContain("Execute");
    expect(document.getElementById("validationIssues").textContent).toContain("UID=1");
    expect(document.getElementById("validationIssues").textContent).toContain("Prepare");
  });

  it("reports validation error when predecessor references a missing task", () => {
    bootPage();

    document.getElementById("xmlInput").value = dependencyXml.replace(
      "<PredecessorUID>1</PredecessorUID>",
      "<PredecessorUID>99</PredecessorUID>"
    );
    parseXmlViaHook();
    document.getElementById("roundTripBtn").click();

    expect(document.getElementById("validationIssues").textContent).toContain("PredecessorUID");
    expect(document.getElementById("validationIssues").textContent).toContain("UID=2");
    expect(document.getElementById("validationIssues").textContent).toContain("Execute");
    expect(document.getElementById("validationIssues").textContent).toContain("TaskUID=99");
  });

  it("rejects invalid patch json operations without main UI", () => {
    const projectResult = applyPatchWithoutMain(dependencyXml, {
      operations: [{
        op: "update_project",
        fields: {
          name: "",
          start_date: "2026-03-20",
          finish_date: "2026-03-19",
          current_date: "bad-date",
          calendar_uid: "999",
          minutes_per_day: 0,
          schedule_from_start: "yes"
        }
      }]
    });
    const assignmentResult = applyPatchWithoutMain(dependencyXml, {
      operations: [
        { op: "update_assignment", uid: "1", fields: { start: "2026-03-20", finish: "2026-03-19", units: -1, work: "", percent_work_complete: 120 } },
        { op: "add_assignment", uid: "3", task_uid: "999", resource_uid: "1", units: -1 }
      ]
    });

    const projectWarnings = projectResult.warnings.map((warning) => warning.message).join("\n");
    const assignmentWarnings = assignmentResult.warnings.map((warning) => warning.message).join("\n");

    expect(projectResult.changes).toHaveLength(0);
    expect(projectWarnings).toContain("update_project.name は空でない文字列が必要です");
    expect(projectWarnings).toContain("update_project.start_date が finish_date より後です");
    expect(projectWarnings).toContain("update_project.current_date の日付形式が解釈できません");
    expect(projectWarnings).toContain("update_project.calendar_uid が既存 calendar を指していません");
    expect(projectWarnings).toContain("update_project.minutes_per_day は 0 より大きい数値が必要です");
    expect(projectWarnings).toContain("update_project.schedule_from_start は boolean が必要です");
    expect(projectResult.xml).toContain("<Name>Dependency Project</Name>");

    expect(assignmentResult.changes).toHaveLength(0);
    expect(assignmentWarnings).toContain("update_assignment.start が finish より後です");
    expect(assignmentWarnings).toContain("update_assignment.units は 0 以上の数値が必要です");
    expect(assignmentWarnings).toContain("update_assignment.work は空でない文字列が必要です");
    expect(assignmentWarnings).toContain("update_assignment.percent_work_complete は 0 以上 100 以下の数値が必要です");
    expect(assignmentWarnings).toContain("add_assignment.task_uid が既存 task を指していません");
  });

  it("reports patch json warning details without main UI", () => {
    const linked = applyPatchWithoutMain(hierarchyXml, {
      operations: [{ op: "link_tasks", from_uid: "2", to_uid: "3", type: "SS", lag_hours: 4 }]
    });
    const duplicate = applyPatchWithoutMain(linked.xml, {
      operations: [{ op: "link_tasks", from_uid: "2", to_uid: "3", type: "SS", lag_hours: 4 }]
    });
    const missingLag = applyPatchWithoutMain(linked.xml, {
      operations: [{ op: "unlink_tasks", from_uid: "2", to_uid: "3", type: "SS", lag_hours: 8 }]
    });
    const modelResult = applyPatchWithoutMain(dependencyXml, {
      operations: [
        {
          op: "update_resource",
          uid: "1",
          fields: { name: "", calendar_uid: "999", max_units: -1, cost_per_use: -1, percent_work_complete: 120 }
        },
        { op: "update_calendar", uid: "1", fields: { name: "", is_base_calendar: "yes", base_calendar_uid: "1" } },
        { op: "add_calendar", uid: "2", name: "", is_base_calendar: "yes", base_calendar_uid: "999" },
        { op: "add_resource", uid: "2", name: "", calendar_uid: "999", max_units: -1, cost_per_use: -1, percent_work_complete: 120 },
        { op: "delete_calendar", uid: "1" },
        { op: "delete_resource", uid: "1" }
      ]
    });

    expect(duplicate.warnings.map((warning) => warning.message).join("\n")).toContain(
      "link_tasks の依存関係は既に存在します: 2 -> 3 (SS, lag=PT4H0M0S)"
    );
    expect(missingLag.warnings.map((warning) => warning.message).join("\n")).toContain(
      "unlink_tasks の対象依存関係が見つかりません: 2 -> 3 (SS, lag=PT8H0M0S)"
    );

    const modelWarnings = modelResult.warnings.map((warning) => warning.message).join("\n");
    expect(modelWarnings).toContain("update_resource.name は空でない文字列が必要です");
    expect(modelWarnings).toContain("update_resource.calendar_uid が既存 calendar を指していません");
    expect(modelWarnings).toContain("update_resource.max_units は 0 以上の数値が必要です");
    expect(modelWarnings).toContain("update_resource.cost_per_use は 0 以上の数値が必要です");
    expect(modelWarnings).toContain("update_resource.percent_work_complete は 0 以上 100 以下の数値が必要です");
    expect(modelWarnings).toContain("update_calendar.name は空でない文字列が必要です");
    expect(modelWarnings).toContain("update_calendar.is_base_calendar は boolean が必要です");
    expect(modelWarnings).toContain("update_calendar.base_calendar_uid は自身を指せません");
    expect(modelWarnings).toContain("add_calendar.name は空でない文字列が必要です");
    expect(modelWarnings).toContain("add_resource.name は空でない文字列が必要です");
    expect(modelWarnings).toContain("delete_calendar first cut では参照が残っている calendar は削除できません");
    expect(modelWarnings).toContain("delete_resource first cut では assignment がある resource は削除できません");
  });
});
