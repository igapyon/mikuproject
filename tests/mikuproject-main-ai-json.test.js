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
const hierarchyXml = readFileSync(
  path.resolve(__dirname, "../testdata/hierarchy.xml"),
  "utf8"
);
const dependencyXml = readFileSync(
  path.resolve(__dirname, "../testdata/dependency.xml"),
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
  exportWeeklyNativeSvg: () => "<svg data-stub=\\"weekly\\">weekly overview</svg>",
  exportMonthlyWbsCalendarSvgArchive: () => ({
    entries: [{ fileName: "2026-03.svg", svg: "<svg data-stub=\\"monthly\\">2026-03</svg>" }],
    zipBytes: new Uint8Array()
  })
};
`;

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
    <button id="copyAiPromptBtnPane" type="button"></button>
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
      <template id="aiPromptTemplate"># mikuproject AI JSON Spec

あなたはこれから mikuproject とやりとりします。</template>
      <textarea id="projectDraftImportInput"></textarea>
      <div id="xmlSaveState"></div>
    </section>
    <section id="tabPanelTransform" class="md-tab-panel" data-tab-panel="transform" hidden>
      <div id="summaryProjectName"></div>
      <div id="summaryTaskCount"></div>
      <div id="summaryResourceCount"></div>
      <div id="summaryAssignmentCount"></div>
      <div id="summaryCalendarCount"></div>
      <div id="nativeSvgPreview"></div>
      <textarea id="modelOutput"></textarea>
      <div id="projectPreview"></div>
      <div id="taskPreview"></div>
      <div id="resourcePreview"></div>
      <div id="assignmentPreview"></div>
      <div id="calendarPreview"></div>
      <textarea id="mermaidOutput"></textarea>
      <div id="validationIssues" class="md-hidden"></div>
      <div id="importWarnings" class="md-hidden"></div>
      <div id="xlsxImportSummary" class="md-hidden"></div>
    </section>
    <section id="tabPanelOutput" class="md-tab-panel" data-tab-panel="output" hidden>
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
  const copyButton = document.getElementById("copyAiPromptBtnPane");
  if (copyButton) {
    copyButton.id = "copyAiPromptBtn";
  }
}

function bootPage() {
  mountDom();
  new Function(`${typesCode}\n${markdownEscapeCode}\n${aiJsonUtilCode}\n${mainUtilCode}\n${msProjectXmlCode}\n${projectPatchJsonCode}\n${excelIoStubCode}\n${projectXlsxStubCode}\n${projectWorkbookJsonStubCode}\n${wbsXlsxStubCode}\n${wbsMarkdownStubCode}\n${nativeSvgStubCode}\n${mainRenderCode}\n${mainCode}`)();
  document.dispatchEvent(new Event("DOMContentLoaded"));
}

function getMainHooks() {
  return globalThis.__mikuprojectMainTestHooks;
}

function parseXmlViaHook() {
  getMainHooks().parseCurrentXml();
}

async function flushAsyncWork() {
  await Promise.resolve();
  await Promise.resolve();
}

describe("mikuproject main ai json", () => {
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

  it("exports ai projection views", () => {
    bootPage();

    document.getElementById("xmlInput").value = hierarchyXml;
    parseXmlViaHook();
    const createObjectUrlCalls = URL.createObjectURL.mock.calls.length;
    const anchorClickCalls = HTMLAnchorElement.prototype.click.mock.calls.length;

    document.getElementById("exportProjectOverviewBtn").click();
    document.getElementById("exportPhaseDetailFullBtn").click();

    const projectOverview = JSON.parse(document.getElementById("projectOverviewOutput").value);
    const phaseDetail = JSON.parse(document.getElementById("phaseDetailOutput").value);

    expect(projectOverview.view_type).toBe("project_overview_view");
    expect(Array.isArray(projectOverview.phases)).toBe(true);
    expect(projectOverview.phases.length).toBeGreaterThan(0);
    expect(phaseDetail.view_type).toBe("phase_detail_view");
    expect(Array.isArray(phaseDetail.tasks)).toBe(true);
    expect(phaseDetail.phase.uid).toBeTruthy();
    expect(phaseDetail.scope).toEqual({ mode: "full", root_uid: null, max_depth: null });
    const downloads = HTMLAnchorElement.prototype.click.mock.instances
      .slice(anchorClickCalls)
      .map((anchor) => anchor.download);
    expect(downloads).toContain("mikuproject-project-overview-view.editjson");
    expect(downloads).toContain(`mikuproject-phase-detail-view-${phaseDetail.phase.uid}-full.editjson`);
    expect(URL.createObjectURL.mock.calls.length - createObjectUrlCalls).toBeGreaterThanOrEqual(2);
    expect(HTMLAnchorElement.prototype.click.mock.calls.length - anchorClickCalls).toBeGreaterThanOrEqual(2);
  });

  it("exports task_edit_view", () => {
    bootPage();

    document.getElementById("xmlInput").value = hierarchyXml;
    parseXmlViaHook();
    document.getElementById("taskEditUidInput").value = "3";

    document.getElementById("exportTaskEditBtn").click();

    const taskEdit = JSON.parse(document.getElementById("taskEditOutput").value);
    expect(taskEdit.view_type).toBe("task_edit_view");
    expect(taskEdit.target_task.uid).toBe("3");
    expect(taskEdit.parent_task).toBeTruthy();
    expect(Array.isArray(taskEdit.sibling_tasks)).toBe(true);
    expect(Array.isArray(taskEdit.predecessors)).toBe(true);
    expect(Array.isArray(taskEdit.successors)).toBe(true);
    expect(Array.isArray(taskEdit.assignments)).toBe(true);
    expect(taskEdit.rules.allow_patch_ops).toContain("update_task");
    expect(taskEdit.rules.allow_patch_ops).toContain("update_assignment");
    const clickedAnchor = HTMLAnchorElement.prototype.click.mock.instances.at(-1);
    expect(clickedAnchor.download).toBe("mikuproject-task-edit-view-3.editjson");
  });

  it("exports ai projection bundle", () => {
    bootPage();

    document.getElementById("xmlInput").value = hierarchyXml;
    parseXmlViaHook();
    const createObjectUrlCalls = URL.createObjectURL.mock.calls.length;
    const anchorClickCalls = HTMLAnchorElement.prototype.click.mock.calls.length;

    document.getElementById("exportAiBundleBtn").click();

    const bundle = JSON.parse(document.getElementById("aiBundleOutput").value);
    expect(bundle.view_type).toBe("ai_projection_bundle");
    expect(bundle.project_overview_view.view_type).toBe("project_overview_view");
    expect(Array.isArray(bundle.project_overview_view.phases)).toBe(true);
    expect(Array.isArray(bundle.phase_detail_views_full)).toBe(true);
    expect(bundle.phase_detail_views_full.length).toBeGreaterThan(0);
    expect(bundle.phase_detail_views_full.every((item) => item.view_type === "phase_detail_view")).toBe(true);
    expect(bundle.phase_detail_views_full.every((item) => item.scope?.mode === "full")).toBe(true);
    expect(Array.isArray(bundle.task_edit_views_full)).toBe(true);
    expect(bundle.task_edit_views_full.length).toBeGreaterThan(0);
    expect(bundle.task_edit_views_full.every((item) => item.view_type === "task_edit_view")).toBe(true);
    const clickedAnchor = HTMLAnchorElement.prototype.click.mock.instances.at(-1);
    expect(clickedAnchor.download).toBe("mikuproject-full-bundle.editjson");
    expect(URL.createObjectURL.mock.calls.length - createObjectUrlCalls).toBeGreaterThanOrEqual(1);
    expect(HTMLAnchorElement.prototype.click.mock.calls.length - anchorClickCalls).toBeGreaterThanOrEqual(1);
  });

  it("exports scoped phase_detail_view", () => {
    bootPage();

    document.getElementById("xmlInput").value = hierarchyXml;
    parseXmlViaHook();
    document.getElementById("phaseDetailUidInput").value = "1";
    document.getElementById("phaseDetailRootUidInput").value = "2";
    document.getElementById("phaseDetailMaxDepthInput").value = "1";

    document.getElementById("exportPhaseDetailBtn").click();

    const phaseDetail = JSON.parse(document.getElementById("phaseDetailOutput").value);
    expect(phaseDetail.scope).toEqual({ mode: "scoped", root_uid: "2", max_depth: 1 });
    expect(phaseDetail.tasks.every((task) => ["2", "3", "4", "5", "18"].includes(task.uid))).toBe(true);
    expect(phaseDetail.tasks.some((task) => task.uid === "19")).toBe(false);
    const clickedAnchor = HTMLAnchorElement.prototype.click.mock.instances.at(-1);
    expect(clickedAnchor.download).toBe("mikuproject-phase-detail-view-1-scoped-root-2-depth-1.editjson");
  });

  it("imports project_draft_view", async () => {
    bootPage();

    document.getElementById("projectDraftImportInput").value = [
      "説明文",
      "```json",
      JSON.stringify({
        view_type: "project_draft_view",
        project: {
          name: "新規基幹刷新",
          planned_start: "2026-04-01",
          schedule_from_start: true,
          minutes_per_day: 480,
          minutes_per_week: 2400,
          days_per_month: 20
        },
        tasks: [
          { uid: "draft-1", name: "要件定義", parent_uid: null, position: 0, is_summary: true, percent_complete: 100 },
          { uid: "draft-2", name: "ヒアリング", parent_uid: "draft-1", position: 0, percent_complete: 50, planned_finish: "2026-04-01" },
          { uid: "draft-3", name: "整理期間", parent_uid: "draft-1", position: 1, planned_start: "2026-04-02", planned_finish: "2026-04-03" },
          { uid: "draft-4", name: "要件確定", parent_uid: "draft-1", position: 2, is_milestone: true, predecessors: ["draft-2"], planned_start: "2026-04-08T18:00:00", planned_finish: "2026-04-08T18:00:00" }
        ],
        resources: [
          { uid: "res-1", name: "Mikuku", initials: "M", group: "PMO", max_units: 1, calendar_uid: "1" }
        ],
        assignments: [
          { uid: "asg-1", task_uid: "draft-2", resource_uid: "res-1", start: "2026-04-01T09:00:00", finish: "2026-04-01T18:00:00", units: 1, work: "PT8H0M0S", percent_work_complete: 50 }
        ]
      }, null, 2),
      "```"
    ].join("\n");

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("summaryProjectName").textContent).toBe("新規基幹刷新");
    expect(document.getElementById("summaryTaskCount").textContent).toBe("4");
    expect(document.getElementById("summaryResourceCount").textContent).toBe("1");
    expect(document.getElementById("summaryAssignmentCount").textContent).toBe("1");
    expect(document.getElementById("summaryCalendarCount").textContent).toBe("1");
    expect(document.getElementById("xmlInput").value).toContain("<Name>新規基幹刷新</Name>");
    expect(document.getElementById("xmlInput").value).toContain("<Title>新規基幹刷新</Title>");
    expect(document.getElementById("xmlInput").value).toContain("<CalendarUID>1</CalendarUID>");
    expect(document.getElementById("xmlInput").value).toContain("<ScheduleFromStart>1</ScheduleFromStart>");
    expect(document.getElementById("xmlInput").value).toContain("<MinutesPerDay>480</MinutesPerDay>");
    expect(document.getElementById("xmlInput").value).toContain("<MinutesPerWeek>2400</MinutesPerWeek>");
    expect(document.getElementById("xmlInput").value).toContain("<DaysPerMonth>20</DaysPerMonth>");
    expect(document.getElementById("xmlInput").value).toContain("<Name>Standard</Name>");
    expect(document.getElementById("xmlInput").value).toContain("<UID>3</UID>");
    expect(document.getElementById("modelOutput").value).toContain("\"title\": \"新規基幹刷新\"");
    expect(document.getElementById("modelOutput").value).toContain("\"scheduleFromStart\": true");
    expect(document.getElementById("modelOutput").value).toContain("\"minutesPerDay\": 480");
    expect(document.getElementById("modelOutput").value).toContain("\"minutesPerWeek\": 2400");
    expect(document.getElementById("modelOutput").value).toContain("\"daysPerMonth\": 20");
    expect(document.getElementById("modelOutput").value).toContain("\"name\": \"ヒアリング\"");
    expect(document.getElementById("modelOutput").value).toContain("\"milestone\": false");
    expect(document.getElementById("modelOutput").value).toContain("\"percentComplete\": 100");
    expect(document.getElementById("modelOutput").value).toContain("\"percentComplete\": 50");
    expect(document.getElementById("modelOutput").value).toContain("\"start\": \"2026-04-01T09:00:00\"");
    expect(document.getElementById("modelOutput").value).toContain("\"finish\": \"2026-04-01T18:00:00\"");
    expect(document.getElementById("modelOutput").value).toContain("\"name\": \"整理期間\"");
    expect(document.getElementById("modelOutput").value).toContain("\"start\": \"2026-04-02T09:00:00\"");
    expect(document.getElementById("modelOutput").value).toContain("\"finish\": \"2026-04-03T18:00:00\"");
    expect(document.getElementById("modelOutput").value).toContain("\"name\": \"要件確定\"");
    expect(document.getElementById("modelOutput").value).toContain("\"milestone\": true");
    expect(document.getElementById("modelOutput").value).toContain("\"uid\": \"4\"");
    expect(document.getElementById("modelOutput").value).not.toContain("\"uid\": \"draft-4\"");
    expect(document.getElementById("modelOutput").value).toContain("\"name\": \"Standard\"");
    expect(document.getElementById("modelOutput").value).toContain("\"name\": \"Mikuku\"");
    expect(document.getElementById("modelOutput").value).toContain("\"initials\": \"M\"");
    expect(document.getElementById("modelOutput").value).toContain("\"taskUid\": \"2\"");
    expect(document.getElementById("modelOutput").value).toContain("\"resourceUid\": \"1\"");
  });

  it("fills default project minute settings when project_draft_view omits them", async () => {
    bootPage();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      view_type: "project_draft_view",
      project: {
        name: "既定時間補完確認",
        planned_start: "2026-04-01"
      },
      tasks: [
        { uid: "draft-1", name: "確認", parent_uid: null, position: 0, planned_start: "2026-04-01", planned_finish: "2026-04-01" }
      ],
      resources: [],
      assignments: []
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("xmlInput").value).toContain("<MinutesPerDay>480</MinutesPerDay>");
    expect(document.getElementById("xmlInput").value).toContain("<MinutesPerWeek>2400</MinutesPerWeek>");
    expect(document.getElementById("xmlInput").value).toContain("<DaysPerMonth>20</DaysPerMonth>");
    expect(document.getElementById("modelOutput").value).toContain("\"minutesPerDay\": 480");
    expect(document.getElementById("modelOutput").value).toContain("\"minutesPerWeek\": 2400");
    expect(document.getElementById("modelOutput").value).toContain("\"daysPerMonth\": 20");
  });

  it("imports patch json edits into the current model and xml", async () => {
    bootPage();

    document.getElementById("projectDraftImportInput").value = [
      "説明文",
      "```json",
      JSON.stringify({
        operations: [
          {
            op: "update_task",
            uid: "10",
            fields: {
              planned_start: "2026-03-24",
              planned_finish: "2026-03-26"
            }
          },
          {
            op: "update_task",
            uid: "11",
            fields: {
              name: "XLSXレイアウト最終調整"
            }
          }
        ]
      }, null, 2),
      "```"
    ].join("\n");

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("xmlInput").value).toContain("<Name>XLSXレイアウト最終調整</Name>");
    expect(document.getElementById("xmlInput").value).toContain("<Start>2026-03-24T09:00:00</Start>");
    expect(document.getElementById("xmlInput").value).toContain("<Finish>2026-03-26T18:00:00</Finish>");
    expect(document.getElementById("modelOutput").value).toContain("\"name\": \"XLSXレイアウト最終調整\"");
    expect(document.getElementById("modelOutput").value).toContain("\"start\": \"2026-03-24T09:00:00\"");
    expect(document.getElementById("modelOutput").value).toContain("\"finish\": \"2026-03-26T18:00:00\"");
    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON を読み込んで 3 件の変更を反映しました");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Patch JSON 反映結果");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Patch JSON の部分適用結果です");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("Start");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("Finish");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("Name");
  });

  it("reports patch json warnings for unsupported operations", async () => {
    bootPage();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "rename_task",
          uid: "3"
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON に反映対象の変更はありませんでした");
    expect(document.getElementById("statusMessage").textContent).toContain("warning を無視しました");
    expect(document.getElementById("importWarnings").textContent).toContain("Patch JSON warning");
    expect(document.getElementById("importWarnings").textContent).toContain("未対応の op は無視します");
    expect(document.getElementById("importWarnings").textContent).toContain("operations[0].op = rename_task");
  });

  it("imports patch json add_task into the current model and xml", async () => {
    bootPage();

    document.getElementById("xmlInput").value = hierarchyXml;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "add_task",
          uid: "4",
          name: "Child C",
          new_parent_uid: "1",
          new_index: 1,
          planned_start: "2026-03-17",
          planned_finish: "2026-03-17"
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON を読み込んで 5 件の変更を反映しました");
    expect(document.getElementById("xmlInput").value).toContain("<UID>4</UID>");
    expect(document.getElementById("xmlInput").value).toContain("<Name>Child C</Name>");
    expect(document.getElementById("xmlInput").value).toContain("<OutlineLevel>2</OutlineLevel>");
    expect(document.getElementById("xmlInput").value).toContain("<OutlineNumber>1.2</OutlineNumber>");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("ParentUID");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("Position");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Child C");
  });

  it("rejects add_task when planned_start is after planned_finish", async () => {
    bootPage();

    document.getElementById("xmlInput").value = hierarchyXml;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "add_task",
          uid: "4",
          name: "Broken Task",
          new_parent_uid: "1",
          new_index: 1,
          planned_start: "2026-03-18",
          planned_finish: "2026-03-17"
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON に反映対象の変更はありませんでした");
    expect(document.getElementById("importWarnings").textContent).toContain("add_task.planned_start が planned_finish より後です");
    expect(document.getElementById("xmlInput").value).not.toContain("<UID>4</UID>");
  });

  it("imports summary add_task into the current model and xml", async () => {
    bootPage();

    document.getElementById("xmlInput").value = hierarchyXml;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "add_task",
          uid: "4",
          name: "New Summary",
          is_summary: true,
          new_parent_uid: null,
          new_index: 1
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON を読み込んで 3 件の変更を反映しました");
    expect(document.getElementById("xmlInput").value).toContain("<UID>4</UID>");
    expect(document.getElementById("xmlInput").value).toContain("<Name>New Summary</Name>");
    expect(document.getElementById("xmlInput").value).toContain("<Summary>1</Summary>");
    expect(document.getElementById("xmlInput").value).toContain("<OutlineLevel>1</OutlineLevel>");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("New Summary");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("ParentUID");
  });

  it("rejects add_task when is_summary and is_milestone are both true", async () => {
    bootPage();

    document.getElementById("xmlInput").value = hierarchyXml;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "add_task",
          uid: "4",
          name: "Broken Summary Gate",
          is_summary: true,
          is_milestone: true,
          new_parent_uid: null,
          new_index: 1
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON に反映対象の変更はありませんでした");
    expect(document.getElementById("importWarnings").textContent).toContain("add_task では is_summary と is_milestone を同時に true にできません");
    expect(document.getElementById("xmlInput").value).not.toContain("<UID>4</UID>");
  });

  it("normalizes milestone add_task finish and duration", async () => {
    bootPage();

    document.getElementById("xmlInput").value = hierarchyXml;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "add_task",
          uid: "4",
          name: "Gate",
          new_parent_uid: "1",
          new_index: 1,
          is_milestone: true,
          planned_start: "2026-03-17",
          planned_finish: "2026-03-18",
          planned_duration_hours: 8,
          extra_key: "ignored"
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("warning を無視しました");
    expect(document.getElementById("xmlInput").value).toContain("<UID>4</UID>");
    expect(document.getElementById("xmlInput").value).toContain("<Milestone>1</Milestone>");
    expect(document.getElementById("xmlInput").value).toContain("<Finish>2026-03-17T09:00:00</Finish>");
    expect(document.getElementById("xmlInput").value).toContain("<Duration>PT0H0M0S</Duration>");
    expect(document.getElementById("importWarnings").textContent).toContain("add_task.is_milestone=true のため planned_finish は planned_start に揃えます");
    expect(document.getElementById("importWarnings").textContent).toContain("add_task.is_milestone=true のため planned_duration は 0 に揃えます");
    expect(document.getElementById("importWarnings").textContent).toContain("add_task の未対応 key は無視します: extra_key");
  });

  it("imports patch json move_task into the current model and xml", async () => {
    bootPage();

    document.getElementById("xmlInput").value = hierarchyXml;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "move_task",
          uid: "3",
          new_parent_uid: null,
          new_index: 1
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON を読み込んで 1 件の変更を反映しました");
    expect(document.getElementById("xmlInput").value).toContain("<UID>3</UID>");
    expect(document.getElementById("xmlInput").value).toContain("<OutlineLevel>1</OutlineLevel>");
    expect(document.getElementById("xmlInput").value).toContain("<OutlineNumber>2</OutlineNumber>");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("ParentUID");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("(root)");
  });

  it("imports patch json delete_task into the current model and xml", async () => {
    bootPage();

    document.getElementById("xmlInput").value = hierarchyXml;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "delete_task",
          uid: "3"
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON を読み込んで 3 件の変更を反映しました");
    expect(document.getElementById("xmlInput").value).not.toContain("<UID>3</UID>");
    expect(document.getElementById("xmlInput").value).not.toContain("<Name>Child B</Name>");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("ParentUID");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("Position");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("(deleted)");
  });

  it("rejects delete_task for summary task in first cut", async () => {
    bootPage();

    document.getElementById("xmlInput").value = hierarchyXml;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "delete_task",
          uid: "1"
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON に反映対象の変更はありませんでした");
    expect(document.getElementById("importWarnings").textContent).toContain("delete_task first cut では summary task や子を持つ task は削除できません");
    expect(document.getElementById("importWarnings").textContent).toContain("children=2");
    expect(document.getElementById("xmlInput").value).toContain("<UID>1</UID>");
  });

  it("rejects delete_task when assignments still reference the task", async () => {
    bootPage();

    const xmlWithAssignment = hierarchyXml
      .replace("<Resources />", [
        "<Resources>",
        "  <Resource>",
        "    <UID>1</UID>",
        "    <ID>1</ID>",
        "    <Name>Owner</Name>",
        "  </Resource>",
        "</Resources>"
      ].join("\n"))
      .replace("<Assignments />", [
        "<Assignments>",
        "  <Assignment>",
        "    <UID>1</UID>",
        "    <TaskUID>3</TaskUID>",
        "    <ResourceUID>1</ResourceUID>",
        "  </Assignment>",
        "</Assignments>"
      ].join("\n"));
    document.getElementById("xmlInput").value = xmlWithAssignment;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "delete_task",
          uid: "3"
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON に反映対象の変更はありませんでした");
    expect(document.getElementById("importWarnings").textContent).toContain("delete_task first cut では assignment がある task は削除できません");
    expect(document.getElementById("importWarnings").textContent).toContain("assignments=1");
    expect(document.getElementById("xmlInput").value).toContain("<UID>3</UID>");
  });

  it("rejects delete_task when successors still reference the task", async () => {
    bootPage();

    const xmlWithDependency = hierarchyXml.replace(
      "</Notes>\n    </Task>",
      [
        "</Notes>",
        "      <PredecessorLink>",
        "        <PredecessorUID>2</PredecessorUID>",
        "        <Type>1</Type>",
        "      </PredecessorLink>",
        "    </Task>"
      ].join("\n")
    );
    document.getElementById("xmlInput").value = xmlWithDependency;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "delete_task",
          uid: "2"
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON に反映対象の変更はありませんでした");
    expect(document.getElementById("importWarnings").textContent).toContain("delete_task first cut では後続依存がある task は削除できません");
    expect(document.getElementById("importWarnings").textContent).toContain("successors=3");
    expect(document.getElementById("xmlInput").value).toContain("<UID>2</UID>");
  });

  it("ignores no-op patch json move_task", async () => {
    bootPage();

    document.getElementById("xmlInput").value = hierarchyXml;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "move_task",
          uid: "3",
          new_parent_uid: "1",
          new_index: 1
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON に反映対象の変更はありませんでした");
    expect(document.getElementById("statusMessage").textContent).toContain("warning を無視しました");
    expect(document.getElementById("importWarnings").textContent).toContain("move_task は結果が変わらないため無視します");
    expect(document.getElementById("importWarnings").textContent).toContain("parent=1 index=1");
  });

  it("imports patch json link_tasks into the current model and xml", async () => {
    bootPage();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "link_tasks",
          from_uid: "2",
          to_uid: "3",
          type: "SS",
          lag_hours: 4
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON を読み込んで 1 件の変更を反映しました");
    expect(document.getElementById("xmlInput").value).toContain("<PredecessorUID>2</PredecessorUID>");
    expect(document.getElementById("xmlInput").value).toContain("<Type>2</Type>");
    expect(document.getElementById("xmlInput").value).toContain("<LinkLag>PT4H0M0S</LinkLag>");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Predecessors");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("UID=3");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("2(SS, lag=PT4H0M0S)");
  });

  it("imports patch json unlink_tasks into the current model and xml", async () => {
    bootPage();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "link_tasks",
          from_uid: "2",
          to_uid: "3",
          type: "SS",
          lag_hours: 4
        }
      ]
    }, null, 2);
    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "unlink_tasks",
          from_uid: "2",
          to_uid: "3",
          type: "SS"
        }
      ]
    }, null, 2);
    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON を読み込んで 1 件の変更を反映しました");
    expect(document.getElementById("xmlInput").value).not.toContain("<PredecessorUID>2</PredecessorUID>");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Predecessors");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("2(SS, lag=PT4H0M0S)");
  });

  it("imports patch json unlink_tasks with lag filter into the current model and xml", async () => {
    bootPage();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "link_tasks",
          from_uid: "2",
          to_uid: "3",
          type: "SS",
          lag_hours: 4
        }
      ]
    }, null, 2);
    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "unlink_tasks",
          from_uid: "2",
          to_uid: "3",
          type: "SS",
          lag: "PT4H0M0S"
        }
      ]
    }, null, 2);
    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON を読み込んで 1 件の変更を反映しました");
    expect(document.getElementById("xmlInput").value).not.toContain("<PredecessorUID>2</PredecessorUID>");
  });

  it("reports link_tasks warnings with type and lag details", async () => {
    bootPage();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "link_tasks",
          from_uid: "2",
          to_uid: "3",
          type: "SS",
          lag_hours: 4
        },
        {
          op: "link_tasks",
          from_uid: "2",
          to_uid: "3",
          type: "SS",
          lag_hours: 4
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("warning を無視しました");
    expect(document.getElementById("importWarnings").textContent).toContain("link_tasks の依存関係は既に存在します");
    expect(document.getElementById("importWarnings").textContent).toContain("2 -> 3 (SS, lag=PT4H0M0S)");
  });

  it("reports unlink_tasks warnings with requested type details", async () => {
    bootPage();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "unlink_tasks",
          from_uid: "2",
          to_uid: "3",
          type: "FF"
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("warning を無視しました");
    expect(document.getElementById("importWarnings").textContent).toContain("unlink_tasks の対象依存関係が見つかりません");
    expect(document.getElementById("importWarnings").textContent).toContain("2 -> 3 (FF)");
  });

  it("reports unlink_tasks warnings with requested lag details", async () => {
    bootPage();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "link_tasks",
          from_uid: "2",
          to_uid: "3",
          type: "SS",
          lag_hours: 4
        }
      ]
    }, null, 2);
    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "unlink_tasks",
          from_uid: "2",
          to_uid: "3",
          type: "SS",
          lag_hours: 8
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("warning を無視しました");
    expect(document.getElementById("importWarnings").textContent).toContain("unlink_tasks の対象依存関係が見つかりません");
    expect(document.getElementById("importWarnings").textContent).toContain("2 -> 3 (SS, lag=PT8H0M0S)");
  });

  it("reports when unlink_tasks removes multiple matching dependencies", async () => {
    bootPage();

    const xmlWithDuplicateLinks = hierarchyXml.replace(
      "</Notes>\n    </Task>",
      [
        "</Notes>",
        "      <PredecessorLink>",
        "        <PredecessorUID>2</PredecessorUID>",
        "        <Type>1</Type>",
        "        <LinkLag>PT0H0M0S</LinkLag>",
        "      </PredecessorLink>",
        "      <PredecessorLink>",
        "        <PredecessorUID>2</PredecessorUID>",
        "        <Type>1</Type>",
        "        <LinkLag>PT0H0M0S</LinkLag>",
        "      </PredecessorLink>",
        "    </Task>"
      ].join("\n")
    );
    document.getElementById("xmlInput").value = xmlWithDuplicateLinks;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "unlink_tasks",
          from_uid: "2",
          to_uid: "3",
          type: "FS",
          lag: "PT0H0M0S"
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("warning を無視しました");
    expect(document.getElementById("importWarnings").textContent).toContain("一致した依存関係 2 件をすべて解除しました");
    expect(document.getElementById("importWarnings").textContent).toContain("2 -> 3 (FS)");
    expect(document.getElementById("xmlInput").value).not.toContain("<PredecessorUID>2</PredecessorUID>");
  });

  it("rejects link_tasks that would create a dependency cycle", async () => {
    bootPage();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "link_tasks",
          from_uid: "2",
          to_uid: "3",
          type: "FS"
        },
        {
          op: "link_tasks",
          from_uid: "3",
          to_uid: "2",
          type: "FS"
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("warning を無視しました");
    expect(document.getElementById("importWarnings").textContent).toContain("link_tasks で循環依存になるため無視します");
    expect(document.getElementById("importWarnings").textContent).toContain("3 -> 2 (FS)");
    expect(document.getElementById("xmlInput").value).toContain("<PredecessorUID>2</PredecessorUID>");
    expect(document.getElementById("xmlInput").value).not.toContain("<PredecessorUID>3</PredecessorUID>");
  });

  it("rejects invalid link_tasks lag text", async () => {
    bootPage();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "link_tasks",
          from_uid: "2",
          to_uid: "3",
          type: "FS",
          lag: "4hours"
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("warning を無視しました");
    expect(document.getElementById("importWarnings").textContent).toContain("link_tasks.lag は ISO 8601 duration 形式が必要です");
    expect(document.getElementById("xmlInput").value).toContain("<PredecessorUID>2</PredecessorUID>");
    expect(document.getElementById("xmlInput").value).not.toContain("<LinkLag>4hours</LinkLag>");
  });

  it("imports patch json update_task is_milestone into the current model and xml", async () => {
    bootPage();

    document.getElementById("xmlInput").value = hierarchyXml;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "update_task",
          uid: "3",
          fields: {
            is_milestone: true
          }
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON を読み込んで 3 件の変更を反映しました");
    expect(document.getElementById("xmlInput").value).toContain("<UID>3</UID>");
    expect(document.getElementById("xmlInput").value).toContain("<Milestone>1</Milestone>");
    expect(document.getElementById("xmlInput").value).toContain("<Finish>2026-03-17T09:00:00</Finish>");
    expect(document.getElementById("xmlInput").value).toContain("<Duration>PT0H0M0S</Duration>");
    expect(document.getElementById("importWarnings").textContent).toContain("update_task.is_milestone=true のため planned_finish は planned_start に揃えます");
    expect(document.getElementById("importWarnings").textContent).toContain("update_task.is_milestone=true のため planned_duration は 0 に揃えます");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("Milestone");
  });

  it("rejects update_task is_milestone on summary task", async () => {
    bootPage();

    document.getElementById("xmlInput").value = hierarchyXml;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "update_task",
          uid: "1",
          fields: {
            is_milestone: true
          }
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON に反映対象の変更はありませんでした");
    expect(document.getElementById("importWarnings").textContent).toContain("update_task では summary task を milestone にできません");
    expect(document.getElementById("xmlInput").value).toContain("<UID>1</UID>");
    expect(document.getElementById("xmlInput").value).toContain("<Milestone>0</Milestone>");
  });

  it("imports patch json update_task notes into the current model and xml", async () => {
    bootPage();

    document.getElementById("xmlInput").value = hierarchyXml;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "update_task",
          uid: "3",
          fields: {
            notes: "Patched note"
          }
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON を読み込んで 1 件の変更を反映しました");
    expect(document.getElementById("xmlInput").value).toContain("<Notes>Patched note</Notes>");
    expect(document.getElementById("modelOutput").value).toContain("\"notes\": \"Patched note\"");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("Notes");
  });

  it("clears task notes with patch json update_task notes empty string", async () => {
    bootPage();

    document.getElementById("xmlInput").value = hierarchyXml;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "update_task",
          uid: "3",
          fields: {
            notes: ""
          }
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON を読み込んで 1 件の変更を反映しました");
    expect(document.getElementById("xmlInput").value).not.toContain("<Notes>Second child task</Notes>");
    expect(document.getElementById("modelOutput").value).not.toContain("\"notes\": \"Second child task\"");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("Notes");
  });

  it("imports patch json update_task calendar_uid into the current model and xml", async () => {
    bootPage();

    document.getElementById("xmlInput").value = dependencyXml;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "update_task",
          uid: "2",
          fields: {
            calendar_uid: "1"
          }
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON を読み込んで 1 件の変更を反映しました");
    expect(document.getElementById("xmlInput").value).toContain("<UID>2</UID>");
    expect(document.getElementById("xmlInput").value).toContain("<CalendarUID>1</CalendarUID>");
    expect(document.getElementById("modelOutput").value).toContain("\"calendarUID\": \"1\"");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("CalendarUID");
  });

  it("clears task calendar_uid with patch json update_task calendar_uid empty string", async () => {
    bootPage();

    const xmlWithTaskCalendar = dependencyXml.replace(
      "<PercentComplete>0</PercentComplete>",
      "<PercentComplete>0</PercentComplete>\n      <CalendarUID>1</CalendarUID>"
    );
    document.getElementById("xmlInput").value = xmlWithTaskCalendar;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "update_task",
          uid: "2",
          fields: {
            calendar_uid: ""
          }
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON を読み込んで 1 件の変更を反映しました");
    expect((document.getElementById("xmlInput").value.match(/<CalendarUID>1<\/CalendarUID>/g) || []).length).toBe(1);
    expect(JSON.parse(document.getElementById("modelOutput").value).tasks.find((task) => task.uid === "2").calendarUID).toBeUndefined();
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("CalendarUID");
  });

  it("rejects unknown patch json update_task calendar_uid", async () => {
    bootPage();

    document.getElementById("xmlInput").value = dependencyXml;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "update_task",
          uid: "2",
          fields: {
            calendar_uid: "999"
          }
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON に反映対象の変更はありませんでした");
    expect(document.getElementById("importWarnings").textContent).toContain("update_task.calendar_uid が既存 calendar を指していません");
    expect(JSON.parse(document.getElementById("modelOutput").value).tasks.find((task) => task.uid === "2").calendarUID).toBeUndefined();
  });

  it("imports patch json update_task progress fields into the current model and xml", async () => {
    bootPage();

    document.getElementById("xmlInput").value = hierarchyXml;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "update_task",
          uid: "3",
          fields: {
            percent_complete: 40,
            percent_work_complete: 60
          }
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON を読み込んで 2 件の変更を反映しました");
    expect(document.getElementById("xmlInput").value).toContain("<PercentComplete>40</PercentComplete>");
    expect(document.getElementById("xmlInput").value).toContain("<PercentWorkComplete>60</PercentWorkComplete>");
    expect(document.getElementById("modelOutput").value).toContain("\"percentComplete\": 40");
    expect(document.getElementById("modelOutput").value).toContain("\"percentWorkComplete\": 60");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("PercentComplete");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("PercentWorkComplete");
  });

  it("rejects out-of-range patch json progress fields", async () => {
    bootPage();

    document.getElementById("xmlInput").value = hierarchyXml;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "update_task",
          uid: "3",
          fields: {
            percent_complete: 120,
            percent_work_complete: -1
          }
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON に反映対象の変更はありませんでした");
    expect(document.getElementById("importWarnings").textContent).toContain("update_task.percent_complete は 0 以上 100 以下の数値が必要です");
    expect(document.getElementById("importWarnings").textContent).toContain("update_task.percent_work_complete は 0 以上 100 以下の数値が必要です");
  });

  it("imports patch json update_task critical into the current model and xml", async () => {
    bootPage();

    document.getElementById("xmlInput").value = hierarchyXml;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "update_task",
          uid: "3",
          fields: {
            critical: true
          }
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON を読み込んで 1 件の変更を反映しました");
    expect(document.getElementById("xmlInput").value).toContain("<Critical>1</Critical>");
    expect(document.getElementById("modelOutput").value).toContain("\"critical\": true");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("Critical");
  });

  it("rejects invalid patch json critical type", async () => {
    bootPage();

    document.getElementById("xmlInput").value = hierarchyXml;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "update_task",
          uid: "3",
          fields: {
            critical: "yes"
          }
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON に反映対象の変更はありませんでした");
    expect(document.getElementById("importWarnings").textContent).toContain("update_task.critical は boolean が必要です");
  });

  it("imports patch json update_project into the current model and xml", async () => {
    bootPage();

    const xmlWithExtraCalendar = dependencyXml.replace(
      "</Calendars>",
      `  <Calendar>\n      <UID>2</UID>\n      <Name>Alt</Name>\n      <IsBaseCalendar>0</IsBaseCalendar>\n      <BaseCalendarUID>1</BaseCalendarUID>\n    </Calendar>\n  </Calendars>`
    );
    document.getElementById("xmlInput").value = xmlWithExtraCalendar;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "update_project",
          fields: {
            name: "Dependency Project Prime",
            title: "Prime Title",
            author: "Prime Author",
            company: "Prime Company",
            start_date: "2026-03-15",
            finish_date: "2026-03-20",
            current_date: "2026-03-17",
            status_date: "2026-03-18",
            calendar_uid: "2",
            minutes_per_day: 420,
            minutes_per_week: 2100,
            days_per_month: 18,
            schedule_from_start: false
          }
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON を読み込んで 13 件の変更を反映しました");
    expect(document.getElementById("xmlInput").value).toContain("<Name>Dependency Project Prime</Name>");
    expect(document.getElementById("xmlInput").value).toContain("<Title>Prime Title</Title>");
    expect(document.getElementById("xmlInput").value).toContain("<Author>Prime Author</Author>");
    expect(document.getElementById("xmlInput").value).toContain("<Company>Prime Company</Company>");
    expect(document.getElementById("xmlInput").value).toContain("<StartDate>2026-03-15T09:00:00</StartDate>");
    expect(document.getElementById("xmlInput").value).toContain("<FinishDate>2026-03-20T18:00:00</FinishDate>");
    expect(document.getElementById("xmlInput").value).toContain("<CurrentDate>2026-03-17T09:00:00</CurrentDate>");
    expect(document.getElementById("xmlInput").value).toContain("<StatusDate>2026-03-18T09:00:00</StatusDate>");
    expect(document.getElementById("xmlInput").value).toContain("<CalendarUID>2</CalendarUID>");
    expect(document.getElementById("xmlInput").value).toContain("<MinutesPerDay>420</MinutesPerDay>");
    expect(document.getElementById("xmlInput").value).toContain("<MinutesPerWeek>2100</MinutesPerWeek>");
    expect(document.getElementById("xmlInput").value).toContain("<DaysPerMonth>18</DaysPerMonth>");
    expect(document.getElementById("xmlInput").value).toContain("<ScheduleFromStart>0</ScheduleFromStart>");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Project");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("Title");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("StartDate");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("MinutesPerDay");
  });

  it("rejects invalid patch json update_project fields", async () => {
    bootPage();

    document.getElementById("xmlInput").value = dependencyXml;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
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
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON に反映対象の変更はありませんでした");
    expect(document.getElementById("importWarnings").textContent).toContain("update_project.name は空でない文字列が必要です");
    expect(document.getElementById("importWarnings").textContent).toContain("update_project.start_date が finish_date より後です");
    expect(document.getElementById("importWarnings").textContent).toContain("update_project.current_date の日付形式が解釈できません");
    expect(document.getElementById("importWarnings").textContent).toContain("update_project.calendar_uid が既存 calendar を指していません");
    expect(document.getElementById("importWarnings").textContent).toContain("update_project.minutes_per_day は 0 より大きい数値が必要です");
    expect(document.getElementById("importWarnings").textContent).toContain("update_project.schedule_from_start は boolean が必要です");
  });

  it("imports patch json update_resource into the current model and xml", async () => {
    bootPage();

    document.getElementById("xmlInput").value = dependencyXml;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "update_resource",
          uid: "1",
          fields: {
            name: "Miku Prime",
            initials: "MP",
            group: "Platform",
            calendar_uid: "1",
            max_units: 0.75,
            standard_rate: "1000",
            overtime_rate: "1500",
            cost_per_use: 250,
            percent_work_complete: 60
          }
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON を読み込んで 9 件の変更を反映しました");
    expect(document.getElementById("xmlInput").value).toContain("<Name>Miku Prime</Name>");
    expect(document.getElementById("xmlInput").value).toContain("<Initials>MP</Initials>");
    expect(document.getElementById("xmlInput").value).toContain("<Group>Platform</Group>");
    expect(document.getElementById("xmlInput").value).toContain("<CalendarUID>1</CalendarUID>");
    expect(document.getElementById("xmlInput").value).toContain("<MaxUnits>0.75</MaxUnits>");
    expect(document.getElementById("xmlInput").value).toContain("<StandardRate>1000</StandardRate>");
    expect(document.getElementById("xmlInput").value).toContain("<OvertimeRate>1500</OvertimeRate>");
    expect(document.getElementById("xmlInput").value).toContain("<CostPerUse>250</CostPerUse>");
    expect(document.getElementById("xmlInput").value).toContain("<PercentWorkComplete>60</PercentWorkComplete>");
    expect(document.getElementById("modelOutput").value).toContain("\"initials\": \"MP\"");
    expect(document.getElementById("modelOutput").value).toContain("\"maxUnits\": 0.75");
    expect(document.getElementById("modelOutput").value).toContain("\"costPerUse\": 250");
    expect(document.getElementById("modelOutput").value).toContain("\"percentWorkComplete\": 60");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Resources");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("Initials");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("MaxUnits");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("StandardRate");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("CostPerUse");
  });

  it("rejects invalid patch json update_resource fields", async () => {
    bootPage();

    document.getElementById("xmlInput").value = dependencyXml;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "update_resource",
          uid: "1",
          fields: {
            name: "",
            calendar_uid: "999",
            max_units: -1,
            cost_per_use: -1,
            percent_work_complete: 120
          }
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON に反映対象の変更はありませんでした");
    expect(document.getElementById("importWarnings").textContent).toContain("update_resource.name は空でない文字列が必要です");
    expect(document.getElementById("importWarnings").textContent).toContain("update_resource.calendar_uid が既存 calendar を指していません");
    expect(document.getElementById("importWarnings").textContent).toContain("update_resource.max_units は 0 以上の数値が必要です");
    expect(document.getElementById("importWarnings").textContent).toContain("update_resource.cost_per_use は 0 以上の数値が必要です");
    expect(document.getElementById("importWarnings").textContent).toContain("update_resource.percent_work_complete は 0 以上 100 以下の数値が必要です");
  });

  it("imports patch json update_calendar into the current model and xml", async () => {
    bootPage();

    const xmlWithExtraCalendar = dependencyXml.replace(
      "</Calendars>",
      `  <Calendar>\n      <UID>2</UID>\n      <Name>Alt</Name>\n      <IsBaseCalendar>0</IsBaseCalendar>\n      <BaseCalendarUID>1</BaseCalendarUID>\n      <WeekDays />\n    </Calendar>\n  </Calendars>`
    );
    document.getElementById("xmlInput").value = xmlWithExtraCalendar;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "update_calendar",
          uid: "2",
          fields: {
            name: "Alt Prime",
            is_base_calendar: true,
            base_calendar_uid: ""
          }
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON を読み込んで 3 件の変更を反映しました");
    expect(document.getElementById("xmlInput").value).toContain("<UID>2</UID>");
    expect(document.getElementById("xmlInput").value).toContain("<Name>Alt Prime</Name>");
    expect(document.getElementById("xmlInput").value).toContain("<IsBaseCalendar>1</IsBaseCalendar>");
    expect(document.getElementById("modelOutput").value).toContain("\"isBaseCalendar\": true");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Calendars");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("IsBaseCalendar");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("BaseCalendarUID");
  });

  it("rejects invalid patch json update_calendar fields", async () => {
    bootPage();

    document.getElementById("xmlInput").value = dependencyXml;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "update_calendar",
          uid: "1",
          fields: {
            name: "",
            is_base_calendar: "yes",
            base_calendar_uid: "1"
          }
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON に反映対象の変更はありませんでした");
    expect(document.getElementById("importWarnings").textContent).toContain("update_calendar.name は空でない文字列が必要です");
    expect(document.getElementById("importWarnings").textContent).toContain("update_calendar.is_base_calendar は boolean が必要です");
    expect(document.getElementById("importWarnings").textContent).toContain("update_calendar.base_calendar_uid は自身を指せません");
  });

  it("imports patch json add_calendar into the current model and xml", async () => {
    bootPage();

    document.getElementById("xmlInput").value = dependencyXml;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "add_calendar",
          uid: "2",
          name: "Night Shift",
          is_base_calendar: false,
          base_calendar_uid: "1"
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON を読み込んで 2 件の変更を反映しました");
    expect(document.getElementById("xmlInput").value).toContain("<UID>2</UID>");
    expect(document.getElementById("xmlInput").value).toContain("<Name>Night Shift</Name>");
    expect(document.getElementById("xmlInput").value).toContain("<BaseCalendarUID>1</BaseCalendarUID>");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Calendars");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("BaseCalendarUID");
  });

  it("rejects invalid patch json add_calendar fields", async () => {
    bootPage();

    document.getElementById("xmlInput").value = dependencyXml;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "add_calendar",
          uid: "2",
          name: "",
          is_base_calendar: "yes",
          base_calendar_uid: "999"
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON に反映対象の変更はありませんでした");
    expect(document.getElementById("importWarnings").textContent).toContain("add_calendar.name は空でない文字列が必要です");
  });

  it("imports patch json delete_calendar into the current model and xml", async () => {
    bootPage();

    const xmlWithExtraCalendar = dependencyXml.replace(
      "</Calendars>",
      `  <Calendar>\n      <UID>2</UID>\n      <Name>Alt</Name>\n      <IsBaseCalendar>0</IsBaseCalendar>\n      <WeekDays />\n    </Calendar>\n  </Calendars>`
    );
    document.getElementById("xmlInput").value = xmlWithExtraCalendar;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "delete_calendar",
          uid: "2"
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON を読み込んで 1 件の変更を反映しました");
    expect(document.getElementById("xmlInput").value).not.toContain("<Name>Alt</Name>");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("(deleted)");
  });

  it("rejects delete_calendar when references still exist", async () => {
    bootPage();

    document.getElementById("xmlInput").value = dependencyXml;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "delete_calendar",
          uid: "1"
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON に反映対象の変更はありませんでした");
    expect(document.getElementById("importWarnings").textContent).toContain("delete_calendar first cut では参照が残っている calendar は削除できません");
  });

  it("imports patch json add_resource into the current model and xml", async () => {
    bootPage();

    document.getElementById("xmlInput").value = dependencyXml;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "add_resource",
          uid: "2",
          name: "Mikuku",
          initials: "M",
          group: "PMO",
          calendar_uid: "1",
          max_units: 1,
          standard_rate: "1200",
          overtime_rate: "1800",
          cost_per_use: 300,
          percent_work_complete: 25
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON を読み込んで 9 件の変更を反映しました");
    expect(document.getElementById("xmlInput").value).toContain("<UID>2</UID>");
    expect(document.getElementById("xmlInput").value).toContain("<Name>Mikuku</Name>");
    expect(document.getElementById("xmlInput").value).toContain("<Initials>M</Initials>");
    expect(document.getElementById("xmlInput").value).toContain("<Group>PMO</Group>");
    expect(document.getElementById("xmlInput").value).toContain("<CalendarUID>1</CalendarUID>");
    expect(document.getElementById("xmlInput").value).toContain("<MaxUnits>1</MaxUnits>");
    expect(document.getElementById("xmlInput").value).toContain("<StandardRate>1200</StandardRate>");
    expect(document.getElementById("xmlInput").value).toContain("<OvertimeRate>1800</OvertimeRate>");
    expect(document.getElementById("xmlInput").value).toContain("<CostPerUse>300</CostPerUse>");
    expect(document.getElementById("xmlInput").value).toContain("<PercentWorkComplete>25</PercentWorkComplete>");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Resources");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("Initials");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("StandardRate");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("CostPerUse");
  });

  it("rejects invalid patch json add_resource fields", async () => {
    bootPage();

    document.getElementById("xmlInput").value = dependencyXml;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "add_resource",
          uid: "2",
          name: "",
          calendar_uid: "999",
          max_units: -1,
          cost_per_use: -1,
          percent_work_complete: 120
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON に反映対象の変更はありませんでした");
    expect(document.getElementById("importWarnings").textContent).toContain("add_resource.name は空でない文字列が必要です");
  });

  it("imports patch json delete_resource into the current model and xml", async () => {
    bootPage();

    const xmlWithoutAssignments = dependencyXml.replace(/<Assignments>[\s\S]*<\/Assignments>/, "<Assignments />");
    document.getElementById("xmlInput").value = xmlWithoutAssignments;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "delete_resource",
          uid: "1"
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON を読み込んで 1 件の変更を反映しました");
    expect(document.getElementById("xmlInput").value).not.toContain("<Name>Miku</Name>");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("(deleted)");
  });

  it("rejects delete_resource when assignments still reference the resource", async () => {
    bootPage();

    document.getElementById("xmlInput").value = dependencyXml;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "delete_resource",
          uid: "1"
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON に反映対象の変更はありませんでした");
    expect(document.getElementById("importWarnings").textContent).toContain("delete_resource first cut では assignment がある resource は削除できません");
  });

  it("imports patch json update_assignment into the current model and xml", async () => {
    bootPage();

    document.getElementById("xmlInput").value = dependencyXml;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "update_assignment",
          uid: "1",
          fields: {
            start: "2026-03-17",
            finish: "2026-03-18",
            units: 0.5,
            work: "PT12H0M0S",
            percent_work_complete: 75
          }
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON を読み込んで 5 件の変更を反映しました");
    expect(document.getElementById("xmlInput").value).toContain("<Assignment>");
    expect(document.getElementById("xmlInput").value).toContain("<UID>1</UID>");
    expect(document.getElementById("xmlInput").value).toContain("<Start>2026-03-17T09:00:00</Start>");
    expect(document.getElementById("xmlInput").value).toContain("<Finish>2026-03-18T18:00:00</Finish>");
    expect(document.getElementById("xmlInput").value).toContain("<Units>0.5</Units>");
    expect(document.getElementById("xmlInput").value).toContain("<Work>PT12H0M0S</Work>");
    expect(document.getElementById("xmlInput").value).toContain("<PercentWorkComplete>75</PercentWorkComplete>");
    expect(document.getElementById("modelOutput").value).toContain("\"units\": 0.5");
    expect(document.getElementById("modelOutput").value).toContain("\"percentWorkComplete\": 75");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Assignments");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("Units");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("Work");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("PercentWorkComplete");
  });

  it("rejects invalid patch json update_assignment range and values", async () => {
    bootPage();

    document.getElementById("xmlInput").value = dependencyXml;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "update_assignment",
          uid: "1",
          fields: {
            start: "2026-03-20",
            finish: "2026-03-19",
            units: -1,
            work: "",
            percent_work_complete: 120
          }
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON に反映対象の変更はありませんでした");
    expect(document.getElementById("importWarnings").textContent).toContain("update_assignment.start が finish より後です");
    expect(document.getElementById("importWarnings").textContent).toContain("update_assignment.units は 0 以上の数値が必要です");
    expect(document.getElementById("importWarnings").textContent).toContain("update_assignment.work は空でない文字列が必要です");
    expect(document.getElementById("importWarnings").textContent).toContain("update_assignment.percent_work_complete は 0 以上 100 以下の数値が必要です");
    expect(document.getElementById("xmlInput").value).toContain("<Start>2026-03-16T09:00:00</Start>");
    expect(document.getElementById("xmlInput").value).toContain("<Finish>2026-03-17T18:00:00</Finish>");
  });

  it("imports patch json add_assignment into the current model and xml", async () => {
    bootPage();

    document.getElementById("xmlInput").value = dependencyXml;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "add_assignment",
          uid: "3",
          task_uid: "2",
          resource_uid: "1",
          start: "2026-03-19",
          finish: "2026-03-19",
          units: 0.25,
          work: "PT2H0M0S",
          percent_work_complete: 10
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON を読み込んで 7 件の変更を反映しました");
    expect(document.getElementById("xmlInput").value).toContain("<UID>3</UID>");
    expect(document.getElementById("xmlInput").value).toContain("<TaskUID>2</TaskUID>");
    expect(document.getElementById("xmlInput").value).toContain("<ResourceUID>1</ResourceUID>");
    expect(document.getElementById("xmlInput").value).toContain("<Units>0.25</Units>");
    expect(document.getElementById("xmlInput").value).toContain("<Work>PT2H0M0S</Work>");
    expect(document.getElementById("xmlInput").value).toContain("<PercentWorkComplete>10</PercentWorkComplete>");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Assignments");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("TaskUID");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("ResourceUID");
  });

  it("rejects invalid patch json add_assignment references and values", async () => {
    bootPage();

    document.getElementById("xmlInput").value = dependencyXml;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "add_assignment",
          uid: "3",
          task_uid: "999",
          resource_uid: "1",
          units: -1
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON に反映対象の変更はありませんでした");
    expect(document.getElementById("importWarnings").textContent).toContain("add_assignment.task_uid が既存 task を指していません");
    expect(document.getElementById("xmlInput").value).not.toContain("<UID>3</UID>");
  });

  it("imports patch json delete_assignment into the current model and xml", async () => {
    bootPage();

    document.getElementById("xmlInput").value = dependencyXml;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "delete_assignment",
          uid: "1"
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON を読み込んで 2 件の変更を反映しました");
    expect(document.getElementById("xmlInput").value).not.toContain("<Assignment>");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Assignments");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("TaskUID");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("ResourceUID");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("(deleted)");
  });

  it("rejects missing patch json delete_assignment target", async () => {
    bootPage();

    document.getElementById("xmlInput").value = dependencyXml;
    parseXmlViaHook();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "delete_assignment",
          uid: "999"
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON に反映対象の変更はありませんでした");
    expect(document.getElementById("importWarnings").textContent).toContain("delete_assignment の uid が既存 assignment を指していません");
    expect(document.getElementById("xmlInput").value).toContain("<Assignment>");
  });

  it("groups patch json task warnings by uid", async () => {
    bootPage();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "update_task",
          uid: "3",
          fields: {
            unknown_field: "ignored"
          }
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("warning を無視しました");
    expect(document.getElementById("importWarnings").textContent).toContain("UID=3");
    expect(document.getElementById("importWarnings").textContent).toContain("初期実装");
    expect(document.getElementById("importWarnings").textContent).toContain("未対応の field は無視します");
    expect(document.getElementById("xlsxImportSummary").classList.contains("md-hidden")).toBe(true);
  });

  it("loads sample project_draft_view into the input area", () => {
    bootPage();

    document.getElementById("loadProjectDraftSampleBtn").click();

    const draftText = document.getElementById("projectDraftImportInput").value;
    expect(draftText).toContain("\"view_type\": \"project_draft_view\"");
    expect(draftText).toContain("\"name\": \"mikuproject開発\"");
    expect(draftText).toContain("架空検討フェーズ【架空】");
    expect(draftText).toContain("\"resources\"");
    expect(draftText).toContain("\"Mikuku\"");
    expect(draftText).toContain("\"initials\": \"M\"");
    expect(document.getElementById("statusMessage").textContent).toContain("サンプル project_draft_view");
  });

  it("copies ai prompt to clipboard", async () => {
    bootPage();

    document.getElementById("copyAiPromptBtn").click();
    await flushAsyncWork();

    expect(globalThis.navigator.clipboard.writeText.mock.calls.length).toBeGreaterThan(0);
    expect(globalThis.navigator.clipboard.writeText.mock.calls.at(-1)[0]).toContain("# mikuproject AI JSON Spec");
    expect(document.getElementById("statusMessage").textContent).toContain("生成AIプロンプトをクリップボードにコピーしました");
  });
});
