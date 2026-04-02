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
    <button id="exportPhaseDetailFullBtn" type="button"></button>
    <button id="exportPhaseDetailBtn" type="button"></button>
    <button id="loadProjectDraftSampleBtn" type="button"></button>
    <button id="importProjectDraftBtn" type="button"></button>
    <button id="downloadXmlBtn" type="button"></button>
    <button id="roundTripBtn" type="button"></button>
    <button id="copyAiPromptBtnPane" type="button"></button>
    <input id="importFileInput" type="file" />
    <input id="phaseDetailUidInput" type="text" />
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
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("planned_start");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("planned_finish");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("name");
  });

  it("reports patch json warnings for unsupported operations", async () => {
    bootPage();

    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [
        {
          op: "link_tasks",
          from_uid: "1",
          to_uid: "3",
          type: "FS"
        }
      ]
    }, null, 2);

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON に反映対象の変更はありませんでした");
    expect(document.getElementById("statusMessage").textContent).toContain("warning を無視しました");
    expect(document.getElementById("importWarnings").textContent).toContain("未対応の op は無視します");
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
