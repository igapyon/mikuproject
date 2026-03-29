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
const excelIoCode = readFileSync(
  path.resolve(__dirname, "../src/js/excel-io.js"),
  "utf8"
);
const msProjectXmlCode = readFileSync(
  path.resolve(__dirname, "../src/js/msproject-xml.js"),
  "utf8"
);
const projectXlsxCode = readFileSync(
  path.resolve(__dirname, "../src/js/project-xlsx.js"),
  "utf8"
);
const wbsXlsxCode = readFileSync(
  path.resolve(__dirname, "../src/js/wbs-xlsx.js"),
  "utf8"
);
const mainCode = readFileSync(
  path.resolve(__dirname, "../src/js/main.js"),
  "utf8"
);
const minimalXml = readFileSync(
  path.resolve(__dirname, "../testdata/minimal.xml"),
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

function mountDom() {
  document.body.innerHTML = `
    <button id="importXmlBtn" type="button">MS Project XML</button>
    <button id="importXlsxBtn" type="button">XLSX</button>
    <button id="parseCsvBtn" type="button">CSV</button>
    <button id="loadSampleBtn" type="button">サンプル</button>
    <button id="exportXlsxBtn" type="button">XLSX</button>
    <button id="exportWbsXlsxBtn" type="button">WBS XLSX</button>
    <button id="downloadMermaidSvgBtn" type="button" disabled>SVG</button>
    <button id="exportCsvBtn" type="button">CSV</button>
    <button id="exportAiBundleBtn" type="button">project_overview + all phase_detail full</button>
    <button id="exportProjectOverviewBtn" type="button">project_overview_view</button>
    <button id="exportPhaseDetailFullBtn" type="button">phase_detail_view full</button>
    <button id="exportPhaseDetailBtn" type="button">phase_detail_view</button>
    <button id="loadProjectDraftSampleBtn" type="button">サンプル draft</button>
    <button id="importProjectDraftFileBtn" type="button">project_draft_view JSON</button>
    <button id="importProjectDraftBtn" type="button">project_draft_view を取り込む</button>
    <button id="downloadXmlBtn" type="button">MS Project XML</button>
    <input id="importXmlInput" type="file" />
    <input id="importXlsxInput" type="file" />
    <input id="importCsvInput" type="file" />
    <input id="importProjectDraftInput" type="file" />
    <input id="phaseDetailUidInput" type="text" />
    <input id="phaseDetailRootUidInput" type="text" />
    <input id="phaseDetailMaxDepthInput" type="text" />
    <div id="statusMessage"></div>
    <div class="md-top-tabs" role="tablist" aria-label="mikuproject sections">
      <button type="button" class="md-top-tab is-active" data-tab="input" role="tab" aria-selected="true" aria-controls="tabPanelInput">
        <span class="md-top-tab-no">1</span>
        <span class="md-top-tab-label">Input</span>
      </button>
      <button type="button" class="md-top-tab" data-tab="transform" role="tab" aria-selected="false" aria-controls="tabPanelTransform">
        <span class="md-top-tab-no">2</span>
        <span class="md-top-tab-label">Transform</span>
      </button>
      <button type="button" class="md-top-tab" data-tab="output" role="tab" aria-selected="false" aria-controls="tabPanelOutput">
        <span class="md-top-tab-no">3</span>
        <span class="md-top-tab-label">Output</span>
      </button>
    </div>
    <section id="tabPanelInput" class="md-flow-section md-tab-panel" data-tab-panel="input">
      <span class="md-button-with-help">
        <button id="importXlsxBtnWithHelp" type="button">XLSX help host</button>
        <span class="md-button-help-anchor">
          <lht-help-tooltip label="XLSX 編集の扱い" placement="right" wide>
            <p>XLSX は MS Project XML の代替正本ではなく、確認と限定編集のための周辺表現として扱います。</p>
          </lht-help-tooltip>
        </span>
      </span>
      <details class="md-debug-accordion">
        <summary class="md-debug-accordion__summary">デバッグ情報</summary>
        <div class="md-debug-accordion__body">
          <textarea id="xmlInput"></textarea>
        </div>
      </details>
      <div id="xmlSaveState" class="md-save-state md-save-state--dirty">XML 保存状態: 未保存</div>
    </section>
    <section id="tabPanelTransform" class="md-flow-section md-tab-panel" data-tab-panel="transform" hidden>
      <div id="summaryProjectName"></div>
      <div id="summaryTaskCount"></div>
      <div id="summaryResourceCount"></div>
      <div id="summaryAssignmentCount"></div>
      <div id="summaryCalendarCount"></div>
      <div id="mermaidSvgError" class="md-hidden"></div>
      <div id="mermaidSvgPreview"></div>
      <details class="md-debug-accordion">
        <summary class="md-debug-accordion__summary">デバッグ情報</summary>
        <div class="md-debug-accordion__body">
          <button id="roundTripBtn" type="button">再読込テスト</button>
          <textarea id="modelOutput"></textarea>
          <textarea id="mermaidOutput"></textarea>
          <div id="projectPreview"></div>
          <div id="taskPreview"></div>
          <div id="resourcePreview"></div>
          <div id="assignmentPreview"></div>
          <div id="calendarPreview"></div>
        </div>
      </details>
      <div class="md-feedback-stack md-hidden">
        <div class="md-feedback-stack__title">取込結果</div>
        <div class="md-feedback-stack__text">XLSX Import 後は、ここで差分反映と検証結果を確認します。</div>
        <div class="md-feedback-stack__label md-hidden">検証メッセージ</div>
        <div id="validationIssues" class="md-hidden"></div>
        <div class="md-feedback-stack__label md-hidden">差分要約</div>
        <div id="xlsxImportSummary" class="md-hidden"></div>
      </div>
    </section>
    <section id="tabPanelOutput" class="md-flow-section md-tab-panel" data-tab-panel="output" hidden>
      <details class="md-debug-accordion">
        <summary class="md-debug-accordion__summary">設定</summary>
        <div class="md-debug-accordion__body">
          <section class="md-note-card">
            <h3 class="md-note-card__title">WBS XLSX の祝日指定</h3>
            <p class="md-note-card__text">WBS XLSX Export では、ProjectModel から補完した既定祝日を WBS 日付帯へ反映します。</p>
            <p class="md-note-card__text">既定祝日は、現在の ProjectModel に含まれる Calendar.Exceptions の非稼働日例外から補完します。画面では Calendars / Exceptions を直接編集せず、必要な変更は MS Project XML または XLSX Import 側で扱います。表示期間を空欄にすると全期間、数値を入れると BaseDate 前後の営業日で切り出します。進捗帯も営業日基準で計算します。</p>
          </section>
          <input id="wbsDisplayDaysBeforeInput" />
          <input id="wbsDisplayDaysAfterInput" />
          <input id="wbsBusinessDayRangeInput" type="checkbox" />
          <input id="wbsBusinessDayProgressInput" type="checkbox" />
          <div id="wbsHolidaySummary"></div>
          <textarea id="wbsHolidayDatesInput"></textarea>
        </div>
      </details>
      <details class="md-debug-accordion">
        <summary class="md-debug-accordion__summary">デバッグ情報</summary>
        <div class="md-debug-accordion__body">
          <textarea id="projectDraftImportInput"></textarea>
          <textarea id="aiBundleOutput"></textarea>
          <textarea id="projectOverviewOutput"></textarea>
          <input id="phaseDetailUidInput" />
          <input id="phaseDetailRootUidInput" />
          <input id="phaseDetailMaxDepthInput" />
          <textarea id="phaseDetailOutput"></textarea>
        </div>
      </details>
    </section>
    <div id="toast"></div>
  `;
  const toast = document.getElementById("toast");
  toast.show = vi.fn();
}

function bootPage() {
  mountDom();
  globalThis.mermaid = {
    initialize: vi.fn(),
    render: vi.fn(async (_id, source) => ({
      svg: `<svg data-source="${String(source).includes("gantt") ? "gantt" : "other"}"></svg>`
    }))
  };
  new Function(`${typesCode}\n${excelIoCode}\n${msProjectXmlCode}\n${projectXlsxCode}\n${wbsXlsxCode}\n${mainCode}`)();
  document.dispatchEvent(new Event("DOMContentLoaded"));
}

function bootXmlModule() {
  new Function(`${typesCode}\n${msProjectXmlCode}`)();
  return globalThis.__mikuprojectXml;
}

function getMainHooks() {
  return globalThis.__mikuprojectMainTestHooks;
}

function parseXmlViaHook() {
  getMainHooks().parseCurrentXml();
}

async function exportMermaidViaHook() {
  await getMainHooks().exportCurrentMermaid();
}

const SAMPLE_HOLIDAY_COUNT = 90;
const SAMPLE_FIRST_HOLIDAY_NAME = "春分の日";
const SAMPLE_FIRST_HOLIDAY_DATE = "2026-03-20";

function getDefaultSampleHolidayDates() {
  return globalThis.__mikuprojectWbsXlsx.collectWbsHolidayDates(
    globalThis.__mikuprojectXml.importMsProjectXml(globalThis.__mikuprojectXml.SAMPLE_XML)
  );
}

async function flushAsyncWork() {
  await Promise.resolve();
  await Promise.resolve();
}

describe("mikuproject main", () => {
  beforeEach(() => {
    document.body.innerHTML = "";
    delete globalThis.mermaid;
    Object.defineProperty(URL, "createObjectURL", {
      value: vi.fn(() => "blob:mock"),
      configurable: true
    });
    Object.defineProperty(URL, "revokeObjectURL", {
      value: vi.fn(),
      configurable: true
    });
    HTMLAnchorElement.prototype.click = vi.fn();
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2026-03-16T23:12:00+09:00"));
  });

  it("loads sample xml on startup", () => {
    bootPage();

    expect(document.getElementById("xmlInput").value).toContain("<Project");
    expect(document.getElementById("statusMessage").textContent).toContain("サンプル XML");
    expect(document.getElementById("xmlSaveState").textContent).toContain("XML 保存状態: 未保存");
    expect(document.querySelector('lht-help-tooltip[label="XLSX 編集の扱い"]')).not.toBeNull();
    expect(document.body.textContent).toContain("取込結果");
    expect(document.body.textContent).toContain("検証メッセージ");
    expect(document.body.textContent).toContain("差分要約");
    expect(document.body.textContent).toContain("差分反映と検証結果を確認します");
    expect(document.body.textContent).toContain("デバッグ情報");
    expect(document.querySelector(".md-debug-accordion")?.hasAttribute("open")).toBe(false);
    expect(document.querySelectorAll(".md-debug-accordion")[1]?.hasAttribute("open")).toBe(false);
    expect(document.body.textContent).toContain("WBS XLSX の祝日指定");
    expect(document.body.textContent).toContain("設定");
    expect(document.querySelectorAll(".md-debug-accordion")[2]?.hasAttribute("open")).toBe(false);
    expect(document.querySelectorAll(".md-debug-accordion")[3]?.hasAttribute("open")).toBe(false);
    expect(document.body.textContent).toContain("ProjectModel から補完した既定祝日");
    expect(document.body.textContent).toContain("非稼働日例外から補完");
    expect(document.body.textContent).toContain("必要な変更は MS Project XML または XLSX Import 側で扱います");
    expect(document.body.textContent).toContain("BaseDate 前後の営業日で切り出します");
    expect(document.body.textContent).toContain("進捗帯も営業日基準で計算します");
    expect(document.getElementById("wbsHolidayDatesInput").value).toBe("");
    expect(document.getElementById("wbsHolidaySummary").textContent).toBe("既定祝日: 0 件");
    expect(document.getElementById("wbsDisplayDaysBeforeInput").value).toBe("");
    expect(document.getElementById("wbsDisplayDaysAfterInput").value).toBe("");
    expect(document.querySelector(".md-feedback-stack")?.classList.contains("md-hidden")).toBe(true);
  });

  it("switches top tabs and toggles panels", async () => {
    bootPage();

    expect(document.getElementById("tabPanelInput").hidden).toBe(false);
    expect(document.getElementById("tabPanelTransform").hidden).toBe(true);
    expect(document.getElementById("tabPanelOutput").hidden).toBe(true);

    document.querySelector('.md-top-tab[data-tab="transform"]').click();
    await flushAsyncWork();
    await flushAsyncWork();
    expect(document.getElementById("tabPanelInput").hidden).toBe(true);
    expect(document.getElementById("tabPanelTransform").hidden).toBe(false);
    expect(document.getElementById("tabPanelOutput").hidden).toBe(true);
    expect(document.getElementById("summaryProjectName").textContent).toBe("mikuproject開発");
    expect(document.getElementById("mermaidOutput").value).toContain("gantt");

    document.querySelector('.md-top-tab[data-tab="output"]').click();
    expect(document.getElementById("tabPanelInput").hidden).toBe(true);
    expect(document.getElementById("tabPanelTransform").hidden).toBe(true);
    expect(document.getElementById("tabPanelOutput").hidden).toBe(false);
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
          planned_start: "2026-04-01"
        },
        tasks: [
          { uid: "draft-1", name: "要件定義", parent_uid: null, position: 0, is_summary: true },
          { uid: "draft-2", name: "ヒアリング", parent_uid: "draft-1", position: 0, planned_finish: "2026-04-01" },
          { uid: "draft-3", name: "整理期間", parent_uid: "draft-1", position: 1, planned_start: "2026-04-02", planned_finish: "2026-04-03" },
          { uid: "draft-4", name: "要件確定", parent_uid: "draft-1", position: 2, is_milestone: true, predecessors: ["draft-2"], planned_start: "2026-04-08T18:00:00", planned_finish: "2026-04-08T18:00:00" }
        ]
      }, null, 2),
      "```"
    ].join("\n");

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("summaryProjectName").textContent).toBe("新規基幹刷新");
    expect(document.getElementById("summaryTaskCount").textContent).toBe("4");
    expect(document.getElementById("summaryCalendarCount").textContent).toBe("1");
    expect(document.getElementById("xmlInput").value).toContain("<Name>新規基幹刷新</Name>");
    expect(document.getElementById("xmlInput").value).toContain("<CalendarUID>1</CalendarUID>");
    expect(document.getElementById("xmlInput").value).toContain("<Name>Standard</Name>");
    expect(document.getElementById("xmlInput").value).toContain("<UID>3</UID>");
    expect(document.getElementById("modelOutput").value).toContain("\"name\": \"ヒアリング\"");
    expect(document.getElementById("modelOutput").value).toContain("\"milestone\": false");
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
  });

  it("loads sample project_draft_view into the input area", () => {
    bootPage();

    document.getElementById("loadProjectDraftSampleBtn").click();

    const draftText = document.getElementById("projectDraftImportInput").value;
    expect(draftText).toContain("\"view_type\": \"project_draft_view\"");
    expect(draftText).toContain("\"name\": \"mikuproject開発\"");
    expect(draftText).toContain("架空検討フェーズ【架空】");
    expect(document.getElementById("statusMessage").textContent).toContain("サンプル project_draft_view");
  });

  it("parses xml into internal model summary", () => {
    bootPage();

    parseXmlViaHook();

    expect(document.getElementById("summaryProjectName").textContent).toBe("mikuproject開発");
    expect(document.getElementById("summaryTaskCount").textContent).toBe("13");
    expect(document.getElementById("summaryResourceCount").textContent).toBe("0");
    expect(document.getElementById("summaryAssignmentCount").textContent).toBe("0");
    expect(document.getElementById("summaryCalendarCount").textContent).toBe("1");
    expect(document.getElementById("modelOutput").value).toContain("\"name\": \"mikuproject開発\"");
    expect(document.getElementById("modelOutput").value).toContain("\"name\": \"基盤整備\"");
    expect(document.getElementById("modelOutput").value).toContain("\"name\": \"架空検討フェーズ【架空】\"");
    expect(document.getElementById("modelOutput").value).toContain("\"name\": \"Standard\"");
    expect(document.getElementById("projectPreview").textContent).toContain("mikuproject開発");
    expect(document.getElementById("projectPreview").textContent).toContain("Calendar=1 (Standard)");
    expect(document.getElementById("taskPreview").textContent).toContain("初期実装（MS Project XML 調査・基軸フォーマット選定・内部モデルの概要確定）");
    expect(document.getElementById("resourcePreview").textContent).toContain("まだ表示できる項目がありません");
    expect(document.getElementById("assignmentPreview").textContent).toContain("まだ表示できる項目がありません");
    expect(document.getElementById("calendarPreview").textContent).toContain(`WeekDays=7 / Exceptions=${SAMPLE_HOLIDAY_COUNT} / WorkWeeks=0`);
  });

  it("exports xml from the current model", () => {
    bootPage();

    parseXmlViaHook();
    document.getElementById("downloadXmlBtn").click();

    const xmlText = document.getElementById("xmlInput").value;
    expect(xmlText).toContain("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
    expect(xmlText).toContain("\n<Project xmlns=\"http://schemas.microsoft.com/project\">\n");
    expect(xmlText).toContain("<Name>mikuproject開発</Name>");
    expect(xmlText).toContain("<StartDate>2026-03-16</StartDate>");
    expect(xmlText).toContain("<FinishDate>2026-04-01</FinishDate>");
    expect(xmlText).toContain("<CalendarUID>1</CalendarUID>");
    expect(xmlText).toContain("<Name>Standard</Name>");
    expect(xmlText).toContain("<Name>架空検討フェーズ【架空】</Name>");
    expect(xmlText).toContain("<Name>v1.0 リリース</Name>");
  });

  it("exports mermaid gantt from the current model", async () => {
    bootPage();

    parseXmlViaHook();
    await exportMermaidViaHook();

    const mermaidText = document.getElementById("mermaidOutput").value;
    expect(mermaidText).toContain("gantt");
    expect(mermaidText).toContain("title mikuproject開発");
    expect(mermaidText).toContain("section 基盤整備");
    expect(mermaidText).toContain("section 架空検討フェーズ【架空】");
    expect(mermaidText).toContain("初期実装");
    expect(document.getElementById("mermaidSvgPreview").innerHTML).toContain("<svg");
    expect(document.getElementById("downloadMermaidSvgBtn").disabled).toBe(false);
    expect(document.getElementById("statusMessage").textContent).toContain("SVG プレビューを更新しました");
  });

  it("keeps complex mermaid dependencies as comments", () => {
    const xmlTools = bootXmlModule();
    const model = {
      project: {
        name: "Mermaid Complex",
        startDate: "2026-03-16T09:00:00",
        finishDate: "2026-03-20T18:00:00",
        scheduleFromStart: true,
        outlineCodes: [],
        wbsMasks: [],
        extendedAttributes: []
      },
      tasks: [
        {
          uid: "1",
          id: "1",
          name: "Prep",
          outlineLevel: 1,
          outlineNumber: "1",
          start: "2026-03-16T09:00:00",
          finish: "2026-03-16T18:00:00",
          duration: "PT8H0M0S",
          milestone: false,
          summary: false,
          percentComplete: 100,
          predecessors: [],
          extendedAttributes: [],
          baselines: [],
          timephasedData: []
        },
        {
          uid: "2",
          id: "2",
          name: "Review",
          outlineLevel: 1,
          outlineNumber: "2",
          start: "2026-03-17T09:00:00",
          finish: "2026-03-17T18:00:00",
          duration: "PT8H0M0S",
          milestone: false,
          summary: false,
          percentComplete: 0,
          predecessors: [],
          extendedAttributes: [],
          baselines: [],
          timephasedData: []
        },
        {
          uid: "3",
          id: "3",
          name: "Ship",
          outlineLevel: 1,
          outlineNumber: "3",
          start: "2026-03-18T09:00:00",
          finish: "2026-03-18T18:00:00",
          duration: "PT8H0M0S",
          milestone: false,
          summary: false,
          percentComplete: 0,
          predecessors: [
            { predecessorUid: "1", type: 1, linkLag: "PT2H0M0S" },
            { predecessorUid: "2", type: 4 }
          ],
          extendedAttributes: [],
          baselines: [],
          timephasedData: []
        }
      ],
      resources: [],
      assignments: [],
      calendars: []
    };

    const mermaidText = xmlTools.exportMermaidGantt(model);

    expect(mermaidText).toContain("Ship :task_3, 2026-03-18T09:00:00, 2026-03-18T18:00:00");
    expect(mermaidText).toContain("%% dependency: Ship after Prep (type=FS, lag=2h) [task_3 after task_1]");
    expect(mermaidText).toContain("%% dependency(pseudo): Ship ~= after Prep + 2h");
    expect(mermaidText).toContain("%% dependency: Ship after Review (type=SS) [task_3 after task_2]");
    expect(mermaidText).toContain("%% dependency(note): Ship has multiple predecessors");
  });

  it("sanitizes date-leading mermaid gantt labels", () => {
    const xmlTools = bootXmlModule();
    const model = {
      project: {
        name: "2026-03 mikuproject開発",
        startDate: "2026-03-16T09:00:00",
        finishDate: "2026-03-16T18:00:00",
        scheduleFromStart: true,
        outlineCodes: [],
        wbsMasks: [],
        extendedAttributes: []
      },
      tasks: [
        {
          uid: "1",
          id: "1",
          name: "2026-03-16 初期実装（42513dd：XML import/export）",
          outlineLevel: 1,
          outlineNumber: "1",
          start: "2026-03-16T09:00:00",
          finish: "2026-03-16T18:00:00",
          duration: "PT8H0M0S",
          milestone: false,
          summary: false,
          percentComplete: 0,
          predecessors: [],
          extendedAttributes: [],
          baselines: [],
          timephasedData: []
        }
      ],
      resources: [],
      assignments: [],
      calendars: []
    };

    const mermaidText = xmlTools.exportMermaidGantt(model);

    expect(mermaidText).toContain("title Project 2026-03 mikuproject開発");
    expect(mermaidText).toContain("Task 2026-03-16 初期実装（42513dd XML import/export） :task_1, 2026-03-16T09:00:00, 2026-03-16T18:00:00");
  });

  it("exports csv with parent id from the current model", async () => {
    bootPage();

    parseXmlViaHook();
    document.getElementById("downloadXmlBtn").click();
    const createObjectUrlCalls = URL.createObjectURL.mock.calls.length;
    const anchorClickCalls = HTMLAnchorElement.prototype.click.mock.calls.length;
    const xmlInput = document.getElementById("xmlInput");
    xmlInput.value = `${xmlInput.value}\n<!-- edited -->`;
    xmlInput.dispatchEvent(new Event("input"));
    document.getElementById("exportCsvBtn").click();

    expect(URL.createObjectURL.mock.calls.length - createObjectUrlCalls).toBeGreaterThan(0);
    expect(HTMLAnchorElement.prototype.click.mock.calls.length - anchorClickCalls).toBeGreaterThan(0);
    const csvBlob = URL.createObjectURL.mock.calls.at(-1)?.[0];
    expect(csvBlob).toBeTruthy();
    expect(csvBlob.type).toBe("text/csv;charset=utf-8");
    expect(document.getElementById("xmlInput").value).not.toContain("<!-- edited -->");
    expect(document.getElementById("xmlSaveState").textContent).toContain("XML 保存状態: 保存済み (2026-03-16 23:12)");
    expect(document.getElementById("statusMessage").textContent).toContain("CSV + ParentID を生成して保存しました");
  });

  it("parses csv with parent id into internal model summary", async () => {
    bootPage();

    const file = new File([[
      "ID,ParentID,WBS,Name,Start,Finish,PredecessorID,Resource,PercentComplete",
      "1,,1,Project Summary,2026-03-16T09:00:00,2026-03-20T18:00:00,,,50",
      "2,1,1.1,Design,2026-03-16T09:00:00,2026-03-17T18:00:00,,Miku,100",
      "3,1,1.2,Implementation,2026-03-18T09:00:00,2026-03-20T18:00:00,2,Miku,0"
    ].join("\n")], "sample.csv", { type: "text/csv" });
    Object.defineProperty(file, "text", {
      configurable: true,
      value: () => Promise.resolve([
        "ID,ParentID,WBS,Name,Start,Finish,PredecessorID,Resource,PercentComplete",
        "1,,1,Project Summary,2026-03-16T09:00:00,2026-03-20T18:00:00,,,50",
        "2,1,1.1,Design,2026-03-16T09:00:00,2026-03-17T18:00:00,,Miku,100",
        "3,1,1.2,Implementation,2026-03-18T09:00:00,2026-03-20T18:00:00,2,Miku,0"
      ].join("\n"))
    });

    const csvInput = document.getElementById("importCsvInput");
    Object.defineProperty(csvInput, "files", {
      value: [file],
      configurable: true
    });
    csvInput.dispatchEvent(new Event("change"));
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("summaryProjectName").textContent).toBe("CSV Imported Project");
    expect(document.getElementById("summaryTaskCount").textContent).toBe("3");
    expect(document.getElementById("summaryResourceCount").textContent).toBe("1");
    expect(document.getElementById("summaryAssignmentCount").textContent).toBe("2");
    expect(document.getElementById("summaryCalendarCount").textContent).toBe("1");
    expect(document.getElementById("modelOutput").value).toContain("\"name\": \"CSV Imported Project\"");
    expect(document.getElementById("modelOutput").value).toContain("\"name\": \"Standard\"");
    expect(document.getElementById("taskPreview").textContent).toContain("Implementation");
    expect(document.getElementById("taskPreview").textContent).toContain("Predecessors=2");
    expect(document.getElementById("resourcePreview").textContent).toContain("Miku");
    expect(document.getElementById("assignmentPreview").textContent).toContain("Task=2 (Design)");
    expect(document.getElementById("assignmentPreview").textContent).toContain("Resource=1 (Miku)");
    expect(document.getElementById("mermaidOutput").value).toContain("gantt");
    expect(document.getElementById("mermaidSvgPreview").innerHTML).toContain("<svg");
    expect(document.getElementById("statusMessage").textContent).toContain("CSV + ParentID を内部モデルへ変換しました");
  });

  it("passes round-trip check", () => {
    bootPage();

    parseXmlViaHook();
    document.getElementById("roundTripBtn").click();

    expect(document.getElementById("statusMessage").textContent).toContain("再読込テストに成功");
    expect(document.getElementById("modelOutput").value).toContain("\"extendedAttributes\": [");
  });

  it("imports xml from a file into the textarea", async () => {
    bootPage();

    const importInput = document.getElementById("importXmlInput");
    const file = new File(["<Project><Name>Imported</Name></Project>"], "sample.xml", { type: "application/xml" });
    Object.defineProperty(file, "text", {
      configurable: true,
      value: () => Promise.resolve("<Project><Name>Imported</Name></Project>")
    });
    Object.defineProperty(importInput, "files", {
      configurable: true,
      value: [file]
    });

    importInput.dispatchEvent(new Event("change"));
    await Promise.resolve();
    await Promise.resolve();

    expect(document.getElementById("xmlInput").value).toContain("<Name>Imported</Name>");
    expect(document.getElementById("summaryProjectName").textContent).toBe("Imported");
    expect(document.getElementById("summaryCalendarCount").textContent).toBe("1");
    expect(document.getElementById("modelOutput").value).toContain("\"name\": \"Imported\"");
    expect(document.getElementById("modelOutput").value).toContain("\"name\": \"Standard\"");
    expect(document.getElementById("mermaidOutput").value).toContain("gantt");
    expect(document.getElementById("mermaidSvgPreview").innerHTML).toContain("<svg");
    expect(document.getElementById("statusMessage").textContent).toContain("XML ファイルを読み込んで解析しました");
  });

  it("downloads current xlsx", () => {
    bootPage();

    parseXmlViaHook();
    document.getElementById("downloadXmlBtn").click();
    const xmlInput = document.getElementById("xmlInput");
    xmlInput.value = `${xmlInput.value}\n<!-- edited -->`;
    xmlInput.dispatchEvent(new Event("input"));
    document.getElementById("exportXlsxBtn").click();

    expect(URL.createObjectURL).toHaveBeenCalled();
    expect(HTMLAnchorElement.prototype.click).toHaveBeenCalled();
    const clickedAnchor = HTMLAnchorElement.prototype.click.mock.instances.at(-1);
    expect(clickedAnchor.download).toBe("mikuproject-export-202603162312.xlsx");
    expect(document.getElementById("xmlInput").value).not.toContain("<!-- edited -->");
    expect(document.getElementById("xmlSaveState").textContent).toContain("XML 保存状態: 保存済み (2026-03-16 23:12)");
    expect(document.getElementById("statusMessage").textContent).toContain("XLSX ファイルをエクスポートしました");
  });

  it("downloads current wbs xlsx", () => {
    bootPage();
    const exportSpy = vi.spyOn(globalThis.__mikuprojectWbsXlsx, "exportWbsWorkbook");

    parseXmlViaHook();
    document.getElementById("downloadXmlBtn").click();
    const xmlInput = document.getElementById("xmlInput");
    xmlInput.value = `${xmlInput.value}\n<!-- edited -->`;
    xmlInput.dispatchEvent(new Event("input"));
    document.getElementById("exportWbsXlsxBtn").click();
    const defaultHolidayDates = getDefaultSampleHolidayDates();

    expect(URL.createObjectURL).toHaveBeenCalled();
    expect(HTMLAnchorElement.prototype.click).toHaveBeenCalled();
    const clickedAnchor = HTMLAnchorElement.prototype.click.mock.instances.at(-1);
    expect(clickedAnchor.download).toBe("mikuproject-wbs-202603162312.xlsx");
    expect(document.getElementById("wbsHolidayDatesInput").value.split("\n")).toEqual(defaultHolidayDates);
    expect(exportSpy.mock.calls.at(-1)?.[1]).toEqual({
      holidayDates: defaultHolidayDates,
      displayDaysBeforeBaseDate: undefined,
      displayDaysAfterBaseDate: undefined,
      useBusinessDaysForDisplayRange: true,
      useBusinessDaysForProgressBand: true
    });
    expect(document.getElementById("xmlInput").value).not.toContain("<!-- edited -->");
    expect(document.getElementById("xmlSaveState").textContent).toContain("XML 保存状態: 保存済み (2026-03-16 23:12)");
    expect(document.getElementById("statusMessage").textContent).toContain("WBS XLSX ファイルをエクスポートしました");
    expect(document.getElementById("statusMessage").textContent).toContain(`祝日 ${SAMPLE_HOLIDAY_COUNT} 件`);
  });

  it("downloads current wbs xlsx with configured display range", async () => {
    bootPage();
    const exportSpy = vi.spyOn(globalThis.__mikuprojectWbsXlsx, "exportWbsWorkbook");

    parseXmlViaHook();
    document.getElementById("wbsDisplayDaysBeforeInput").value = "1";
    document.getElementById("wbsDisplayDaysAfterInput").value = "2";
    document.getElementById("exportWbsXlsxBtn").click();
    const defaultHolidayDates = getDefaultSampleHolidayDates();

    expect(exportSpy).toHaveBeenCalled();
    expect(exportSpy.mock.calls.at(-1)?.[1]).toEqual({
      holidayDates: defaultHolidayDates,
      displayDaysBeforeBaseDate: 1,
      displayDaysAfterBaseDate: 2,
      useBusinessDaysForDisplayRange: true,
      useBusinessDaysForProgressBand: true
    });
    expect(document.getElementById("statusMessage").textContent).toContain("表示期間 営業日 基準日前 1 日, 基準日後 2 日");
  });

  it("downloads current wbs xlsx with business-day display range", async () => {
    bootPage();
    const exportSpy = vi.spyOn(globalThis.__mikuprojectWbsXlsx, "exportWbsWorkbook");

    parseXmlViaHook();
    document.getElementById("wbsDisplayDaysBeforeInput").value = "1";
    document.getElementById("wbsDisplayDaysAfterInput").value = "2";
    document.getElementById("exportWbsXlsxBtn").click();
    const defaultHolidayDates = getDefaultSampleHolidayDates();

    expect(exportSpy.mock.calls.at(-1)?.[1]).toEqual({
      holidayDates: defaultHolidayDates,
      displayDaysBeforeBaseDate: 1,
      displayDaysAfterBaseDate: 2,
      useBusinessDaysForDisplayRange: true,
      useBusinessDaysForProgressBand: true
    });
    expect(document.getElementById("statusMessage").textContent).toContain("表示期間 営業日 基準日前 1 日, 基準日後 2 日");
  });

  it("downloads current wbs xlsx with business-day progress band", async () => {
    bootPage();
    const exportSpy = vi.spyOn(globalThis.__mikuprojectWbsXlsx, "exportWbsWorkbook");

    parseXmlViaHook();
    document.getElementById("exportWbsXlsxBtn").click();
    const defaultHolidayDates = getDefaultSampleHolidayDates();

    expect(exportSpy.mock.calls.at(-1)?.[1]).toEqual({
      holidayDates: defaultHolidayDates,
      displayDaysBeforeBaseDate: undefined,
      displayDaysAfterBaseDate: undefined,
      useBusinessDaysForDisplayRange: true,
      useBusinessDaysForProgressBand: true
    });
    expect(document.getElementById("statusMessage").textContent).toContain("進捗帯 営業日");
  });

  it("fills wbs holiday input from model defaults when xml is parsed", () => {
    bootPage();

    parseXmlViaHook();
    const defaultHolidayDates = getDefaultSampleHolidayDates();

    expect(document.getElementById("wbsHolidayDatesInput").value.split("\n")).toEqual(defaultHolidayDates);
    expect(document.getElementById("wbsHolidaySummary").textContent).toContain(`既定祝日: ${SAMPLE_HOLIDAY_COUNT} 件`);
    expect(document.getElementById("wbsHolidaySummary").textContent).toContain("2026-03-20");
  });

  it("imports xlsx edits back into the current model and xml", async () => {
    bootPage();
    parseXmlViaHook();
    document.getElementById("downloadXmlBtn").click();
    expect(document.getElementById("xmlSaveState").textContent).toContain("XML 保存状態: 保存済み (2026-03-16 23:12)");

    const codec = new globalThis.__mikuprojectExcelIo.XlsxWorkbookCodec();
    const workbook = globalThis.__mikuprojectProjectXlsx.exportProjectWorkbook(
      globalThis.__mikuprojectXml.importMsProjectXml(document.getElementById("xmlInput").value)
    );
    const tasksSheet = workbook.sheets.find((sheet) => sheet.name === "Tasks");
    tasksSheet.rows[5].cells[2].value = "初期実装 Imported From XLSX";
    tasksSheet.rows[5].cells[9].value = 77;
    const bytes = codec.exportWorkbook(workbook);

    const importInput = document.getElementById("importXlsxInput");
    const file = new File([bytes], "sample.xlsx", {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    Object.defineProperty(file, "arrayBuffer", {
      configurable: true,
      value: () => Promise.resolve(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength))
    });
    Object.defineProperty(importInput, "files", {
      configurable: true,
      value: [file]
    });

    importInput.dispatchEvent(new Event("change"));
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("modelOutput").value).toContain("\"name\": \"初期実装 Imported From XLSX\"");
    expect(document.getElementById("modelOutput").value).toContain("\"percentComplete\": 77");
    expect(document.getElementById("xmlInput").value).toContain("<Name>初期実装 Imported From XLSX</Name>");
    expect(document.getElementById("statusMessage").textContent).toContain("XLSX を読み込んで 2 件の変更を反映しました");
    expect(document.getElementById("statusMessage").textContent).toContain("XML Export で保存できます");
    expect(document.getElementById("xmlSaveState").textContent).toContain("XML 保存状態: 未保存");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Tasks");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Tasks 1");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("変更なし: Project, Resources, Assignments, Calendars");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("UID=3 初期実装（MS Project XML 調査・基軸フォーマット選定・内部モデルの概要確定）");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Name: 初期実装（MS Project XML 調査・基軸フォーマット選定・内部モデルの概要確定） -> 初期実装 Imported From XLSX");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("PercentComplete: 0 -> 77");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("反映後の XML は更新済みです");
    expect(document.querySelectorAll("#xlsxImportSummary .md-xlsx-summary__section")).toHaveLength(1);
    expect(document.querySelectorAll("#xlsxImportSummary .md-xlsx-summary__item")).toHaveLength(1);
  });

  it("imports project sheet edits back into the current model and xml", async () => {
    bootPage();
    parseXmlViaHook();

    const codec = new globalThis.__mikuprojectExcelIo.XlsxWorkbookCodec();
    const workbook = globalThis.__mikuprojectProjectXlsx.exportProjectWorkbook(
      globalThis.__mikuprojectXml.importMsProjectXml(document.getElementById("xmlInput").value)
    );
    const projectSheet = workbook.sheets.find((sheet) => sheet.name === "Project");
    projectSheet.rows[3].cells[1].value = "Project From XLSX";
    projectSheet.rows[13].cells[1].value = 420;
    const bytes = codec.exportWorkbook(workbook);

    const importInput = document.getElementById("importXlsxInput");
    const file = new File([bytes], "project-sheet.xlsx", {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    Object.defineProperty(file, "arrayBuffer", {
      configurable: true,
      value: () => Promise.resolve(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength))
    });
    Object.defineProperty(importInput, "files", {
      configurable: true,
      value: [file]
    });

    importInput.dispatchEvent(new Event("change"));
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("modelOutput").value).toContain("\"name\": \"Project From XLSX\"");
    expect(document.getElementById("modelOutput").value).toContain("\"minutesPerDay\": 420");
    expect(document.getElementById("xmlInput").value).toContain("<Name>Project From XLSX</Name>");
    expect(document.getElementById("statusMessage").textContent).toContain("XLSX を読み込んで 2 件の変更を反映しました");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Project 1");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("変更なし: Tasks, Resources, Assignments, Calendars");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("UID=project mikuproject開発");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Name: mikuproject開発 -> Project From XLSX");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("MinutesPerDay: (empty) -> 420");
    expect(document.querySelectorAll("#xlsxImportSummary .md-xlsx-summary__section")).toHaveLength(1);
    expect(document.querySelectorAll("#xlsxImportSummary .md-xlsx-summary__item")).toHaveLength(1);
  });

  it("renders project import summary content without xlsx import wiring", () => {
    bootPage();

    getMainHooks().renderXlsxImportSummary([
      { scope: "project", uid: "project", label: "mikuproject開発", field: "CalendarUID", before: "1", after: "2" },
      { scope: "project", uid: "project", label: "mikuproject開発", field: "ScheduleFromStart", before: true, after: false },
      { scope: "project", uid: "project", label: "mikuproject開発", field: "Author", before: undefined, after: "Author From XLSX" }
    ]);

    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Project 1");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("変更なし: Tasks, Resources, Assignments, Calendars");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("UID=project mikuproject開発");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("CalendarUID: 1 -> 2");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("ScheduleFromStart: true -> false");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Author: (empty) -> Author From XLSX");
    expect(document.querySelectorAll("#xlsxImportSummary .md-xlsx-summary__section")).toHaveLength(1);
    expect(document.querySelectorAll("#xlsxImportSummary .md-xlsx-summary__item")).toHaveLength(1);
  });

  it("reports when xlsx import has no applicable changes", async () => {
    bootPage();
    parseXmlViaHook();

    const codec = new globalThis.__mikuprojectExcelIo.XlsxWorkbookCodec();
    const originalXml = document.getElementById("xmlInput").value;
    const workbook = globalThis.__mikuprojectProjectXlsx.exportProjectWorkbook(
      globalThis.__mikuprojectXml.importMsProjectXml(originalXml)
    );
    const bytes = codec.exportWorkbook(workbook);

    const importInput = document.getElementById("importXlsxInput");
    const file = new File([bytes], "no-change.xlsx", {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    Object.defineProperty(file, "arrayBuffer", {
      configurable: true,
      value: () => Promise.resolve(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength))
    });
    Object.defineProperty(importInput, "files", {
      configurable: true,
      value: [file]
    });

    importInput.dispatchEvent(new Event("change"));
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("XLSX に反映対象の変更はありませんでした");
    expect(document.getElementById("statusMessage").textContent).toContain("XML は未変更です");
    expect(
      globalThis.__mikuprojectXml.normalizeProjectModel(
        globalThis.__mikuprojectXml.importMsProjectXml(document.getElementById("xmlInput").value)
      )
    ).toEqual(
      globalThis.__mikuprojectXml.normalizeProjectModel(
        globalThis.__mikuprojectXml.importMsProjectXml(originalXml)
      )
    );
    expect(document.getElementById("xlsxImportSummary").textContent).toBe("");
    expect(document.getElementById("xlsxImportSummary").classList.contains("md-hidden")).toBe(true);
  });

  it("ignores edits in unsupported xlsx columns and sheets", async () => {
    bootPage();
    parseXmlViaHook();

    const codec = new globalThis.__mikuprojectExcelIo.XlsxWorkbookCodec();
    const originalXml = document.getElementById("xmlInput").value;
    const workbook = globalThis.__mikuprojectProjectXlsx.exportProjectWorkbook(
      globalThis.__mikuprojectXml.importMsProjectXml(originalXml)
    );
    const tasksSheet = workbook.sheets.find((sheet) => sheet.name === "Tasks");
    const calendarsSheet = workbook.sheets.find((sheet) => sheet.name === "Calendars");

    tasksSheet.rows[4].cells[8].value = "PT99H0M0S";
    calendarsSheet.rows[3].cells[4].value = 99;

    const bytes = codec.exportWorkbook(workbook);
    const importInput = document.getElementById("importXlsxInput");
    const file = new File([bytes], "unsupported-columns.xlsx", {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    Object.defineProperty(file, "arrayBuffer", {
      configurable: true,
      value: () => Promise.resolve(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength))
    });
    Object.defineProperty(importInput, "files", {
      configurable: true,
      value: [file]
    });

    importInput.dispatchEvent(new Event("change"));
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("XLSX に反映対象の変更はありませんでした");
    expect(document.getElementById("statusMessage").textContent).toContain("XML は未変更です");
    expect(document.getElementById("modelOutput").value).not.toContain("\"duration\": \"PT99H0M0S\"");
    expect(document.getElementById("modelOutput").value).not.toContain("\"weekDays\": 99");
    expect(document.getElementById("xmlInput").value).toBe(originalXml);
    expect(document.getElementById("xlsxImportSummary").textContent).toBe("");
    expect(document.getElementById("xlsxImportSummary").classList.contains("md-hidden")).toBe(true);
  });

  it("ignores calendar WeekDays, Exceptions, and WorkWeeks edits in xlsx import", async () => {
    bootPage();
    parseXmlViaHook();

    const codec = new globalThis.__mikuprojectExcelIo.XlsxWorkbookCodec();
    const originalXml = document.getElementById("xmlInput").value;
    const workbook = globalThis.__mikuprojectProjectXlsx.exportProjectWorkbook(
      globalThis.__mikuprojectXml.importMsProjectXml(originalXml)
    );
    const calendarsSheet = workbook.sheets.find((sheet) => sheet.name === "Calendars");

    calendarsSheet.rows[3].cells[4].value = 77;
    calendarsSheet.rows[3].cells[5].value = 88;
    calendarsSheet.rows[3].cells[6].value = 99;

    const bytes = codec.exportWorkbook(workbook);
    const importInput = document.getElementById("importXlsxInput");
    const file = new File([bytes], "ignored-calendar-structure.xlsx", {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    Object.defineProperty(file, "arrayBuffer", {
      configurable: true,
      value: () => Promise.resolve(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength))
    });
    Object.defineProperty(importInput, "files", {
      configurable: true,
      value: [file]
    });

    importInput.dispatchEvent(new Event("change"));
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("statusMessage").textContent).toContain("XLSX に反映対象の変更はありませんでした");
    expect(document.getElementById("statusMessage").textContent).toContain("XML は未変更です");
    expect(document.getElementById("modelOutput").value).not.toContain("\"weekDays\": 77");
    expect(document.getElementById("modelOutput").value).not.toContain("\"exceptions\": 88");
    expect(document.getElementById("modelOutput").value).not.toContain("\"workWeeks\": 99");
    expect(document.getElementById("xmlInput").value).toBe(originalXml);
    expect(document.getElementById("xlsxImportSummary").textContent).toBe("");
    expect(document.getElementById("xlsxImportSummary").classList.contains("md-hidden")).toBe(true);
  }, 10000);

  it("renders grouped xlsx import summary content without xlsx import wiring", () => {
    bootPage();

    getMainHooks().renderXlsxImportSummary([
      { scope: "calendars", uid: "1", label: "Standard", field: "Name", before: "Standard", after: "Standard Updated" },
      { scope: "tasks", uid: "3", label: "初期実装（MS Project XML 調査・基軸フォーマット選定・内部モデルの概要確定）", field: "Start", before: "2026-03-16", after: "2026-03-17" },
      { scope: "tasks", uid: "3", label: "初期実装（MS Project XML 調査・基軸フォーマット選定・内部モデルの概要確定）", field: "Finish", before: "2026-03-16", after: "2026-03-18" },
      { scope: "resources", uid: "1", label: "Miku", field: "Name", before: "Miku", after: "Miku Renamed" },
      { scope: "assignments", uid: "1", label: "TaskUID=2", field: "Work", before: "PT16H0M0S", after: "PT12H0M0S" }
    ]);

    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Tasks 1");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Resources 1");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Assignments 1");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Calendars 1");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("変更なし: Project");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("UID=1 Standard");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("UID=3 初期実装（MS Project XML 調査・基軸フォーマット選定・内部モデルの概要確定）");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("UID=1 Miku");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("UID=1 TaskUID=2");
    expect(document.querySelectorAll("#xlsxImportSummary .md-xlsx-summary__section")).toHaveLength(4);
    expect(document.querySelectorAll("#xlsxImportSummary .md-xlsx-summary__item")).toHaveLength(4);
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


  it("downloads current xml", () => {
    bootPage();

    document.getElementById("downloadXmlBtn").click();

    expect(URL.createObjectURL).toHaveBeenCalled();
    expect(HTMLAnchorElement.prototype.click).toHaveBeenCalled();
    const clickedAnchor = HTMLAnchorElement.prototype.click.mock.instances.at(-1);
    expect(clickedAnchor.download).toBe("mikuproject-export-202603162312.xml");
    expect(document.getElementById("statusMessage").textContent).toContain("XML ファイルをエクスポートしました");
    expect(document.getElementById("xmlSaveState").textContent).toContain("XML 保存状態: 保存済み (2026-03-16 23:12)");
    expect(document.getElementById("xmlSaveState").classList.contains("md-save-state--clean")).toBe(true);
  });

  it("returns xml save state to unsaved after manual xml edit", async () => {
    bootPage();

    document.getElementById("downloadXmlBtn").click();
    expect(document.getElementById("xmlSaveState").textContent).toContain("XML 保存状態: 保存済み (2026-03-16 23:12)");

    const xmlInput = document.getElementById("xmlInput");
    xmlInput.value = `${xmlInput.value}\n<!-- edited -->`;
    xmlInput.dispatchEvent(new Event("input"));
    await flushAsyncWork();

    expect(document.getElementById("xmlSaveState").textContent).toContain("XML 保存状態: 未保存");
    expect(document.getElementById("xmlSaveState").classList.contains("md-save-state--dirty")).toBe(true);
  });

  it("exports regenerated xml instead of manual textarea edits when a model exists", async () => {
    bootPage();

    parseXmlViaHook();
    const xmlInput = document.getElementById("xmlInput");
    xmlInput.value = `${xmlInput.value}\n<!-- edited -->`;
    xmlInput.dispatchEvent(new Event("input"));
    await flushAsyncWork();

    const OriginalBlob = Blob;
    class InspectableBlob extends OriginalBlob {
      __parts;

      constructor(parts, options) {
        super(parts, options);
        this.__parts = parts;
      }

      text() {
        return Promise.resolve(this.__parts.join(""));
      }
    }
    globalThis.Blob = InspectableBlob;
    URL.createObjectURL.mockClear();
    HTMLAnchorElement.prototype.click.mockClear();

    try {
      document.getElementById("downloadXmlBtn").click();

      expect(URL.createObjectURL).toHaveBeenCalled();
      const exportedBlob = URL.createObjectURL.mock.calls.at(-1)?.[0];
      expect(exportedBlob).toBeInstanceOf(InspectableBlob);
      await expect(exportedBlob.text()).resolves.not.toContain("<!-- edited -->");
      expect(document.getElementById("xmlInput").value).not.toContain("<!-- edited -->");
      expect(document.getElementById("xmlSaveState").textContent).toContain("XML 保存状態: 保存済み (2026-03-16 23:12)");
    } finally {
      globalThis.Blob = OriginalBlob;
    }
  });


  it("downloads rendered mermaid svg", async () => {
    bootPage();

    parseXmlViaHook();
    document.getElementById("downloadXmlBtn").click();
    const xmlInput = document.getElementById("xmlInput");
    xmlInput.value = `${xmlInput.value}\n<!-- edited -->`;
    xmlInput.dispatchEvent(new Event("input"));
    URL.createObjectURL.mockClear();
    HTMLAnchorElement.prototype.click.mockClear();
    document.getElementById("downloadMermaidSvgBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(URL.createObjectURL).toHaveBeenCalled();
    expect(HTMLAnchorElement.prototype.click).toHaveBeenCalled();
    const clickedAnchor = HTMLAnchorElement.prototype.click.mock.instances.at(-1);
    expect(clickedAnchor.download).toBe("mikuproject-mermaid.svg");
    expect(document.getElementById("xmlInput").value).not.toContain("<!-- edited -->");
    expect(document.getElementById("xmlSaveState").textContent).toContain("XML 保存状態: 保存済み (2026-03-16 23:12)");
    expect(document.getElementById("mermaidOutput").value).toContain("gantt");
    expect(document.getElementById("statusMessage").textContent).toContain("Mermaid SVG を保存しました");
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
  }, 10000);

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
  }, 10000);

  it("round-trips the minimal xml sample", () => {
    const xmlTools = bootXmlModule();

    const model = xmlTools.importMsProjectXml(minimalXml);
    const exportedXml = xmlTools.exportMsProjectXml(model);
    const reparsedModel = xmlTools.importMsProjectXml(exportedXml);

    expect(model.project.name).toBe("Minimal Project");
    expect(reparsedModel.project.name).toBe("Minimal Project");
    expect(reparsedModel.tasks).toHaveLength(1);
    expect(reparsedModel.tasks[0].name).toBe("Single Task");
    expect(xmlTools.validateProjectModel(reparsedModel)).toHaveLength(0);
  });

  it("round-trips project metadata fields", () => {
    const xmlTools = bootXmlModule();
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<Project xmlns="http://schemas.microsoft.com/project">
  <Name>Metadata Project</Name>
  <Title>Metadata Title</Title>
  <Company>Example Company</Company>
  <Author>Example Author</Author>
  <CreationDate>2026-03-16T07:00:00</CreationDate>
  <LastSaved>2026-03-16T10:15:00</LastSaved>
  <SaveVersion>14</SaveVersion>
  <CurrencyCode>JPY</CurrencyCode>
  <CurrencyDigits>0</CurrencyDigits>
  <CurrencySymbol>¥</CurrencySymbol>
  <CurrencySymbolPosition>0</CurrencySymbolPosition>
  <FYStartDate>2026-04-01T00:00:00</FYStartDate>
  <FiscalYearStart>1</FiscalYearStart>
  <CriticalSlackLimit>0</CriticalSlackLimit>
  <DefaultTaskType>1</DefaultTaskType>
  <DefaultFixedCostAccrual>2</DefaultFixedCostAccrual>
  <DefaultStandardRate>5000/h</DefaultStandardRate>
  <DefaultOvertimeRate>7000/h</DefaultOvertimeRate>
  <DefaultTaskEVMethod>0</DefaultTaskEVMethod>
  <NewTaskStartDate>0</NewTaskStartDate>
  <NewTasksAreManual>0</NewTasksAreManual>
  <NewTasksEffortDriven>1</NewTasksEffortDriven>
  <NewTasksEstimated>1</NewTasksEstimated>
  <ActualsInSync>0</ActualsInSync>
  <EditableActualCosts>1</EditableActualCosts>
  <HonorConstraints>1</HonorConstraints>
  <InsertedProjectsLikeSummary>1</InsertedProjectsLikeSummary>
  <MultipleCriticalPaths>0</MultipleCriticalPaths>
  <TaskUpdatesResource>1</TaskUpdatesResource>
  <UpdateManuallyScheduledTasksWhenEditingLinks>0</UpdateManuallyScheduledTasksWhenEditingLinks>
  <OutlineCodes>
    <OutlineCode>
      <FieldID>188743731</FieldID>
      <FieldName>Outline Code1</FieldName>
      <Alias>Phase</Alias>
      <OnlyTableValues>1</OnlyTableValues>
      <Masks>
        <Mask>
          <Level>1</Level>
          <Mask>*</Mask>
          <Length>0</Length>
          <Sequence>0</Sequence>
        </Mask>
      </Masks>
      <Values>
        <Value>
          <Value>PLAN</Value>
          <Description>Planning</Description>
        </Value>
      </Values>
    </OutlineCode>
  </OutlineCodes>
  <WBSMasks>
    <WBSMask>
      <Level>1</Level>
      <Mask>A</Mask>
      <Length>1</Length>
      <Sequence>1</Sequence>
    </WBSMask>
  </WBSMasks>
  <ExtendedAttributes>
    <ExtendedAttribute>
      <FieldID>188743734</FieldID>
      <FieldName>Text1</FieldName>
      <Alias>Owner</Alias>
      <CalculationType>0</CalculationType>
      <RestrictValues>0</RestrictValues>
      <AppendNewValues>1</AppendNewValues>
    </ExtendedAttribute>
  </ExtendedAttributes>
  <StartDate>2026-03-16T09:00:00</StartDate>
  <FinishDate>2026-03-16T18:00:00</FinishDate>
  <ScheduleFromStart>1</ScheduleFromStart>
  <Tasks />
  <Resources />
  <Assignments />
</Project>`;

    const model = xmlTools.importMsProjectXml(xml);
    const exportedXml = xmlTools.exportMsProjectXml(model);
    const reparsedModel = xmlTools.importMsProjectXml(exportedXml);

    expect(reparsedModel.project.title).toBe("Metadata Title");
    expect(reparsedModel.project.company).toBe("Example Company");
    expect(reparsedModel.project.author).toBe("Example Author");
    expect(reparsedModel.project.creationDate).toBe("2026-03-16T07:00:00");
    expect(reparsedModel.project.lastSaved).toBe("2026-03-16T10:15:00");
    expect(reparsedModel.project.saveVersion).toBe(14);
    expect(reparsedModel.project.currencyCode).toBe("JPY");
    expect(reparsedModel.project.currencyDigits).toBe(0);
    expect(reparsedModel.project.currencySymbol).toBe("¥");
    expect(reparsedModel.project.currencySymbolPosition).toBe(0);
    expect(reparsedModel.project.fyStartDate).toBe("2026-04-01T00:00:00");
    expect(reparsedModel.project.fiscalYearStart).toBe(true);
    expect(reparsedModel.project.criticalSlackLimit).toBe(0);
    expect(reparsedModel.project.defaultTaskType).toBe(1);
    expect(reparsedModel.project.defaultFixedCostAccrual).toBe(2);
    expect(reparsedModel.project.defaultStandardRate).toBe("5000/h");
    expect(reparsedModel.project.defaultOvertimeRate).toBe("7000/h");
    expect(reparsedModel.project.defaultTaskEVMethod).toBe(0);
    expect(reparsedModel.project.newTaskStartDate).toBe(0);
    expect(reparsedModel.project.newTasksAreManual).toBe(false);
    expect(reparsedModel.project.newTasksEffortDriven).toBe(true);
    expect(reparsedModel.project.newTasksEstimated).toBe(true);
    expect(reparsedModel.project.actualsInSync).toBe(false);
    expect(reparsedModel.project.editableActualCosts).toBe(true);
    expect(reparsedModel.project.honorConstraints).toBe(true);
    expect(reparsedModel.project.insertedProjectsLikeSummary).toBe(true);
    expect(reparsedModel.project.multipleCriticalPaths).toBe(false);
    expect(reparsedModel.project.taskUpdatesResource).toBe(true);
    expect(reparsedModel.project.updateManuallyScheduledTasksWhenEditingLinks).toBe(false);
    expect(reparsedModel.project.outlineCodes).toHaveLength(1);
    expect(reparsedModel.project.outlineCodes[0].fieldID).toBe("188743731");
    expect(reparsedModel.project.outlineCodes[0].alias).toBe("Phase");
    expect(reparsedModel.project.outlineCodes[0].values[0].value).toBe("PLAN");
    expect(reparsedModel.project.wbsMasks).toHaveLength(1);
    expect(reparsedModel.project.wbsMasks[0].mask).toBe("A");
    expect(reparsedModel.project.extendedAttributes).toHaveLength(1);
    expect(reparsedModel.project.extendedAttributes[0].fieldName).toBe("Text1");
    expect(reparsedModel.project.extendedAttributes[0].alias).toBe("Owner");
    expect(reparsedModel.project.extendedAttributes[0].appendNewValues).toBe(true);
  });

  it("round-trips project scheduling metadata fields", () => {
    const xmlTools = bootXmlModule();
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<Project xmlns="http://schemas.microsoft.com/project">
  <Name>Schedule Metadata Project</Name>
  <StatusDate>2026-03-17T09:00:00</StatusDate>
  <WeekStartDay>2</WeekStartDay>
  <WorkFormat>2</WorkFormat>
  <DurationFormat>7</DurationFormat>
  <StartDate>2026-03-16T09:00:00</StartDate>
  <FinishDate>2026-03-18T18:00:00</FinishDate>
  <ScheduleFromStart>1</ScheduleFromStart>
  <Tasks />
  <Resources />
  <Assignments />
</Project>`;

    const model = xmlTools.importMsProjectXml(xml);
    const exportedXml = xmlTools.exportMsProjectXml(model);
    const reparsedModel = xmlTools.importMsProjectXml(exportedXml);

    expect(reparsedModel.project.statusDate).toBe("2026-03-17T09:00:00");
    expect(reparsedModel.project.weekStartDay).toBe(2);
    expect(reparsedModel.project.workFormat).toBe(2);
    expect(reparsedModel.project.durationFormat).toBe(7);
  });

  it("round-trips calendar base and weekday fields", () => {
    const xmlTools = bootXmlModule();
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<Project xmlns="http://schemas.microsoft.com/project">
  <Name>Calendar Detail Project</Name>
  <StartDate>2026-03-16T09:00:00</StartDate>
  <FinishDate>2026-03-16T18:00:00</FinishDate>
  <ScheduleFromStart>1</ScheduleFromStart>
  <Calendars>
    <Calendar>
      <UID>1</UID>
      <Name>Standard</Name>
      <IsBaseCalendar>1</IsBaseCalendar>
      <IsBaselineCalendar>1</IsBaselineCalendar>
    </Calendar>
    <Calendar>
      <UID>2</UID>
      <Name>Night Shift</Name>
      <IsBaseCalendar>0</IsBaseCalendar>
      <BaseCalendarUID>1</BaseCalendarUID>
      <WeekDays>
        <WeekDay>
          <DayType>7</DayType>
          <DayWorking>1</DayWorking>
          <WorkingTimes>
            <WorkingTime>
              <FromTime>18:00:00</FromTime>
              <ToTime>22:00:00</ToTime>
            </WorkingTime>
          </WorkingTimes>
        </WeekDay>
      </WeekDays>
    </Calendar>
  </Calendars>
  <Tasks />
  <Resources />
  <Assignments />
</Project>`;

    const model = xmlTools.importMsProjectXml(xml);
    const exportedXml = xmlTools.exportMsProjectXml(model);
    const reparsedModel = xmlTools.importMsProjectXml(exportedXml);

    expect(reparsedModel.calendars).toHaveLength(2);
    expect(reparsedModel.calendars[0].isBaselineCalendar).toBe(true);
    expect(reparsedModel.calendars[1].baseCalendarUID).toBe("1");
    expect(reparsedModel.calendars[1].weekDays[0].dayType).toBe(7);
    expect(reparsedModel.calendars[1].weekDays[0].dayWorking).toBe(true);
    expect(reparsedModel.calendars[1].weekDays[0].workingTimes[0].fromTime).toBe("18:00:00");
    expect(reparsedModel.calendars[1].weekDays[0].workingTimes[0].toTime).toBe("22:00:00");
  });

  it("round-trips calendar exceptions and workweeks", () => {
    const xmlTools = bootXmlModule();
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<Project xmlns="http://schemas.microsoft.com/project">
  <Name>Calendar Exception Project</Name>
  <StartDate>2026-03-16T09:00:00</StartDate>
  <FinishDate>2026-03-16T18:00:00</FinishDate>
  <ScheduleFromStart>1</ScheduleFromStart>
  <Calendars>
    <Calendar>
      <UID>1</UID>
      <Name>Standard</Name>
      <IsBaseCalendar>1</IsBaseCalendar>
      <Exceptions>
        <Exception>
          <Name>Holiday</Name>
          <FromDate>2026-03-20T00:00:00</FromDate>
          <ToDate>2026-03-20T23:59:59</ToDate>
          <DayWorking>0</DayWorking>
          <WorkingTimes>
            <WorkingTime>
              <FromTime>09:00:00</FromTime>
              <ToTime>12:00:00</ToTime>
            </WorkingTime>
          </WorkingTimes>
        </Exception>
      </Exceptions>
      <WorkWeeks>
        <WorkWeek>
          <Name>Sprint 1</Name>
          <FromDate>2026-03-16T00:00:00</FromDate>
          <ToDate>2026-03-31T23:59:59</ToDate>
          <WeekDays>
            <WeekDay>
              <DayType>2</DayType>
              <DayWorking>1</DayWorking>
              <WorkingTimes>
                <WorkingTime>
                  <FromTime>09:00:00</FromTime>
                  <ToTime>17:00:00</ToTime>
                </WorkingTime>
              </WorkingTimes>
            </WeekDay>
          </WeekDays>
        </WorkWeek>
      </WorkWeeks>
    </Calendar>
  </Calendars>
  <Tasks />
  <Resources />
  <Assignments />
</Project>`;

    const model = xmlTools.importMsProjectXml(xml);
    const exportedXml = xmlTools.exportMsProjectXml(model);
    const reparsedModel = xmlTools.importMsProjectXml(exportedXml);

    expect(reparsedModel.calendars[0].exceptions[0].name).toBe("Holiday");
    expect(reparsedModel.calendars[0].exceptions[0].dayWorking).toBe(false);
    expect(reparsedModel.calendars[0].exceptions[0].workingTimes[0].fromTime).toBe("09:00:00");
    expect(reparsedModel.calendars[0].exceptions[0].workingTimes[0].toTime).toBe("12:00:00");
    expect(reparsedModel.calendars[0].workWeeks[0].name).toBe("Sprint 1");
    expect(reparsedModel.calendars[0].workWeeks[0].weekDays[0].dayType).toBe(2);
    expect(reparsedModel.calendars[0].workWeeks[0].weekDays[0].workingTimes[0].toTime).toBe("17:00:00");
  });

  it("warns when calendar baseCalendarUID points to itself", () => {
    const xmlTools = bootXmlModule();
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<Project xmlns="http://schemas.microsoft.com/project">
  <Name>Self Base Calendar Project</Name>
  <StartDate>2026-03-16T09:00:00</StartDate>
  <FinishDate>2026-03-16T18:00:00</FinishDate>
  <ScheduleFromStart>1</ScheduleFromStart>
  <Calendars>
    <Calendar>
      <UID>1</UID>
      <Name>Loop Calendar</Name>
      <IsBaseCalendar>0</IsBaseCalendar>
      <BaseCalendarUID>1</BaseCalendarUID>
    </Calendar>
  </Calendars>
  <Tasks>
    <Task>
      <UID>1</UID>
      <ID>1</ID>
      <Name>Task</Name>
      <OutlineLevel>1</OutlineLevel>
      <OutlineNumber>1</OutlineNumber>
      <Start>2026-03-16T09:00:00</Start>
      <Finish>2026-03-16T18:00:00</Finish>
      <Duration>PT8H0M0S</Duration>
      <Milestone>0</Milestone>
      <Summary>0</Summary>
      <PercentComplete>0</PercentComplete>
    </Task>
  </Tasks>
</Project>`;

    const issues = xmlTools.validateProjectModel(xmlTools.importMsProjectXml(xml));

    expect(issues.some((issue) => issue.message.includes("BaseCalendarUID が自身を指しています"))).toBe(true);
  });

  it("round-trips resource and assignment practical fields", () => {
    const xmlTools = bootXmlModule();
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<Project xmlns="http://schemas.microsoft.com/project">
  <Name>Resource Assignment Project</Name>
  <StartDate>2026-03-16T09:00:00</StartDate>
  <FinishDate>2026-03-17T18:00:00</FinishDate>
  <ScheduleFromStart>1</ScheduleFromStart>
  <Calendars>
    <Calendar>
      <UID>1</UID>
      <Name>Standard</Name>
      <IsBaseCalendar>1</IsBaseCalendar>
    </Calendar>
  </Calendars>
  <Tasks>
    <Task>
      <UID>1</UID>
      <ID>1</ID>
      <Name>Assigned Task</Name>
      <OutlineLevel>1</OutlineLevel>
      <OutlineNumber>1</OutlineNumber>
      <Start>2026-03-16T09:00:00</Start>
      <Finish>2026-03-17T18:00:00</Finish>
      <Duration>PT16H0M0S</Duration>
      <Milestone>0</Milestone>
      <Summary>0</Summary>
      <PercentComplete>0</PercentComplete>
    </Task>
  </Tasks>
  <Resources>
    <Resource>
      <UID>1</UID>
      <ID>1</ID>
      <Name>Worker</Name>
      <Type>1</Type>
      <WorkGroup>0</WorkGroup>
      <CalendarUID>1</CalendarUID>
      <StandardRate>8000/h</StandardRate>
      <StandardRateFormat>2</StandardRateFormat>
      <OvertimeRate>12000/h</OvertimeRate>
      <OvertimeRateFormat>2</OvertimeRateFormat>
      <CostPerUse>1500</CostPerUse>
      <Work>PT24H0M0S</Work>
      <ActualWork>PT8H0M0S</ActualWork>
      <RemainingWork>PT16H0M0S</RemainingWork>
      <Cost>180000</Cost>
      <ActualCost>60000</ActualCost>
      <RemainingCost>120000</RemainingCost>
      <PercentWorkComplete>33</PercentWorkComplete>
    </Resource>
  </Resources>
  <Assignments>
    <Assignment>
      <UID>1</UID>
      <TaskUID>1</TaskUID>
      <ResourceUID>1</ResourceUID>
      <Start>2026-03-16T09:00:00</Start>
      <Finish>2026-03-17T18:00:00</Finish>
      <StartVariance>PT1H0M0S</StartVariance>
      <FinishVariance>PT2H0M0S</FinishVariance>
      <Delay>PT3H0M0S</Delay>
      <Milestone>0</Milestone>
      <WorkContour>1</WorkContour>
      <Units>1</Units>
      <Work>PT16H0M0S</Work>
      <Cost>100000</Cost>
      <ActualCost>30000</ActualCost>
      <RemainingCost>70000</RemainingCost>
      <PercentWorkComplete>50</PercentWorkComplete>
      <OvertimeWork>PT2H0M0S</OvertimeWork>
      <ActualOvertimeWork>PT1H0M0S</ActualOvertimeWork>
      <ActualWork>PT6H0M0S</ActualWork>
      <RemainingWork>PT10H0M0S</RemainingWork>
    </Assignment>
  </Assignments>
</Project>`;

    const model = xmlTools.importMsProjectXml(xml);
    const exportedXml = xmlTools.exportMsProjectXml(model);
    const reparsedModel = xmlTools.importMsProjectXml(exportedXml);

    expect(reparsedModel.resources[0].calendarUID).toBe("1");
    expect(reparsedModel.resources[0].workGroup).toBe(0);
    expect(reparsedModel.resources[0].standardRate).toBe("8000/h");
    expect(reparsedModel.resources[0].standardRateFormat).toBe(2);
    expect(reparsedModel.resources[0].overtimeRate).toBe("12000/h");
    expect(reparsedModel.resources[0].overtimeRateFormat).toBe(2);
    expect(reparsedModel.resources[0].costPerUse).toBe(1500);
    expect(reparsedModel.resources[0].work).toBe("PT24H0M0S");
    expect(reparsedModel.resources[0].actualWork).toBe("PT8H0M0S");
    expect(reparsedModel.resources[0].remainingWork).toBe("PT16H0M0S");
    expect(reparsedModel.resources[0].cost).toBe(180000);
    expect(reparsedModel.resources[0].actualCost).toBe(60000);
    expect(reparsedModel.resources[0].remainingCost).toBe(120000);
    expect(reparsedModel.resources[0].percentWorkComplete).toBe(33);
    expect(reparsedModel.assignments[0].startVariance).toBe("PT1H0M0S");
    expect(reparsedModel.assignments[0].finishVariance).toBe("PT2H0M0S");
    expect(reparsedModel.assignments[0].delay).toBe("PT3H0M0S");
    expect(reparsedModel.assignments[0].milestone).toBe(false);
    expect(reparsedModel.assignments[0].workContour).toBe(1);
    expect(reparsedModel.assignments[0].cost).toBe(100000);
    expect(reparsedModel.assignments[0].actualCost).toBe(30000);
    expect(reparsedModel.assignments[0].remainingCost).toBe(70000);
    expect(reparsedModel.assignments[0].percentWorkComplete).toBe(50);
    expect(reparsedModel.assignments[0].overtimeWork).toBe("PT2H0M0S");
    expect(reparsedModel.assignments[0].actualOvertimeWork).toBe("PT1H0M0S");
    expect(reparsedModel.assignments[0].actualWork).toBe("PT6H0M0S");
    expect(reparsedModel.assignments[0].remainingWork).toBe("PT10H0M0S");
  });

  it("round-trips task and assignment cost fields", () => {
    const xmlTools = bootXmlModule();
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<Project xmlns="http://schemas.microsoft.com/project">
  <Name>Cost Project</Name>
  <StartDate>2026-03-16T09:00:00</StartDate>
  <FinishDate>2026-03-18T18:00:00</FinishDate>
  <ScheduleFromStart>1</ScheduleFromStart>
  <Tasks>
    <Task>
      <UID>1</UID>
      <ID>1</ID>
      <Name>Cost Task</Name>
      <OutlineLevel>1</OutlineLevel>
      <OutlineNumber>1</OutlineNumber>
      <Start>2026-03-16T09:00:00</Start>
      <Finish>2026-03-18T18:00:00</Finish>
      <Duration>PT24H0M0S</Duration>
      <Work>PT24H0M0S</Work>
      <Cost>150000</Cost>
      <ActualCost>50000</ActualCost>
      <RemainingCost>100000</RemainingCost>
      <Milestone>0</Milestone>
      <Summary>0</Summary>
      <PercentComplete>0</PercentComplete>
    </Task>
  </Tasks>
  <Resources />
  <Assignments>
    <Assignment>
      <UID>1</UID>
      <TaskUID>1</TaskUID>
      <ResourceUID>-65535</ResourceUID>
      <Start>2026-03-16T09:00:00</Start>
      <Finish>2026-03-18T18:00:00</Finish>
      <Units>1</Units>
      <Work>PT24H0M0S</Work>
      <Cost>150000</Cost>
      <ActualCost>50000</ActualCost>
      <RemainingCost>100000</RemainingCost>
    </Assignment>
  </Assignments>
</Project>`;

    const model = xmlTools.importMsProjectXml(xml);
    const exportedXml = xmlTools.exportMsProjectXml(model);
    const reparsedModel = xmlTools.importMsProjectXml(exportedXml);

    expect(reparsedModel.tasks[0].cost).toBe(150000);
    expect(reparsedModel.tasks[0].actualCost).toBe(50000);
    expect(reparsedModel.tasks[0].remainingCost).toBe(100000);
    expect(reparsedModel.assignments[0].cost).toBe(150000);
    expect(reparsedModel.assignments[0].actualCost).toBe(50000);
    expect(reparsedModel.assignments[0].remainingCost).toBe(100000);
  });

  it("round-trips task deadline and variance fields", () => {
    const xmlTools = bootXmlModule();
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<Project xmlns="http://schemas.microsoft.com/project">
  <Name>Task Variance Project</Name>
  <StartDate>2026-03-16T09:00:00</StartDate>
  <FinishDate>2026-03-18T18:00:00</FinishDate>
  <ScheduleFromStart>1</ScheduleFromStart>
  <Tasks>
    <Task>
      <UID>1</UID>
      <ID>1</ID>
      <Name>Variance Task</Name>
      <OutlineLevel>1</OutlineLevel>
      <OutlineNumber>1</OutlineNumber>
      <Start>2026-03-16T09:00:00</Start>
      <Finish>2026-03-18T18:00:00</Finish>
      <Deadline>2026-03-19T18:00:00</Deadline>
      <Duration>PT24H0M0S</Duration>
      <StartVariance>PT1H0M0S</StartVariance>
      <FinishVariance>PT2H0M0S</FinishVariance>
      <Milestone>0</Milestone>
      <Summary>0</Summary>
      <PercentComplete>0</PercentComplete>
    </Task>
  </Tasks>
  <Resources />
  <Assignments />
</Project>`;

    const model = xmlTools.importMsProjectXml(xml);
    const exportedXml = xmlTools.exportMsProjectXml(model);
    const reparsedModel = xmlTools.importMsProjectXml(exportedXml);

    expect(reparsedModel.tasks[0].deadline).toBe("2026-03-19T18:00:00");
    expect(reparsedModel.tasks[0].startVariance).toBe("PT1H0M0S");
    expect(reparsedModel.tasks[0].finishVariance).toBe("PT2H0M0S");
  });

  it("round-trips extended task work fields", () => {
    const xmlTools = bootXmlModule();
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<Project xmlns="http://schemas.microsoft.com/project">
  <Name>Task Detail Project</Name>
  <StartDate>2026-03-16T09:00:00</StartDate>
  <FinishDate>2026-03-17T18:00:00</FinishDate>
  <ScheduleFromStart>1</ScheduleFromStart>
  <Tasks>
    <Task>
      <UID>1</UID>
      <ID>1</ID>
      <Name>Detailed Task</Name>
      <OutlineLevel>1</OutlineLevel>
      <OutlineNumber>1</OutlineNumber>
      <WBS>1</WBS>
      <Type>1</Type>
      <CalendarUID>1</CalendarUID>
      <Priority>700</Priority>
      <Start>2026-03-16T09:00:00</Start>
      <Finish>2026-03-17T18:00:00</Finish>
      <Duration>PT16H0M0S</Duration>
      <Work>PT16H0M0S</Work>
      <WorkVariance>PT1H0M0S</WorkVariance>
      <TotalSlack>PT4H0M0S</TotalSlack>
      <FreeSlack>PT2H0M0S</FreeSlack>
      <RemainingWork>PT8H0M0S</RemainingWork>
      <ActualWork>PT8H0M0S</ActualWork>
      <Milestone>0</Milestone>
      <Summary>0</Summary>
      <Critical>1</Critical>
      <PercentComplete>50</PercentComplete>
      <PercentWorkComplete>50</PercentWorkComplete>
    </Task>
  </Tasks>
  <Resources />
  <Assignments />
</Project>`;

    const model = xmlTools.importMsProjectXml(xml);
    const exportedXml = xmlTools.exportMsProjectXml(model);
    const reparsedModel = xmlTools.importMsProjectXml(exportedXml);

    expect(reparsedModel.tasks[0].wbs).toBe("1");
    expect(reparsedModel.tasks[0].type).toBe(1);
    expect(reparsedModel.tasks[0].calendarUID).toBe("1");
    expect(reparsedModel.tasks[0].priority).toBe(700);
    expect(reparsedModel.tasks[0].work).toBe("PT16H0M0S");
    expect(reparsedModel.tasks[0].workVariance).toBe("PT1H0M0S");
    expect(reparsedModel.tasks[0].totalSlack).toBe("PT4H0M0S");
    expect(reparsedModel.tasks[0].freeSlack).toBe("PT2H0M0S");
    expect(reparsedModel.tasks[0].remainingWork).toBe("PT8H0M0S");
    expect(reparsedModel.tasks[0].actualWork).toBe("PT8H0M0S");
    expect(reparsedModel.tasks[0].critical).toBe(true);
    expect(reparsedModel.tasks[0].percentWorkComplete).toBe(50);
  });

  it("round-trips the hierarchy xml sample", () => {
    const xmlTools = bootXmlModule();

    const model = xmlTools.importMsProjectXml(hierarchyXml);
    const exportedXml = xmlTools.exportMsProjectXml(model);
    const reparsedModel = xmlTools.importMsProjectXml(exportedXml);

    expect(reparsedModel.tasks).toHaveLength(3);
    expect(reparsedModel.tasks[0].summary).toBe(true);
    expect(reparsedModel.tasks[1].outlineNumber).toBe("1.1");
    expect(reparsedModel.tasks[2].notes).toBe("Second child task");
    expect(xmlTools.validateProjectModel(reparsedModel)).toHaveLength(0);
  });

  it("round-trips the dependency xml sample", () => {
    const xmlTools = bootXmlModule();

    const model = xmlTools.importMsProjectXml(dependencyXml);
    const exportedXml = xmlTools.exportMsProjectXml(model);
    const reparsedModel = xmlTools.importMsProjectXml(exportedXml);

    expect(reparsedModel.calendars).toHaveLength(1);
    expect(reparsedModel.tasks[1].predecessors).toHaveLength(1);
    expect(reparsedModel.tasks[1].predecessors[0].predecessorUid).toBe("1");
    expect(reparsedModel.assignments).toHaveLength(1);
    expect(reparsedModel.assignments[0].taskUid).toBe("2");
    expect(xmlTools.validateProjectModel(reparsedModel)).toHaveLength(0);
  });

  it("imports csv with parent id into a minimal project model", () => {
    const xmlTools = bootXmlModule();
    const csv = `ID,ParentID,WBS,Name,Start,Finish,PredecessorID,Resource,PercentComplete
1,,1,Project Summary,2026-03-16T09:00:00,2026-03-20T18:00:00,,,50
2,1,1.1,Design,2026-03-16T09:00:00,2026-03-17T18:00:00,,Miku,100
3,1,1.2,Implementation,2026-03-18T09:00:00,2026-03-20T18:00:00,2,Miku|Rin,0
4,3,1.2.1,Coding,2026-03-18T09:00:00,2026-03-19T18:00:00,2,Rin,20
`;

    const model = xmlTools.importCsvParentId(csv);

    expect(model.project.name).toBe("CSV Imported Project");
    expect(model.project.calendarUID).toBe("1");
    expect(model.tasks).toHaveLength(4);
    expect(model.tasks[0].summary).toBe(true);
    expect(model.tasks[1].outlineNumber).toBe("1.1");
    expect(model.tasks[2].predecessors[0].predecessorUid).toBe("2");
    expect(model.tasks[3].outlineLevel).toBe(3);
    expect(model.resources.map((item) => item.name)).toEqual(["Miku", "Rin"]);
    expect(model.assignments).toHaveLength(4);
    expect(model.assignments[1].resourceUid).toBe("1");
    expect(model.assignments[2].resourceUid).toBe("2");
    expect(model.calendars).toHaveLength(1);
    expect(model.calendars[0].name).toBe("Standard");
    expect(model.calendars[0].weekDays).toHaveLength(7);
    expect(model.calendars[0].exceptions.length).toBeGreaterThan(0);
  });

  it("imports extended task fields in csv with parent id", () => {
    const xmlTools = bootXmlModule();
    const csv = `ID,ParentID,WBS,Name,Start,Finish,PredecessorID,Resource,PercentComplete,PercentWorkComplete,Milestone,Summary,Critical,Type,Priority,Work,CalendarUID,ConstraintType,ConstraintDate,Deadline,Notes
1,,1,Project Summary,2026-03-16T09:00:00,2026-03-20T18:00:00,,,,,1,1,0,1,500,PT40H0M0S,1,,,,Root note
2,1,1.1,Design,2026-03-16T09:00:00,2026-03-17T18:00:00,,Miku,100,100,0,0,0,1,600,PT16H0M0S,1,,,,Design done
3,1,1.2,Release,2026-03-20T18:00:00,2026-03-20T18:00:00,2,Miku,100,100,1,0,1,1,700,PT0H0M0S,2,4,2026-03-20T09:00:00,2026-03-21T18:00:00,Release gate
`;

    const model = xmlTools.importCsvParentId(csv);

    expect(model.tasks[0].summary).toBe(true);
    expect(model.tasks[0].notes).toBe("Root note");
    expect(model.tasks[1].percentWorkComplete).toBe(100);
    expect(model.tasks[1].critical).toBe(false);
    expect(model.tasks[1].priority).toBe(600);
    expect(model.tasks[1].work).toBe("PT16H0M0S");
    expect(model.tasks[2].milestone).toBe(true);
    expect(model.tasks[2].critical).toBe(true);
    expect(model.tasks[2].type).toBe(1);
    expect(model.tasks[2].calendarUID).toBe("2");
    expect(model.tasks[2].constraintType).toBe(4);
    expect(model.tasks[2].constraintDate).toBe("2026-03-20T09:00:00");
    expect(model.tasks[2].deadline).toBe("2026-03-21T18:00:00");
    expect(model.tasks[2].work).toBe("PT0H0M0S");
    expect(model.tasks[2].notes).toBe("Release gate");
  });

  it("normalizes predecessor and resource separators in csv import", () => {
    const xmlTools = bootXmlModule();
    const csv = `ID,ParentID,WBS,Name,Start,Finish,PredecessorID,Resource,PercentComplete
1,,1,Project Summary,2026-03-16T09:00:00,2026-03-20T18:00:00,,,50
2,1,1.1,Design,2026-03-16T09:00:00,2026-03-17T18:00:00,,Miku,100
3,1,1.2,Implementation,2026-03-18T09:00:00,2026-03-20T18:00:00,"2, 4; 2","Miku; Rin、Luka| Rin",0
4,1,1.3,Review,2026-03-20T09:00:00,2026-03-20T18:00:00,,Luka,0
`;

    const model = xmlTools.importCsvParentId(csv);

    expect(model.tasks[2].predecessors.map((item) => item.predecessorUid)).toEqual(["2", "4"]);
    expect(model.resources.map((item) => item.name)).toEqual(["Miku", "Rin", "Luka"]);
    expect(model.assignments.filter((item) => item.taskUid === "3")).toHaveLength(3);
  });

  it("rejects duplicate id in csv import", () => {
    const xmlTools = bootXmlModule();
    const csv = `ID,ParentID,Name
1,,Root
1,,Duplicate
`;

    expect(() => xmlTools.importCsvParentId(csv)).toThrow("CSV の ID が重複しています");
  });

  it("rejects missing parent id in csv import", () => {
    const xmlTools = bootXmlModule();
    const csv = `ID,ParentID,Name
1,,Root
2,99,Child
`;

    expect(() => xmlTools.importCsvParentId(csv)).toThrow("CSV の ParentID が既存 ID を指していません");
  });

  it("rejects cyclic parent id in csv import", () => {
    const xmlTools = bootXmlModule();
    const csv = `ID,ParentID,Name
1,2,Root
2,1,Child
`;

    expect(() => xmlTools.importCsvParentId(csv)).toThrow("CSV の ParentID が循環しています");
  });

  it("allows placeholder UID=0 and unassigned ResourceUID=-65535", () => {
    const xmlTools = bootXmlModule();
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<Project xmlns="http://schemas.microsoft.com/project">
  <Name>Placeholder Project</Name>
  <StartDate>2026-03-16T09:00:00</StartDate>
  <FinishDate>2026-03-16T18:00:00</FinishDate>
  <ScheduleFromStart>1</ScheduleFromStart>
  <Tasks>
    <Task>
      <UID>0</UID>
      <ID>0</ID>
      <Name></Name>
      <OutlineLevel>0</OutlineLevel>
      <OutlineNumber></OutlineNumber>
      <Start>2026-03-16T09:00:00</Start>
      <Finish>2026-03-16T18:00:00</Finish>
      <Duration>PT8H0M0S</Duration>
      <Milestone>0</Milestone>
      <Summary>0</Summary>
      <PercentComplete>0</PercentComplete>
    </Task>
  </Tasks>
  <Resources>
    <Resource>
      <UID>0</UID>
      <ID>0</ID>
      <Name></Name>
      <Type>1</Type>
    </Resource>
  </Resources>
  <Assignments>
    <Assignment>
      <UID>1</UID>
      <TaskUID>0</TaskUID>
      <ResourceUID>-65535</ResourceUID>
      <Start>2026-03-16T09:00:00</Start>
      <Finish>2026-03-16T18:00:00</Finish>
      <Units>1</Units>
      <Work>PT8H0M0S</Work>
    </Assignment>
  </Assignments>
</Project>`;

    const model = xmlTools.importMsProjectXml(xml);
    const issues = xmlTools.validateProjectModel(model);

    expect(issues.some((issue) => issue.message.includes("OutlineLevel"))).toBe(false);
    expect(issues.some((issue) => issue.message.includes("ResourceUID が既存 Resource"))).toBe(false);
    expect(issues.some((issue) => issue.message.includes("Resource Name が空"))).toBe(false);
  });
});
