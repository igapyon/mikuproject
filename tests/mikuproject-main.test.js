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
    <button id="exportProjectOverviewBtn" type="button">project_overview_view</button>
    <button id="exportPhaseDetailFullBtn" type="button">phase_detail_view full</button>
    <button id="exportPhaseDetailBtn" type="button">phase_detail_view</button>
    <button id="importProjectDraftFileBtn" type="button">project_draft_view JSON</button>
    <button id="importProjectDraftBtn" type="button">project_draft_view を取り込む</button>
    <button id="resetWbsHolidayDatesBtn" type="button">WBS 祝日を既定値へ戻す</button>
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
            <p class="md-note-card__text">WBS XLSX Export では、ProjectModel から補完した既定祝日と、YYYY-MM-DD 形式で指定した追加祝日を合成して WBS 日付帯へ反映します。</p>
            <p class="md-note-card__text">既定祝日は、現在の ProjectModel に含まれる Calendar.Exceptions の非稼働日例外から補完します。追加祝日は改行またはカンマ区切りで入力できます。表示期間を空欄にすると全期間、数値を入れると BaseDate 前後の日数で切り出します。営業日ベースを選ぶと土日祝を飛ばします。進捗帯も営業日基準へ切り替えられます。</p>
          </section>
          <input id="wbsDisplayDaysBeforeInput" />
          <input id="wbsDisplayDaysAfterInput" />
          <input id="wbsBusinessDayRangeInput" type="checkbox" />
          <input id="wbsBusinessDayProgressInput" type="checkbox" />
          <div id="wbsHolidaySummary"></div>
          <textarea id="wbsHolidayDatesInput"></textarea>
          <textarea id="wbsExtraHolidayDatesInput"></textarea>
        </div>
      </details>
      <details class="md-debug-accordion">
        <summary class="md-debug-accordion__summary">デバッグ情報</summary>
        <div class="md-debug-accordion__body">
          <textarea id="projectDraftImportInput"></textarea>
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

const SAMPLE_HOLIDAY_COUNT = 89;
const SAMPLE_FIRST_HOLIDAY_NAME = "元日（公式）";
const SAMPLE_FIRST_HOLIDAY_DATE = "2026-01-01";

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
    expect(document.body.textContent).toContain("既定祝日と");
    expect(document.body.textContent).toContain("追加祝日を合成");
    expect(document.body.textContent).toContain("非稼働日例外から補完");
    expect(document.body.textContent).toContain("BaseDate 前後の日数で切り出します");
    expect(document.body.textContent).toContain("営業日ベースを選ぶと土日祝を飛ばします");
    expect(document.body.textContent).toContain("進捗帯も営業日基準へ切り替えられます");
    expect(document.getElementById("wbsHolidayDatesInput").value).toBe("");
    expect(document.getElementById("wbsExtraHolidayDatesInput").value).toBe("");
    expect(document.getElementById("wbsHolidaySummary").textContent).toBe("既定祝日: 0 件");
    expect(document.getElementById("wbsDisplayDaysBeforeInput").value).toBe("");
    expect(document.getElementById("wbsDisplayDaysAfterInput").value).toBe("");
    expect(document.getElementById("wbsBusinessDayRangeInput").checked).toBe(false);
    expect(document.getElementById("wbsBusinessDayProgressInput").checked).toBe(false);
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
    expect(document.getElementById("summaryProjectName").textContent).toBe("Sample Project");
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
          planned_start: "2026-04-01T09:00:00"
        },
        tasks: [
          { uid: "draft-1", name: "要件定義", parent_uid: null, position: 0, is_summary: true },
          { uid: "draft-2", name: "ヒアリング", parent_uid: "draft-1", position: 0, planned_duration_hours: 40, planned_start: "2026-04-01T09:00:00" },
          { uid: "draft-3", name: "要件確定", parent_uid: "draft-1", position: 1, is_milestone: true, predecessors: ["draft-2"], planned_start: "2026-04-08T18:00:00", planned_finish: "2026-04-08T18:00:00" }
        ]
      }, null, 2),
      "```"
    ].join("\n");

    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();

    expect(document.getElementById("summaryProjectName").textContent).toBe("新規基幹刷新");
    expect(document.getElementById("summaryTaskCount").textContent).toBe("3");
    expect(document.getElementById("xmlInput").value).toContain("<Name>新規基幹刷新</Name>");
    expect(document.getElementById("modelOutput").value).toContain("\"uid\": \"draft-3\"");
  });

  it("parses xml into internal model summary", () => {
    bootPage();

    parseXmlViaHook();

    expect(document.getElementById("summaryProjectName").textContent).toBe("Sample Project");
    expect(document.getElementById("summaryTaskCount").textContent).toBe("3");
    expect(document.getElementById("summaryResourceCount").textContent).toBe("1");
    expect(document.getElementById("summaryAssignmentCount").textContent).toBe("2");
    expect(document.getElementById("summaryCalendarCount").textContent).toBe("2");
    expect(document.getElementById("modelOutput").value).toContain("\"name\": \"Sample Project\"");
    expect(document.getElementById("modelOutput").value).toContain("\"title\": \"Sample Project Title\"");
    expect(document.getElementById("modelOutput").value).toContain("\"author\": \"Toshiki Iga\"");
    expect(document.getElementById("modelOutput").value).toContain("\"company\": \"Local HTML Tools\"");
    expect(document.getElementById("modelOutput").value).toContain("\"creationDate\": \"2026-03-16T08:30:00\"");
    expect(document.getElementById("modelOutput").value).toContain("\"lastSaved\": \"2026-03-16T09:10:00\"");
    expect(document.getElementById("modelOutput").value).toContain("\"saveVersion\": 14");
    expect(document.getElementById("modelOutput").value).toContain("\"currentDate\": \"2026-03-16T09:00:00\"");
    expect(document.getElementById("modelOutput").value).toContain("\"defaultStartTime\": \"09:00:00\"");
    expect(document.getElementById("modelOutput").value).toContain("\"minutesPerDay\": 480");
    expect(document.getElementById("modelOutput").value).toContain("\"statusDate\": \"2026-03-19T09:00:00\"");
    expect(document.getElementById("modelOutput").value).toContain("\"weekStartDay\": 1");
    expect(document.getElementById("modelOutput").value).toContain("\"workFormat\": 2");
    expect(document.getElementById("modelOutput").value).toContain("\"durationFormat\": 7");
    expect(document.getElementById("modelOutput").value).toContain("\"currencyCode\": \"JPY\"");
    expect(document.getElementById("modelOutput").value).toContain("\"currencyDigits\": 0");
    expect(document.getElementById("modelOutput").value).toContain("\"currencySymbol\": \"¥\"");
    expect(document.getElementById("modelOutput").value).toContain("\"currencySymbolPosition\": 0");
    expect(document.getElementById("modelOutput").value).toContain("\"fyStartDate\": \"2026-04-01T00:00:00\"");
    expect(document.getElementById("modelOutput").value).toContain("\"fiscalYearStart\": true");
    expect(document.getElementById("modelOutput").value).toContain("\"criticalSlackLimit\": 0");
    expect(document.getElementById("modelOutput").value).toContain("\"defaultTaskType\": 1");
    expect(document.getElementById("modelOutput").value).toContain("\"defaultFixedCostAccrual\": 2");
    expect(document.getElementById("modelOutput").value).toContain("\"defaultStandardRate\": \"5000/h\"");
    expect(document.getElementById("modelOutput").value).toContain("\"defaultOvertimeRate\": \"7000/h\"");
    expect(document.getElementById("modelOutput").value).toContain("\"defaultTaskEVMethod\": 0");
    expect(document.getElementById("modelOutput").value).toContain("\"newTaskStartDate\": 0");
    expect(document.getElementById("modelOutput").value).toContain("\"newTasksAreManual\": false");
    expect(document.getElementById("modelOutput").value).toContain("\"newTasksEffortDriven\": true");
    expect(document.getElementById("modelOutput").value).toContain("\"newTasksEstimated\": true");
    expect(document.getElementById("modelOutput").value).toContain("\"actualsInSync\": false");
    expect(document.getElementById("modelOutput").value).toContain("\"editableActualCosts\": true");
    expect(document.getElementById("modelOutput").value).toContain("\"honorConstraints\": true");
    expect(document.getElementById("modelOutput").value).toContain("\"insertedProjectsLikeSummary\": true");
    expect(document.getElementById("modelOutput").value).toContain("\"multipleCriticalPaths\": false");
    expect(document.getElementById("modelOutput").value).toContain("\"taskUpdatesResource\": true");
    expect(document.getElementById("modelOutput").value).toContain("\"updateManuallyScheduledTasksWhenEditingLinks\": false");
    expect(document.getElementById("modelOutput").value).toContain("\"calendarUID\": \"1\"");
    expect(document.getElementById("modelOutput").value).toContain("\"fieldID\": \"188743731\"");
    expect(document.getElementById("modelOutput").value).toContain("\"fieldName\": \"Outline Code1\"");
    expect(document.getElementById("modelOutput").value).toContain("\"alias\": \"Phase\"");
    expect(document.getElementById("modelOutput").value).toContain("\"onlyTableValues\": true");
    expect(document.getElementById("modelOutput").value).toContain("\"value\": \"PLAN\"");
    expect(document.getElementById("modelOutput").value).toContain("\"description\": \"Planning\"");
    expect(document.getElementById("modelOutput").value).toContain("\"level\": 2");
    expect(document.getElementById("modelOutput").value).toContain("\"mask\": \"00\"");
    expect(document.getElementById("modelOutput").value).toContain("\"fieldName\": \"Text1\"");
    expect(document.getElementById("modelOutput").value).toContain("\"alias\": \"Owner\"");
    expect(document.getElementById("modelOutput").value).toContain("\"appendNewValues\": true");
    expect(document.getElementById("modelOutput").value).toContain("\"isBaselineCalendar\": true");
    expect(document.getElementById("modelOutput").value).toContain("\"baseCalendarUID\": \"1\"");
    expect(document.getElementById("modelOutput").value).toContain("\"dayType\": 2");
    expect(document.getElementById("modelOutput").value).toContain("\"fromTime\": \"10:00:00\"");
    expect(document.getElementById("modelOutput").value).toContain(`"name": "${SAMPLE_FIRST_HOLIDAY_NAME}"`);
    expect(document.getElementById("modelOutput").value).toContain("\"workingTimes\": [");
    expect(document.getElementById("modelOutput").value).toContain("\"name\": \"Spring Sprint\"");
    expect(document.getElementById("modelOutput").value).toContain("\"wbs\": \"1.2\"");
    expect(document.getElementById("modelOutput").value).toContain("\"priority\": 700");
    expect(document.getElementById("modelOutput").value).toContain("\"type\": 1");
    expect(document.getElementById("modelOutput").value).toContain("\"work\": \"PT24H0M0S\"");
    expect(document.getElementById("modelOutput").value).toContain("\"workVariance\": \"PT0H0M0S\"");
    expect(document.getElementById("modelOutput").value).toContain("\"totalSlack\": \"PT4H0M0S\"");
    expect(document.getElementById("modelOutput").value).toContain("\"freeSlack\": \"PT2H0M0S\"");
    expect(document.getElementById("modelOutput").value).toContain("\"cost\": 120000");
    expect(document.getElementById("modelOutput").value).toContain("\"actualCost\": 0");
    expect(document.getElementById("modelOutput").value).toContain("\"remainingCost\": 120000");
    expect(document.getElementById("modelOutput").value).toContain("\"remainingWork\": \"PT24H0M0S\"");
    expect(document.getElementById("modelOutput").value).toContain("\"actualWork\": \"PT0H0M0S\"");
    expect(document.getElementById("modelOutput").value).toContain("\"critical\": true");
    expect(document.getElementById("modelOutput").value).toContain("\"percentWorkComplete\": 0");
    expect(document.getElementById("modelOutput").value).toContain("\"predecessorUid\": \"2\"");
    expect(document.getElementById("modelOutput").value).toContain("\"name\": \"Standard\"");
    expect(document.getElementById("modelOutput").value).toContain("\"actualStart\": \"2026-03-16T09:00:00\"");
    expect(document.getElementById("modelOutput").value).toContain("\"constraintType\": 4");
    expect(document.getElementById("modelOutput").value).toContain("\"notes\": \"Implementation starts after design\"");
    expect(document.getElementById("modelOutput").value).toContain("\"deadline\": \"2026-03-21T18:00:00\"");
    expect(document.getElementById("modelOutput").value).toContain("\"startVariance\": \"PT0H0M0S\"");
    expect(document.getElementById("modelOutput").value).toContain("\"finishVariance\": \"PT0H0M0S\"");
    expect(document.getElementById("modelOutput").value).toContain("\"value\": \"Miku\"");
    expect(document.getElementById("modelOutput").value).toContain("\"number\": 0");
    expect(document.getElementById("modelOutput").value).toContain("\"unit\": 2");
    expect(document.getElementById("modelOutput").value).toContain("\"value\": \"PT8H0M0S\"");
    expect(document.getElementById("modelOutput").value).toContain("\"initials\": \"MK\"");
    expect(document.getElementById("modelOutput").value).toContain("\"group\": \"Engineering\"");
    expect(document.getElementById("modelOutput").value).toContain("\"workGroup\": 0");
    expect(document.getElementById("modelOutput").value).toContain("\"calendarUID\": \"2\"");
    expect(document.getElementById("modelOutput").value).toContain("\"standardRate\": \"5000/h\"");
    expect(document.getElementById("modelOutput").value).toContain("\"standardRateFormat\": 2");
    expect(document.getElementById("modelOutput").value).toContain("\"overtimeRate\": \"7000/h\"");
    expect(document.getElementById("modelOutput").value).toContain("\"overtimeRateFormat\": 2");
    expect(document.getElementById("modelOutput").value).toContain("\"costPerUse\": 1000");
    expect(document.getElementById("modelOutput").value).toContain("\"work\": \"PT40H0M0S\"");
    expect(document.getElementById("modelOutput").value).toContain("\"actualWork\": \"PT20H0M0S\"");
    expect(document.getElementById("modelOutput").value).toContain("\"remainingWork\": \"PT20H0M0S\"");
    expect(document.getElementById("modelOutput").value).toContain("\"cost\": 200000");
    expect(document.getElementById("modelOutput").value).toContain("\"actualCost\": 100000");
    expect(document.getElementById("modelOutput").value).toContain("\"remainingCost\": 100000");
    expect(document.getElementById("modelOutput").value).toContain("\"percentWorkComplete\": 50");
    expect(document.getElementById("modelOutput").value).toContain("\"value\": \"Platform Team\"");
    expect(document.getElementById("modelOutput").value).toContain("\"work\": \"PT40H0M0S\"");
    expect(document.getElementById("modelOutput").value).toContain("\"unit\": 2");
    expect(document.getElementById("modelOutput").value).toContain("\"start\": \"2026-03-16T09:00:00\"");
    expect(document.getElementById("modelOutput").value).toContain("\"startVariance\": \"PT0H0M0S\"");
    expect(document.getElementById("modelOutput").value).toContain("\"finishVariance\": \"PT0H0M0S\"");
    expect(document.getElementById("modelOutput").value).toContain("\"delay\": \"PT0H0M0S\"");
    expect(document.getElementById("modelOutput").value).toContain("\"milestone\": false");
    expect(document.getElementById("modelOutput").value).toContain("\"workContour\": 0");
    expect(document.getElementById("modelOutput").value).toContain("\"percentWorkComplete\": 50");
    expect(document.getElementById("modelOutput").value).toContain("\"overtimeWork\": \"PT2H0M0S\"");
    expect(document.getElementById("modelOutput").value).toContain("\"actualOvertimeWork\": \"PT1H0M0S\"");
    expect(document.getElementById("modelOutput").value).toContain("\"actualWork\": \"PT8H0M0S\"");
    expect(document.getElementById("modelOutput").value).toContain("\"remainingWork\": \"PT8H0M0S\"");
    expect(document.getElementById("modelOutput").value).toContain("\"value\": \"Design Slot\"");
    expect(document.getElementById("modelOutput").value).toContain("\"cost\": 80000");
    expect(document.getElementById("modelOutput").value).toContain("\"unit\": 2");
    expect(document.getElementById("projectPreview").textContent).toContain("Sample Project");
    expect(document.getElementById("projectPreview").textContent).toContain("Title=Sample Project Title");
    expect(document.getElementById("projectPreview").textContent).toContain("Author=Toshiki Iga / Company=Local HTML Tools");
    expect(document.getElementById("projectPreview").textContent).toContain("Calendar=1 (Standard)");
    expect(document.getElementById("projectPreview").textContent).toContain("OutlineCodes=1 / WBSMasks=2 / Ext=1");
    expect(document.getElementById("projectPreview").textContent).toContain("OutlineCode1=FieldID=188743731 / FieldName=Outline Code1 / Alias=Phase");
    expect(document.getElementById("projectPreview").textContent).toContain("WBSMask1=Level=1 / Mask=A / Length=1 / Sequence=1");
    expect(document.getElementById("projectPreview").textContent).toContain("Ext1=FieldID=188743734 / FieldName=Text1 / Alias=Owner");
    expect(document.getElementById("taskPreview").textContent).toContain("Implementation");
    expect(document.getElementById("taskPreview").textContent).toContain("Calendar=1 (Standard)");
    expect(document.getElementById("taskPreview").textContent).toContain("Ext=1 / Baselines=1 / Timephased=1");
    expect(document.getElementById("taskPreview").textContent).toContain("Ext1=FieldID=188743734 / Value=Miku");
    expect(document.getElementById("taskPreview").textContent).toContain("Baseline1=#0 2026-03-16T09:00:00 -> 2026-03-17T18:00:00");
    expect(document.getElementById("taskPreview").textContent).toContain("Timephased1=Type=1 2026-03-16T09:00:00 -> 2026-03-16T18:00:00");
    expect(document.getElementById("resourcePreview").textContent).toContain("Engineering");
    expect(document.getElementById("resourcePreview").textContent).toContain("Calendar=2 (Development)");
    expect(document.getElementById("resourcePreview").textContent).toContain("Ext=1 / Baselines=1 / Timephased=1");
    expect(document.getElementById("resourcePreview").textContent).toContain("Ext1=FieldID=188743737 / Value=Platform Team");
    expect(document.getElementById("resourcePreview").textContent).toContain("Baseline1=#0 2026-03-16T09:00:00 -> 2026-03-20T18:00:00");
    expect(document.getElementById("resourcePreview").textContent).toContain("Timephased1=Type=1 2026-03-16T09:00:00 -> 2026-03-16T18:00:00");
    expect(document.getElementById("assignmentPreview").textContent).toContain("Task=2 (Design)");
    expect(document.getElementById("assignmentPreview").textContent).toContain("Resource=1 (Miku)");
    expect(document.getElementById("assignmentPreview").textContent).toContain("Ext=1 / Baselines=1 / Timephased=1");
    expect(document.getElementById("assignmentPreview").textContent).toContain("Ext1=FieldID=255852547 / Value=Design Slot");
    expect(document.getElementById("assignmentPreview").textContent).toContain("Baseline1=#0 2026-03-16T09:00:00 -> 2026-03-17T18:00:00");
    expect(document.getElementById("assignmentPreview").textContent).toContain("Timephased1=Type=1 2026-03-16T09:00:00 -> 2026-03-16T18:00:00");
    expect(document.getElementById("calendarPreview").textContent).toContain("Standard");
    expect(document.getElementById("calendarPreview").textContent).toContain("Base=1 / Baseline=1 / BaseCalendarUID=-");
    expect(document.getElementById("calendarPreview").textContent).toContain(`WeekDays=1 / Exceptions=${SAMPLE_HOLIDAY_COUNT} / WorkWeeks=0`);
    expect(document.getElementById("calendarPreview").textContent).toContain("Refs=Project=1 / Tasks=3 / Resources=0 / BaseOf=1");
    expect(document.getElementById("calendarPreview").textContent).toContain("WeekDay1=DayType=2 / Working=1 / Times=09:00:00-12:00:00, 13:00:00-18:00:00");
    expect(document.getElementById("calendarPreview").textContent).toContain(`Exception1=${SAMPLE_FIRST_HOLIDAY_NAME} ${SAMPLE_FIRST_HOLIDAY_DATE}T00:00:00 -> ${SAMPLE_FIRST_HOLIDAY_DATE}T23:59:59 / Working=0`);
    expect(document.getElementById("calendarPreview").textContent).toContain("Development");
    expect(document.getElementById("calendarPreview").textContent).toContain("Base=0 / Baseline=0 / BaseCalendarUID=1");
    expect(document.getElementById("calendarPreview").textContent).toContain("Refs=Project=0 / Tasks=0 / Resources=1 / BaseOf=0");
    expect(document.getElementById("calendarPreview").textContent).toContain("WorkWeek1=Spring Sprint 2026-03-16T00:00:00 -> 2026-03-31T23:59:59 / WeekDays=1");
  });

  it("exports xml from the current model", () => {
    bootPage();

    parseXmlViaHook();
    document.getElementById("downloadXmlBtn").click();

    const xmlText = document.getElementById("xmlInput").value;
    expect(xmlText).toContain("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
    expect(xmlText).toContain("\n<Project xmlns=\"http://schemas.microsoft.com/project\">\n");
    expect(xmlText).toContain("<Title>Sample Project Title</Title>");
    expect(xmlText).toContain("<Company>Local HTML Tools</Company>");
    expect(xmlText).toContain("<Author>Toshiki Iga</Author>");
    expect(xmlText).toContain("<CreationDate>2026-03-16T08:30:00</CreationDate>");
    expect(xmlText).toContain("<LastSaved>2026-03-16T09:10:00</LastSaved>");
    expect(xmlText).toContain("<SaveVersion>14</SaveVersion>");
    expect(xmlText).toContain("<CurrentDate>2026-03-16T09:00:00</CurrentDate>");
    expect(xmlText).toContain("<StatusDate>2026-03-19T09:00:00</StatusDate>");
    expect(xmlText).toContain("<WeekStartDay>1</WeekStartDay>");
    expect(xmlText).toContain("<WorkFormat>2</WorkFormat>");
    expect(xmlText).toContain("<DurationFormat>7</DurationFormat>");
    expect(xmlText).toContain("<CurrencyCode>JPY</CurrencyCode>");
    expect(xmlText).toContain("<CurrencyDigits>0</CurrencyDigits>");
    expect(xmlText).toContain("<CurrencySymbol>¥</CurrencySymbol>");
    expect(xmlText).toContain("<CurrencySymbolPosition>0</CurrencySymbolPosition>");
    expect(xmlText).toContain("<FYStartDate>2026-04-01T00:00:00</FYStartDate>");
    expect(xmlText).toContain("<FiscalYearStart>1</FiscalYearStart>");
    expect(xmlText).toContain("<CriticalSlackLimit>0</CriticalSlackLimit>");
    expect(xmlText).toContain("<DefaultTaskType>1</DefaultTaskType>");
    expect(xmlText).toContain("<DefaultFixedCostAccrual>2</DefaultFixedCostAccrual>");
    expect(xmlText).toContain("<DefaultStandardRate>5000/h</DefaultStandardRate>");
    expect(xmlText).toContain("<DefaultOvertimeRate>7000/h</DefaultOvertimeRate>");
    expect(xmlText).toContain("<DefaultTaskEVMethod>0</DefaultTaskEVMethod>");
    expect(xmlText).toContain("<NewTaskStartDate>0</NewTaskStartDate>");
    expect(xmlText).toContain("<NewTasksAreManual>0</NewTasksAreManual>");
    expect(xmlText).toContain("<NewTasksEffortDriven>1</NewTasksEffortDriven>");
    expect(xmlText).toContain("<NewTasksEstimated>1</NewTasksEstimated>");
    expect(xmlText).toContain("<ActualsInSync>0</ActualsInSync>");
    expect(xmlText).toContain("<EditableActualCosts>1</EditableActualCosts>");
    expect(xmlText).toContain("<HonorConstraints>1</HonorConstraints>");
    expect(xmlText).toContain("<InsertedProjectsLikeSummary>1</InsertedProjectsLikeSummary>");
    expect(xmlText).toContain("<MultipleCriticalPaths>0</MultipleCriticalPaths>");
    expect(xmlText).toContain("<TaskUpdatesResource>1</TaskUpdatesResource>");
    expect(xmlText).toContain("<UpdateManuallyScheduledTasksWhenEditingLinks>0</UpdateManuallyScheduledTasksWhenEditingLinks>");
    expect(xmlText).toContain("<OutlineCodes>");
    expect(xmlText).toContain("<FieldID>188743731</FieldID>");
    expect(xmlText).toContain("<FieldName>Outline Code1</FieldName>");
    expect(xmlText).toContain("<Alias>Phase</Alias>");
    expect(xmlText).toContain("<OnlyTableValues>1</OnlyTableValues>");
    expect(xmlText).toContain("<Values>");
    expect(xmlText).toContain("<Value>PLAN</Value>");
    expect(xmlText).toContain("<Description>Planning</Description>");
    expect(xmlText).toContain("<WBSMasks>");
    expect(xmlText).toContain("<WBSMask>");
    expect(xmlText).toContain("<Level>2</Level>");
    expect(xmlText).toContain("<Mask>00</Mask>");
    expect(xmlText).toContain("<ExtendedAttributes>");
    expect(xmlText).toContain("<ExtendedAttribute>");
    expect(xmlText).toContain("<FieldName>Text1</FieldName>");
    expect(xmlText).toContain("<Alias>Owner</Alias>");
    expect(xmlText).toContain("<AppendNewValues>1</AppendNewValues>");
    expect(xmlText).toContain("<WBS>1.2</WBS>");
    expect(xmlText).toContain("<Priority>700</Priority>");
    expect(xmlText).toContain("<CalendarUID>1</CalendarUID>");
    expect(xmlText).toContain("<Work>PT24H0M0S</Work>");
    expect(xmlText).toContain("<WorkVariance>PT0H0M0S</WorkVariance>");
    expect(xmlText).toContain("<TotalSlack>PT4H0M0S</TotalSlack>");
    expect(xmlText).toContain("<FreeSlack>PT2H0M0S</FreeSlack>");
    expect(xmlText).toContain("<Cost>120000</Cost>");
    expect(xmlText).toContain("<ActualCost>0</ActualCost>");
    expect(xmlText).toContain("<RemainingCost>120000</RemainingCost>");
    expect(xmlText).toContain("<RemainingWork>PT24H0M0S</RemainingWork>");
    expect(xmlText).toContain("<ActualWork>PT0H0M0S</ActualWork>");
    expect(xmlText).toContain("<PercentWorkComplete>0</PercentWorkComplete>");
    expect(xmlText).toContain("<DefaultStartTime>09:00:00</DefaultStartTime>");
    expect(xmlText).toContain("<MinutesPerDay>480</MinutesPerDay>");
    expect(xmlText).toContain("<CalendarUID>1</CalendarUID>");
    expect(xmlText).toContain("<Calendars>");
    expect(xmlText).toContain("\n  <Calendars>\n");
    expect(xmlText).toContain("<IsBaselineCalendar>1</IsBaselineCalendar>");
    expect(xmlText).toContain("<BaseCalendarUID>1</BaseCalendarUID>");
    expect(xmlText).toContain("<Exceptions>");
    expect(xmlText).toContain("<WorkWeeks>");
    expect(xmlText).toContain(`<Name>${SAMPLE_FIRST_HOLIDAY_NAME}</Name>`);
    expect(xmlText).toContain("<WorkingTimes>");
    expect(xmlText).toContain("<Name>Spring Sprint</Name>");
    expect(xmlText).toContain("<WeekDays>");
    expect(xmlText).toContain("<DayType>2</DayType>");
    expect(xmlText).toContain("<FromTime>10:00:00</FromTime>");
    expect(xmlText).toContain("<Tasks>");
    expect(xmlText).toContain("<Assignments>");
    expect(xmlText).toContain("<LinkLag>PT0H0M0S</LinkLag>");
    expect(xmlText).toContain("<ActualStart>2026-03-16T09:00:00</ActualStart>");
    expect(xmlText).toContain("<Deadline>2026-03-21T18:00:00</Deadline>");
    expect(xmlText).toContain("<StartVariance>PT0H0M0S</StartVariance>");
    expect(xmlText).toContain("<FinishVariance>PT0H0M0S</FinishVariance>");
    expect(xmlText).toContain("<ConstraintType>4</ConstraintType>");
    expect(xmlText).toContain("<Notes>Implementation starts after design</Notes>");
    expect(xmlText).toContain("<ExtendedAttribute>");
    expect(xmlText).toContain("<FieldID>188743734</FieldID>");
    expect(xmlText).toContain("<Value>Miku</Value>");
    expect(xmlText).toContain("<Baseline>");
    expect(xmlText).toContain("<Number>0</Number>");
    expect(xmlText).toContain("<Work>PT16H0M0S</Work>");
    expect(xmlText).toContain("<TimephasedData>");
    expect(xmlText).toContain("<Unit>2</Unit>");
    expect(xmlText).toContain("<Value>PT8H0M0S</Value>");
    expect(xmlText).toContain("<Critical>1</Critical>");
    expect(xmlText).toContain("<Initials>MK</Initials>");
    expect(xmlText).toContain("<Group>Engineering</Group>");
    expect(xmlText).toContain("<WorkGroup>0</WorkGroup>");
    expect(xmlText).toContain("<StandardRate>5000/h</StandardRate>");
    expect(xmlText).toContain("<StandardRateFormat>2</StandardRateFormat>");
    expect(xmlText).toContain("<OvertimeRate>7000/h</OvertimeRate>");
    expect(xmlText).toContain("<OvertimeRateFormat>2</OvertimeRateFormat>");
    expect(xmlText).toContain("<CostPerUse>1000</CostPerUse>");
    expect(xmlText).toContain("<Work>PT40H0M0S</Work>");
    expect(xmlText).toContain("<ActualWork>PT20H0M0S</ActualWork>");
    expect(xmlText).toContain("<RemainingWork>PT20H0M0S</RemainingWork>");
    expect(xmlText).toContain("<Cost>200000</Cost>");
    expect(xmlText).toContain("<ActualCost>100000</ActualCost>");
    expect(xmlText).toContain("<RemainingCost>100000</RemainingCost>");
    expect(xmlText).toContain("<PercentWorkComplete>50</PercentWorkComplete>");
    expect(xmlText).toContain("<ExtendedAttribute>");
    expect(xmlText).toContain("<FieldID>188743737</FieldID>");
    expect(xmlText).toContain("<Value>Platform Team</Value>");
    expect(xmlText).toContain("<Baseline>");
    expect(xmlText).toContain("<Work>PT40H0M0S</Work>");
    expect(xmlText).toContain("<TimephasedData>");
    expect(xmlText).toContain("<Unit>2</Unit>");
    expect(xmlText).toContain("<Start>2026-03-16T09:00:00</Start>");
    expect(xmlText).toContain("<StartVariance>PT0H0M0S</StartVariance>");
    expect(xmlText).toContain("<FinishVariance>PT0H0M0S</FinishVariance>");
    expect(xmlText).toContain("<Delay>PT0H0M0S</Delay>");
    expect(xmlText).toContain("<Milestone>0</Milestone>");
    expect(xmlText).toContain("<WorkContour>0</WorkContour>");
    expect(xmlText).toContain("<PercentWorkComplete>50</PercentWorkComplete>");
    expect(xmlText).toContain("<OvertimeWork>PT2H0M0S</OvertimeWork>");
    expect(xmlText).toContain("<ActualOvertimeWork>PT1H0M0S</ActualOvertimeWork>");
    expect(xmlText).toContain("<ActualWork>PT8H0M0S</ActualWork>");
    expect(xmlText).toContain("<RemainingWork>PT8H0M0S</RemainingWork>");
    expect(xmlText).toContain("<FieldID>255852547</FieldID>");
    expect(xmlText).toContain("<Value>Design Slot</Value>");
    expect(xmlText).toContain("<Baseline>");
    expect(xmlText).toContain("<Number>0</Number>");
    expect(xmlText).toContain("<TimephasedData>");
    expect(xmlText).toContain("<Unit>2</Unit>");
  });

  it("exports mermaid gantt from the current model", async () => {
    bootPage();

    parseXmlViaHook();
    await exportMermaidViaHook();

    const mermaidText = document.getElementById("mermaidOutput").value;
    expect(mermaidText).toContain("gantt");
    expect(mermaidText).toContain("title Sample Project");
    expect(mermaidText).toContain("dateFormat YYYY-MM-DDTHH:mm:ss");
    expect(mermaidText).toContain("section Project Summary");
    expect(mermaidText).toContain("Design :done, task_2, 2026-03-16T09:00:00, 2026-03-17T18:00:00");
    expect(mermaidText).toContain("Implementation :crit, task_3, after task_2, 24h");
    expect(mermaidText).toContain("%% dependency(native): Implementation after Design (task_3 after task_2)");
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
    expect(document.getElementById("modelOutput").value).toContain("\"name\": \"CSV Imported Project\"");
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
    expect(document.getElementById("modelOutput").value).toContain("\"name\": \"Imported\"");
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
    expect(document.getElementById("wbsExtraHolidayDatesInput").value).toBe("");
    expect(exportSpy.mock.calls.at(-1)?.[1]).toEqual({
      holidayDates: defaultHolidayDates,
      displayDaysBeforeBaseDate: undefined,
      displayDaysAfterBaseDate: undefined,
      useBusinessDaysForDisplayRange: false,
      useBusinessDaysForProgressBand: false
    });
    expect(document.getElementById("xmlInput").value).not.toContain("<!-- edited -->");
    expect(document.getElementById("xmlSaveState").textContent).toContain("XML 保存状態: 保存済み (2026-03-16 23:12)");
    expect(document.getElementById("statusMessage").textContent).toContain("WBS XLSX ファイルをエクスポートしました");
    expect(document.getElementById("statusMessage").textContent).toContain(`祝日 ${SAMPLE_HOLIDAY_COUNT} 件`);
  });

  it("downloads current wbs xlsx with configured holidays", async () => {
    bootPage();
    const exportSpy = vi.spyOn(globalThis.__mikuprojectWbsXlsx, "exportWbsWorkbook");

    parseXmlViaHook();
    document.getElementById("wbsExtraHolidayDatesInput").value = "2026-03-20, 2026-03-20\n2026-03-21";
    document.getElementById("exportWbsXlsxBtn").click();
    const defaultHolidayDates = getDefaultSampleHolidayDates();

    expect(exportSpy).toHaveBeenCalled();
    expect(exportSpy.mock.calls.at(-1)?.[1]).toEqual({
      holidayDates: [...defaultHolidayDates, "2026-03-21"],
      displayDaysBeforeBaseDate: undefined,
      displayDaysAfterBaseDate: undefined,
      useBusinessDaysForDisplayRange: false,
      useBusinessDaysForProgressBand: false
    });
    const workbook = exportSpy.mock.results.at(-1)?.value;
    const sheet = workbook.sheets[0];
    const projectInfoHeaderIndex = sheet.rows.findIndex((row) => row.cells[0]?.value === "プロジェクト");
    expect(sheet.rows[projectInfoHeaderIndex + 7].cells[2].value).toBe(SAMPLE_HOLIDAY_COUNT + 1);
    const headerRowIndex = sheet.rows.findIndex((row) => row.cells[0]?.value === "UID");
    const dateRowIndex = headerRowIndex - 1;
    const holidayColumnIndex = sheet.rows[dateRowIndex].cells.findIndex((cell) => cell.value === "3/20");
    expect(sheet.rows[dateRowIndex].cells[holidayColumnIndex].fillColor).toBe("#FCE4EC");
    expect(document.getElementById("statusMessage").textContent).toContain(`祝日 ${SAMPLE_HOLIDAY_COUNT + 1} 件`);
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
      useBusinessDaysForDisplayRange: false,
      useBusinessDaysForProgressBand: false
    });
    expect(document.getElementById("statusMessage").textContent).toContain("表示期間 暦日 基準日前 1 日, 基準日後 2 日");
  });

  it("downloads current wbs xlsx with business-day display range", async () => {
    bootPage();
    const exportSpy = vi.spyOn(globalThis.__mikuprojectWbsXlsx, "exportWbsWorkbook");

    parseXmlViaHook();
    document.getElementById("wbsDisplayDaysBeforeInput").value = "1";
    document.getElementById("wbsDisplayDaysAfterInput").value = "2";
    document.getElementById("wbsBusinessDayRangeInput").checked = true;
    document.getElementById("exportWbsXlsxBtn").click();
    const defaultHolidayDates = getDefaultSampleHolidayDates();

    expect(exportSpy.mock.calls.at(-1)?.[1]).toEqual({
      holidayDates: defaultHolidayDates,
      displayDaysBeforeBaseDate: 1,
      displayDaysAfterBaseDate: 2,
      useBusinessDaysForDisplayRange: true,
      useBusinessDaysForProgressBand: false
    });
    expect(document.getElementById("statusMessage").textContent).toContain("表示期間 営業日 基準日前 1 日, 基準日後 2 日");
  });

  it("downloads current wbs xlsx with business-day progress band", async () => {
    bootPage();
    const exportSpy = vi.spyOn(globalThis.__mikuprojectWbsXlsx, "exportWbsWorkbook");

    parseXmlViaHook();
    document.getElementById("wbsBusinessDayProgressInput").checked = true;
    document.getElementById("exportWbsXlsxBtn").click();
    const defaultHolidayDates = getDefaultSampleHolidayDates();

    expect(exportSpy.mock.calls.at(-1)?.[1]).toEqual({
      holidayDates: defaultHolidayDates,
      displayDaysBeforeBaseDate: undefined,
      displayDaysAfterBaseDate: undefined,
      useBusinessDaysForDisplayRange: false,
      useBusinessDaysForProgressBand: true
    });
    expect(document.getElementById("statusMessage").textContent).toContain("進捗帯 営業日");
  });

  it("fills wbs holiday input from model defaults when xml is parsed", () => {
    bootPage();

    parseXmlViaHook();
    const defaultHolidayDates = getDefaultSampleHolidayDates();

    expect(document.getElementById("wbsHolidayDatesInput").value.split("\n")).toEqual(defaultHolidayDates);
    expect(document.getElementById("wbsExtraHolidayDatesInput").value).toBe("");
    expect(document.getElementById("wbsHolidaySummary").textContent).toContain(`既定祝日: ${SAMPLE_HOLIDAY_COUNT} 件`);
    expect(document.getElementById("wbsHolidaySummary").textContent).toContain("2026-03-20");
  });

  it("resets wbs holiday input back to model defaults", () => {
    bootPage();

    parseXmlViaHook();
    document.getElementById("wbsExtraHolidayDatesInput").value = "2026-03-25\n2026-03-26";
    document.getElementById("resetWbsHolidayDatesBtn").click();
    const defaultHolidayDates = getDefaultSampleHolidayDates();

    expect(document.getElementById("wbsHolidayDatesInput").value.split("\n")).toEqual(defaultHolidayDates);
    expect(document.getElementById("wbsExtraHolidayDatesInput").value).toBe("");
    expect(document.getElementById("wbsHolidaySummary").textContent).toContain(`既定祝日: ${SAMPLE_HOLIDAY_COUNT} 件`);
    expect(document.getElementById("statusMessage").textContent).toContain("WBS 祝日入力を既定値へ戻しました");
    expect(document.getElementById("statusMessage").textContent).toContain(`${SAMPLE_HOLIDAY_COUNT} 件`);
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
    tasksSheet.rows[4].cells[2].value = "Design Imported From XLSX";
    tasksSheet.rows[4].cells[9].value = 77;
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

    expect(document.getElementById("modelOutput").value).toContain("\"name\": \"Design Imported From XLSX\"");
    expect(document.getElementById("modelOutput").value).toContain("\"percentComplete\": 77");
    expect(document.getElementById("xmlInput").value).toContain("<Name>Design Imported From XLSX</Name>");
    expect(document.getElementById("statusMessage").textContent).toContain("XLSX を読み込んで 2 件の変更を反映しました");
    expect(document.getElementById("statusMessage").textContent).toContain("XML Export で保存できます");
    expect(document.getElementById("xmlSaveState").textContent).toContain("XML 保存状態: 未保存");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Tasks");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Tasks 1");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("変更なし: Project, Resources, Assignments, Calendars");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("UID=2 Design");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Name: Design -> Design Imported From XLSX");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("PercentComplete: 100 -> 77");
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
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("UID=project Sample Project");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Name: Sample Project -> Project From XLSX");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("MinutesPerDay: 480 -> 420");
    expect(document.querySelectorAll("#xlsxImportSummary .md-xlsx-summary__section")).toHaveLength(1);
    expect(document.querySelectorAll("#xlsxImportSummary .md-xlsx-summary__item")).toHaveLength(1);
  });

  it("renders project import summary content without xlsx import wiring", () => {
    bootPage();

    getMainHooks().renderXlsxImportSummary([
      { scope: "project", uid: "project", label: "Sample Project", field: "CalendarUID", before: "1", after: "2" },
      { scope: "project", uid: "project", label: "Sample Project", field: "ScheduleFromStart", before: true, after: false },
      { scope: "project", uid: "project", label: "Sample Project", field: "Author", before: "Toshiki Iga", after: "Author From XLSX" }
    ]);

    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Project 1");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("変更なし: Tasks, Resources, Assignments, Calendars");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("UID=project Sample Project");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("CalendarUID: 1 -> 2");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("ScheduleFromStart: true -> false");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Author: Toshiki Iga -> Author From XLSX");
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
      { scope: "calendars", uid: "2", label: "Development", field: "IsBaseCalendar", before: false, after: true },
      { scope: "tasks", uid: "2", label: "Design", field: "Start", before: "2026-03-16T09:00:00", after: "2026-03-17T09:00:00" },
      { scope: "tasks", uid: "2", label: "Design", field: "Finish", before: "2026-03-17T18:00:00", after: "2026-03-18T18:00:00" },
      { scope: "resources", uid: "1", label: "Miku", field: "Name", before: "Miku", after: "Miku Renamed" },
      { scope: "assignments", uid: "1", label: "TaskUID=2", field: "Work", before: "PT16H0M0S", after: "PT12H0M0S" }
    ]);

    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Tasks 1");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Resources 1");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Assignments 1");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("Calendars 2");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("変更なし: Project");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("UID=1 Standard");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("UID=2 Development");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("UID=2 Design");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("UID=1 Miku");
    expect(document.getElementById("xlsxImportSummary").textContent).toContain("UID=1 TaskUID=2");
    expect(document.querySelectorAll("#xlsxImportSummary .md-xlsx-summary__section")).toHaveLength(4);
    expect(document.querySelectorAll("#xlsxImportSummary .md-xlsx-summary__item")).toHaveLength(5);
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

    document.getElementById("xmlInput").value = document.getElementById("xmlInput").value.replace(
      "<ResourceUID>1</ResourceUID>",
      "<ResourceUID>99</ResourceUID>"
    );
    parseXmlViaHook();
    document.getElementById("roundTripBtn").click();

    expect(document.getElementById("statusMessage").textContent).toContain("Assignment ResourceUID");
    expect(document.getElementById("validationIssues").textContent).toContain("Assignment ResourceUID");
    expect(document.getElementById("validationIssues").textContent).toContain("UID=1");
    expect(document.getElementById("validationIssues").textContent).toContain("TaskUID=2");
    expect(document.getElementById("validationIssues").textContent).toContain("Design");
    expect(document.getElementById("validationIssues").textContent).toContain("ResourceUID=99");
  }, 10000);

  it("reports validation error when project calendar does not exist", () => {
    bootPage();

    document.getElementById("xmlInput").value = document.getElementById("xmlInput").value.replace(
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

    document.getElementById("xmlInput").value = document.getElementById("xmlInput").value.replace(
      "<CalendarUID>1</CalendarUID>\n      <Priority>700</Priority>",
      "<CalendarUID>99</CalendarUID>\n      <Priority>700</Priority>"
    );
    parseXmlViaHook();

    expect(document.getElementById("statusMessage").textContent).toContain("検証で");
    expect(document.getElementById("validationIssues").textContent).toContain("Task CalendarUID");
    expect(document.getElementById("validationIssues").textContent).toContain("UID=3");
    expect(document.getElementById("validationIssues").textContent).toContain("Implementation");
  });

  it("reports validation warning when percent complete is out of range", () => {
    bootPage();

    document.getElementById("xmlInput").value = document.getElementById("xmlInput").value.replace(
      "<PercentComplete>100</PercentComplete>",
      "<PercentComplete>120</PercentComplete>"
    );
    parseXmlViaHook();

    expect(document.getElementById("validationIssues").textContent).toContain("PercentComplete");
  });

  it("reports validation warning when task start is after finish", () => {
    bootPage();

    document.getElementById("xmlInput").value = document.getElementById("xmlInput").value.replace(
      "<Start>2026-03-18T09:00:00</Start>\n      <Finish>2026-03-20T18:00:00</Finish>",
      "<Start>2026-03-21T09:00:00</Start>\n      <Finish>2026-03-20T18:00:00</Finish>"
    );
    parseXmlViaHook();

    expect(document.getElementById("validationIssues").textContent).toContain("Task Start が Finish より後");
  });

  it("reports validation error when predecessor references a missing task", () => {
    bootPage();

    document.getElementById("xmlInput").value = document.getElementById("xmlInput").value.replace(
      "<PredecessorUID>2</PredecessorUID>",
      "<PredecessorUID>99</PredecessorUID>"
    );
    parseXmlViaHook();
    document.getElementById("roundTripBtn").click();

    expect(document.getElementById("statusMessage").textContent).toContain("PredecessorUID");
    expect(document.getElementById("validationIssues").textContent).toContain("PredecessorUID");
    expect(document.getElementById("validationIssues").textContent).toContain("UID=3");
    expect(document.getElementById("validationIssues").textContent).toContain("Implementation");
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
    expect(model.tasks).toHaveLength(4);
    expect(model.tasks[0].summary).toBe(true);
    expect(model.tasks[1].outlineNumber).toBe("1.1");
    expect(model.tasks[2].predecessors[0].predecessorUid).toBe("2");
    expect(model.tasks[3].outlineLevel).toBe(3);
    expect(model.resources.map((item) => item.name)).toEqual(["Miku", "Rin"]);
    expect(model.assignments).toHaveLength(4);
    expect(model.assignments[1].resourceUid).toBe("1");
    expect(model.assignments[2].resourceUid).toBe("2");
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
