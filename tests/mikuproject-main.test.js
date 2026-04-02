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
const projectPatchJsonCode = readFileSync(
  path.resolve(__dirname, "../src/js/project-patch-json.js"),
  "utf8"
);
const wbsXlsxCode = readFileSync(
  path.resolve(__dirname, "../src/js/wbs-xlsx.js"),
  "utf8"
);
const wbsMarkdownCode = readFileSync(
  path.resolve(__dirname, "../src/js/wbs-markdown.js"),
  "utf8"
);
const nativeSvgCode = readFileSync(
  path.resolve(__dirname, "../src/js/wbs-svg.js"),
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
function mountFullDom() {
  document.body.innerHTML = `
    <button id="importFileBtn" type="button">Load from file</button>
    <button id="loadSampleBtn" type="button">サンプル</button>
    <button id="downloadAllOutputsBtn" type="button">All</button>
    <button id="exportXlsxBtn" type="button">XLSX</button>
    <button id="exportWorkbookJsonBtn" type="button">JSON</button>
    <button id="exportCsvBtn" type="button">CSV</button>
    <button id="exportWbsXlsxBtn" type="button">WBS XLSX</button>
    <button id="exportWbsMdBtn" type="button">WBS Markdown</button>
    <button id="downloadWeeklySvgBtn" type="button" disabled>Weekly SVG</button>
    <button id="downloadMonthlyCalendarSvgBtn" type="button" disabled>Monthly Calendar SVG</button>
    <button id="exportMermaidMdBtn" type="button">Mermaid</button>
    <button id="downloadSvgBtn" type="button" disabled>Daily SVG</button>
    <button id="previewDailySvgBtn" type="button">Daily SVG</button>
    <button id="previewWeeklySvgBtn" type="button">Weekly SVG</button>
    <button id="previewMonthlySvgBtn" type="button">Monthly Calendar SVG</button>
    <button id="exportAiBundleBtn" type="button">project_overview + all phase_detail full</button>
    <button id="exportProjectOverviewBtn" type="button">project_overview_view</button>
    <button id="exportTaskEditBtn" type="button">task_edit_view</button>
    <button id="exportPhaseDetailFullBtn" type="button">phase_detail_view full</button>
    <button id="exportPhaseDetailBtn" type="button">phase_detail_view</button>
    <button id="loadProjectDraftSampleBtn" type="button">サンプル draft</button>
    <button id="importProjectDraftBtn" type="button">project_draft_view を取り込む</button>
    <button id="downloadXmlBtn" type="button">MS Project XML</button>
    <input id="importFileInput" type="file" />
    <input id="phaseDetailUidInput" type="text" />
    <input id="taskEditUidInput" type="text" />
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
      <div class="md-flow-section__header">
        <h2 class="md-flow-section__title">
          <span class="md-flow-section__step">1</span>
          <span>Input</span>
          <lht-help-tooltip label="Input の説明" wide>
            <p>MS Project XML、XLSX、mikuproject_workbook_json (.json)、生成AI 向け編集用 JSON (.editjson)、CSV を読み込みます。</p>
          </lht-help-tooltip>
        </h2>
      </div>
      <span class="md-button-with-help">
        <button id="importFileBtnWithHelp" type="button">Load from file help host</button>
        <span class="md-button-help-anchor">
          <lht-help-tooltip label="Load from file の説明" placement="right" wide>
            <p>Load from file は .xml / .xlsx / .json / .editjson / .csv を拡張子で判別して取り込みます。</p>
          </lht-help-tooltip>
        </span>
      </span>
      <div hidden>
          <textarea id="xmlInput"></textarea>
      </div>
      <details class="md-note-accordion">
        <summary class="md-note-accordion__summary">新規生成AI連携</summary>
        <div class="md-note-accordion__body">
          <section class="md-note-card md-note-card--accent">
            <div class="md-panel-actions">
              <button id="copyAiPromptBtnPane" type="button">mikuproject 用の生成AIプロンプト</button>
              <button id="importProjectDraftBtn" type="button">貼り付けた JSON を取り込む</button>
              <lht-help-tooltip label="新規生成AI連携の説明" wide>
                <p>(i) まず生成AIに 生成AIプロンプト を読み込ませます。</p>
              </lht-help-tooltip>
              <button id="loadProjectDraftSampleBtn" type="button">サンプル</button>
            </div>
            <template id="aiPromptTemplate"># mikuproject AI JSON Spec

あなたはこれから mikuproject とやりとりします。</template>
            <div class="md-form-grid">
              <textarea id="projectDraftImportInput"></textarea>
            </div>
          </section>
        </div>
      </details>
      <div id="xmlSaveState" class="md-save-state md-save-state--dirty">XML 保存状態: 未保存</div>
    </section>
    <section id="tabPanelTransform" class="md-flow-section md-tab-panel" data-tab-panel="transform" hidden>
      <div class="md-flow-section__header">
        <h2 class="md-flow-section__title">
          <span class="md-flow-section__step">2</span>
          <span>Overview</span>
          <lht-help-tooltip label="Overview の説明" wide>
            <p>内部モデル、検証結果、SVG preview を確認します。取込後の warning と差分要約もここに表示します。</p>
          </lht-help-tooltip>
        </h2>
      </div>
      <div id="summaryProjectName"></div>
      <div id="summaryTaskCount"></div>
      <div id="summaryResourceCount"></div>
      <div id="summaryAssignmentCount"></div>
      <div id="summaryCalendarCount"></div>
      <div id="nativeSvgPreview"></div>
      <details class="md-debug-accordion">
        <summary class="md-debug-accordion__summary">デバッグ情報</summary>
        <div class="md-debug-accordion__body">
          <button id="roundTripBtn" type="button">再読込テスト</button>
          <textarea id="modelOutput"></textarea>
          <div id="projectPreview"></div>
          <div id="taskPreview"></div>
          <div id="resourcePreview"></div>
          <div id="assignmentPreview"></div>
          <div id="calendarPreview"></div>
        </div>
      </details>
      <div hidden>
        <textarea id="mermaidOutput"></textarea>
      </div>
      <div class="md-feedback-stack md-hidden">
        <div class="md-feedback-stack__title">取込結果</div>
        <div class="md-feedback-stack__text">取込後は、ここで差分反映・warning・検証結果を確認します。</div>
        <div class="md-feedback-stack__label md-hidden">検証メッセージ</div>
        <div id="validationIssues" class="md-hidden"></div>
        <div class="md-feedback-stack__label md-hidden">取込 warning</div>
        <div id="importWarnings" class="md-hidden"></div>
        <div class="md-feedback-stack__label md-hidden">差分要約</div>
        <div id="xlsxImportSummary" class="md-hidden"></div>
      </div>
    </section>
    <section id="tabPanelOutput" class="md-flow-section md-tab-panel" data-tab-panel="output" hidden>
      <div class="md-flow-section__header">
        <h2 class="md-flow-section__title">
          <span class="md-flow-section__step">3</span>
          <span>Output</span>
          <lht-help-tooltip label="Output の説明" wide>
            <p>MS Project XML、XLSX、JSON、CSV、WBS XLSX、Mermaid テキスト、SVG、生成AI 向け .editjson を保存します。</p>
          </lht-help-tooltip>
        </h2>
      </div>
      <details class="md-debug-accordion">
        <summary class="md-debug-accordion__summary">設定</summary>
        <div class="md-debug-accordion__body">
          <section class="md-note-card">
            <h3 class="md-note-card__title">
              <span>WBS XLSX 表示設定</span>
              <lht-help-tooltip label="WBS XLSX 表示設定の説明" placement="right" wide>
                <p>WBS XLSX Export では、ProjectModel から補完した既定祝日を WBS 日付帯へ反映します。</p>
                <p>既定祝日は、現在の ProjectModel に含まれる Calendar.Exceptions の非稼働日例外から内部で直接補完します。画面では Calendars / Exceptions を直接編集せず、必要な変更は MS Project XML または XLSX Import 側で扱います。表示期間を空欄にすると全期間、数値を入れると BaseDate 前後の営業日で切り出します。進捗帯も営業日基準で計算します。</p>
              </lht-help-tooltip>
            </h3>
          </section>
          <input id="wbsDisplayDaysBeforeInput" />
          <input id="wbsDisplayDaysAfterInput" />
          <input id="wbsBusinessDayRangeInput" type="checkbox" />
          <input id="wbsBusinessDayProgressInput" type="checkbox" />
        </div>
      </details>
      <div hidden>
          <textarea id="workbookJsonOutput"></textarea>
          <textarea id="aiBundleOutput"></textarea>
          <textarea id="projectOverviewOutput"></textarea>
          <input id="phaseDetailUidInput" />
          <input id="phaseDetailRootUidInput" />
          <input id="phaseDetailMaxDepthInput" />
          <textarea id="phaseDetailOutput"></textarea>
      </div>
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

function mountMinimalDom() {
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
      <div id="validationIssues" class="md-hidden"></div>
      <div id="importWarnings" class="md-hidden"></div>
      <div id="xlsxImportSummary" class="md-hidden"></div>
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

function bootPage(mount = mountFullDom) {
  mount();
  new Function(`${typesCode}\n${markdownEscapeCode}\n${aiJsonUtilCode}\n${mainUtilCode}\n${excelIoCode}\n${msProjectXmlCode}\n${projectWorkbookSchemaCode}\n${projectXlsxCode}\n${projectWorkbookJsonCode}\n${projectPatchJsonCode}\n${wbsXlsxCode}\n${wbsMarkdownCode}\n${nativeSvgCode}\n${mainRenderCode}\n${mainCode}`)();
  document.dispatchEvent(new Event("DOMContentLoaded"));
}

function getMainHooks() {
  return globalThis.__mikuprojectMainTestHooks;
}

function parseXmlViaHook() {
  getMainHooks().parseCurrentXml();
}

const SAMPLE_HOLIDAY_COUNT = 1;

async function flushAsyncWork() {
  await Promise.resolve();
  await Promise.resolve();
}


describe("mikuproject main", () => {
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

  it("loads sample xml on startup", () => {
    bootPage();

    expect(document.getElementById("xmlInput").value).toContain("<Project");
    expect(document.getElementById("xmlInput").value).toContain("<ScheduleFromStart>1</ScheduleFromStart>");
    expect(document.getElementById("xmlInput").value).toContain("<MinutesPerDay>480</MinutesPerDay>");
    expect(document.getElementById("xmlInput").value).toContain("<MinutesPerWeek>2400</MinutesPerWeek>");
    expect(document.getElementById("xmlInput").value).toContain("<DaysPerMonth>20</DaysPerMonth>");
    expect(document.getElementById("statusMessage").textContent).toContain("サンプル XML");
    expect(document.getElementById("xmlSaveState").textContent).toContain("XML 保存状態: 未保存");
    expect(document.querySelector('lht-help-tooltip[label="Load from file の説明"]')).not.toBeNull();
    expect(document.querySelector('lht-help-tooltip[label="Input の説明"]')).not.toBeNull();
    expect(document.querySelector('lht-help-tooltip[label="Overview の説明"]')).not.toBeNull();
    expect(document.querySelector('lht-help-tooltip[label="Output の説明"]')).not.toBeNull();
    expect(document.body.textContent).toContain("取込結果");
    expect(document.body.textContent).toContain("検証メッセージ");
    expect(document.body.textContent).toContain("取込 warning");
    expect(document.body.textContent).toContain("差分要約");
    expect(document.body.textContent).toContain("差分反映・warning・検証結果を確認します");
    expect(document.body.textContent).not.toContain("Workbook JSON");
    expect(document.body.textContent).not.toContain("入力、読込、編集元データの確認をここにまとめます。");
    expect(document.body.textContent).not.toContain("内部モデル、要約、検証結果、プレビューをここで確認します。");
    expect(document.body.textContent).not.toContain("出力設定と生成結果の確認をここにまとめます。");
    expect(document.body.textContent).toContain("デバッグ情報");
    expect(document.querySelectorAll(".md-debug-accordion")).toHaveLength(2);
    expect(document.querySelector(".md-debug-accordion")?.hasAttribute("open")).toBe(false);
    expect(document.querySelector(".md-note-accordion")?.hasAttribute("open")).toBe(false);
    expect(document.body.textContent).toContain("WBS XLSX 表示設定");
    expect(document.body.textContent).toContain("設定");
    expect(document.querySelectorAll(".md-debug-accordion")[1]?.hasAttribute("open")).toBe(false);
    expect(document.querySelector('lht-help-tooltip[label="WBS XLSX 表示設定の説明"]')).not.toBeNull();
    expect(document.body.textContent).toContain("ProjectModel から補完した既定祝日");
    expect(document.body.textContent).toContain("非稼働日例外から内部で直接補完");
    expect(document.body.textContent).toContain("必要な変更は MS Project XML または XLSX Import 側で扱います");
    expect(document.body.textContent).toContain("BaseDate 前後の営業日で切り出します");
    expect(document.body.textContent).toContain("進捗帯も営業日基準で計算します");
    expect(document.getElementById("wbsDisplayDaysBeforeInput").value).toBe("");
    expect(document.getElementById("wbsDisplayDaysAfterInput").value).toBe("");
    expect(document.querySelector(".md-feedback-stack")?.classList.contains("md-hidden")).toBe(true);
  });

  it("switches top tabs and toggles panels", async () => {
    bootPage(mountMinimalDom);

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
    expect(document.getElementById("nativeSvgPreview").innerHTML).toContain("<svg");

    document.querySelector('.md-top-tab[data-tab="output"]').click();
    expect(document.getElementById("tabPanelInput").hidden).toBe(true);
    expect(document.getElementById("tabPanelTransform").hidden).toBe(true);
    expect(document.getElementById("tabPanelOutput").hidden).toBe(false);
  });

  it("parses xml into internal model summary", () => {
    bootPage(mountMinimalDom);

    parseXmlViaHook();

    expect(document.getElementById("summaryProjectName").textContent).toBe("mikuproject開発");
    expect(document.getElementById("summaryTaskCount").textContent).toBe("13");
    expect(document.getElementById("summaryResourceCount").textContent).toBe("1");
    expect(document.getElementById("summaryAssignmentCount").textContent).toBe("2");
    expect(document.getElementById("summaryCalendarCount").textContent).toBe("1");
    expect(document.getElementById("modelOutput").value).toContain("\"name\": \"mikuproject開発\"");
    expect(document.getElementById("modelOutput").value).toContain("\"name\": \"基盤整備\"");
    expect(document.getElementById("modelOutput").value).toContain("\"name\": \"架空検討フェーズ【架空】\"");
    expect(document.getElementById("modelOutput").value).toContain("\"name\": \"Mikuku\"");
    expect(document.getElementById("modelOutput").value).toContain("\"initials\": \"M\"");
    expect(document.getElementById("modelOutput").value).toContain("\"name\": \"Standard\"");
    expect(document.getElementById("projectPreview").textContent).toContain("mikuproject開発");
    expect(document.getElementById("projectPreview").textContent).toContain("Calendar=1 (Standard)");
    expect(document.getElementById("taskPreview").textContent).toContain("初期実装（MS Project XML 調査・基軸フォーマット選定・内部モデルの概要確定）");
    expect(document.getElementById("resourcePreview").textContent).toContain("Mikuku");
    expect(document.getElementById("resourcePreview").textContent).toContain("Initials=M");
    expect(document.getElementById("assignmentPreview").textContent).toContain("Task=3");
    expect(document.getElementById("assignmentPreview").textContent).toContain("Resource=1 (Mikuku)");
    expect(document.getElementById("assignmentPreview").textContent).toContain("Task=4");
  expect(document.getElementById("calendarPreview").textContent).toContain(`WeekDays=7 / Exceptions=${SAMPLE_HOLIDAY_COUNT} / WorkWeeks=0`);
  });

  it("passes round-trip check", () => {
    bootPage(mountMinimalDom);

    parseXmlViaHook();
    document.getElementById("roundTripBtn").click();

    expect(document.getElementById("statusMessage").textContent).toContain("再読込テストに成功");
    expect(document.getElementById("modelOutput").value).toContain("\"extendedAttributes\": [");
  });

  it("imports xml from a file into the textarea", async () => {
    bootPage(mountMinimalDom);

    const importInput = document.getElementById("importFileInput");
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
    expect(document.getElementById("nativeSvgPreview").innerHTML).toContain("<svg");
    expect(document.getElementById("statusMessage").textContent).toContain("XML ファイルを読み込んで解析しました");
  });

});
