// @vitest-environment jsdom

import { beforeEach, describe, expect, it } from "vitest";

import {
  bootPage,
  flushAsyncWork,
  hierarchyXml,
  parseXmlViaHook,
  setupMainAiJsonTestDom
} from "./helpers/main-ai-json-harness.js";

describe("mikuproject main ai json", () => {
  beforeEach(() => {
    setupMainAiJsonTestDom();
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
