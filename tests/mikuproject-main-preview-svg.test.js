// @vitest-environment jsdom

import { beforeEach, describe, expect, it } from "vitest";

import {
  bootPage,
  bootXmlModule,
  exportMermaidViaHook,
  hierarchyXml,
  parseXmlViaHook,
  setupMainPreviewExportDom
} from "./helpers/main-preview-export-harness.js";

describe("mikuproject main preview svg", () => {
  beforeEach(() => {
    setupMainPreviewExportDom();
  });

  it("exports mermaid gantt from the current model", async () => {
    bootPage();
    parseXmlViaHook();
    await exportMermaidViaHook();
    const mermaidText = document.getElementById("mermaidOutput").value;
    expect(mermaidText).toContain("gantt");
    expect(mermaidText).toContain("title mikuproject開発");
    expect(mermaidText).toContain("section 基盤整備");
    expect(document.getElementById("nativeSvgPreview").innerHTML).toContain("<svg");
    expect(document.getElementById("downloadSvgBtn").disabled).toBe(false);
    expect(document.getElementById("downloadWeeklySvgBtn").disabled).toBe(false);
  });

  it("switches overview svg preview between daily, weekly, and monthly", async () => {
    bootPage();
    parseXmlViaHook();
    await exportMermaidViaHook();
    document.getElementById("previewWeeklySvgBtn").click();
    expect(document.getElementById("nativeSvgPreview").innerHTML).toContain("weekly overview");
    document.getElementById("previewMonthlySvgBtn").click();
    expect(document.getElementById("nativeSvgPreview").innerHTML).toContain("2026-03");
    expect(document.getElementById("nativeSvgPreview").innerHTML).toContain("2026-04");
    document.getElementById("previewDailySvgBtn").click();
    expect(document.getElementById("nativeSvgPreview").innerHTML).toContain("<svg");
    expect(document.getElementById("nativeSvgPreview").innerHTML).not.toContain("weekly overview");
  });

  it("keeps long daily task labels outside narrow on-bar labels", async () => {
    bootPage();
    parseXmlViaHook();
    await exportMermaidViaHook();
    const dailySvg = document.getElementById("nativeSvgPreview").innerHTML;
    expect(dailySvg).not.toContain('text-anchor="middle">round-trip拡張（MS Project XML → 内部JSON形式 → MS Project XML の往復対応）</text>');
    expect(dailySvg).not.toContain('text-anchor="middle">MS Project XML と XLSX の相互変換・round-trip実装</text>');
  });

  it("places daily labels on the left after they pass the chart midpoint", async () => {
    bootPage();
    parseXmlViaHook();
    await exportMermaidViaHook();
    const dailySvg = document.getElementById("nativeSvgPreview").innerHTML;
    expect(dailySvg).toContain('text-anchor="start">架空検討フェーズ【架空】</text>');
    expect(dailySvg).toContain('text-anchor="end">MS Project XML と XLSX の相互変換・round-trip実装</text>');
  });

  it("renders dependency connectors in daily and weekly native svg", () => {
    bootPage();
    const model = {
      project: {
        name: "Dependency Demo",
        startDate: "2026-03-16T09:00:00",
        finishDate: "2026-03-20T18:00:00",
        currentDate: "2026-03-18T12:00:00",
        scheduleFromStart: true,
        outlineCodes: [],
        wbsMasks: [],
        extendedAttributes: []
      },
      tasks: [
        { uid: "1", id: "1", name: "Prep", outlineLevel: 1, outlineNumber: "1", start: "2026-03-16T09:00:00", finish: "2026-03-17T18:00:00", duration: "PT16H0M0S", milestone: false, summary: false, percentComplete: 100, predecessors: [], extendedAttributes: [], baselines: [], timephasedData: [] },
        { uid: "2", id: "2", name: "Ship", outlineLevel: 1, outlineNumber: "2", start: "2026-03-18T09:00:00", finish: "2026-03-19T18:00:00", duration: "PT16H0M0S", milestone: false, summary: false, percentComplete: 0, predecessors: [{ predecessorUid: "1", type: 1 }], extendedAttributes: [], baselines: [], timephasedData: [] }
      ],
      resources: [],
      assignments: [],
      calendars: []
    };
    const dailySvg = globalThis.__mikuprojectNativeSvg.exportNativeSvg(model);
    const weeklySvg = globalThis.__mikuprojectNativeSvg.exportWeeklyNativeSvg(model);
    expect(dailySvg).toContain('class="dependencyPath"');
    expect(dailySvg).toContain('marker-end="url(#dependencyArrow)"');
    expect(dailySvg).toContain('data-from-uid="1"');
    expect(dailySvg).toContain('data-to-uid="2"');
    expect(weeklySvg).toContain('class="dependencyPath"');
    expect(weeklySvg).toContain('marker-end="url(#dependencyArrow)"');
  });

  it("keeps complex mermaid dependencies as comments", () => {
    const xml = bootXmlModule();
    const model = xml.importMsProjectXml(hierarchyXml);
    model.tasks.find((task) => task.uid === "2").predecessors = [
      { predecessorUid: "1", type: 1 },
      { predecessorUid: "3", type: 2, linkLag: "PT4H0M0S" }
    ];
    const mermaid = xml.exportMermaidGantt(model);
    expect(mermaid).toContain("%% dependency:");
    expect(mermaid).toContain("type=FF");
    expect(mermaid).toContain("lag=4h");
  });

  it("sanitizes date-leading mermaid gantt labels", () => {
    const xml = bootXmlModule();
    const model = xml.importMsProjectXml(hierarchyXml);
    model.tasks[0].name = "2026-03-16 初期実装（42513dd：XML import/export）";
    const mermaidText = xml.exportMermaidGantt(model);
    expect(mermaidText).toContain("title Hierarchy Project");
    expect(mermaidText).toContain("section Section 2026-03-16 初期実装（42513dd XML import/export）");
  });
});
