// @vitest-environment jsdom

import { beforeEach, describe, expect, it } from "vitest";

import {
  bootPage,
  dependencyXml,
  flushAsyncWork,
  hierarchyXml,
  parseXmlViaHook,
  setupMainAiJsonTestDom
} from "./helpers/main-ai-json-harness.js";

describe("mikuproject main ai json patch task", () => {
  beforeEach(() => {
    setupMainAiJsonTestDom();
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

  it("imports patch json update_task is_milestone into the current model and xml", async () => {
    bootPage();
    document.getElementById("xmlInput").value = hierarchyXml;
    parseXmlViaHook();
    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [{ op: "update_task", uid: "3", fields: { is_milestone: true } }]
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
      operations: [{ op: "update_task", uid: "1", fields: { is_milestone: true } }]
    }, null, 2);
    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();
    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON に反映対象の変更はありませんでした");
    expect(document.getElementById("importWarnings").textContent).toContain("update_task では summary task を milestone にできません");
    expect(document.getElementById("xmlInput").value).toContain("<UID>1</UID>");
    expect(document.getElementById("xmlInput").value).toContain("<Milestone>0</Milestone>");
  });

  it("imports and clears patch json update_task notes", async () => {
    bootPage();
    document.getElementById("xmlInput").value = hierarchyXml;
    parseXmlViaHook();
    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [{ op: "update_task", uid: "3", fields: { notes: "Patched note" } }]
    }, null, 2);
    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();
    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON を読み込んで 1 件の変更を反映しました");
    expect(document.getElementById("xmlInput").value).toContain("<Notes>Patched note</Notes>");
    expect(document.getElementById("modelOutput").value).toContain("\"notes\": \"Patched note\"");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("Notes");
    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [{ op: "update_task", uid: "3", fields: { notes: "" } }]
    }, null, 2);
    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();
    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON を読み込んで 1 件の変更を反映しました");
    expect(document.getElementById("xmlInput").value).not.toContain("<Notes>Second child task</Notes>");
    expect(document.getElementById("modelOutput").value).not.toContain("\"notes\": \"Second child task\"");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("Notes");
  });

  it("imports and clears patch json update_task calendar_uid", async () => {
    bootPage();
    document.getElementById("xmlInput").value = dependencyXml;
    parseXmlViaHook();
    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [{ op: "update_task", uid: "2", fields: { calendar_uid: "1" } }]
    }, null, 2);
    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();
    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON を読み込んで 1 件の変更を反映しました");
    expect(document.getElementById("xmlInput").value).toContain("<UID>2</UID>");
    expect(document.getElementById("xmlInput").value).toContain("<CalendarUID>1</CalendarUID>");
    expect(document.getElementById("modelOutput").value).toContain("\"calendarUID\": \"1\"");
    expect(document.getElementById("xlsxImportSummary").innerHTML).toContain("CalendarUID");
    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [{ op: "update_task", uid: "2", fields: { calendar_uid: "" } }]
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
      operations: [{ op: "update_task", uid: "2", fields: { calendar_uid: "999" } }]
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
      operations: [{ op: "update_task", uid: "3", fields: { percent_complete: 40, percent_work_complete: 60 } }]
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
      operations: [{ op: "update_task", uid: "3", fields: { percent_complete: 120, percent_work_complete: -1 } }]
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
      operations: [{ op: "update_task", uid: "3", fields: { critical: true } }]
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
      operations: [{ op: "update_task", uid: "3", fields: { critical: "yes" } }]
    }, null, 2);
    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();
    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON に反映対象の変更はありませんでした");
    expect(document.getElementById("importWarnings").textContent).toContain("update_task.critical は boolean が必要です");
  });

  it("groups patch json task warnings by uid", async () => {
    bootPage();
    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [{ op: "update_task", uid: "3", fields: { unknown_field: "ignored" } }]
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
});
