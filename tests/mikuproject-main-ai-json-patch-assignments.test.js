// @vitest-environment jsdom

import { beforeEach, describe, expect, it } from "vitest";

import {
  bootPage,
  dependencyXml,
  flushAsyncWork,
  parseXmlViaHook,
  setupMainAiJsonTestDom
} from "./helpers/main-ai-json-harness.js";

describe("mikuproject main ai json patch assignments", () => {
  beforeEach(() => {
    setupMainAiJsonTestDom();
  });

  it("imports patch json update_assignment into the current model and xml", async () => {
    bootPage();
    document.getElementById("xmlInput").value = dependencyXml;
    parseXmlViaHook();
    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [{ op: "update_assignment", uid: "1", fields: { start: "2026-03-17", finish: "2026-03-18", units: 0.5, work: "PT12H0M0S", percent_work_complete: 75 } }]
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

  it("imports patch json add_assignment into the current model and xml", async () => {
    bootPage();
    document.getElementById("xmlInput").value = dependencyXml;
    parseXmlViaHook();
    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [{ op: "add_assignment", uid: "3", task_uid: "2", resource_uid: "1", start: "2026-03-19", finish: "2026-03-19", units: 0.25, work: "PT2H0M0S", percent_work_complete: 10 }]
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

  it("imports patch json delete_assignment into the current model and xml", async () => {
    bootPage();
    document.getElementById("xmlInput").value = dependencyXml;
    parseXmlViaHook();
    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [{ op: "delete_assignment", uid: "1" }]
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
      operations: [{ op: "delete_assignment", uid: "999" }]
    }, null, 2);
    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();
    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON に反映対象の変更はありませんでした");
    expect(document.getElementById("importWarnings").textContent).toContain("delete_assignment の uid が既存 assignment を指していません");
    expect(document.getElementById("xmlInput").value).toContain("<Assignment>");
  });
});
