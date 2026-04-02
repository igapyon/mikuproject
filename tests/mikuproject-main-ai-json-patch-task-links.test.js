// @vitest-environment jsdom

import { beforeEach, describe, expect, it } from "vitest";

import {
  bootPage,
  flushAsyncWork,
  hierarchyXml,
  parseXmlViaHook,
  setupMainAiJsonTestDom
} from "./helpers/main-ai-json-harness.js";

describe("mikuproject main ai json patch task links", () => {
  beforeEach(() => {
    setupMainAiJsonTestDom();
  });

  it("imports patch json link_tasks into the current model and xml", async () => {
    bootPage();
    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [{ op: "link_tasks", from_uid: "2", to_uid: "3", type: "SS", lag_hours: 4 }]
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
      operations: [{ op: "link_tasks", from_uid: "2", to_uid: "3", type: "SS", lag_hours: 4 }]
    }, null, 2);
    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();
    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [{ op: "unlink_tasks", from_uid: "2", to_uid: "3", type: "SS" }]
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
      operations: [{ op: "link_tasks", from_uid: "2", to_uid: "3", type: "SS", lag_hours: 4 }]
    }, null, 2);
    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();
    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [{ op: "unlink_tasks", from_uid: "2", to_uid: "3", type: "SS", lag: "PT4H0M0S" }]
    }, null, 2);
    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();
    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON を読み込んで 1 件の変更を反映しました");
    expect(document.getElementById("xmlInput").value).not.toContain("<PredecessorUID>2</PredecessorUID>");
  });

  it("reports unlink_tasks warnings with requested type details", async () => {
    bootPage();
    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [{ op: "unlink_tasks", from_uid: "2", to_uid: "3", type: "FF" }]
    }, null, 2);
    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();
    expect(document.getElementById("statusMessage").textContent).toContain("warning を無視しました");
    expect(document.getElementById("importWarnings").textContent).toContain("unlink_tasks の対象依存関係が見つかりません");
    expect(document.getElementById("importWarnings").textContent).toContain("2 -> 3 (FF)");
  });

  it("reports when unlink_tasks removes multiple matching dependencies", async () => {
    bootPage();
    const xmlWithDuplicateLinks = hierarchyXml.replace("</Notes>\n    </Task>", ["</Notes>", "      <PredecessorLink>", "        <PredecessorUID>2</PredecessorUID>", "        <Type>1</Type>", "        <LinkLag>PT0H0M0S</LinkLag>", "      </PredecessorLink>", "      <PredecessorLink>", "        <PredecessorUID>2</PredecessorUID>", "        <Type>1</Type>", "        <LinkLag>PT0H0M0S</LinkLag>", "      </PredecessorLink>", "    </Task>"].join("\n"));
    document.getElementById("xmlInput").value = xmlWithDuplicateLinks;
    parseXmlViaHook();
    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [{ op: "unlink_tasks", from_uid: "2", to_uid: "3", type: "FS", lag: "PT0H0M0S" }]
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
        { op: "link_tasks", from_uid: "2", to_uid: "3", type: "FS" },
        { op: "link_tasks", from_uid: "3", to_uid: "2", type: "FS" }
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
      operations: [{ op: "link_tasks", from_uid: "2", to_uid: "3", type: "FS", lag: "4hours" }]
    }, null, 2);
    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();
    expect(document.getElementById("statusMessage").textContent).toContain("warning を無視しました");
    expect(document.getElementById("importWarnings").textContent).toContain("link_tasks.lag は ISO 8601 duration 形式が必要です");
    expect(document.getElementById("xmlInput").value).toContain("<PredecessorUID>2</PredecessorUID>");
    expect(document.getElementById("xmlInput").value).not.toContain("<LinkLag>4hours</LinkLag>");
  });
});
