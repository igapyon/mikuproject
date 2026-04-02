// @vitest-environment jsdom

import { beforeEach, describe, expect, it } from "vitest";

import {
  bootPage,
  flushAsyncWork,
  hierarchyXml,
  parseXmlViaHook,
  setupMainAiJsonTestDom
} from "./helpers/main-ai-json-harness.js";

describe("mikuproject main ai json patch task structure", () => {
  beforeEach(() => {
    setupMainAiJsonTestDom();
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
      operations: [{ op: "move_task", uid: "3", new_parent_uid: null, new_index: 1 }]
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
      operations: [{ op: "delete_task", uid: "3" }]
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
      operations: [{ op: "delete_task", uid: "1" }]
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
      .replace("<Resources />", ["<Resources>", "  <Resource>", "    <UID>1</UID>", "    <ID>1</ID>", "    <Name>Owner</Name>", "  </Resource>", "</Resources>"].join("\n"))
      .replace("<Assignments />", ["<Assignments>", "  <Assignment>", "    <UID>1</UID>", "    <TaskUID>3</TaskUID>", "    <ResourceUID>1</ResourceUID>", "  </Assignment>", "</Assignments>"].join("\n"));
    document.getElementById("xmlInput").value = xmlWithAssignment;
    parseXmlViaHook();
    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [{ op: "delete_task", uid: "3" }]
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
    const xmlWithDependency = hierarchyXml.replace("</Notes>\n    </Task>", ["</Notes>", "      <PredecessorLink>", "        <PredecessorUID>2</PredecessorUID>", "        <Type>1</Type>", "      </PredecessorLink>", "    </Task>"].join("\n"));
    document.getElementById("xmlInput").value = xmlWithDependency;
    parseXmlViaHook();
    document.getElementById("projectDraftImportInput").value = JSON.stringify({
      operations: [{ op: "delete_task", uid: "2" }]
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
      operations: [{ op: "move_task", uid: "3", new_parent_uid: "1", new_index: 1 }]
    }, null, 2);
    document.getElementById("importProjectDraftBtn").click();
    await flushAsyncWork();
    await flushAsyncWork();
    expect(document.getElementById("statusMessage").textContent).toContain("Patch JSON に反映対象の変更はありませんでした");
    expect(document.getElementById("statusMessage").textContent).toContain("warning を無視しました");
    expect(document.getElementById("importWarnings").textContent).toContain("move_task は結果が変わらないため無視します");
    expect(document.getElementById("importWarnings").textContent).toContain("parent=1 index=1");
  });

});
