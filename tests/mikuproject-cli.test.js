import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { spawnSync } from "node:child_process";
import { gunzipSync } from "node:zlib";

import { afterEach, describe, expect, it } from "vitest";

import { loadMikuprojectCoreApi } from "../scripts/lib/core-api-loader.mjs";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..");
const cliPath = path.resolve(repoRoot, "scripts/miku-project-cli.mjs");
const cliBundleBuildPath = path.resolve(repoRoot, "scripts/build-cli-bundle.mjs");
const cliAiWorkflowExamplePath = path.resolve(repoRoot, "scripts/cli-ai-workflow-example.mjs");
const cliAiStdioExamplePath = path.resolve(repoRoot, "scripts/cli-ai-stdio-example.mjs");
const packageVersion = JSON.parse(readFileSync(path.resolve(repoRoot, "package.json"), "utf8")).version;

const tempDirs = [];
const disposers = [];

afterEach(() => {
  while (disposers.length > 0) {
    disposers.pop()();
  }
  while (tempDirs.length > 0) {
    rmSync(tempDirs.pop(), { recursive: true, force: true });
  }
});

describe("miku-project cli", () => {
  it("prints the AI spec", () => {
    const result = runCli(["ai", "spec"]);

    expect(result.status).toBe(0);
    expect(result.stdout).toContain("# miku-project AI JSON Prompt / Spec");
    expect(result.stderr).toBe("");
  });

  it("prints help for AI local projection exports", () => {
    const result = runCli(["--help"]);

    expect(result.status).toBe(0);
    expect(result.stdout).toContain("Description:");
    expect(result.stdout).toContain("Default behavior:");
    expect(result.stdout).toContain("Inputs:");
    expect(result.stdout).toContain("Outputs:");
    expect(result.stdout).toContain("Generated artifacts:");
    expect(result.stdout).toContain("Overwrite behavior:");
    expect(result.stdout).toContain("Machine-readable output contract:");
    expect(result.stdout).toContain("Diagnostics / warnings:");
    expect(result.stdout).toContain("Exit codes:");
    expect(result.stdout).toContain("References:");
    expect(result.stdout).toContain("miku-project --version");
    expect(result.stdout).toContain("miku-project ai export project-overview");
    expect(result.stdout).toContain("miku-project ai export task-edit");
    expect(result.stdout).toContain("miku-project ai export phase-detail");
    expect(result.stdout).toContain("miku-project ai export bundle");
    expect(result.stdout).toContain("--select auto|first-task|uid");
    expect(result.stdout).toContain("--select auto|first-phase|uid");
    expect(result.stdout).toContain("miku-project ai detect-kind");
    expect(result.stdout).toContain("miku-project ai validate-patch");
    expect(result.stdout).toContain("miku-project state summarize");
    expect(result.stdout).toContain("miku-project state diff");
  });

  it("prints the cli version", () => {
    const result = runCli(["--version"]);

    expect(result.status).toBe(0);
    expect(result.stdout).toBe(`miku-project ${packageVersion}\n`);
    expect(result.stderr).toBe("");
  });

  it("creates workbook json from project_draft_view", () => {
    const draftPath = writeTempJson("draft.json", {
      view_type: "project_draft_view",
      project: {
        name: "CLI draft import",
        planned_start: "2026-04-01"
      },
      tasks: [
        {
          uid: "draft-1",
          name: "開始",
          parent_uid: null,
          position: 0,
          is_milestone: true,
          planned_start: "2026-04-01",
          planned_finish: "2026-04-01"
        }
      ],
      resources: [],
      assignments: []
    });

    const result = runCli(["state", "from-draft", "--in", draftPath]);

    expect(result.status).toBe(0);
    const workbook = JSON.parse(result.stdout);
    expect(workbook.format).toBe("mikuproject_workbook_json");
    expect(workbook.sheets.Project.find((row) => row.Field === "Name").Value).toBe("CLI draft import");
  });

  it("reads input JSON from stdin when --in - is specified", () => {
    const result = runCli(["ai", "detect-kind", "--in", "-"], {
      input: `${JSON.stringify({ operations: [] }, null, 2)}\n`
    });

    expect(result.status).toBe(0);
    expect(result.stdout).toBe("patch_json\n");
  });

  it("detects patch_json kind", () => {
    const patchPath = writeTempJson("patch.json", {
      operations: []
    });

    const result = runCli(["ai", "detect-kind", "--in", patchPath]);

    expect(result.status).toBe(0);
    expect(result.stdout).toBe("patch_json\n");
    expect(result.stderr).toBe("");
  });

  it("writes detect-kind diagnostics as json when requested", () => {
    const patchPath = writeTempJson("patch.json", {
      operations: []
    });

    const result = runCli(["ai", "detect-kind", "--in", patchPath, "--diagnostics", "json"]);

    expect(result.status).toBe(0);
    expect(result.stdout).toBe("patch_json\n");
    const diagnostics = JSON.parse(result.stderr);
    expect(diagnostics.ok).toBe(true);
    expect(diagnostics.diagnostics_version).toBe(1);
    expect(diagnostics.command).toBe("detect-kind");
    expect(diagnostics.context).toBe("detect-kind");
    expect(diagnostics.status).toBe("success");
    expect(diagnostics.exit_code).toBe(0);
    expect(diagnostics.warning_count).toBe(0);
    expect(diagnostics.error_count).toBe(0);
    expect(diagnostics.io).toEqual({
      inputs: [{ option: "--in", mode: "file", path: patchPath }],
      output: { mode: "stdout" }
    });
    expect(diagnostics.warnings).toEqual([]);
    expect(diagnostics.errors).toEqual([]);
    expect(diagnostics.detected_kind).toBe("patch_json");
  });

  it("applies patch_json to workbook state", () => {
    const statePath = writeTempJson("workbook.json", buildWorkbookState("CLI patch base"));
    const patchPath = writeTempJson("patch.json", {
      operations: [
        {
          op: "update_project",
          fields: {
            name: "CLI patch updated"
          }
        }
      ]
    });

    const result = runCli(["state", "apply-patch", "--state", statePath, "--in", patchPath]);

    expect(result.status).toBe(0);
    const workbook = JSON.parse(result.stdout);
    expect(workbook.sheets.Project.find((row) => row.Field === "Name").Value).toBe("CLI patch updated");
    expect(result.stderr).toContain("apply-patch patch_json");
  });

  it("exports workbook json canonically", () => {
    const workbookPath = writeTempJson("workbook.json", buildWorkbookState("CLI export workbook"));

    const result = runCli(["export", "workbook-json", "--in", workbookPath]);

    expect(result.status).toBe(0);
    const workbook = JSON.parse(result.stdout);
    expect(workbook.format).toBe("mikuproject_workbook_json");
    expect(workbook.sheets.Project.find((row) => row.Field === "Name").Value).toBe("CLI export workbook");
  });

  it("writes output to stdout when --out - is specified", () => {
    const workbookPath = writeTempJson("workbook.json", buildWorkbookState("CLI stdout output"));

    const result = runCli(["export", "workbook-json", "--in", workbookPath, "--out", "-"]);

    expect(result.status).toBe(0);
    const workbook = JSON.parse(result.stdout);
    expect(workbook.sheets.Project.find((row) => row.Field === "Name").Value).toBe("CLI stdout output");
  });

  it("summarizes workbook state", () => {
    const loaded = loadMikuprojectCoreApi({ rootDir: repoRoot });
    disposers.push(() => loaded.dispose());
    const model = loaded.api.msProject.importFromXml(readFixture("hierarchy.xml"));
    const workbookPath = writeTempJson("workbook.json", loaded.api.workbookJson.exportDocument(model));

    const result = runCli(["state", "summarize", "--in", workbookPath]);

    expect(result.status).toBe(0);
    const summary = JSON.parse(result.stdout);
    expect(summary.kind).toBe("state_summary");
    expect(summary.project.name).toBe("Hierarchy Project");
    expect(summary.phase_count).toBe(1);
    expect(summary.summary.task_count).toBe(3);
  });

  it("writes summarize diagnostics as json when requested", () => {
    const loaded = loadMikuprojectCoreApi({ rootDir: repoRoot });
    disposers.push(() => loaded.dispose());
    const model = loaded.api.msProject.importFromXml(readFixture("hierarchy.xml"));
    const workbookPath = writeTempJson("workbook.json", loaded.api.workbookJson.exportDocument(model));

    const result = runCli(["state", "summarize", "--in", workbookPath, "--diagnostics", "json"]);

    expect(result.status).toBe(0);
    const diagnostics = JSON.parse(result.stderr);
    expect(diagnostics.ok).toBe(true);
    expect(diagnostics.command).toBe("state summarize");
    expect(diagnostics.context).toBe("state summarize");
    expect(diagnostics.status).toBe("success");
    expect(diagnostics.exit_code).toBe(0);
    expect(diagnostics.warnings).toEqual([]);
    expect(diagnostics.errors).toEqual([]);
    expect(diagnostics.project_name).toBe("Hierarchy Project");
    expect(diagnostics.phase_count).toBe(1);
    expect(diagnostics.task_count).toBe(3);
  });

  it("exports project_overview_view from workbook state", () => {
    const loaded = loadMikuprojectCoreApi({ rootDir: repoRoot });
    disposers.push(() => loaded.dispose());
    const model = loaded.api.msProject.importFromXml(readFixture("hierarchy.xml"));
    const workbookPath = writeTempJson("workbook.json", loaded.api.workbookJson.exportDocument(model));

    const result = runCli(["ai", "export", "project-overview", "--in", workbookPath]);

    expect(result.status).toBe(0);
    expect(JSON.parse(result.stdout)).toEqual(loaded.api.aiViews.exportProjectOverviewView(model));
  });

  it("exports ai projection bundle from workbook state", () => {
    const loaded = loadMikuprojectCoreApi({ rootDir: repoRoot });
    disposers.push(() => loaded.dispose());
    const model = loaded.api.msProject.importFromXml(readFixture("hierarchy.xml"));
    const workbookPath = writeTempJson("workbook.json", loaded.api.workbookJson.exportDocument(model));

    const result = runCli(["ai", "export", "bundle", "--in", workbookPath]);

    expect(result.status).toBe(0);
    const bundle = JSON.parse(result.stdout);
    expect(bundle.view_type).toBe("ai_projection_bundle");
    expect(bundle.project_overview_view.view_type).toBe("project_overview_view");
    expect(Array.isArray(bundle.phase_detail_views_full)).toBe(true);
    expect(bundle.phase_detail_views_full.every((item) => item.scope?.mode === "full")).toBe(true);
    expect(Array.isArray(bundle.task_edit_views_full)).toBe(true);
    expect(bundle.task_edit_views_full.every((item) => item.view_type === "task_edit_view")).toBe(true);
  });

  it("writes ai export bundle diagnostics as json when requested", () => {
    const loaded = loadMikuprojectCoreApi({ rootDir: repoRoot });
    disposers.push(() => loaded.dispose());
    const model = loaded.api.msProject.importFromXml(readFixture("hierarchy.xml"));
    const workbookPath = writeTempJson("workbook.json", loaded.api.workbookJson.exportDocument(model));

    const result = runCli(["ai", "export", "bundle", "--in", workbookPath, "--diagnostics", "json"]);

    expect(result.status).toBe(0);
    const diagnostics = JSON.parse(result.stderr);
    expect(diagnostics.ok).toBe(true);
    expect(diagnostics.command).toBe("ai export bundle");
    expect(diagnostics.context).toBe("ai export bundle");
    expect(diagnostics.status).toBe("success");
    expect(diagnostics.exit_code).toBe(0);
    expect(diagnostics.warnings).toEqual([]);
    expect(diagnostics.errors).toEqual([]);
    expect(diagnostics.output_kind).toBe("ai_projection_bundle");
    expect(diagnostics.phase_count).toBeGreaterThan(0);
    expect(diagnostics.task_count).toBeGreaterThan(0);
  });

  it("exports task_edit_view from workbook state", () => {
    const loaded = loadMikuprojectCoreApi({ rootDir: repoRoot });
    disposers.push(() => loaded.dispose());
    const model = loaded.api.msProject.importFromXml(readFixture("dependency.xml"));
    const workbookPath = writeTempJson("workbook.json", loaded.api.workbookJson.exportDocument(model));

    const result = runCli(["ai", "export", "task-edit", "--in", workbookPath, "--task-uid", "2"]);

    expect(result.status).toBe(0);
    expect(JSON.parse(result.stdout)).toEqual(loaded.api.aiViews.exportTaskEditView(model, "2"));
  });

  it("exports task_edit_view with the default task selection when task uid is omitted", () => {
    const loaded = loadMikuprojectCoreApi({ rootDir: repoRoot });
    disposers.push(() => loaded.dispose());
    const model = loaded.api.msProject.importFromXml(readFixture("hierarchy.xml"));
    const workbookPath = writeTempJson("workbook.json", loaded.api.workbookJson.exportDocument(model));

    const result = runCli(["ai", "export", "task-edit", "--in", workbookPath]);

    expect(result.status).toBe(0);
    expect(JSON.parse(result.stdout)).toEqual(loaded.api.aiViews.exportTaskEditView(model));
  });

  it("exports task_edit_view with --select first-task", () => {
    const loaded = loadMikuprojectCoreApi({ rootDir: repoRoot });
    disposers.push(() => loaded.dispose());
    const model = loaded.api.msProject.importFromXml(readFixture("hierarchy.xml"));
    const workbookPath = writeTempJson("workbook.json", loaded.api.workbookJson.exportDocument(model));

    const result = runCli(["ai", "export", "task-edit", "--in", workbookPath, "--select", "first-task"]);

    expect(result.status).toBe(0);
    expect(JSON.parse(result.stdout)).toEqual(loaded.api.aiViews.exportTaskEditView(model, "2"));
  });

  it("requires --task-uid when task-edit uses --select uid", () => {
    const workbookPath = writeTempJson("workbook.json", buildWorkbookState("CLI task select uid"));

    const result = runCli(["ai", "export", "task-edit", "--in", workbookPath, "--select", "uid"]);

    expect(result.status).toBe(2);
    expect(result.stderr).toContain("--task-uid");
  });

  it("exports scoped phase_detail_view from workbook state", () => {
    const loaded = loadMikuprojectCoreApi({ rootDir: repoRoot });
    disposers.push(() => loaded.dispose());
    const model = loaded.api.msProject.importFromXml(readFixture("hierarchy.xml"));
    const workbookPath = writeTempJson("workbook.json", loaded.api.workbookJson.exportDocument(model));

    const result = runCli([
      "ai",
      "export",
      "phase-detail",
      "--in",
      workbookPath,
      "--phase-uid",
      "1",
      "--root-task-uid",
      "2",
      "--max-depth",
      "1"
    ]);

    expect(result.status).toBe(0);
    expect(JSON.parse(result.stdout)).toEqual(loaded.api.aiViews.exportPhaseDetailView(model, "1", {
      mode: "scoped",
      rootUid: "2",
      maxDepth: 1
    }));
  });

  it("exports phase_detail_view with --select first-phase", () => {
    const loaded = loadMikuprojectCoreApi({ rootDir: repoRoot });
    disposers.push(() => loaded.dispose());
    const model = loaded.api.msProject.importFromXml(readFixture("hierarchy.xml"));
    const workbookPath = writeTempJson("workbook.json", loaded.api.workbookJson.exportDocument(model));

    const result = runCli(["ai", "export", "phase-detail", "--in", workbookPath, "--select", "first-phase"]);

    expect(result.status).toBe(0);
    expect(JSON.parse(result.stdout)).toEqual(loaded.api.aiViews.exportPhaseDetailView(model, "1", {
      mode: "scoped",
      rootUid: undefined,
      maxDepth: undefined
    }));
  });

  it("requires --phase-uid when phase-detail uses --select uid", () => {
    const workbookPath = writeTempJson("workbook.json", buildWorkbookState("CLI phase select uid"));

    const result = runCli(["ai", "export", "phase-detail", "--in", workbookPath, "--select", "uid"]);

    expect(result.status).toBe(2);
    expect(result.stderr).toContain("--phase-uid");
  });

  it("returns exit code 2 for missing required state option in validate-patch", () => {
    const patchPath = writeTempJson("patch.json", {
      operations: []
    });

    const result = runCli(["ai", "validate-patch", "--in", patchPath]);

    expect(result.status).toBe(2);
    expect(result.stderr).toContain("--state");
  });

  it("writes usage errors as json diagnostics when requested", () => {
    const patchPath = writeTempJson("patch.json", {
      operations: []
    });

    const result = runCli(["ai", "validate-patch", "--in", patchPath, "--diagnostics", "json"]);

    expect(result.status).toBe(2);
    const diagnostics = JSON.parse(result.stderr);
    expect(diagnostics.ok).toBe(false);
    expect(diagnostics.diagnostics_version).toBe(1);
    expect(diagnostics.command).toBe("ai validate-patch");
    expect(diagnostics.context).toBe("ai validate-patch");
    expect(diagnostics.status).toBe("error");
    expect(diagnostics.exit_code).toBe(2);
    expect(diagnostics.warning_count).toBe(0);
    expect(diagnostics.error_count).toBe(1);
    expect(diagnostics.io).toEqual({
      inputs: [{ option: "--in", mode: "file", path: patchPath }],
      output: { mode: "stdout" }
    });
    expect(diagnostics.error_type).toBe("usage_error");
    expect(diagnostics.error_code).toBe("missing_state_option");
    expect(diagnostics.error_details).toEqual({ option: "--state" });
    expect(diagnostics.errors[0].code).toBe("missing_state_option");
    expect(diagnostics.errors[0].details).toEqual({ option: "--state" });
    expect(diagnostics.errors[0].message).toContain("--state");
  });

  it("returns exit code 2 for invalid diagnostics option value", () => {
    const patchPath = writeTempJson("patch.json", {
      operations: []
    });

    const result = runCli(["ai", "detect-kind", "--in", patchPath, "--diagnostics", "yaml"]);

    expect(result.status).toBe(2);
    expect(result.stderr).toContain("--diagnostics");
  });

  it("writes processing errors as json diagnostics when requested", () => {
    const result = runCli(["ai", "detect-kind", "--diagnostics", "json"], {
      input: "{ invalid json"
    });

    expect(result.status).toBe(1);
    const diagnostics = JSON.parse(result.stderr);
    expect(diagnostics.ok).toBe(false);
    expect(diagnostics.command).toBe("ai detect-kind");
    expect(diagnostics.context).toBe("ai detect-kind");
    expect(diagnostics.status).toBe("error");
    expect(diagnostics.exit_code).toBe(1);
    expect(diagnostics.io).toEqual({
      inputs: [{ option: "--in", mode: "stdin_implicit" }],
      output: { mode: "stdout" }
    });
    expect(diagnostics.error_type).toBe("processing_error");
    expect(diagnostics.error_code).toBe("invalid_json_input");
    expect(diagnostics.error_details).toEqual({ context: "ai detect-kind" });
    expect(Array.isArray(diagnostics.errors)).toBe(true);
    expect(diagnostics.errors[0].code).toBe("invalid_json_input");
    expect(diagnostics.errors[0].details).toEqual({ context: "ai detect-kind" });
  });

  it("returns exit code 2 when validate-patch would read stdin from both state and patch", () => {
    const result = runCli(["ai", "validate-patch", "--state", "-", "--diagnostics", "json"], {
      input: `${JSON.stringify({ operations: [] }, null, 2)}\n`
    });

    expect(result.status).toBe(2);
    expect(result.stderr).toContain("--state");
    expect(result.stderr).toContain("--in");
  });

  it("returns exit code 2 when state diff would read stdin from both before and after", () => {
    const result = runCli(["state", "diff", "--before", "-", "--after", "-"], {
      input: `${JSON.stringify(buildWorkbookState("stdin diff"), null, 2)}\n`
    });

    expect(result.status).toBe(2);
    expect(result.stderr).toContain("--before");
    expect(result.stderr).toContain("--after");
  });

  it("validates patch_json with text output before apply", () => {
    const statePath = writeTempJson("workbook.json", buildWorkbookState("CLI validate patch"));
    const patchPath = writeTempJson("patch.json", {
      operations: [
        {
          op: "update_project",
          fields: {
            name: "CLI validate patch updated"
          }
        }
      ]
    });

    const result = runCli(["ai", "validate-patch", "--state", statePath, "--in", patchPath]);

    expect(result.status).toBe(0);
    expect(result.stdout).toContain("validate-patch ok=true");
    expect(result.stdout).toContain("status=success");
    expect(result.stdout).toContain("changes=1");
    expect(result.stderr).toBe("");
  });

  it("keeps validate-patch text/json diagnostics aligned across statuses", () => {
    const successStatePath = writeTempJson("validate-success-workbook.json", buildWorkbookState("CLI validate align success"));
    const successPatchPath = writeTempJson("validate-success-patch.json", {
      operations: [
        {
          op: "update_project",
          fields: {
            name: "CLI validate align success updated"
          }
        }
      ]
    });
    const warningStatePath = writeTempJson("validate-warning-workbook.json", buildWorkbookState("CLI validate align warning"));
    const warningPatchPath = writeTempJson("validate-warning-patch.json", {
      operations: [
        {
          op: "unknown_op"
        }
      ]
    });
    const noopStatePath = writeTempJson("validate-noop-workbook.json", buildWorkbookState("CLI validate align noop"));
    const noopPatchPath = writeTempJson("validate-noop-patch.json", {
      operations: []
    });
    const errorStatePath = writeTempJson("validate-error-workbook.json", buildWorkbookState("CLI validate align error"));
    const errorPatchPath = writeTempJson("validate-error-patch.json", {
      invalid: true
    });

    const cases = [
      {
        label: "success",
        statePath: successStatePath,
        patchPath: successPatchPath,
        expectedStatus: "success",
        expectedExitCode: 0
      },
      {
        label: "warning",
        statePath: warningStatePath,
        patchPath: warningPatchPath,
        expectedStatus: "warning",
        expectedExitCode: 0
      },
      {
        label: "noop",
        statePath: noopStatePath,
        patchPath: noopPatchPath,
        expectedStatus: "noop",
        expectedExitCode: 0
      },
      {
        label: "error",
        statePath: errorStatePath,
        patchPath: errorPatchPath,
        expectedStatus: "error",
        expectedExitCode: 1
      }
    ];

    for (const testCase of cases) {
      const textResult = runCli([
        "ai",
        "validate-patch",
        "--state",
        testCase.statePath,
        "--in",
        testCase.patchPath
      ]);
      const jsonResult = runCli([
        "ai",
        "validate-patch",
        "--state",
        testCase.statePath,
        "--in",
        testCase.patchPath,
        "--diagnostics",
        "json"
      ]);

      expect(textResult.status, `${testCase.label} text exit code`).toBe(testCase.expectedExitCode);
      expect(textResult.stdout, `${testCase.label} text status`).toContain(`status=${testCase.expectedStatus}`);
      expect(jsonResult.status, `${testCase.label} json exit code`).toBe(testCase.expectedExitCode);

      const jsonReport = JSON.parse(jsonResult.stdout);
      expect(jsonReport.status, `${testCase.label} json status`).toBe(testCase.expectedStatus);
      expect(jsonReport.exit_code, `${testCase.label} json report exit_code`).toBe(testCase.expectedExitCode);
    }
  });

  it("reports warning status for validate-patch when patch has warnings", () => {
    const statePath = writeTempJson("workbook.json", buildWorkbookState("CLI validate patch warning"));
    const patchPath = writeTempJson("patch.json", {
      operations: [
        {
          op: "unknown_op"
        }
      ]
    });

    const result = runCli(["ai", "validate-patch", "--state", statePath, "--in", patchPath, "--diagnostics", "json"]);

    expect(result.status).toBe(0);
    const report = JSON.parse(result.stdout);
    expect(report.ok).toBe(true);
    expect(report.command).toBe("ai validate-patch");
    expect(report.status).toBe("warning");
    expect(report.exit_code).toBe(0);
    expect(report.warning_count).toBeGreaterThan(0);
    expect(report.error_count).toBe(0);
    expect(report.warnings.length).toBeGreaterThan(0);
  });

  it("reports noop status for validate-patch when patch has no changes", () => {
    const statePath = writeTempJson("workbook.json", buildWorkbookState("CLI validate patch noop"));
    const patchPath = writeTempJson("patch.json", {
      operations: []
    });

    const result = runCli(["ai", "validate-patch", "--state", statePath, "--in", patchPath, "--diagnostics", "json"]);

    expect(result.status).toBe(0);
    const report = JSON.parse(result.stdout);
    expect(report.ok).toBe(true);
    expect(report.command).toBe("ai validate-patch");
    expect(report.status).toBe("noop");
    expect(report.exit_code).toBe(0);
    expect(report.changes_summary.total_changes).toBe(0);
  });

  it("validates patch_json with json output and nonzero exit on errors", () => {
    const statePath = writeTempJson("workbook.json", buildWorkbookState("CLI validate patch error"));
    const patchPath = writeTempJson("patch.json", {
      invalid: true
    });

    const result = runCli(["ai", "validate-patch", "--state", statePath, "--in", patchPath, "--diagnostics", "json"]);

    expect(result.status).toBe(1);
    const report = JSON.parse(result.stdout);
    expect(report.ok).toBe(false);
    expect(report.diagnostics_version).toBe(1);
    expect(report.command).toBe("ai validate-patch");
    expect(report.status).toBe("error");
    expect(report.exit_code).toBe(1);
    expect(report.warning_count).toBe(0);
    expect(report.error_count).toBe(1);
    expect(report.io).toEqual({
      inputs: [
        { option: "--state", mode: "file", path: statePath },
        { option: "--in", mode: "file", path: patchPath }
      ],
      output: { mode: "stdout" }
    });
    expect(report.error_type).toBe("processing_error");
    expect(report.error_code).toBe("invalid_patch_kind");
    expect(report.error_details).toEqual({
      context: "ai validate-patch",
      expected_kind: "patch_json",
      actual_kind: null
    });
    expect(Array.isArray(report.errors)).toBe(true);
    expect(report.errors[0].code).toBe("invalid_patch_kind");
    expect(report.errors[0].details).toEqual({
      context: "ai validate-patch",
      expected_kind: "patch_json",
      actual_kind: null
    });
    expect(report.errors[0].message).toContain("patch_json");
    expect(result.stderr).toBe("");
  });

  it("writes apply-patch diagnostics as json when requested", () => {
    const statePath = writeTempJson("workbook.json", buildWorkbookState("CLI patch json diagnostics"));
    const patchPath = writeTempJson("patch.json", {
      operations: [
        {
          op: "update_project",
          fields: {
            name: "CLI patch json diagnostics updated"
          }
        }
      ]
    });

    const result = runCli([
      "state",
      "apply-patch",
      "--state",
      statePath,
      "--in",
      patchPath,
      "--diagnostics",
      "json"
    ]);

    expect(result.status).toBe(0);
    const diagnostics = JSON.parse(result.stderr);
    expect(diagnostics.ok).toBe(true);
    expect(diagnostics.command).toBe("apply-patch");
    expect(diagnostics.context).toBe("apply-patch");
    expect(diagnostics.status).toBe("success");
    expect(diagnostics.exit_code).toBe(0);
    expect(diagnostics.warnings).toEqual([]);
    expect(diagnostics.errors).toEqual([]);
    expect(diagnostics.changes_summary.total_changes).toBe(1);
  });

  it("reports noop status for apply-patch diagnostics when patch has no changes", () => {
    const statePath = writeTempJson("workbook.json", buildWorkbookState("CLI patch noop diagnostics"));
    const patchPath = writeTempJson("patch.json", {
      operations: []
    });

    const result = runCli([
      "state",
      "apply-patch",
      "--state",
      statePath,
      "--in",
      patchPath,
      "--diagnostics",
      "json"
    ]);

    expect(result.status).toBe(0);
    const diagnostics = JSON.parse(result.stderr);
    expect(diagnostics.status).toBe("noop");
    expect(diagnostics.exit_code).toBe(0);
    expect(diagnostics.changes_summary.total_changes).toBe(0);
  });

  it("summarizes state diff between two workbook states", () => {
    const beforePath = writeTempJson("before.json", buildWorkbookState("CLI diff before"));
    const afterPath = writeTempJson("after.json", buildWorkbookState("CLI diff after"));

    const result = runCli(["state", "diff", "--before", beforePath, "--after", afterPath]);

    expect(result.status).toBe(0);
    const diff = JSON.parse(result.stdout);
    expect(diff.kind).toBe("state_diff_summary");
    expect(diff.changes_summary.total_changes).toBe(2);
    expect(diff.changes_summary.by_scope.project).toBe(2);
    expect(diff.changed_items.project.some((change) => change.field === "Name")).toBe(true);
  });

  it("writes state diff diagnostics as json when requested", () => {
    const beforePath = writeTempJson("before.json", buildWorkbookState("CLI diff json before"));
    const afterPath = writeTempJson("after.json", buildWorkbookState("CLI diff json after"));

    const result = runCli([
      "state",
      "diff",
      "--before",
      beforePath,
      "--after",
      afterPath,
      "--diagnostics",
      "json"
    ]);

    expect(result.status).toBe(0);
    const diagnostics = JSON.parse(result.stderr);
    expect(diagnostics.ok).toBe(true);
    expect(diagnostics.command).toBe("state diff");
    expect(diagnostics.context).toBe("state diff");
    expect(diagnostics.status).toBe("success");
    expect(diagnostics.exit_code).toBe(0);
    expect(Array.isArray(diagnostics.warnings)).toBe(true);
    expect(diagnostics.errors).toEqual([]);
    expect(diagnostics.changes_summary.total_changes).toBe(2);
  });

  it("exports xml from workbook state", () => {
    const workbookPath = writeTempJson("workbook.json", buildWorkbookState("CLI export xml"));

    const result = runCli(["export", "xml", "--in", workbookPath]);

    expect(result.status).toBe(0);
    expect(result.stdout).toContain("<Project");
    expect(result.stdout).toContain("<Name>CLI export xml</Name>");
  });

  it("writes export xml diagnostics as json when requested", () => {
    const workbookPath = writeTempJson("workbook.json", buildWorkbookState("CLI export xml diagnostics"));

    const result = runCli(["export", "xml", "--in", workbookPath, "--diagnostics", "json"]);

    expect(result.status).toBe(0);
    const diagnostics = JSON.parse(result.stderr);
    expect(diagnostics.ok).toBe(true);
    expect(diagnostics.command).toBe("export xml");
    expect(diagnostics.context).toBe("export xml");
    expect(diagnostics.status).toBe("success");
    expect(diagnostics.exit_code).toBe(0);
    expect(diagnostics.warnings).toEqual([]);
    expect(diagnostics.errors).toEqual([]);
    expect(diagnostics.output_kind).toBe("ms_project_xml");
    expect(typeof diagnostics.output_length).toBe("number");
  });

  it("exports xlsx bytes from workbook state", () => {
    const workbookPath = writeTempJson("workbook.json", buildWorkbookState("CLI export xlsx"));
    const outputPath = path.join(createTempDir("miku-project-cli-out-"), "project.xlsx");

    const result = runCli(["export", "xlsx", "--in", workbookPath, "--out", outputPath], { encoding: "buffer" });
    const output = readFileSync(outputPath);

    expect(result.status).toBe(0);
    expect(result.stdout.length).toBe(0);
    expect(output.subarray(0, 2).toString("utf8")).toBe("PK");
  });

  it("rejects binary stdout output for xlsx export", () => {
    const workbookPath = writeTempJson("workbook.json", buildWorkbookState("CLI export xlsx stdout"));

    const result = runCli(["export", "xlsx", "--in", workbookPath, "--out", "-", "--diagnostics", "json"]);

    expect(result.status).toBe(2);
    expect(result.stdout).toBe("");
    const diagnostics = JSON.parse(result.stderr);
    expect(diagnostics.error_code).toBe("binary_stdout_not_supported");
    expect(diagnostics.errors[0].code).toBe("binary_stdout_not_supported");
  });

  it("exports xlsx bytes as base64 when --out-base64 - is specified", () => {
    const workbookPath = writeTempJson("workbook.json", buildWorkbookState("CLI export xlsx base64"));

    const result = runCli(["export", "xlsx", "--in", workbookPath, "--out-base64", "-"]);
    const output = Buffer.from(result.stdout.trim(), "base64");

    expect(result.status).toBe(0);
    expect(output.subarray(0, 2).toString("utf8")).toBe("PK");
    expect(result.stderr).toBe("");
  });

  it("imports xlsx bytes from base64 stdin as workbook json", () => {
    const loaded = loadMikuprojectCoreApi({ rootDir: repoRoot });
    disposers.push(() => loaded.dispose());
    const model = loaded.api.workbookJson.importAsProjectModel(buildWorkbookState("CLI import xlsx base64 base")).model;
    const xlsxWorkbook = loaded.api.xlsx.exportWorkbook(model);
    xlsxWorkbook.sheets
      .find((sheet) => sheet.name === "Project")
      .rows.find((row) => row.cells[0]?.value === "Name")
      .cells[1].value = "CLI import xlsx base64";
    const bytes = loaded.api.xlsx.encodeWorkbook(xlsxWorkbook);

    const result = runCli(["import", "xlsx", "--in-base64", "-", "--out", "-"], {
      input: `${Buffer.from(bytes).toString("base64")}\n`
    });

    expect(result.status).toBe(0);
    const workbook = JSON.parse(result.stdout);
    expect(workbook.format).toBe("mikuproject_workbook_json");
    expect(workbook.sheets.Project.find((row) => row.Field === "Name").Value).toBe("CLI import xlsx base64");
    expect(result.stderr).toBe("");
  });

  it("exports report wbs-xlsx bytes from workbook state", () => {
    const workbookPath = writeTempJson("workbook.json", buildWorkbookState("CLI report wbs-xlsx"));
    const outputPath = path.join(createTempDir("miku-project-cli-out-"), "wbs.xlsx");

    const result = runCli(["report", "wbs-xlsx", "--in", workbookPath, "--out", outputPath], { encoding: "buffer" });
    const output = readFileSync(outputPath);

    expect(result.status).toBe(0);
    expect(result.stdout.length).toBe(0);
    expect(output.subarray(0, 2).toString("utf8")).toBe("PK");
  });

  it("exports report all as base64 when --out-base64 - is specified", () => {
    const workbookPath = writeTempJson("workbook.json", buildWorkbookState("CLI report all base64"));
    const loaded = loadMikuprojectCoreApi({ rootDir: repoRoot });
    disposers.push(() => loaded.dispose());

    const result = runCli(["report", "all", "--in", workbookPath, "--out-base64", "-"]);
    const output = Buffer.from(result.stdout.trim(), "base64");
    const entryNames = new globalThis.__mikuprojectExcelIo.XlsxWorkbookCodec().listEntries(output);

    expect(result.status).toBe(0);
    expect(output.subarray(0, 2).toString("utf8")).toBe("PK");
    expect(entryNames).toContain("wbs.xlsx");
    expect(entryNames).toContain("wbs.md");
    expect(entryNames).toContain("mermaid.mmd");
    expect(result.stderr).toBe("");
  });

  it("exports report daily-svg from workbook state", () => {
    const workbookPath = writeTempJson("workbook.json", buildWorkbookState("CLI report daily-svg"));

    const result = runCli(["report", "daily-svg", "--in", workbookPath]);

    expect(result.status).toBe(0);
    expect(result.stdout).toContain("<svg");
    expect(result.stdout).toContain("CLI report daily-svg");
  });

  it("exports report weekly-svg from workbook state", () => {
    const workbookPath = writeTempJson("workbook.json", buildWorkbookState("CLI report weekly-svg"));

    const result = runCli(["report", "weekly-svg", "--in", workbookPath]);

    expect(result.status).toBe(0);
    expect(result.stdout).toContain("<svg");
    expect(result.stdout).toContain("CLI report weekly-svg");
  });

  it("exports report monthly-calendar-svg as zip bytes from workbook state", () => {
    const workbookPath = writeTempJson("workbook.json", buildWorkbookState("CLI report monthly-calendar-svg"));
    const outputPath = path.join(createTempDir("miku-project-cli-out-"), "monthly.zip");

    const result = runCli(["report", "monthly-calendar-svg", "--in", workbookPath, "--out", outputPath], { encoding: "buffer" });
    const output = readFileSync(outputPath);

    expect(result.status).toBe(0);
    expect(result.stdout.length).toBe(0);
    expect(output.subarray(0, 2).toString("utf8")).toBe("PK");
  });

  it("exports report all as a zip bundle from workbook state", () => {
    const workbookPath = writeTempJson("workbook.json", buildWorkbookState("CLI report all"));
    const loaded = loadMikuprojectCoreApi({ rootDir: repoRoot });
    disposers.push(() => loaded.dispose());
    const outputPath = path.join(createTempDir("miku-project-cli-out-"), "report-bundle.zip");

    const result = runCli(["report", "all", "--in", workbookPath, "--out", outputPath], { encoding: "buffer" });
    const output = readFileSync(outputPath);
    const entryNames = new globalThis.__mikuprojectExcelIo.XlsxWorkbookCodec().listEntries(output);

    expect(result.status).toBe(0);
    expect(result.stdout.length).toBe(0);
    expect(output.subarray(0, 2).toString("utf8")).toBe("PK");
    expect(entryNames).toContain("wbs.xlsx");
    expect(entryNames).toContain("wbs.md");
    expect(entryNames).toContain("mermaid.mmd");
    expect(entryNames).toContain("daily.svg");
    expect(entryNames).toContain("weekly.svg");
    expect(entryNames.some((name) => name.startsWith("monthly-calendar/"))).toBe(true);
    expect(entryNames).not.toContain("monthly-calendar.zip");
  });

  it("exports report wbs-markdown from workbook state", () => {
    const workbookPath = writeTempJson("workbook.json", buildWorkbookState("CLI report wbs-markdown"));

    const result = runCli(["report", "wbs-markdown", "--in", workbookPath]);

    expect(result.status).toBe(0);
    expect(result.stdout).toContain("# プロジェクト情報");
    expect(result.stdout).toContain("CLI report wbs-markdown");
  });

  it("exports report mermaid from workbook state", () => {
    const workbookPath = writeTempJson("workbook.json", buildWorkbookState("CLI report mermaid"));

    const result = runCli(["report", "mermaid", "--in", workbookPath]);

    expect(result.status).toBe(0);
    expect(result.stdout).toContain("gantt");
    expect(result.stdout).toContain("title CLI report mermaid");
  });

  it("writes report mermaid diagnostics as json when requested", () => {
    const workbookPath = writeTempJson("workbook.json", buildWorkbookState("CLI report mermaid diagnostics"));

    const result = runCli(["report", "mermaid", "--in", workbookPath, "--diagnostics", "json"]);

    expect(result.status).toBe(0);
    const diagnostics = JSON.parse(result.stderr);
    expect(diagnostics.ok).toBe(true);
    expect(diagnostics.command).toBe("report mermaid");
    expect(diagnostics.context).toBe("report mermaid");
    expect(diagnostics.status).toBe("success");
    expect(diagnostics.exit_code).toBe(0);
    expect(diagnostics.warnings).toEqual([]);
    expect(diagnostics.errors).toEqual([]);
    expect(diagnostics.output_kind).toBe("mermaid");
    expect(typeof diagnostics.output_length).toBe("number");
  });

  it("builds a single-file cli runtime artifact that runs outside the repo", () => {
    const bundlePath = path.join(createTempDir("miku-project-cli-bundle-test-"), "miku-project.mjs");
    const sourcesPath = path.join(path.dirname(bundlePath), "miku-project-sources.tgz");
    const buildResult = spawnSync(process.execPath, [cliBundleBuildPath, "--out", bundlePath], {
      cwd: repoRoot,
      encoding: "utf8"
    });

    expect(buildResult.status).toBe(0);
    expect(existsSync(bundlePath)).toBe(true);
    expect(existsSync(sourcesPath)).toBe(true);
    const sourceEntries = listTarGzEntries(sourcesPath);
    expect(sourceEntries).toEqual([...sourceEntries].sort());
    expect(sourceEntries).toEqual(expect.arrayContaining([
      "miku-project-sources/package.json",
      "miku-project-sources/README.md",
      "miku-project-sources/scripts/build-cli-bundle.mjs",
      "miku-project-sources/scripts/miku-project-cli.mjs",
      "miku-project-sources/scripts/lib/core-api-loader.mjs",
      "miku-project-sources/src/ts/core-api.ts",
      "miku-project-sources/src/js/core-api.js",
      "miku-project-sources/docs/miku-soft-40-agentskills-design-v20260501.md",
      "miku-project-sources/tests/mikuproject-cli.test.js"
    ]));

    const workbookPath = writeTempJson("bundle-workbook.json", buildWorkbookState("Bundled CLI export xml"));
    const result = spawnSync(process.execPath, [bundlePath, "export", "xml", "--in", workbookPath], {
      cwd: path.dirname(workbookPath),
      encoding: "utf8"
    });

    expect(result.status).toBe(0);
    expect(result.stdout).toContain("<Project");
    expect(result.stdout).toContain("<Name>Bundled CLI export xml</Name>");
  });

  it("prints the version from a single-file cli runtime artifact outside the repo", () => {
    const bundlePath = path.join(createTempDir("miku-project-cli-version-bundle-test-"), "miku-project.mjs");
    const buildResult = spawnSync(process.execPath, [cliBundleBuildPath, "--out", bundlePath], {
      cwd: repoRoot,
      encoding: "utf8"
    });

    expect(buildResult.status).toBe(0);

    const result = spawnSync(process.execPath, [bundlePath, "--version"], {
      cwd: path.dirname(bundlePath),
      encoding: "utf8"
    });

    expect(result.status).toBe(0);
    expect(result.stdout).toBe(`miku-project ${packageVersion}\n`);
    expect(result.stderr).toBe("");
  });

  it("provides a runnable CLI AI workflow example script", () => {
    const result = spawnSync(process.execPath, [cliAiWorkflowExamplePath], {
      cwd: repoRoot,
      encoding: "utf8"
    });

    expect(result.status).toBe(0);
    const output = JSON.parse(result.stdout);
    expect(output.project_name).toBe("Dependency Project");
    expect(output.task_edit_target_uid).toBe("2");
    expect(output.validate_patch.ok).toBe(true);
    expect(output.validate_patch.status).toBe("success");
    expect(output.apply_patch_diagnostics.context).toBe("apply-patch");
    expect(output.diff_summary.kind).toBe("state_diff_summary");
    expect(output.diff_summary.changes_summary.total_changes).toBeGreaterThan(0);
  });

  it("provides a runnable CLI stdio workflow example script", () => {
    const result = spawnSync(process.execPath, [cliAiStdioExamplePath], {
      cwd: repoRoot,
      encoding: "utf8"
    });

    expect(result.status).toBe(0);
    const output = JSON.parse(result.stdout);
    expect(output.validate_patch.ok).toBe(true);
    expect(output.validate_patch.status).toBe("success");
    expect(output.apply_patch_diagnostics.context).toBe("apply-patch");
    expect(output.patched_workbook_format).toBe("mikuproject_workbook_json");
    expect(output.patched_task_name).toContain("renamed through stdin");
  });

  it("reports io metadata for implicit stdin detect-kind diagnostics", () => {
    const result = runCli(["ai", "detect-kind", "--diagnostics", "json"], {
      input: `${JSON.stringify({ operations: [] }, null, 2)}\n`
    });

    expect(result.status).toBe(0);
    const diagnostics = JSON.parse(result.stderr);
    expect(diagnostics.io).toEqual({
      inputs: [{ option: "--in", mode: "stdin_implicit" }],
      output: { mode: "stdout" }
    });
  });
});

function runCli(args, options = {}) {
  return spawnSync(process.execPath, [cliPath, ...args], {
    cwd: repoRoot,
    encoding: options.encoding === "buffer" ? null : "utf8",
    input: options.input
  });
}

function writeTempJson(fileName, documentLike) {
  const dir = createTempDir("miku-project-cli-test-");
  const filePath = path.join(dir, fileName);
  writeFileSync(filePath, `${JSON.stringify(documentLike, null, 2)}\n`, "utf8");
  return filePath;
}

function createTempDir(prefix) {
  const dir = mkdtempSync(path.join(os.tmpdir(), prefix));
  tempDirs.push(dir);
  return dir;
}

function listTarGzEntries(filePath) {
  const tar = gunzipSync(readFileSync(filePath));
  const entries = [];
  for (let offset = 0; offset + 512 <= tar.length;) {
    const header = tar.subarray(offset, offset + 512);
    if (header.every((byte) => byte === 0)) {
      break;
    }
    const name = readTarString(header, 0, 100);
    const prefix = readTarString(header, 345, 155);
    const sizeText = readTarString(header, 124, 12).trim();
    const size = sizeText ? Number.parseInt(sizeText, 8) : 0;
    entries.push(prefix ? `${prefix}/${name}` : name);
    offset += 512 + Math.ceil(size / 512) * 512;
  }
  return entries;
}

function readTarString(buffer, offset, length) {
  const slice = buffer.subarray(offset, offset + length);
  const end = slice.indexOf(0);
  return slice.subarray(0, end === -1 ? slice.length : end).toString("utf8");
}

function buildWorkbookState(projectName) {
  const loaded = loadMikuprojectCoreApi({ rootDir: repoRoot });
  try {
    const documentLike = loaded.api.workbookJson.exportDocument(
      loaded.api.importAiJsonDocument({
        view_type: "project_draft_view",
        project: {
          name: projectName,
          planned_start: "2026-04-01"
        },
        tasks: [
          {
            uid: "draft-1",
            name: "開始",
            parent_uid: null,
            position: 0,
            is_milestone: true,
            planned_start: "2026-04-01",
            planned_finish: "2026-04-01"
          }
        ],
        resources: [],
        assignments: []
      }).model
    );
    return documentLike;
  } finally {
    loaded.dispose();
  }
}

function readFixture(fileName) {
  return readFileSync(path.resolve(repoRoot, "testdata", fileName), "utf8");
}
