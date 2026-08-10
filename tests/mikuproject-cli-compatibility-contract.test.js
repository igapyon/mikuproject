import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { spawnSync } from "node:child_process";

import { afterEach, describe, expect, it } from "vitest";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..");
const cliPath = path.resolve(repoRoot, "scripts/miku-project-cli.mjs");
const packageVersion = JSON.parse(readFileSync(path.resolve(repoRoot, "package.json"), "utf8")).version;
const tempDirs = [];

afterEach(() => {
  while (tempDirs.length > 0) {
    rmSync(tempDirs.pop(), { recursive: true, force: true });
  }
});

describe("legacy CLI compatibility contract", () => {
  it("keeps the documented legacy command surface, help, and version", () => {
    const help = runCli(["--help"]);

    expect(help.status).toBe(0);
    expect(help.stderr).toBe("");
    for (const usage of LEGACY_USAGE_LINES) {
      expect(help.stdout).toContain(usage);
    }
    expect(help.stdout).toContain("--diagnostics defaults to text");
    expect(help.stdout).toContain("--out <path> overwrites an existing file without an interactive prompt.");

    const version = runCli(["--version"]);
    expect(version.status).toBe(0);
    expect(version.stdout).toBe(`miku-project ${packageVersion}\n`);
    expect(version.stderr).toBe("");
  });

  it("keeps ai spec as a successful text command", () => {
    const result = runCli(["ai", "spec"]);

    expect(result.status).toBe(0);
    expect(result.stdout).toContain("# miku-project AI JSON Prompt / Spec");
    expect(result.stderr).toBe("");
  });

  it("keeps JSON primary output and JSON diagnostics separated for stdin", () => {
    const result = runCli(["ai", "detect-kind", "--in", "-", "--diagnostics", "json"], {
      input: "{\"operations\":[]}\n"
    });

    expect(result.status).toBe(0);
    expect(result.stdout).toBe("patch_json\n");
    expect(JSON.parse(result.stderr)).toMatchObject({
      ok: true,
      diagnostics_version: 1,
      command: "detect-kind",
      status: "success",
      exit_code: 0,
      io: {
        inputs: [{ option: "--in", mode: "stdin" }],
        output: { mode: "stdout" }
      },
      detected_kind: "patch_json"
    });
  });

  it("keeps named file output and the legacy overwrite behavior", () => {
    const outputPath = path.join(createTempDir(), "detected-kind.txt");
    writeFileSync(outputPath, "old output\n", "utf8");

    const result = runCli([
      "ai", "detect-kind",
      "--in", "-",
      "--out", outputPath,
      "--diagnostics", "json"
    ], {
      input: "{\"operations\":[]}\n"
    });

    expect(result.status).toBe(0);
    expect(result.stdout).toBe("");
    expect(readFileSync(outputPath, "utf8")).toBe("patch_json\n");
    // The legacy diagnostics envelope does not reflect --out. Keep that
    // observed compatibility behavior separate from the v1 result contract.
    expect(JSON.parse(result.stderr)).toMatchObject({
      ok: true,
      io: {
        inputs: [{ option: "--in", mode: "stdin" }],
        output: { mode: "stdout" }
      }
    });
  });

  it("keeps usage errors structured when JSON diagnostics are requested", () => {
    const result = runCli(["legacy-command", "--diagnostics", "json"]);

    expect(result.status).toBe(2);
    expect(result.stdout).toBe("");
    expect(JSON.parse(result.stderr)).toMatchObject({
      ok: false,
      diagnostics_version: 1,
      command: "legacy-command",
      status: "error",
      exit_code: 2,
      error_type: "usage_error",
      error_code: "unsupported_command",
      io: {
        inputs: [],
        output: { mode: "stdout" }
      },
      errors: [{ code: "unsupported_command" }]
    });
  });

  it("keeps project draft conversion available through stdin", () => {
    const result = runCli(["state", "from-draft", "--in", "-"], {
      input: `${JSON.stringify({
        view_type: "project_draft_view",
        project: { name: "Compatibility contract", planned_start: "2026-04-01" },
        tasks: [{
          uid: "draft-1",
          name: "Start",
          parent_uid: null,
          position: 0,
          is_milestone: true,
          planned_start: "2026-04-01",
          planned_finish: "2026-04-01"
        }],
        resources: [],
        assignments: []
      })}\n`
    });

    expect(result.status).toBe(0);
    expect(result.stderr).toBe("");
    const workbook = JSON.parse(result.stdout);
    expect(workbook.format).toBe("mikuproject_workbook_json");
    expect(workbook.sheets.Project.find((row) => row.Field === "Name").Value).toBe("Compatibility contract");
  });
});

const LEGACY_USAGE_LINES = [
  "miku-project ai spec",
  "miku-project ai export project-overview",
  "miku-project ai export task-edit",
  "miku-project ai export phase-detail",
  "miku-project ai export bundle",
  "miku-project ai detect-kind",
  "miku-project ai validate-patch",
  "miku-project state from-draft",
  "miku-project state summarize",
  "miku-project state diff",
  "miku-project state apply-patch",
  "miku-project import xlsx",
  "miku-project export workbook-json",
  "miku-project export xml",
  "miku-project export xlsx",
  "miku-project report wbs-xlsx",
  "miku-project report daily-svg",
  "miku-project report weekly-svg",
  "miku-project report monthly-calendar-svg",
  "miku-project report all",
  "miku-project report wbs-markdown",
  "miku-project report mermaid"
];

function runCli(args, options = {}) {
  return spawnSync(process.execPath, [cliPath, ...args], {
    cwd: repoRoot,
    encoding: "utf8",
    input: options.input
  });
}

function createTempDir() {
  const dir = mkdtempSync(path.join(os.tmpdir(), "miku-project-cli-compatibility-"));
  tempDirs.push(dir);
  return dir;
}
