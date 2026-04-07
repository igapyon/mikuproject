import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { spawnSync } from "node:child_process";

import { afterEach, describe, expect, it } from "vitest";

import { loadMikuprojectCoreApi } from "../scripts/lib/core-api-loader.mjs";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..");
const cliPath = path.resolve(repoRoot, "scripts/mikuproject-cli.mjs");

const tempDirs = [];

afterEach(() => {
  while (tempDirs.length > 0) {
    rmSync(tempDirs.pop(), { recursive: true, force: true });
  }
});

describe("mikuproject cli", () => {
  it("prints the AI spec", () => {
    const result = runCli(["ai", "spec"]);

    expect(result.status).toBe(0);
    expect(result.stdout).toContain("# mikuproject AI JSON Prompt / Spec");
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
    expect(result.stderr).toContain("applied patch_json");
  });

  it("exports workbook json canonically", () => {
    const workbookPath = writeTempJson("workbook.json", buildWorkbookState("CLI export workbook"));

    const result = runCli(["export", "workbook-json", "--in", workbookPath]);

    expect(result.status).toBe(0);
    const workbook = JSON.parse(result.stdout);
    expect(workbook.format).toBe("mikuproject_workbook_json");
    expect(workbook.sheets.Project.find((row) => row.Field === "Name").Value).toBe("CLI export workbook");
  });

  it("exports xml from workbook state", () => {
    const workbookPath = writeTempJson("workbook.json", buildWorkbookState("CLI export xml"));

    const result = runCli(["export", "xml", "--in", workbookPath]);

    expect(result.status).toBe(0);
    expect(result.stdout).toContain("<Project");
    expect(result.stdout).toContain("<Name>CLI export xml</Name>");
  });

  it("exports xlsx bytes from workbook state", () => {
    const workbookPath = writeTempJson("workbook.json", buildWorkbookState("CLI export xlsx"));

    const result = runCli(["export", "xlsx", "--in", workbookPath], { encoding: "buffer" });

    expect(result.status).toBe(0);
    expect(Buffer.isBuffer(result.stdout)).toBe(true);
    expect(result.stdout.subarray(0, 2).toString("utf8")).toBe("PK");
  });
});

function runCli(args, options = {}) {
  return spawnSync(process.execPath, [cliPath, ...args], {
    cwd: repoRoot,
    encoding: options.encoding === "buffer" ? null : "utf8"
  });
}

function writeTempJson(fileName, documentLike) {
  const dir = mkdtempSync(path.join(os.tmpdir(), "mikuproject-cli-test-"));
  tempDirs.push(dir);
  const filePath = path.join(dir, fileName);
  writeFileSync(filePath, `${JSON.stringify(documentLike, null, 2)}\n`, "utf8");
  return filePath;
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
