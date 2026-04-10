import { existsSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
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
const cliBundleBuildPath = path.resolve(repoRoot, "scripts/build-cli-bundle.mjs");

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

  it("exports report wbs-xlsx bytes from workbook state", () => {
    const workbookPath = writeTempJson("workbook.json", buildWorkbookState("CLI report wbs-xlsx"));

    const result = runCli(["report", "wbs-xlsx", "--in", workbookPath], { encoding: "buffer" });

    expect(result.status).toBe(0);
    expect(Buffer.isBuffer(result.stdout)).toBe(true);
    expect(result.stdout.subarray(0, 2).toString("utf8")).toBe("PK");
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

    const result = runCli(["report", "monthly-calendar-svg", "--in", workbookPath], { encoding: "buffer" });

    expect(result.status).toBe(0);
    expect(Buffer.isBuffer(result.stdout)).toBe(true);
    expect(result.stdout.subarray(0, 2).toString("utf8")).toBe("PK");
  });

  it("exports report all as a zip bundle from workbook state", () => {
    const workbookPath = writeTempJson("workbook.json", buildWorkbookState("CLI report all"));
    const loaded = loadMikuprojectCoreApi({ rootDir: repoRoot });
    disposers.push(() => loaded.dispose());

    const result = runCli(["report", "all", "--in", workbookPath], { encoding: "buffer" });
    const entryNames = new globalThis.__mikuprojectExcelIo.XlsxWorkbookCodec().listEntries(result.stdout);

    expect(result.status).toBe(0);
    expect(Buffer.isBuffer(result.stdout)).toBe(true);
    expect(result.stdout.subarray(0, 2).toString("utf8")).toBe("PK");
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

  it("builds a self-contained cli bundle that runs outside the repo", () => {
    const bundleRoot = path.join(createTempDir("mikuproject-cli-bundle-test-"), "bundle");
    const buildResult = spawnSync(process.execPath, [cliBundleBuildPath, "--out", bundleRoot], {
      cwd: repoRoot,
      encoding: "utf8"
    });

    expect(buildResult.status).toBe(0);
    expect(existsSync(path.join(bundleRoot, "node_modules", "jsdom", "package.json"))).toBe(true);

    const workbookPath = writeTempJson("bundle-workbook.json", buildWorkbookState("Bundled CLI export xml"));
    const bundledCliPath = path.join(bundleRoot, "scripts", "mikuproject-cli.mjs");
    const result = spawnSync(process.execPath, [bundledCliPath, "export", "xml", "--in", workbookPath], {
      cwd: path.dirname(workbookPath),
      encoding: "utf8"
    });

    expect(result.status).toBe(0);
    expect(result.stdout).toContain("<Project");
    expect(result.stdout).toContain("<Name>Bundled CLI export xml</Name>");
  });
});

function runCli(args, options = {}) {
  return spawnSync(process.execPath, [cliPath, ...args], {
    cwd: repoRoot,
    encoding: options.encoding === "buffer" ? null : "utf8"
  });
}

function writeTempJson(fileName, documentLike) {
  const dir = createTempDir("mikuproject-cli-test-");
  const filePath = path.join(dir, fileName);
  writeFileSync(filePath, `${JSON.stringify(documentLike, null, 2)}\n`, "utf8");
  return filePath;
}

function createTempDir(prefix) {
  const dir = mkdtempSync(path.join(os.tmpdir(), prefix));
  tempDirs.push(dir);
  return dir;
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
