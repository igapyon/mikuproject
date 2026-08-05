import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { spawnSync } from "node:child_process";

import { loadMikuprojectCoreApi } from "./lib/core-api-loader.mjs";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..");
const cliPath = path.resolve(repoRoot, "scripts/miku-project-cli.mjs");
const dependencyXmlPath = path.resolve(repoRoot, "testdata/dependency.xml");

const loaded = loadMikuprojectCoreApi({ rootDir: repoRoot });
const tempDir = mkdtempSync(path.join(os.tmpdir(), "miku-project-cli-ai-workflow-"));

try {
  const model = loaded.api.msProject.importFromXml(readFileSync(dependencyXmlPath, "utf8"));
  const workbookPath = path.join(tempDir, "workbook.json");
  writeFileSync(workbookPath, `${JSON.stringify(loaded.api.workbookJson.exportDocument(model), null, 2)}\n`, "utf8");

  const projectOverview = runCli(["ai", "export", "project-overview", "--in", workbookPath]);
  const taskEdit = runCli(["ai", "export", "task-edit", "--in", workbookPath, "--task-uid", "2"]);

  const patchPath = path.join(tempDir, "patch.json");
  writeFileSync(patchPath, `${JSON.stringify({
    operations: [
      {
        op: "update_task",
        uid: "2",
        fields: {
          name: "Task 2 renamed by CLI workflow example"
        }
      }
    ]
  }, null, 2)}\n`, "utf8");

  const validatePatch = runCli([
    "ai",
    "validate-patch",
    "--state",
    workbookPath,
    "--in",
    patchPath,
    "--diagnostics",
    "json"
  ]);
  const applyPatch = runCli([
    "state",
    "apply-patch",
    "--state",
    workbookPath,
    "--in",
    patchPath,
    "--diagnostics",
    "json"
  ]);
  const diff = runCli([
    "state",
    "diff",
    "--before",
    workbookPath,
    "--after",
    path.join(tempDir, "after.json")
  ], {
    prepare() {
      writeFileSync(path.join(tempDir, "after.json"), applyPatch.stdout, "utf8");
    }
  });

  process.stdout.write(`${JSON.stringify({
    project_name: JSON.parse(projectOverview.stdout).project?.name,
    task_edit_target_uid: JSON.parse(taskEdit.stdout).target_task?.uid,
    validate_patch: JSON.parse(validatePatch.stdout),
    apply_patch_diagnostics: JSON.parse(applyPatch.stderr),
    diff_summary: JSON.parse(diff.stdout)
  }, null, 2)}\n`);
} finally {
  loaded.dispose();
  rmSync(tempDir, { recursive: true, force: true });
}

function runCli(args, options = {}) {
  if (typeof options.prepare === "function") {
    options.prepare();
  }
  const result = spawnSync(process.execPath, [cliPath, ...args], {
    cwd: repoRoot,
    encoding: "utf8"
  });
  if (result.status !== 0) {
    throw new Error(`CLI failed: ${args.join(" ")}\n${result.stderr}`);
  }
  return result;
}
