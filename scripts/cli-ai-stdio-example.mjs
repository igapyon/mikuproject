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
const tempDir = mkdtempSync(path.join(os.tmpdir(), "miku-project-cli-stdio-"));

try {
  const model = loaded.api.msProject.importFromXml(readFileSync(dependencyXmlPath, "utf8"));
  const workbookDocument = loaded.api.workbookJson.exportDocument(model);
  const workbookPath = path.join(tempDir, "workbook.json");
  writeFileSync(workbookPath, `${JSON.stringify(workbookDocument, null, 2)}\n`, "utf8");

  const patchJsonText = `${JSON.stringify({
    operations: [
      {
        op: "update_task",
        uid: "2",
        fields: {
          name: "Task 2 renamed through stdin"
        }
      }
    ]
  }, null, 2)}\n`;

  const validatePatch = runCli([
    "ai",
    "validate-patch",
    "--state",
    workbookPath,
    "--in",
    "-",
    "--diagnostics",
    "json"
  ], { input: patchJsonText });

  const applyPatch = runCli([
    "state",
    "apply-patch",
    "--state",
    workbookPath,
    "--in",
    "-",
    "--diagnostics",
    "json",
    "--out",
    "-"
  ], { input: patchJsonText });

  process.stdout.write(`${JSON.stringify({
    validate_patch: JSON.parse(validatePatch.stdout),
    apply_patch_diagnostics: JSON.parse(applyPatch.stderr),
    patched_workbook_format: JSON.parse(applyPatch.stdout).format,
    patched_task_name: JSON.parse(applyPatch.stdout).sheets.Tasks.find((row) => row.UID === "2")?.Name || ""
  }, null, 2)}\n`);
} finally {
  loaded.dispose();
  rmSync(tempDir, { recursive: true, force: true });
}

function runCli(args, options = {}) {
  const result = spawnSync(process.execPath, [cliPath, ...args], {
    cwd: repoRoot,
    encoding: "utf8",
    input: options.input
  });
  if (result.status !== 0) {
    throw new Error(`CLI failed: ${args.join(" ")}\n${result.stderr}`);
  }
  return result;
}
