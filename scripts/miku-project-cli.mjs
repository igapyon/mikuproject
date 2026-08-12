#!/usr/bin/env node

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { loadMikuprojectCoreApi } from "./lib/core-api-loader.mjs";
import { CliUsageError } from "./lib/cli-errors.mjs";
import { detectRequestedDiagnosticsFormat, parseArgs } from "./lib/cli-argv.mjs";
import { writeOutput } from "./lib/cli-io.mjs";
import { buildErrorDiagnostics, writeDiagnostics } from "./lib/cli-diagnostics.mjs";
import { writeHelp } from "./lib/cli-presentation.mjs";
import { runLegacyCommand } from "./lib/cli-legacy-router.mjs";
import { recognizesV1Workflow, rejectUnreleasedV1Workflow } from "./lib/v1/cli-v1-router.mjs";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..");

async function main() {
  const rawArgv = process.argv.slice(2);
  if (recognizesV1Workflow(rawArgv)) {
    const result = rejectUnreleasedV1Workflow(rawArgv, {
      version: getCliVersion(),
      stdout: process.stdout
    });
    process.exitCode = result.exit_code;
    return;
  }
  const { command, options } = parseArgs(rawArgv);
  if (options.version) {
    process.stdout.write(`miku-project ${getCliVersion()}\n`);
    return;
  }
  if (options.help || command.length === 0) {
    writeHelp(process.stdout);
    return;
  }

  const loaded = loadMikuprojectCoreApi({ rootDir: repoRoot });
  try {
    const result = await runLegacyCommand(command, options, loaded.api);
    if (Array.isArray(result.diagnostics) && result.diagnostics.length > 0) {
      writeDiagnostics(process.stderr, result.diagnostics);
    } else if (result.diagnostics && !Array.isArray(result.diagnostics)) {
      writeDiagnostics(process.stderr, result.diagnostics);
    }
    writeOutput(result.output, options);
    if (typeof result.exitCode === "number") {
      process.exitCode = result.exitCode;
    }
  } finally {
    loaded.dispose();
  }
}

function getCliVersion() {
  if (typeof BUNDLED_PACKAGE_VERSION === "string" && BUNDLED_PACKAGE_VERSION) {
    return BUNDLED_PACKAGE_VERSION;
  }
  try {
    const packageJson = JSON.parse(fs.readFileSync(path.join(repoRoot, "package.json"), "utf8"));
    if (typeof packageJson.version === "string" && packageJson.version) {
      return packageJson.version;
    }
  } catch {
    // Fall through to an explicit unknown marker when package metadata is unavailable.
  }
  return "unknown";
}

main().catch((error) => {
  const rawArgv = process.argv.slice(2);
  const diagnosticsFormat = detectRequestedDiagnosticsFormat(rawArgv);
  const exitCode = error instanceof CliUsageError ? 2 : 1;
  if (diagnosticsFormat === "json") {
    process.stderr.write(`${JSON.stringify(buildErrorDiagnostics(rawArgv, error, exitCode), null, 2)}\n`);
  } else {
    process.stderr.write(`[miku-project-cli] ${error.message}\n`);
  }
  process.exit(exitCode);
});
