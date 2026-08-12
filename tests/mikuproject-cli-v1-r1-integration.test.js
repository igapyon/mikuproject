import { existsSync, mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { spawnSync } from "node:child_process";
import { gunzipSync } from "node:zlib";

import { afterEach, describe, expect, it } from "vitest";

import { canonicalJsonText, sha256CanonicalJson } from "../scripts/lib/v1/cli-v1-canonical-json.mjs";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const sourceCliPath = path.join(repoRoot, "scripts/miku-project-cli.mjs");
const r1HarnessPath = path.join(repoRoot, "tests/helpers/mikuproject-cli-v1-r1-harness.mjs");
const cliBundleBuildPath = path.join(repoRoot, "scripts/build-cli-bundle.mjs");
const canonicalFixturePath = path.join(repoRoot, "testdata/conformance/v1/fixtures/project/dependency-canonical.xml");
const changeRequestTemplatePath = path.join(repoRoot, "testdata/conformance/v1/fixtures/change/set-task-2-percent-0-to-50.template.json");
const temporaryDirectories = [];

afterEach(() => {
  while (temporaryDirectories.length > 0) {
    rmSync(temporaryDirectories.pop(), { recursive: true, force: true });
  }
});

describe("v1 R1 subprocess integration and provisional public boundary", () => {
  it("runs the fixed-binding R1 harness through stdout/file transports and is byte deterministic", () => {
    const directory = createTemporaryDirectory("miku-project-v1-r1-subprocess-");
    const resultPath = path.join(directory, "validate.result.json");
    const fileResult = runR1Harness([
      "validate", "--project", canonicalFixturePath, "--result", resultPath
    ]);

    expect(fileResult.status).toBe(0);
    expect(fileResult.stdout).toBe("");
    expect(fileResult.stderr).toBe("");
    const validation = readJson(resultPath);
    expect(validation).toMatchObject({
      command: "validate",
      status: "succeeded",
      exit_code: 0,
      runtime: { binding_status: "verified", capability_profile: "miku-project-cli-core/v1" },
      io: { result: { target: "file", path: realpathSync(resultPath) } }
    });

    const input = readFileSync(canonicalFixturePath);
    const first = runR1Harness([
      "inspect", "--project", "-", "--purpose", "project_overview"
    ], { input });
    const second = runR1Harness([
      "inspect", "--project", "-", "--purpose", "project_overview"
    ], { input });
    expect(first.status).toBe(0);
    expect(first.stderr).toBe("");
    expect(first.stdout).toBe(second.stdout);
    expect(readJsonText(first.stdout)).toMatchObject({
      command: "inspect",
      status: "succeeded",
      io: { stdin_option: "--project", result: { target: "stdout", path: null } },
      data: { projection: { purpose: "project_overview" } }
    });
  });

  it("keeps invalid grammar and an existing result path ahead of project input reads", () => {
    const directory = createTemporaryDirectory("miku-project-v1-r1-preflight-");
    const existingResultPath = path.join(directory, "already-exists.json");
    writeFileSync(existingResultPath, "keep this result file\n", "utf8");

    const resultPathConflict = runR1Harness([
      "validate", "--project", path.join(directory, "not-read.xml"), "--result", existingResultPath
    ]);
    expect(resultPathConflict.status).toBe(1);
    expect(resultPathConflict.stderr).toBe("");
    expect(readJsonText(resultPathConflict.stdout)).toMatchObject({
      command: "validate",
      status: "rejected",
      diagnostics: [{ code: "io.result-path-exists" }],
      io: {
        inputs: [{ role: "project", path: path.join(directory, "not-read.xml"), digest: null }],
        result: { target: "stdout", path: null }
      },
      data: { validation: { valid: false, format_profile: null, state_digest: null } }
    });
    expect(readFileSync(existingResultPath, "utf8")).toBe("keep this result file\n");

    const applyResultConflict = runR1Harness([
      "apply-change",
      "--project", path.join(directory, "apply-project-must-not-be-read.xml"),
      "--request", path.join(directory, "apply-request-must-not-be-read.json"),
      "--plan-result", path.join(directory, "apply-plan-must-not-be-read.json"),
      "--approval", path.join(directory, "apply-approval-must-not-be-read.json"),
      "--result", existingResultPath
    ]);
    expect(applyResultConflict.status).toBe(1);
    expect(readJsonText(applyResultConflict.stdout)).toMatchObject({
      command: "apply-change",
      status: "rejected",
      diagnostics: [{ code: "io.result-path-exists" }],
      io: {
        inputs: [
          { role: "project", path: path.join(directory, "apply-project-must-not-be-read.xml"), digest: null },
          { role: "change_request", path: path.join(directory, "apply-request-must-not-be-read.json"), digest: null },
          { role: "plan_result", path: path.join(directory, "apply-plan-must-not-be-read.json"), digest: null },
          { role: "approval", path: path.join(directory, "apply-approval-must-not-be-read.json"), digest: null }
        ],
        result: { target: "stdout", path: null },
        destination: null
      },
      effects: { project_artifact: null, cleanup: { status: "not-needed", path: null } },
      data: null
    });

    const invalidGrammar = runR1Harness([
      "validate", "--project", path.join(directory, "also-not-read.xml"), "--unknown", "value"
    ]);
    expect(invalidGrammar.status).toBe(2);
    expect(invalidGrammar.stderr).toBe("");
    expect(readJsonText(invalidGrammar.stdout)).toMatchObject({
      command: "cli",
      status: "usage-error",
      diagnostics: [{ code: "cli.unknown-option" }],
      io: { inputs: [], result: { target: "stdout", path: null } }
    });
  });

  it("runs the fixed-binding C1 task context and planning boundary without publishing a destination", () => {
    const directory = createTemporaryDirectory("miku-project-v1-c1-subprocess-");
    const context = runR1Harness([
      "inspect", "--project", canonicalFixturePath, "--purpose", "task_change_context", "--task-uid", "2"
    ]);
    expect(context.status).toBe(0);
    expect(readJsonText(context.stdout)).toMatchObject({
      command: "inspect",
      status: "succeeded",
      data: { projection: { purpose: "task_change_context", scope: { target_task_uid: "2" } } }
    });

    const requestPath = path.join(directory, "request.json");
    writeFileSync(requestPath, readFileSync(changeRequestTemplatePath, "utf8").replace(
      "${BASE_STATE_DIGEST}", "a98f0c8b560382234572a61f360c9e96911bc75fc1d57b79968d6e60b5d751d0"
    ), "utf8");
    const destination = path.join(directory, "next-project");
    const plan = runR1Harness([
      "plan-change", "--project", canonicalFixturePath, "--request", requestPath, "--destination", destination
    ]);
    expect(plan.status).toBe(0);
    expect(plan.stderr).toBe("");
    expect(readJsonText(plan.stdout)).toMatchObject({
      command: "plan-change",
      status: "succeeded",
      next_action: { action: "request-human-approval" },
      data: {
        semantic_diff: { changes: [{ task_uid: "2", before: 0, after: 50 }] },
        output_plan: { output: { destination: { path: realpathSync(directory) + "/next-project" } } }
      }
    });
    expect(existsSync(destination)).toBe(false);
  });

  it("runs the fixed-binding C1 apply service through a pre-reserved result file and publishes one committed artifact set", () => {
    const directory = createTemporaryDirectory("miku-project-v1-c1-apply-subprocess-");
    const requestPath = path.join(directory, "request.json");
    const planResultPath = path.join(directory, "plan.result.json");
    const approvalPath = path.join(directory, "approval.json");
    const applyResultPath = path.join(directory, "apply.result.json");
    const destination = path.join(directory, "next-project");
    const request = JSON.parse(readFileSync(changeRequestTemplatePath, "utf8").replace(
      "${BASE_STATE_DIGEST}", "a98f0c8b560382234572a61f360c9e96911bc75fc1d57b79968d6e60b5d751d0"
    ));
    writeFileSync(requestPath, `${canonicalJsonText(request)}\n`, "utf8");

    const plan = runR1Harness([
      "plan-change", "--project", canonicalFixturePath, "--request", requestPath,
      "--destination", destination, "--result", planResultPath
    ]);
    expect(plan.status).toBe(0);
    const planResult = readJson(planResultPath);
    const approval = {
      kind: "miku_project_change_approval",
      schema_version: "1",
      semantic_contract_version: "1",
      approved: true,
      base_state_digest: { ...planResult.data.semantic_diff.base_state_digest },
      change_request_digest: sha256CanonicalJson(request),
      semantic_diff_digest: sha256CanonicalJson(planResult.data.semantic_diff),
      output_plan_digest: sha256CanonicalJson(planResult.data.output_plan)
    };
    writeFileSync(approvalPath, `${canonicalJsonText(approval)}\n`, "utf8");

    const applied = runR1Harness([
      "apply-change", "--project", canonicalFixturePath, "--request", requestPath,
      "--plan-result", planResultPath, "--approval", approvalPath, "--result", applyResultPath
    ]);
    expect(applied.status).toBe(0);
    expect(applied.stdout).toBe("");
    expect(applied.stderr).toBe("");
    const applyResult = readJson(applyResultPath);
    const canonicalDestination = planResult.io.destination.path;
    expect(applyResult).toMatchObject({
      command: "apply-change",
      status: "succeeded",
      io: {
        result: { target: "file", path: realpathSync(applyResultPath) },
        destination: { path: canonicalDestination }
      },
      effects: {
        project_input_modified: false,
        project_artifact: {
          path: canonicalDestination,
          publication_state: "committed",
          created_by_invocation: true
        },
        cleanup: { status: "prohibited-after-commit", path: null }
      },
      data: { artifact_set: { path: canonicalDestination, publication_state: "committed" } },
      next_action: { action: "verify-artifact", command: "verify-artifact" }
    });
    expect(existsSync(path.join(canonicalDestination, "project.xml"))).toBe(true);
    expect(existsSync(path.join(canonicalDestination, "provenance.json"))).toBe(true);
    expect(existsSync(path.join(canonicalDestination, "COMMITTED"))).toBe(true);

    const rerun = runR1Harness([
      "apply-change", "--project", canonicalFixturePath, "--request", requestPath,
      "--plan-result", planResultPath, "--approval", approvalPath
    ]);
    expect(rerun.status).toBe(1);
    expect(readJsonText(rerun.stdout)).toMatchObject({
      command: "apply-change",
      status: "rejected",
      diagnostics: [{ code: "publication.reservation-conflict" }],
      effects: { project_artifact: null, cleanup: { status: "not-needed", path: null } },
      data: null
    });
  });

  it("places v1 ahead of legacy routing while refusing an incomplete runtime, and bundles all R1 sources", () => {
    const directory = createTemporaryDirectory("miku-project-v1-r1-bundle-");
    const blockedResultPath = path.join(directory, "must-not-be-reserved.json");
    const source = runNode(sourceCliPath, [
      "validate", "--project", canonicalFixturePath, "--result", blockedResultPath
    ]);
    expect(source.status).toBe(3);
    expect(source.stderr).toBe("");
    expect(readJsonText(source.stdout)).toMatchObject({
      command: "cli",
      status: "runtime-error",
      diagnostics: [{
        code: "runtime.capability-missing",
        details: { implementation_state: "R1-service-only", requested_command: "validate" }
      }],
      io: { inputs: [], result: { target: "stdout", path: null } },
      runtime: { binding_status: "unverified", capability_profile: null }
    });
    expect(existsSync(blockedResultPath)).toBe(false);

    const bundlePath = path.join(directory, "miku-project.mjs");
    const build = runNode(cliBundleBuildPath, ["--out", bundlePath]);
    expect(build.status).toBe(0);
    expect(existsSync(bundlePath)).toBe(true);
    const bundle = runNode(bundlePath, ["validate", "--project", canonicalFixturePath]);
    expect(bundle.status).toBe(3);
    expect(bundle.stderr).toBe("");
    expect(readJsonText(bundle.stdout)).toEqual(readJsonText(source.stdout));
    expect(runNode(bundlePath, ["ai", "spec"]).stdout).toContain("# miku-project AI JSON Prompt / Spec");

    const sourceEntries = listTarGzEntries(path.join(directory, "miku-project-sources.tgz"));
    expect(sourceEntries).toEqual(expect.arrayContaining([
      "miku-project-sources/scripts/generated/cli-v1-schema-validators.mjs",
      "miku-project-sources/scripts/lib/v1/cli-v1-router.mjs",
      "miku-project-sources/scripts/lib/v1/cli-v1-r1-commands.mjs",
      "miku-project-sources/scripts/lib/v1/cli-v1-change.mjs",
      "miku-project-sources/scripts/lib/v1/cli-v1-apply.mjs",
      "miku-project-sources/scripts/lib/v1/cli-v1-provenance.mjs",
      "miku-project-sources/scripts/lib/v1/cli-v1-artifact-verifier.mjs",
      "miku-project-sources/scripts/lib/v1/cli-v1-publisher.mjs",
      "miku-project-sources/scripts/lib/v1/cli-v1-xml-encoder.mjs",
      "miku-project-sources/testdata/conformance/v1/golden/projection/dependency.project-overview.json",
      "miku-project-sources/testdata/conformance/v1/golden/projection/dependency.task-change-context.json",
      "miku-project-sources/tests/mikuproject-cli-v1-plan-change.test.js",
      "miku-project-sources/tests/mikuproject-cli-v1-r1-integration.test.js",
      "miku-project-sources/tests/helpers/mikuproject-cli-v1-r1-harness.mjs"
    ]));
  });
});

function runR1Harness(args, options = {}) {
  return runNode(r1HarnessPath, args, options);
}

function runNode(entryPath, args, options = {}) {
  return spawnSync(process.execPath, [entryPath, ...args], {
    cwd: options.cwd ?? repoRoot,
    encoding: "utf8",
    input: options.input
  });
}

function createTemporaryDirectory(prefix) {
  const directory = mkdtempSync(path.join(os.tmpdir(), prefix));
  temporaryDirectories.push(directory);
  return directory;
}

function readJson(filePath) {
  return readJsonText(readFileSync(filePath, "utf8"));
}

function readJsonText(text) {
  return JSON.parse(text);
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
