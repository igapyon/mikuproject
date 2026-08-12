import fs from "node:fs";
import { access, mkdtemp, readFile, realpath, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { parseV1Invocation } from "../scripts/lib/v1/cli-v1-argv.mjs";
import { planV1SetTaskPercentComplete, validateV1PlanChangeBindings } from "../scripts/lib/v1/cli-v1-change.mjs";
import { sha256RawBytes } from "../scripts/lib/v1/cli-v1-canonical-json.mjs";
import { reserveV1ResultTransport } from "../scripts/lib/v1/cli-v1-io.mjs";
import { runV1PlanChange } from "../scripts/lib/v1/cli-v1-r1-commands.mjs";
import { serializeV1Result } from "../scripts/lib/v1/cli-v1-result.mjs";
import { decodeMsProjectXmlSubset } from "../scripts/lib/v1/cli-v1-xml-adapter.mjs";
import { validateV1SemanticState } from "../scripts/lib/v1/cli-v1-semantic-validator.mjs";
import { validateArtifact, validateCliResult } from "../scripts/generated/cli-v1-schema-validators.mjs";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const fixturePath = path.join(repoRoot, "testdata/conformance/v1/fixtures/project/dependency-canonical.xml");
const requestTemplatePath = path.join(repoRoot, "testdata/conformance/v1/fixtures/change/set-task-2-percent-0-to-50.template.json");
const semanticGolden = readJson("testdata/conformance/v1/golden/semantic/dependency-percent-50.state.json");
const semanticDiffGolden = readJson("docs/examples/cli-v1/plan-change-succeeded.result.json").data.semantic_diff;
const testRuntime = Object.freeze({
  binding_status: "verified",
  family: "node",
  version: "1.0.2",
  artifact_digest: digest("0".repeat(64)),
  manifest_digest: digest("1".repeat(64)),
  capability_profile: "miku-project-cli-core/v1",
  fixture_suite_version: "1"
});

describe("v1 C1 task-change context and plan-change service", () => {
  it("runs CP-CHANGE-001 without creating the destination and binds diff/plan/runtime/request exactly", async () => {
    const directory = await mkdtemp(path.join(os.tmpdir(), "miku-project-v1-plan-change-"));
    const requestPath = path.join(directory, "change-request.json");
    const destination = path.join(directory, "next-project");
    await writeRequest(requestPath);

    const { result, output } = await invokePlanChange([
      "plan-change", "--project", fixturePath, "--request", requestPath, "--destination", destination, "--result", "plan.result.json"
    ], { cwd: directory });
    const canonicalDirectory = await realpath(directory);
    const canonicalRequestPath = await realpath(requestPath);
    const canonicalFixturePath = await realpath(fixturePath);

    expect(validateCliResult(result)).toBe(true);
    expect(validateArtifact(result.data.semantic_diff)).toBe(true);
    expect(validateArtifact(result.data.output_plan)).toBe(true);
    expect(result).toMatchObject({
      command: "plan-change",
      status: "succeeded",
      exit_code: 0,
      next_action: { action: "request-human-approval", command: null, source_retryability: null },
      io: {
        stdin_option: null,
        inputs: [
          { role: "project", option: "--project", source: "file", path: canonicalFixturePath, digest: sha256RawBytes(await readFile(fixturePath)) },
          { role: "change_request", option: "--request", source: "file", path: canonicalRequestPath, digest: sha256RawBytes(await readFile(requestPath)) }
        ],
        result: { target: "file", path: path.join(canonicalDirectory, "plan.result.json") },
        destination: { requested_path: destination, path: path.join(canonicalDirectory, "next-project") }
      },
      effects: { project_input_modified: false, project_artifact: null, cleanup: { status: "not-needed", path: null } },
      diagnostics: []
    });
    expect(result.data.semantic_diff).toEqual(semanticDiffGolden);
    expect(validateV1PlanChangeBindings({
      changeRequest: readJsonFromPath(requestPath),
      semanticDiff: result.data.semantic_diff,
      outputPlan: result.data.output_plan,
      runtime: testRuntime,
      destination: result.io.destination
    })).toBe(true);
    await expect(access(destination)).rejects.toMatchObject({ code: "ENOENT" });
    expect(output).toEqual([]);
    expect(await readFile(result.io.result.path, "utf8")).toBe(serializeV1Result(result));

    const stdinRequestBytes = Buffer.from(await requestText(), "utf8");
    const stdinDestination = path.join(directory, "stdin-next-project");
    const stdinPlan = await invokePlanChange([
      "plan-change", "--project", fixturePath, "--request", "-", "--destination", stdinDestination
    ], { cwd: directory, stdin: stdinRequestBytes });
    expect(validateCliResult(stdinPlan.result)).toBe(true);
    expect(stdinPlan.result).toMatchObject({
      status: "succeeded",
      io: {
        stdin_option: "--request",
        inputs: [
          { role: "project", source: "file" },
          { role: "change_request", source: "stdin", path: null, digest: sha256RawBytes(stdinRequestBytes) }
        ]
      }
    });
    await expect(access(stdinDestination)).rejects.toMatchObject({ code: "ENOENT" });
  });

  it("keeps planned state internal but validates the semantic diff and encode/redecode result against its golden", () => {
    const decoded = decodeMsProjectXmlSubset(fs.readFileSync(fixturePath));
    const plan = planV1SetTaskPercentComplete({
      state: decoded.state,
      changeRequest: readJson("docs/examples/artifacts-v1/change-request.example.json"),
      runtime: testRuntime,
      destination: { path: "/tmp/miku-project-v1-plan-change-golden" }
    });
    const repeated = planV1SetTaskPercentComplete({
      state: decoded.state,
      changeRequest: readJson("docs/examples/artifacts-v1/change-request.example.json"),
      runtime: testRuntime,
      destination: { path: "/tmp/miku-project-v1-plan-change-golden" }
    });
    const redecoded = decodeMsProjectXmlSubset(plan.preflight_project_xml);

    expect(validateV1SemanticState(redecoded.state, { adapterIssues: redecoded.adapter_issues }).valid).toBe(true);
    expect(plan.planned_state).toEqual(semanticGolden);
    expect(redecoded.state).toEqual(semanticGolden);
    expect(plan.semantic_diff).toEqual(semanticDiffGolden);
    expect(repeated).toEqual(plan);
  });

  it("rejects stale preconditions, duplicate request keys, and an existing destination without creating output", async () => {
    const directory = await mkdtemp(path.join(os.tmpdir(), "miku-project-v1-plan-change-reject-"));
    const staleRequest = path.join(directory, "stale.json");
    const validRequest = path.join(directory, "valid.json");
    const staleBaseRequest = path.join(directory, "stale-base.json");
    const noOpRequest = path.join(directory, "no-op.json");
    const unsupportedRequest = path.join(directory, "unsupported.json");
    const bomRequest = path.join(directory, "bom.json");
    const nonDirectoryParent = path.join(directory, "not-a-directory");
    const duplicateRequest = path.join(directory, "duplicate.json");
    const existingDestination = path.join(directory, "already-exists");
    await writeRequest(staleRequest, { expectedPercentComplete: 1 });
    await writeRequest(validRequest);
    await writeRequest(staleBaseRequest, { baseStateDigest: "e".repeat(64) });
    await writeRequest(noOpRequest, { valuePercentComplete: 0 });
    await writeRequest(unsupportedRequest, { operationKind: "set_task_name" });
    await writeFile(bomRequest, Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from(await requestText(), "utf8")]));
    await writeFile(nonDirectoryParent, "not a directory\n", "utf8");
    await writeFile(duplicateRequest, '{"kind":"miku_project_change_request","kind":"miku_project_change_request"}\n', "utf8");
    await writeFile(existingDestination, "preserve me\n", "utf8");

    const stale = await invokePlanChange([
      "plan-change", "--project", fixturePath, "--request", staleRequest, "--destination", path.join(directory, "stale-output")
    ], { cwd: directory });
    expect(validateCliResult(stale.result)).toBe(true);
    expect(stale.result).toMatchObject({
      status: "rejected",
      exit_code: 1,
      next_action: { action: "replan-and-request-human-approval", command: "plan-change", source_retryability: "after-replan-and-approval" },
      diagnostics: [{ code: "change.precondition-failed" }],
      data: null
    });

    const staleBase = await invokePlanChange([
      "plan-change", "--project", fixturePath, "--request", staleBaseRequest, "--destination", path.join(directory, "stale-base-output")
    ], { cwd: directory });
    expect(validateCliResult(staleBase.result)).toBe(true);
    expect(staleBase.result.diagnostics).toMatchObject([{ code: "change.precondition-failed" }]);

    const noOp = await invokePlanChange([
      "plan-change", "--project", fixturePath, "--request", noOpRequest, "--destination", path.join(directory, "no-op-output")
    ], { cwd: directory });
    expect(validateCliResult(noOp.result)).toBe(true);
    expect(noOp.result).toMatchObject({
      status: "rejected",
      next_action: { action: "revise-invocation-or-input", command: null, source_retryability: "after-input-change" },
      diagnostics: [{ code: "change.no-op" }]
    });

    const unsupported = await invokePlanChange([
      "plan-change", "--project", fixturePath, "--request", unsupportedRequest, "--destination", path.join(directory, "unsupported-output")
    ], { cwd: directory });
    expect(validateCliResult(unsupported.result)).toBe(true);
    expect(unsupported.result.diagnostics).toMatchObject([{ code: "change.operation-unsupported" }]);

    const duplicate = await invokePlanChange([
      "plan-change", "--project", fixturePath, "--request", duplicateRequest, "--destination", path.join(directory, "duplicate-output")
    ], { cwd: directory });
    expect(validateCliResult(duplicate.result)).toBe(true);
    expect(duplicate.result).toMatchObject({
      status: "rejected",
      diagnostics: [{ code: "json.duplicate-key" }],
      data: null
    });

    const bom = await invokePlanChange([
      "plan-change", "--project", fixturePath, "--request", bomRequest, "--destination", path.join(directory, "bom-output")
    ], { cwd: directory });
    expect(validateCliResult(bom.result)).toBe(true);
    expect(bom.result.diagnostics).toMatchObject([{ code: "json.bom-not-allowed" }]);

    const existing = await invokePlanChange([
      "plan-change", "--project", fixturePath, "--request", validRequest, "--destination", existingDestination
    ], { cwd: directory });
    expect(validateCliResult(existing.result)).toBe(true);
    expect(existing.result).toMatchObject({
      status: "rejected",
      diagnostics: [{ code: "publication.destination-exists" }],
      data: null
    });
    expect(await readFile(existingDestination, "utf8")).toBe("preserve me\n");

    const unsafe = await invokePlanChange([
      "plan-change", "--project", fixturePath, "--request", validRequest, "--destination", path.join(nonDirectoryParent, "child")
    ], { cwd: directory });
    expect(validateCliResult(unsafe.result)).toBe(true);
    expect(unsafe.result.diagnostics).toMatchObject([{ code: "publication.destination-unsafe" }]);
  });
});

async function invokePlanChange(argv, { cwd = repoRoot, stdin = Buffer.alloc(0) } = {}) {
  const output = [];
  const invocation = parseV1Invocation(argv);
  const resultTransport = await reserveV1ResultTransport(invocation.options.result, {
    cwd,
    stdout: { write(value) { output.push(value); } }
  });
  const result = await runV1PlanChange({ invocation, resultTransport, runtime: testRuntime, cwd, stdin });
  return { result, output };
}

async function writeRequest(filePath, {
  expectedPercentComplete = 0,
  valuePercentComplete = 50,
  baseStateDigest = "a98f0c8b560382234572a61f360c9e96911bc75fc1d57b79968d6e60b5d751d0",
  operationKind = "set_task_percent_complete"
} = {}) {
  await writeFile(filePath, await requestText({ expectedPercentComplete, valuePercentComplete, baseStateDigest, operationKind }), "utf8");
}

async function requestText({
  expectedPercentComplete = 0,
  valuePercentComplete = 50,
  baseStateDigest = "a98f0c8b560382234572a61f360c9e96911bc75fc1d57b79968d6e60b5d751d0",
  operationKind = "set_task_percent_complete"
} = {}) {
  const template = await readFile(requestTemplatePath, "utf8");
  return template
    .replace("${BASE_STATE_DIGEST}", baseStateDigest)
    .replace('"kind": "set_task_percent_complete"', `"kind": "${operationKind}"`)
    .replace('"expected_percent_complete": 0', `"expected_percent_complete": ${expectedPercentComplete}`)
    .replace('"percent_complete": 50', `"percent_complete": ${valuePercentComplete}`);
}

function readJson(relativePath) {
  return readJsonFromPath(path.join(repoRoot, relativePath));
}

function readJsonFromPath(filePath) {
  return JSON.parse(fs.readFileSync(filePath, "utf8"));
}

function digest(value) {
  return { algorithm: "sha-256", value };
}
