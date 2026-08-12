import fs from "node:fs";
import { mkdtemp, readFile, realpath } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { parseV1Invocation } from "../scripts/lib/v1/cli-v1-argv.mjs";
import { sha256RawBytes } from "../scripts/lib/v1/cli-v1-canonical-json.mjs";
import { reserveV1ResultTransport } from "../scripts/lib/v1/cli-v1-io.mjs";
import {
  validateV1ProjectOverviewBinding,
  validateV1TaskChangeContextBinding
} from "../scripts/lib/v1/cli-v1-projection.mjs";
import { runV1Inspect } from "../scripts/lib/v1/cli-v1-r1-commands.mjs";
import { serializeV1Result } from "../scripts/lib/v1/cli-v1-result.mjs";
import { validateArtifact, validateCliResult } from "../scripts/generated/cli-v1-schema-validators.mjs";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const fixtureRoot = path.join(repoRoot, "testdata/conformance/v1/fixtures/project");
const projectionGolden = readJson("testdata/conformance/v1/golden/projection/dependency.project-overview.json");
const taskChangeContextGolden = readJson("testdata/conformance/v1/golden/projection/dependency.task-change-context.json");
const semanticGolden = readJson("testdata/conformance/v1/golden/semantic/dependency.state.json");
const suiteCases = new Map(readJson("testdata/conformance/v1/suite-index.json").cases.map((item) => [item.id, item]));
const testRuntime = Object.freeze({
  binding_status: "verified",
  family: "node",
  version: "1.0.2",
  artifact_digest: digest("0".repeat(64)),
  manifest_digest: digest("1".repeat(64)),
  capability_profile: "miku-project-cli-core/v1",
  fixture_suite_version: "1"
});

describe("v1 R1 inspect project_overview service", () => {
  it("runs CI-OVERVIEW-001 through the validate pipeline and fixes the exact Projection golden", async () => {
    const temporaryDirectory = await mkdtemp(path.join(os.tmpdir(), "miku-project-v1-inspect-overview-"));
    const canonicalTemporaryDirectory = await realpath(temporaryDirectory);
    const fixturePath = path.join(fixtureRoot, "dependency-canonical.xml");
    const inputBytes = await readFile(fixturePath);
    const { result, output } = await invokeInspect([
      "inspect", "--project", fixturePath, "--purpose", "project_overview", "--result", "overview.result.json"
    ], { cwd: temporaryDirectory });
    const expected = suiteCases.get("CI-OVERVIEW-001");

    expect(validateCliResult(result)).toBe(true);
    expect(validateArtifact(result.data.projection)).toBe(true);
    expect(result).toMatchObject({
      command: "inspect",
      status: expected.expected_status,
      exit_code: expected.expected_exit_code,
      next_action: expected.expected_next_action,
      io: {
        stdin_option: null,
        inputs: [{
          role: "project",
          option: "--project",
          source: "file",
          path: await realpath(fixturePath),
          digest: sha256RawBytes(inputBytes)
        }],
        result: { target: "file", path: path.join(canonicalTemporaryDirectory, "overview.result.json") },
        destination: null
      },
      effects: { project_input_modified: false, project_artifact: null, cleanup: { status: "not-needed", path: null } },
      observations: { normalizations: [], losses: [], unsupported: [] },
      diagnostics: []
    });
    expect(result.data).toEqual({ projection: projectionGolden });
    expect(result.data.projection.source_state_digest).toEqual(digest("a98f0c8b560382234572a61f360c9e96911bc75fc1d57b79968d6e60b5d751d0"));
    expect(validateV1ProjectOverviewBinding({ state: semanticGolden, projection: result.data.projection })).toBe(true);
    expect(output).toEqual([]);
    expect(await readFile(result.io.result.path, "utf8")).toBe(serializeV1Result(result));
  });

  it("does not project invalid or unsupported semantic input", async () => {
    const invalidInput = readFixture("dependency-percent-101.xml");
    const invalid = await invokeInspect([
      "inspect", "--project", "-", "--purpose", "project_overview"
    ], { stdin: invalidInput });
    const invalidExpected = suiteCases.get("CV-INVALID-001");

    expect(validateCliResult(invalid.result)).toBe(true);
    expect(invalid.result).toMatchObject({
      status: invalidExpected.expected_status,
      exit_code: invalidExpected.expected_exit_code,
      next_action: invalidExpected.expected_next_action,
      data: null
    });
    expect(invalid.result.diagnostics.map((item) => item.code)).toEqual(invalidExpected.expected_diagnostic_codes);
    expect(invalid.result.diagnostics.map((item) => item.location.rule_id)).toEqual(invalidExpected.expected_rule_ids);
    expect(invalid.output).toEqual([serializeV1Result(invalid.result)]);

    const unsupported = await invokeInspect([
      "inspect", "--project", path.join(fixtureRoot, "dependency-unsupported-actual.xml"), "--purpose", "project_overview"
    ]);
    const unsupportedExpected = suiteCases.get("CV-UNSUPPORTED-001");
    expect(validateCliResult(unsupported.result)).toBe(true);
    expect(unsupported.result).toMatchObject({
      status: unsupportedExpected.expected_status,
      exit_code: unsupportedExpected.expected_exit_code,
      next_action: unsupportedExpected.expected_next_action,
      data: null,
      observations: {
        losses: [],
        unsupported: [{ code: "semantic.unsupported", path: "tasks[uid=2].actual_start" }]
      }
    });
    expect(unsupported.result.diagnostics.map((item) => item.code)).toEqual(unsupportedExpected.expected_diagnostic_codes);
    expect(unsupported.result.diagnostics.map((item) => item.location.rule_id)).toEqual(unsupportedExpected.expected_rule_ids);
  });

  it("runs CI-CONTEXT-001 with only the target task decision context", async () => {
    const fixturePath = path.join(fixtureRoot, "dependency-canonical.xml");
    const { result } = await invokeInspect([
      "inspect", "--project", fixturePath, "--purpose", "task_change_context", "--task-uid", "2"
    ]);
    const expected = suiteCases.get("CI-CONTEXT-001");

    expect(validateCliResult(result)).toBe(true);
    expect(validateArtifact(result.data.projection)).toBe(true);
    expect(result).toMatchObject({
      command: "inspect",
      status: expected.expected_status,
      exit_code: expected.expected_exit_code,
      next_action: expected.expected_next_action,
      diagnostics: []
    });
    expect(result.data).toEqual({ projection: taskChangeContextGolden });
    expect(validateV1TaskChangeContextBinding({ state: semanticGolden, projection: result.data.projection })).toBe(true);

    const changedContent = structuredClone(taskChangeContextGolden);
    changedContent.resources.push({ uid: "not-a-context-resource" });
    expect(validateV1TaskChangeContextBinding({ state: semanticGolden, projection: changedContent })).toBe(false);
  });

  it("fails closed when task_change_context is requested for no task or a non-leaf task", async () => {
    const fixturePath = path.join(fixtureRoot, "dependency-canonical.xml");
    const missing = await invokeInspect([
      "inspect", "--project", fixturePath, "--purpose", "task_change_context", "--task-uid", "999"
    ]);
    expect(validateCliResult(missing.result)).toBe(true);
    expect(missing.result).toMatchObject({
      status: "rejected",
      diagnostics: [{ code: "change.request-invalid", location: { option: "--task-uid" } }],
      data: null
    });
  });

  it("is byte-deterministic for the same runtime, input, and invocation", async () => {
    const fixturePath = path.join(fixtureRoot, "dependency-canonical.xml");
    const first = await invokeInspect(["inspect", "--project", fixturePath, "--purpose", "project_overview"]);
    const second = await invokeInspect(["inspect", "--project", fixturePath, "--purpose", "project_overview"]);

    expect(first.result).toEqual(second.result);
    expect(first.output).toEqual(second.output);
  });

  it("rejects a source digest, scope, or content divergence from the semantic state", () => {
    const changedDigest = structuredClone(projectionGolden);
    changedDigest.source_state_digest.value = "e".repeat(64);
    const changedScope = structuredClone(projectionGolden);
    changedScope.scope.omitted_domains[2] = "different_domain";
    const changedContent = structuredClone(projectionGolden);
    changedContent.tasks[1].percent_complete = 50;

    expect(validateV1ProjectOverviewBinding({ state: semanticGolden, projection: projectionGolden })).toBe(true);
    expect(validateV1ProjectOverviewBinding({ state: semanticGolden, projection: changedDigest })).toBe(false);
    expect(validateV1ProjectOverviewBinding({ state: semanticGolden, projection: changedScope })).toBe(false);
    expect(validateV1ProjectOverviewBinding({ state: semanticGolden, projection: changedContent })).toBe(false);
  });
});

async function invokeInspect(argv, { cwd = repoRoot, stdin = Buffer.alloc(0) } = {}) {
  const output = [];
  const invocation = parseV1Invocation(argv);
  const resultTransport = await reserveV1ResultTransport(invocation.options.result, {
    cwd,
    stdout: { write(value) { output.push(value); } }
  });
  const result = await runV1Inspect({ invocation, resultTransport, runtime: testRuntime, cwd, stdin });
  return { result, output };
}

function readFixture(name) {
  return fs.readFileSync(path.join(fixtureRoot, name));
}

function readJson(relativePath) {
  return JSON.parse(fs.readFileSync(path.join(repoRoot, relativePath), "utf8"));
}

function digest(value) {
  return { algorithm: "sha-256", value };
}
