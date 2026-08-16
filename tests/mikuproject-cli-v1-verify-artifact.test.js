import { readFileSync } from "node:fs";
import { mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { runV1ApplyChange } from "../scripts/lib/v1/cli-v1-apply.mjs";
import { parseV1Invocation } from "../scripts/lib/v1/cli-v1-argv.mjs";
import { canonicalJsonText, sha256CanonicalJson, sha256RawBytes } from "../scripts/lib/v1/cli-v1-canonical-json.mjs";
import { reserveV1ResultTransport } from "../scripts/lib/v1/cli-v1-io.mjs";
import { runV1PlanChange } from "../scripts/lib/v1/cli-v1-r1-commands.mjs";
import { runV1R1Harness } from "../scripts/lib/v1/cli-v1-router.mjs";
import { runV1VerifyArtifact, validateV1VerifyArtifactResultBindings } from "../scripts/lib/v1/cli-v1-verify-artifact.mjs";
import { validateCliResult } from "../scripts/generated/cli-v1-schema-validators.mjs";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const fixturePath = path.join(repoRoot, "testdata/conformance/v1/fixtures/project/dependency-canonical.xml");
const requestTemplatePath = path.join(repoRoot, "testdata/conformance/v1/fixtures/change/set-task-2-percent-0-to-50.template.json");
const suiteCases = new Map(JSON.parse(readFileSync(path.join(repoRoot, "testdata/conformance/v1/suite-index.json"), "utf8")).cases
  .map((item) => [item.id, item]));
const testRuntime = Object.freeze({
  binding_status: "verified",
  family: "node",
  version: "1.0.2",
  artifact_digest: digest("0".repeat(64)),
  manifest_digest: digest("1".repeat(64)),
  capability_profile: "miku-project-cli-core/v1",
  fixture_suite_version: "1"
});

describe("v1 verify-artifact service", () => {
  it("returns CVF-COMMITTED-001 through the fixed-binding harness and preserves every artifact member", async () => {
    const material = await createPublishedArtifact();
    const before = await artifactBytes(material.destination);
    const output = [];
    const result = await runV1R1Harness([
      "verify-artifact",
      "--artifact-set", material.destination,
      "--expect-plan-result", material.planResultPath
    ], {
      runtime: testRuntime,
      cwd: material.directory,
      stdin: Buffer.alloc(0),
      stdout: { write(value) { output.push(value); } }
    });

    expect(validateCliResult(result)).toBe(true);
    expect(result).toMatchObject({
      command: "verify-artifact",
      status: "succeeded",
      exit_code: 0,
      io: {
        stdin_option: null,
        inputs: [{
          role: "artifact_set",
          option: "--artifact-set",
          source: "filesystem-path",
          path: material.destination,
          digest: null
        }, {
          role: "expected_plan_result",
          option: "--expect-plan-result",
          source: "file",
          path: material.planResultPath
        }],
        destination: null
      },
      effects: {
        project_input_modified: false,
        project_artifact: {
          path: material.destination,
          publication_state: "committed",
          created_by_invocation: false
        },
        cleanup: { status: "not-needed", path: null }
      },
      data: {
        verification: {
          path: material.destination,
          publication_state: "committed",
          matches_expected_plan: true
        }
      }
    });
    expect(result.io.inputs[1].digest).toEqual(material.planResultInputDigest);
    expect(output).toHaveLength(1);
    expect(await artifactBytes(material.destination)).toEqual(before);

    const resultPath = path.join(material.directory, "verify.result.json");
    const fileOutput = [];
    const fileResult = await runV1R1Harness([
      "verify-artifact",
      "--artifact-set", material.destination,
      "--expect-plan-result", material.planResultPath,
      "--result", resultPath
    ], {
      runtime: testRuntime,
      cwd: material.directory,
      stdin: Buffer.alloc(0),
      stdout: { write(value) { fileOutput.push(value); } }
    });
    expect(validateCliResult(fileResult)).toBe(true);
    expect(fileResult).toMatchObject({
      status: "succeeded",
      io: { result: { target: "file" } }
    });
    expect(fileOutput).toEqual([]);
    expect(JSON.parse(await readFile(resultPath, "utf8"))).toEqual(fileResult);
  });

  it("produces byte-identical project/provenance members for the same C1 input and deterministic verify stdout", async () => {
    const material = await createApprovedApplyMaterial();
    const first = await runV1ApplyChange({
      invocation: material.applyInvocation,
      resultTransport: captureResultTransport(),
      runtime: testRuntime,
      cwd: material.directory,
      stdin: Buffer.alloc(0)
    });
    expect(first.status).toBe("succeeded");
    const firstProjectBytes = await readFile(path.join(material.destination, "project.xml"));
    const firstProvenanceBytes = await readFile(path.join(material.destination, "provenance.json"));
    // The runner, not the CLI, owns this temporary destination between two
    // independent runs. The command itself never overwrites or cleans it.
    await rm(material.destination, { recursive: true, force: false });
    const second = await runV1ApplyChange({
      invocation: material.applyInvocation,
      resultTransport: captureResultTransport(),
      runtime: testRuntime,
      cwd: material.directory,
      stdin: Buffer.alloc(0)
    });
    expect(second.status).toBe("succeeded");
    expect(await readFile(path.join(material.destination, "project.xml"))).toEqual(firstProjectBytes);
    expect(await readFile(path.join(material.destination, "provenance.json"))).toEqual(firstProvenanceBytes);

    const firstOutput = [];
    const secondOutput = [];
    const argv = ["verify-artifact", "--artifact-set", material.destination, "--expect-plan-result", material.planResultPath];
    const firstResult = await runV1R1Harness(argv, {
      runtime: testRuntime,
      cwd: material.directory,
      stdin: Buffer.alloc(0),
      stdout: { write(value) { firstOutput.push(value); } }
    });
    const secondResult = await runV1R1Harness(argv, {
      runtime: testRuntime,
      cwd: material.directory,
      stdin: Buffer.alloc(0),
      stdout: { write(value) { secondOutput.push(value); } }
    });
    expect(firstResult).toEqual(secondResult);
    expect(firstOutput).toEqual(secondOutput);
  });

  it("materializes absent, incomplete, corrupt, and expected-plan-mismatch verification results without repair", async () => {
    const directory = await mkdtemp(path.join(os.tmpdir(), "miku-project-v1-verify-state-"));
    const absentPath = path.join(directory, "absent");
    const incompletePath = path.join(directory, "incomplete");
    await mkdir(incompletePath);
    await writeFile(path.join(incompletePath, "project.xml"), "partial\n", "utf8");

    const absent = await invokeVerify(["verify-artifact", "--artifact-set", absentPath], { cwd: directory });
    const incompleteBefore = await artifactBytes(incompletePath, ["project.xml"]);
    const incomplete = await invokeVerify(["verify-artifact", "--artifact-set", incompletePath], { cwd: directory });
    expectVerificationFailure(absent.result, "publication.artifact-absent", "absent");
    expectVerificationFailure(incomplete.result, "publication.artifact-incomplete", "incomplete");
    expect(await artifactBytes(incompletePath, ["project.xml"])).toEqual(incompleteBefore);

    const corrupt = await createPublishedArtifact();
    const corruptProjectPath = path.join(corrupt.destination, "project.xml");
    const corruptBefore = await artifactBytes(corrupt.destination);
    await writeFile(corruptProjectPath, Buffer.concat([corruptBefore["project.xml"], Buffer.from(" ", "utf8")]));
    const corruptResult = await invokeVerify(["verify-artifact", "--artifact-set", corrupt.destination], { cwd: corrupt.directory });
    expectVerificationFailure(corruptResult.result, "publication.artifact-corrupt", "corrupt");
    expect(await readFile(corruptProjectPath)).toEqual(Buffer.concat([corruptBefore["project.xml"], Buffer.from(" ", "utf8")]));

    const mismatched = await createPublishedArtifact();
    const mismatchedBefore = await artifactBytes(mismatched.destination);
    const otherPlan = structuredClone(mismatched.planResult);
    const otherDestination = path.join(mismatched.directory, "different-project");
    otherPlan.io.destination.path = otherDestination;
    otherPlan.data.output_plan.output.destination.path = otherDestination;
    expect(validateCliResult(otherPlan)).toBe(true);
    const otherPlanPath = path.join(mismatched.directory, "other-plan.result.json");
    await writeFile(otherPlanPath, `${canonicalJsonText(otherPlan)}\n`, "utf8");
    const mismatchResult = await invokeVerify([
      "verify-artifact",
      "--artifact-set", mismatched.destination,
      "--expect-plan-result", otherPlanPath
    ], { cwd: mismatched.directory });
    expect(mismatchResult.result).toMatchObject({
      status: "rejected",
      diagnostics: [{ code: "publication.expected-plan-mismatch", location: { rule_id: "RB-008" } }],
      effects: { project_artifact: { path: mismatched.destination, publication_state: "committed", created_by_invocation: false } },
      data: { verification: { path: mismatched.destination, publication_state: "committed", matches_expected_plan: false } }
    });
    expect(validateCliResult(mismatchResult.result)).toBe(true);
    expect(await artifactBytes(mismatched.destination)).toEqual(mismatchedBefore);
  });

  it("rejects an invalid expected plan envelope after observing a committed artifact without evaluating RB-008", async () => {
    const material = await createPublishedArtifact();
    const expected = suiteCases.get("CVF-EXPECTED-PLAN-INVALID-001");
    const invalidPlanPath = path.join(material.directory, "invalid-plan.json");
    await writeFile(invalidPlanPath, "{}\n", "utf8");
    const { result } = await invokeVerify([
      "verify-artifact",
      "--artifact-set", material.destination,
      "--expect-plan-result", invalidPlanPath
    ], { cwd: material.directory });

    expect(validateCliResult(result)).toBe(true);
    expect(validateV1VerifyArtifactResultBindings({ result })).toBe(true);
    expect(result).toMatchObject({
      status: "rejected",
      exit_code: expected.expected_exit_code,
      next_action: expected.expected_next_action,
      diagnostics: [{ code: "change.binding-mismatch", location: { option: "--expect-plan-result", rule_id: null } }],
      effects: { project_artifact: { path: material.destination, publication_state: "committed", created_by_invocation: false } },
      data: { verification: { path: material.destination, publication_state: "committed", matches_expected_plan: null } }
    });
    expect(result.diagnostics.map((item) => item.code)).toEqual(expected.expected_diagnostic_codes);
    expect(result.diagnostics.map((item) => item.location.rule_id)).toEqual(expected.expected_rule_ids);
  });

  it("keeps RB-008 unobserved for schema-valid results with the wrong command or a non-success plan status", async () => {
    const material = await createPublishedArtifact();
    const before = await artifactBytes(material.destination);
    const observed = await invokeVerify([
      "verify-artifact",
      "--artifact-set", material.destination
    ], { cwd: material.directory });
    expect(validateCliResult(observed.result)).toBe(true);

    const wrongCommandPath = path.join(material.directory, "wrong-command.result.json");
    await writeFile(wrongCommandPath, `${canonicalJsonText(observed.result)}\n`, "utf8");
    const wrongCommand = await invokeVerify([
      "verify-artifact",
      "--artifact-set", material.destination,
      "--expect-plan-result", wrongCommandPath
    ], { cwd: material.directory });
    expect(wrongCommand.result).toMatchObject({
      status: "rejected",
      diagnostics: [{ code: "change.binding-mismatch", location: { option: "--expect-plan-result", rule_id: null } }],
      data: { verification: { publication_state: "committed", matches_expected_plan: null } }
    });
    expect(validateCliResult(wrongCommand.result)).toBe(true);
    expect(validateV1VerifyArtifactResultBindings({ result: wrongCommand.result })).toBe(true);

    const staleRequest = structuredClone(material.request);
    staleRequest.base.state_digest.value = "0".repeat(64);
    const staleRequestPath = path.join(material.directory, "stale-request.json");
    const stalePlanPath = path.join(material.directory, "stale-plan.result.json");
    await writeFile(staleRequestPath, `${canonicalJsonText(staleRequest)}\n`, "utf8");
    const stalePlanInvocation = parseV1Invocation([
      "plan-change",
      "--project", fixturePath,
      "--request", staleRequestPath,
      "--destination", path.join(material.directory, "stale-output"),
      "--result", stalePlanPath
    ]);
    const stalePlan = await runV1PlanChange({
      invocation: stalePlanInvocation,
      resultTransport: await reserveV1ResultTransport(stalePlanInvocation.options.result, { cwd: material.directory }),
      runtime: testRuntime,
      cwd: material.directory,
      stdin: Buffer.alloc(0)
    });
    expect(stalePlan).toMatchObject({ command: "plan-change", status: "rejected" });
    expect(validateCliResult(stalePlan)).toBe(true);

    const nonSuccess = await invokeVerify([
      "verify-artifact",
      "--artifact-set", material.destination,
      "--expect-plan-result", stalePlanPath
    ], { cwd: material.directory });
    expect(nonSuccess.result).toMatchObject({
      status: "rejected",
      diagnostics: [{ code: "change.binding-mismatch", location: { option: "--expect-plan-result", rule_id: null } }],
      data: { verification: { publication_state: "committed", matches_expected_plan: null } }
    });
    expect(validateCliResult(nonSuccess.result)).toBe(true);
    expect(validateV1VerifyArtifactResultBindings({ result: nonSuccess.result })).toBe(true);
    expect(await artifactBytes(material.destination)).toEqual(before);
  });

  it("reports malformed or unreadable expected-plan input separately while retaining its committed observation", async () => {
    const material = await createPublishedArtifact();
    const malformedPlanPath = path.join(material.directory, "malformed-plan.json");
    await writeFile(malformedPlanPath, "{\n", "utf8");
    const { result } = await invokeVerify([
      "verify-artifact",
      "--artifact-set", material.destination,
      "--expect-plan-result", malformedPlanPath
    ], { cwd: material.directory });

    expect(validateCliResult(result)).toBe(true);
    expect(validateV1VerifyArtifactResultBindings({ result })).toBe(true);
    expect(result).toMatchObject({
      status: "rejected",
      diagnostics: [{ code: "json.invalid", location: { option: "--expect-plan-result" } }],
      effects: { project_artifact: { path: material.destination, publication_state: "committed", created_by_invocation: false } },
      data: { verification: { path: material.destination, publication_state: "committed", matches_expected_plan: null } }
    });

    const pathDiverged = structuredClone(result);
    pathDiverged.data.verification.path = path.join(material.directory, "different-artifact-set");
    expect(validateCliResult(pathDiverged)).toBe(true);
    expect(validateV1VerifyArtifactResultBindings({ result: pathDiverged })).toBe(false);

    const stateDiverged = structuredClone(result);
    stateDiverged.effects.project_artifact.publication_state = "corrupt";
    expect(validateCliResult(stateDiverged)).toBe(true);
    expect(validateV1VerifyArtifactResultBindings({ result: stateDiverged })).toBe(false);

    const missingPlan = await invokeVerify([
      "verify-artifact",
      "--artifact-set", material.destination,
      "--expect-plan-result", path.join(material.directory, "missing-plan.result.json")
    ], { cwd: material.directory });
    expect(missingPlan.result).toMatchObject({
      status: "rejected",
      diagnostics: [{ code: "io.input-not-found", location: { option: "--expect-plan-result" } }],
      effects: { project_artifact: { path: material.destination, publication_state: "committed", created_by_invocation: false } },
      data: { verification: { path: material.destination, publication_state: "committed", matches_expected_plan: null } }
    });
    expect(validateCliResult(missingPlan.result)).toBe(true);
    expect(validateV1VerifyArtifactResultBindings({ result: missingPlan.result })).toBe(true);
  });

  it("keeps a pre-existing result file ahead of every verify-artifact input", async () => {
    const directory = await mkdtemp(path.join(os.tmpdir(), "miku-project-v1-verify-result-preflight-"));
    const existingResultPath = path.join(directory, "already-exists.json");
    await writeFile(existingResultPath, "preserve\n", "utf8");
    const output = [];
    const result = await runV1R1Harness([
      "verify-artifact",
      "--artifact-set", path.join(directory, "must-not-be-read"),
      "--expect-plan-result", path.join(directory, "also-must-not-be-read.json"),
      "--result", existingResultPath
    ], {
      runtime: testRuntime,
      cwd: directory,
      stdin: Buffer.alloc(0),
      stdout: { write(value) { output.push(value); } }
    });

    expect(validateCliResult(result)).toBe(true);
    expect(result).toMatchObject({
      command: "verify-artifact",
      status: "rejected",
      diagnostics: [{ code: "io.result-path-exists" }],
      io: {
        stdin_option: null,
        inputs: [{
          role: "artifact_set",
          option: "--artifact-set",
          source: "filesystem-path",
          path: path.join(directory, "must-not-be-read"),
          digest: null
        }, {
          role: "expected_plan_result",
          option: "--expect-plan-result",
          source: "file",
          path: path.join(directory, "also-must-not-be-read.json"),
          digest: null
        }],
        result: { target: "stdout", path: null },
        destination: null
      },
      data: null
    });
    expect(output).toHaveLength(1);
    expect(await readFile(existingResultPath, "utf8")).toBe("preserve\n");
  });

  it("recovers CU-UNKNOWN-OUTCOME-001 by verifying the destination instead of reapplying", async () => {
    const material = await createApprovedApplyMaterial();
    const lostResultTransport = Object.freeze({
      target: Object.freeze({ target: "stdout", path: null }),
      async writeResult() {
        throw new Error("injected result delivery loss after commit");
      }
    });
    await expect(runV1ApplyChange({
      invocation: material.applyInvocation,
      resultTransport: lostResultTransport,
      runtime: testRuntime,
      cwd: material.directory,
      stdin: Buffer.alloc(0)
    })).rejects.toThrow("injected result delivery loss after commit");

    const { result } = await invokeVerify([
      "verify-artifact",
      "--artifact-set", material.destination,
      "--expect-plan-result", material.planResultPath
    ], { cwd: material.directory });
    expect(validateCliResult(result)).toBe(true);
    expect(result).toMatchObject({
      status: "succeeded",
      data: { verification: { path: material.destination, publication_state: "committed", matches_expected_plan: true } },
      effects: { project_artifact: { path: material.destination, publication_state: "committed", created_by_invocation: false } }
    });
  });
});

async function createPublishedArtifact() {
  const material = await createApprovedApplyMaterial();
  const applyResult = await runV1ApplyChange({
    invocation: material.applyInvocation,
    resultTransport: captureResultTransport(),
    runtime: testRuntime,
    cwd: material.directory,
    stdin: Buffer.alloc(0)
  });
  expect(applyResult.status).toBe("succeeded");
  return {
    ...material,
    destination: applyResult.data.artifact_set.path
  };
}

async function createApprovedApplyMaterial() {
  const directory = await mkdtemp(path.join(os.tmpdir(), "miku-project-v1-verify-artifact-"));
  const destinationParent = path.join(directory, "output-parent");
  const destination = path.join(destinationParent, "next-project");
  const requestPath = path.join(directory, "request.json");
  const planResultPath = path.join(directory, "plan.result.json");
  const approvalPath = path.join(directory, "approval.json");
  await mkdir(destinationParent);
  const request = await requestArtifact();
  await writeFile(requestPath, `${canonicalJsonText(request)}\n`, "utf8");
  const planInvocation = parseV1Invocation([
    "plan-change",
    "--project", fixturePath,
    "--request", requestPath,
    "--destination", destination,
    "--result", planResultPath
  ]);
  const planResult = await runV1PlanChange({
    invocation: planInvocation,
    resultTransport: await reserveV1ResultTransport(planInvocation.options.result, { cwd: directory }),
    runtime: testRuntime,
    cwd: directory,
    stdin: Buffer.alloc(0)
  });
  expect(planResult.status).toBe("succeeded");
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
  await writeFile(approvalPath, `${canonicalJsonText(approval)}\n`, "utf8");
  return {
    directory,
    destination: planResult.data.output_plan.output.destination.path,
    planResult,
    planResultPath: planResult.io.result.path,
    planResultInputDigest: sha256RawBytes(await readFile(planResult.io.result.path)),
    request,
    applyInvocation: parseV1Invocation([
      "apply-change",
      "--project", fixturePath,
      "--request", requestPath,
      "--plan-result", planResultPath,
      "--approval", approvalPath
    ])
  };
}

async function invokeVerify(argv, { cwd }) {
  const invocation = parseV1Invocation(argv);
  const transport = captureResultTransport();
  const result = await runV1VerifyArtifact({
    invocation,
    resultTransport: transport,
    runtime: testRuntime,
    cwd,
    stdin: Buffer.alloc(0)
  });
  expect(result).toEqual(transport.result);
  return { result };
}

async function requestArtifact() {
  const template = await readFile(requestTemplatePath, "utf8");
  return JSON.parse(template.replace("${BASE_STATE_DIGEST}", "a98f0c8b560382234572a61f360c9e96911bc75fc1d57b79968d6e60b5d751d0"));
}

function expectVerificationFailure(result, code, state) {
  expect(validateCliResult(result)).toBe(true);
  expect(result).toMatchObject({
    status: "rejected",
    diagnostics: [{ code }],
    effects: { project_artifact: { publication_state: state, created_by_invocation: false } },
    data: { verification: { publication_state: state, matches_expected_plan: null, bindings: null } }
  });
}

async function artifactBytes(directory, names = ["COMMITTED", "project.xml", "provenance.json"]) {
  const values = await Promise.all(names.map(async (name) => [name, await readFile(path.join(directory, name))]));
  return Object.fromEntries(values);
}

function captureResultTransport() {
  let result = null;
  return Object.freeze({
    target: Object.freeze({ target: "stdout", path: null }),
    async writeResult(value) {
      if (result !== null) throw new Error("test result transport was written more than once");
      result = value;
    },
    get result() {
      return result;
    }
  });
}

function digest(value) {
  return { algorithm: "sha-256", value };
}
