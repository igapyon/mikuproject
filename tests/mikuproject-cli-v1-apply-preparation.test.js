import fs from "node:fs";
import { access, mkdir, mkdtemp, readFile, realpath, rename, symlink, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { prepareV1ApplyChange, runV1ApplyChange } from "../scripts/lib/v1/cli-v1-apply.mjs";
import { parseV1Invocation } from "../scripts/lib/v1/cli-v1-argv.mjs";
import { verifyV1ArtifactSet } from "../scripts/lib/v1/cli-v1-artifact-verifier.mjs";
import {
  canonicalJsonText,
  sha256CanonicalJson,
  sha256RawBytes
} from "../scripts/lib/v1/cli-v1-canonical-json.mjs";
import { validateV1ApplyPreparationBindings } from "../scripts/lib/v1/cli-v1-change.mjs";
import { reserveV1ResultTransport } from "../scripts/lib/v1/cli-v1-io.mjs";
import { runV1PlanChange } from "../scripts/lib/v1/cli-v1-r1-commands.mjs";
import { createV1DiagnosticFromError } from "../scripts/lib/v1/cli-v1-result.mjs";
import { validateArtifact, validateCliDiagnostic, validateCliResult } from "../scripts/generated/cli-v1-schema-validators.mjs";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const fixturePath = path.join(repoRoot, "testdata/conformance/v1/fixtures/project/dependency-canonical.xml");
const requestTemplatePath = path.join(repoRoot, "testdata/conformance/v1/fixtures/change/set-task-2-percent-0-to-50.template.json");
const changedStateGolden = readJson("testdata/conformance/v1/golden/semantic/dependency-percent-50.state.json");
const testRuntime = Object.freeze({
  binding_status: "verified",
  family: "node",
  version: "1.0.2",
  artifact_digest: digest("0".repeat(64)),
  manifest_digest: digest("1".repeat(64)),
  capability_profile: "miku-project-cli-core/v1",
  fixture_suite_version: "1"
});

describe("v1 apply-change preparation and approval binding", () => {
  it("recomputes an approved plan, records four inputs in RB-010 order, and creates no destination", async () => {
    const material = await createApprovedPlan();
    const invocation = applyInvocation(material);
    const prepared = await prepareV1ApplyChange({ invocation, runtime: testRuntime, cwd: repoRoot });

    expect(prepared.error).toBeUndefined();
    expect(prepared.inputs).toEqual([
      await expectedFileInput("project", "--project", fixturePath),
      await expectedFileInput("change_request", "--request", material.requestPath),
      await expectedFileInput("plan_result", "--plan-result", material.planResultPath),
      await expectedFileInput("approval", "--approval", material.approvalPath)
    ]);
    expect(prepared.destination).toEqual(material.planResult.io.destination);
    expect(prepared.prepared.semantic_diff).toEqual(material.planResult.data.semantic_diff);
    expect(prepared.prepared.output_plan).toEqual(material.planResult.data.output_plan);
    expect(prepared.prepared.planned_state).toEqual(changedStateGolden);
    expect(validateArtifact(prepared.prepared.approval)).toBe(true);
    expect(validateV1ApplyPreparationBindings({
      changeRequest: material.request,
      planResult: material.planResult,
      approval: material.approval,
      runtime: testRuntime,
      destination: material.planResult.io.destination
    })).toBe(true);
    await expect(access(material.destination)).rejects.toMatchObject({ code: "ENOENT" });

    const approvalBytes = Buffer.from(canonicalJsonText(material.approval), "utf8");
    const stdinPrepared = await prepareV1ApplyChange({
      invocation: applyInvocation(material, { approvalPath: "-" }),
      runtime: testRuntime,
      cwd: repoRoot,
      stdin: approvalBytes
    });
    expect(stdinPrepared.error).toBeUndefined();
    expect(stdinPrepared.inputs[3]).toEqual({
      role: "approval",
      option: "--approval",
      source: "stdin",
      path: null,
      digest: sha256RawBytes(approvalBytes)
    });
    await expect(access(material.destination)).rejects.toMatchObject({ code: "ENOENT" });
  });

  it("rejects schema-invalid and digest-divergent approvals before destination reservation", async () => {
    const invalid = await createApprovedPlan();
    await writeFile(invalid.approvalPath, `${canonicalJsonText({ ...invalid.approval, approved: false })}\n`, "utf8");
    const invalidPrepared = await prepareV1ApplyChange({
      invocation: applyInvocation(invalid),
      runtime: testRuntime,
      cwd: repoRoot
    });
    expect(invalidPrepared.error).toMatchObject({ code: "change.approval-invalid", status: "rejected" });
    expect(validateCliDiagnostic(createV1DiagnosticFromError(invalidPrepared.error))).toBe(true);
    await expect(access(invalid.destination)).rejects.toMatchObject({ code: "ENOENT" });

    const divergent = await createApprovedPlan();
    const divergentApproval = structuredClone(divergent.approval);
    divergentApproval.output_plan_digest.value = "e".repeat(64);
    await writeFile(divergent.approvalPath, `${canonicalJsonText(divergentApproval)}\n`, "utf8");
    const divergentPrepared = await prepareV1ApplyChange({
      invocation: applyInvocation(divergent),
      runtime: testRuntime,
      cwd: repoRoot
    });
    expect(divergentPrepared.error).toMatchObject({
      code: "change.binding-mismatch",
      status: "rejected",
      location: { option: "--approval", artifact_role: "approval", rule_id: "RB-006" }
    });
    expect(validateCliDiagnostic(createV1DiagnosticFromError(divergentPrepared.error))).toBe(true);
    expect(validateV1ApplyPreparationBindings({
      changeRequest: divergent.request,
      planResult: divergent.planResult,
      approval: divergentApproval,
      runtime: testRuntime,
      destination: divergent.planResult.io.destination
    })).toBe(false);
    await expect(access(divergent.destination)).rejects.toMatchObject({ code: "ENOENT" });

    const tamperedPlan = await createApprovedPlan();
    const tamperedPlanResult = structuredClone(tamperedPlan.planResult);
    tamperedPlanResult.data.output_plan.semantic_diff_digest.value = "d".repeat(64);
    await writeFile(tamperedPlan.planResultPath, `${canonicalJsonText(tamperedPlanResult)}\n`, "utf8");
    const tamperedPrepared = await prepareV1ApplyChange({
      invocation: applyInvocation(tamperedPlan),
      runtime: testRuntime,
      cwd: repoRoot
    });
    expect(tamperedPrepared.error).toMatchObject({
      code: "change.binding-mismatch",
      location: { rule_id: "RB-003" }
    });
    await expect(access(tamperedPlan.destination)).rejects.toMatchObject({ code: "ENOENT" });
  });

  it("rejects a changed request, stale project, runtime divergence, and destination race without replacing anything", async () => {
    const changedRequest = await createApprovedPlan();
    const changedRequestValue = structuredClone(changedRequest.request);
    changedRequestValue.operations[0].value.percent_complete = 60;
    await writeFile(changedRequest.requestPath, `${canonicalJsonText(changedRequestValue)}\n`, "utf8");
    const changedRequestPrepared = await prepareV1ApplyChange({
      invocation: applyInvocation(changedRequest),
      runtime: testRuntime,
      cwd: repoRoot
    });
    expect(changedRequestPrepared.error).toMatchObject({
      code: "change.binding-mismatch",
      location: { rule_id: "RB-002" }
    });
    await expect(access(changedRequest.destination)).rejects.toMatchObject({ code: "ENOENT" });

    const staleProject = await createApprovedPlan();
    const staleProjectPath = path.join(staleProject.directory, "stale-project.xml");
    const sourceXml = await readFile(fixturePath, "utf8");
    await writeFile(staleProjectPath, sourceXml.replace("<PercentComplete>0</PercentComplete>", "<PercentComplete>10</PercentComplete>"), "utf8");
    const stalePrepared = await prepareV1ApplyChange({
      invocation: applyInvocation(staleProject, { projectPath: staleProjectPath }),
      runtime: testRuntime,
      cwd: repoRoot
    });
    expect(stalePrepared.error).toMatchObject({ code: "change.precondition-failed" });
    await expect(access(staleProject.destination)).rejects.toMatchObject({ code: "ENOENT" });

    const runtimeMismatch = await createApprovedPlan();
    const otherRuntime = { ...testRuntime, version: "1.0.3" };
    const runtimePrepared = await prepareV1ApplyChange({
      invocation: applyInvocation(runtimeMismatch),
      runtime: otherRuntime,
      cwd: repoRoot
    });
    expect(runtimePrepared.error).toMatchObject({
      code: "change.binding-mismatch",
      location: { rule_id: "RB-004" }
    });
    expect(runtimePrepared.inputs[3].digest).toBeNull();
    await expect(access(runtimeMismatch.destination)).rejects.toMatchObject({ code: "ENOENT" });

    const raced = await createApprovedPlan();
    await writeFile(raced.destination, "preserve concurrent entry\n", "utf8");
    const racedPrepared = await prepareV1ApplyChange({
      invocation: applyInvocation(raced),
      runtime: testRuntime,
      cwd: repoRoot
    });
    expect(racedPrepared.error).toMatchObject({ code: "publication.reservation-conflict" });
    expect(validateCliDiagnostic(createV1DiagnosticFromError(racedPrepared.error))).toBe(true);
    expect(await readFile(raced.destination, "utf8")).toBe("preserve concurrent entry\n");

    const parentChanged = await createApprovedPlan();
    const approvedParent = path.dirname(parentChanged.destination);
    const movedParent = path.join(parentChanged.directory, "moved-output-parent");
    await rename(approvedParent, movedParent);
    await symlink(movedParent, approvedParent, "dir");
    const parentChangedPrepared = await prepareV1ApplyChange({
      invocation: applyInvocation(parentChanged),
      runtime: testRuntime,
      cwd: repoRoot
    });
    expect(parentChangedPrepared.error).toMatchObject({
      code: "change.binding-mismatch",
      location: { option: "--plan-result", rule_id: "RB-005" }
    });
    await expect(access(parentChanged.destination)).rejects.toMatchObject({ code: "ENOENT" });
  });
});

describe("v1 apply-change publication service", () => {
  it("applies one approved C1 change, publishes only a committed descriptor, and will not reapply to the same destination", async () => {
    const material = await createApprovedPlan();
    const canonicalDestination = material.planResult.io.destination.path;
    const firstTransport = captureResultTransport();
    const first = await runV1ApplyChange({
      invocation: applyInvocation(material),
      resultTransport: firstTransport,
      runtime: testRuntime,
      cwd: repoRoot
    });

    expect(first).toEqual(firstTransport.result);
    expect(validateCliResult(first)).toBe(true);
    expect(first).toMatchObject({
      command: "apply-change",
      status: "succeeded",
      exit_code: 0,
      io: { destination: material.planResult.io.destination },
      effects: {
        project_input_modified: false,
        project_artifact: {
          path: canonicalDestination,
          publication_state: "committed",
          created_by_invocation: true
        },
        cleanup: { status: "prohibited-after-commit", path: null }
      },
      next_action: { action: "verify-artifact", command: "verify-artifact" },
      data: { artifact_set: { path: canonicalDestination, publication_state: "committed" } }
    });
    const verification = await verifyV1ArtifactSet(canonicalDestination);
    expect(verification).toMatchObject({
      verification: { publication_state: "committed" },
      artifact_set: first.data.artifact_set
    });
    expect(await readFile(fixturePath)).toEqual(await readFile(first.io.inputs[0].path));

    const second = await runV1ApplyChange({
      invocation: applyInvocation(material),
      resultTransport: captureResultTransport(),
      runtime: testRuntime,
      cwd: repoRoot
    });
    expect(validateCliResult(second)).toBe(true);
    expect(second).toMatchObject({
      status: "rejected",
      diagnostics: [{ code: "publication.reservation-conflict" }],
      effects: {
        project_artifact: null,
        cleanup: { status: "not-needed", path: null }
      },
      data: null
    });
    expect(await verifyV1ArtifactSet(canonicalDestination)).toMatchObject({
      verification: { publication_state: "committed" },
      artifact_set: first.data.artifact_set
    });
  });

  it("maps a pre-marker write failure to an absent/succeeded cleanup result without claiming publication", async () => {
    const material = await createApprovedPlan();
    const canonicalDestination = material.planResult.io.destination.path;
    const targetMember = path.join(canonicalDestination, "project.xml");
    const failingFileSystem = new Proxy(await import("node:fs/promises"), {
      get(target, property, receiver) {
        if (property === "open") {
          return async (candidatePath, flags, ...rest) => {
            if (candidatePath === targetMember && flags === "wx") {
              const error = new Error("injected project.xml write failure");
              error.code = "EIO";
              throw error;
            }
            return target.open(candidatePath, flags, ...rest);
          };
        }
        return Reflect.get(target, property, receiver);
      }
    });
    const result = await runV1ApplyChange({
      invocation: applyInvocation(material),
      resultTransport: captureResultTransport(),
      runtime: testRuntime,
      cwd: repoRoot,
      fileSystem: failingFileSystem
    });

    expect(validateCliResult(result)).toBe(true);
    expect(result).toMatchObject({
      status: "runtime-error",
      diagnostics: [{ code: "publication.write-failed" }],
      effects: {
        project_artifact: {
          path: canonicalDestination,
          publication_state: "absent",
          created_by_invocation: true
        },
        cleanup: { status: "succeeded", path: canonicalDestination }
      },
      data: null
    });
    await expect(access(material.destination)).rejects.toMatchObject({ code: "ENOENT" });
  });

  it("does not synthesize a failure envelope when result delivery fails after commit; verification recovers the outcome", async () => {
    const material = await createApprovedPlan();
    const canonicalDestination = material.planResult.io.destination.path;
    const lostResultTransport = Object.freeze({
      target: Object.freeze({ target: "file", path: path.join(material.directory, "lost-result.json") }),
      async writeResult() {
        throw new Error("injected result delivery loss");
      }
    });
    await expect(runV1ApplyChange({
      invocation: applyInvocation(material),
      resultTransport: lostResultTransport,
      runtime: testRuntime,
      cwd: repoRoot
    })).rejects.toThrow("injected result delivery loss");

    const verification = await verifyV1ArtifactSet(canonicalDestination);
    expect(verification).toMatchObject({
      verification: { path: canonicalDestination, publication_state: "committed" },
      artifact_set: { path: canonicalDestination, publication_state: "committed" }
    });
  });
});

async function createApprovedPlan() {
  const directory = await mkdtemp(path.join(os.tmpdir(), "miku-project-v1-apply-prepare-"));
  const requestPath = path.join(directory, "request.json");
  const planResultPath = path.join(directory, "plan.result.json");
  const approvalPath = path.join(directory, "approval.json");
  const destinationParent = path.join(directory, "output-parent");
  const destination = path.join(destinationParent, "next-project");
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
  const resultTransport = await reserveV1ResultTransport(planInvocation.options.result, { cwd: directory });
  const planResult = await runV1PlanChange({
    invocation: planInvocation,
    resultTransport,
    runtime: testRuntime,
    cwd: directory,
    stdin: Buffer.alloc(0)
  });
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
  return { directory, requestPath, planResultPath, approvalPath, destination, request, planResult, approval };
}

function applyInvocation(material, { projectPath = fixturePath, approvalPath = material.approvalPath } = {}) {
  return parseV1Invocation([
    "apply-change",
    "--project", projectPath,
    "--request", material.requestPath,
    "--plan-result", material.planResultPath,
    "--approval", approvalPath
  ]);
}

async function requestArtifact() {
  const template = await readFile(requestTemplatePath, "utf8");
  return JSON.parse(template.replace("${BASE_STATE_DIGEST}", "a98f0c8b560382234572a61f360c9e96911bc75fc1d57b79968d6e60b5d751d0"));
}

async function expectedFileInput(role, option, filePath) {
  const canonicalPath = await realpath(filePath);
  return {
    role,
    option,
    source: "file",
    path: canonicalPath,
    digest: sha256RawBytes(await readFile(canonicalPath))
  };
}

function readJson(relativePath) {
  return JSON.parse(fs.readFileSync(path.join(repoRoot, relativePath), "utf8"));
}

function digest(value) {
  return { algorithm: "sha-256", value };
}

function captureResultTransport() {
  let writtenResult = null;
  return Object.freeze({
    target: Object.freeze({ target: "stdout", path: null }),
    async writeResult(result) {
      if (writtenResult !== null) throw new Error("test result transport was written more than once");
      writtenResult = result;
    },
    get result() {
      return writtenResult;
    }
  });
}
