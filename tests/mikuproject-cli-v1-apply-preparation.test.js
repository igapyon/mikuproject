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
  sha256RawBytes,
  sha256SemanticState
} from "../scripts/lib/v1/cli-v1-canonical-json.mjs";
import { validateV1ApplyPreparationBindings } from "../scripts/lib/v1/cli-v1-change.mjs";
import { reserveV1ResultTransport } from "../scripts/lib/v1/cli-v1-io.mjs";
import {
  prepareV1ExternalProjectInput,
  runV1Inspect,
  runV1PlanChange,
  runV1Validate
} from "../scripts/lib/v1/cli-v1-r1-commands.mjs";
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

  it("uses a committed artifact-set directory through validate, inspect, plan-change, and apply-change without changing its source members", async () => {
    const firstMaterial = await createApprovedPlan();
    const first = await runV1ApplyChange({
      invocation: applyInvocation(firstMaterial),
      resultTransport: captureResultTransport(),
      runtime: testRuntime,
      cwd: repoRoot
    });
    expect(first).toMatchObject({ status: "succeeded" });

    const artifactSetPath = first.data.artifact_set.path;
    const projectMemberPath = path.join(artifactSetPath, "project.xml");
    const sourceProjectBefore = await readFile(projectMemberPath);
    const sourceProvenanceBefore = await readFile(path.join(artifactSetPath, "provenance.json"));
    const externalXmlPath = path.join(firstMaterial.directory, "same-state-external.xml");
    await writeFile(externalXmlPath, sourceProjectBefore);

    const fromDirectory = await prepareV1ExternalProjectInput(artifactSetPath, {
      cwd: repoRoot,
      stdin: Buffer.alloc(0)
    });
    const fromExternalXml = await prepareV1ExternalProjectInput(externalXmlPath, {
      cwd: repoRoot,
      stdin: Buffer.alloc(0)
    });
    expect(fromDirectory.error).toBeUndefined();
    expect(fromExternalXml.error).toBeUndefined();
    expect(fromDirectory.input).toEqual({
      role: "project",
      option: "--project",
      source: "directory",
      path: artifactSetPath,
      digest: first.data.artifact_set.project_artifact_digest
    });
    expect(fromDirectory.decoded).toEqual(fromExternalXml.decoded);
    expect(fromDirectory.validation).toEqual(fromExternalXml.validation);

    const validated = await runV1Validate({
      invocation: parseV1Invocation(["validate", "--project", artifactSetPath]),
      resultTransport: captureResultTransport(),
      runtime: testRuntime,
      cwd: repoRoot,
      stdin: Buffer.alloc(0)
    });
    expect(validated).toMatchObject({
      status: "succeeded",
      io: { inputs: [{ source: "directory", path: artifactSetPath, digest: first.data.artifact_set.project_artifact_digest }] },
      data: { validation: { state_digest: sha256SemanticState(fromDirectory.decoded.state) } }
    });

    const inspectedDirectory = await runV1Inspect({
      invocation: parseV1Invocation([
        "inspect", "--project", artifactSetPath, "--purpose", "project_overview"
      ]),
      resultTransport: captureResultTransport(),
      runtime: testRuntime,
      cwd: repoRoot,
      stdin: Buffer.alloc(0)
    });
    const inspectedExternalXml = await runV1Inspect({
      invocation: parseV1Invocation([
        "inspect", "--project", externalXmlPath, "--purpose", "project_overview"
      ]),
      resultTransport: captureResultTransport(),
      runtime: testRuntime,
      cwd: repoRoot,
      stdin: Buffer.alloc(0)
    });
    expect(inspectedDirectory).toMatchObject({ status: "succeeded" });
    expect(inspectedExternalXml).toMatchObject({ status: "succeeded" });
    expect(inspectedDirectory.data.projection).toEqual(inspectedExternalXml.data.projection);

    const secondRequest = await requestArtifact({
      baseStateDigest: sha256SemanticState(fromDirectory.decoded.state).value,
      expectedPercentComplete: 50,
      percentComplete: 75
    });
    const secondRequestPath = path.join(firstMaterial.directory, "second-request.json");
    const secondPlanResultPath = path.join(firstMaterial.directory, "second-plan.result.json");
    const secondApprovalPath = path.join(firstMaterial.directory, "second-approval.json");
    const secondDestinationParent = path.join(firstMaterial.directory, "second-output-parent");
    const secondDestination = path.join(secondDestinationParent, "next-project");
    const directPlanResultPath = path.join(firstMaterial.directory, "direct-plan.result.json");
    const directApprovalPath = path.join(firstMaterial.directory, "direct-approval.json");
    const directDestinationParent = path.join(firstMaterial.directory, "direct-output-parent");
    const directDestination = path.join(directDestinationParent, "next-project");
    await mkdir(secondDestinationParent);
    await mkdir(directDestinationParent);
    await writeFile(secondRequestPath, `${canonicalJsonText(secondRequest)}\n`, "utf8");
    const secondPlanInvocation = parseV1Invocation([
      "plan-change",
      "--project", artifactSetPath,
      "--request", secondRequestPath,
      "--destination", secondDestination,
      "--result", secondPlanResultPath
    ]);
    const secondPlan = await runV1PlanChange({
      invocation: secondPlanInvocation,
      resultTransport: await reserveV1ResultTransport(secondPlanInvocation.options.result, { cwd: firstMaterial.directory }),
      runtime: testRuntime,
      cwd: firstMaterial.directory,
      stdin: Buffer.alloc(0)
    });
    expect(secondPlan.status).toBe("succeeded");
    expect(validateCliResult(secondPlan)).toBe(true);
    expect(secondPlan.io.inputs[0]).toEqual({
      role: "project",
      option: "--project",
      source: "directory",
      path: artifactSetPath,
      digest: first.data.artifact_set.project_artifact_digest
    });
    expect(secondPlan.data.semantic_diff.proposed_state_digest).toEqual(sha256SemanticState({
      ...fromDirectory.decoded.state,
      tasks: fromDirectory.decoded.state.tasks.map((task) => task.uid === "2"
        ? { ...task, percent_complete: 75 }
        : task)
    }));

    const directPlanInvocation = parseV1Invocation([
      "plan-change",
      "--project", externalXmlPath,
      "--request", secondRequestPath,
      "--destination", directDestination,
      "--result", directPlanResultPath
    ]);
    const directPlan = await runV1PlanChange({
      invocation: directPlanInvocation,
      resultTransport: await reserveV1ResultTransport(directPlanInvocation.options.result, { cwd: firstMaterial.directory }),
      runtime: testRuntime,
      cwd: firstMaterial.directory,
      stdin: Buffer.alloc(0)
    });
    expect(directPlan.status).toBe("succeeded");
    expect(secondPlan.data.semantic_diff).toEqual(directPlan.data.semantic_diff);
    expect(secondPlan.data.output_plan.preflight).toEqual(directPlan.data.output_plan.preflight);

    const secondApproval = approvalForPlan(secondRequest, secondPlan);
    const directApproval = approvalForPlan(secondRequest, directPlan);
    await writeFile(secondApprovalPath, `${canonicalJsonText(secondApproval)}\n`, "utf8");
    await writeFile(directApprovalPath, `${canonicalJsonText(directApproval)}\n`, "utf8");
    const second = await runV1ApplyChange({
      invocation: parseV1Invocation([
        "apply-change",
        "--project", artifactSetPath,
        "--request", secondRequestPath,
        "--plan-result", secondPlanResultPath,
        "--approval", secondApprovalPath
      ]),
      resultTransport: captureResultTransport(),
      runtime: testRuntime,
      cwd: firstMaterial.directory,
      stdin: Buffer.alloc(0)
    });
    expect(second.status).toBe("succeeded");
    expect(validateCliResult(second)).toBe(true);
    expect(second.io.inputs[0]).toEqual({
      role: "project",
      option: "--project",
      source: "directory",
      path: artifactSetPath,
      digest: first.data.artifact_set.project_artifact_digest
    });
    expect(second.data.artifact_set).toMatchObject({ publication_state: "committed" });
    const direct = await runV1ApplyChange({
      invocation: parseV1Invocation([
        "apply-change",
        "--project", externalXmlPath,
        "--request", secondRequestPath,
        "--plan-result", directPlanResultPath,
        "--approval", directApprovalPath
      ]),
      resultTransport: captureResultTransport(),
      runtime: testRuntime,
      cwd: firstMaterial.directory,
      stdin: Buffer.alloc(0)
    });
    expect(direct.status).toBe("succeeded");
    expect(validateCliResult(direct)).toBe(true);
    expect(await readFile(path.join(second.data.artifact_set.path, "project.xml"))).toEqual(
      await readFile(path.join(direct.data.artifact_set.path, "project.xml"))
    );
    expect(await readFile(projectMemberPath)).toEqual(sourceProjectBefore);
    expect(await readFile(path.join(artifactSetPath, "provenance.json"))).toEqual(sourceProvenanceBefore);
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

  it("materializes CA-CLEANUP-AGGREGATE-001 when a pre-marker write failure cannot clean up its owned directory", async () => {
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
        if (property === "rmdir") {
          return async (candidatePath, ...rest) => {
            if (candidatePath === canonicalDestination) {
              const error = new Error("injected cleanup failure");
              error.code = "EPERM";
              throw error;
            }
            return target.rmdir(candidatePath, ...rest);
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
      exit_code: 3,
      diagnostics: [
        { code: "publication.cleanup-failed", retryability: "not-retryable" },
        { code: "publication.write-failed", retryability: "after-environment-change" }
      ],
      next_action: { action: "abort-and-investigate", command: null, source_retryability: "not-retryable" },
      effects: {
        project_artifact: {
          path: canonicalDestination,
          publication_state: "incomplete",
          created_by_invocation: true
        },
        cleanup: { status: "failed", path: canonicalDestination }
      },
      data: null
    });
    await expect(access(material.destination)).resolves.toBeUndefined();
    expect(await verifyV1ArtifactSet(canonicalDestination)).toMatchObject({
      verification: { publication_state: "incomplete" }
    });
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

async function requestArtifact({
  baseStateDigest = "a98f0c8b560382234572a61f360c9e96911bc75fc1d57b79968d6e60b5d751d0",
  expectedPercentComplete = 0,
  percentComplete = 50
} = {}) {
  const template = await readFile(requestTemplatePath, "utf8");
  const request = JSON.parse(template.replace("${BASE_STATE_DIGEST}", baseStateDigest));
  request.operations[0].preconditions.expected_percent_complete = expectedPercentComplete;
  request.operations[0].value.percent_complete = percentComplete;
  return request;
}

function approvalForPlan(request, planResult) {
  return {
    kind: "miku_project_change_approval",
    schema_version: "1",
    semantic_contract_version: "1",
    approved: true,
    base_state_digest: { ...planResult.data.semantic_diff.base_state_digest },
    change_request_digest: sha256CanonicalJson(request),
    semantic_diff_digest: sha256CanonicalJson(planResult.data.semantic_diff),
    output_plan_digest: sha256CanonicalJson(planResult.data.output_plan)
  };
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
