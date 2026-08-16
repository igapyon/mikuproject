import { mkdtemp, mkdir, readFile, realpath, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { prepareV1ApplyChange } from "../scripts/lib/v1/cli-v1-apply.mjs";
import {
  readV1CommittedArtifactSetProject,
  validateV1ExpectedPlanResultBinding,
  verifyV1ArtifactSet
} from "../scripts/lib/v1/cli-v1-artifact-verifier.mjs";
import { parseV1Invocation } from "../scripts/lib/v1/cli-v1-argv.mjs";
import { canonicalJsonText, sha256CanonicalJson } from "../scripts/lib/v1/cli-v1-canonical-json.mjs";
import { reserveV1ResultTransport } from "../scripts/lib/v1/cli-v1-io.mjs";
import { createV1C1Provenance } from "../scripts/lib/v1/cli-v1-provenance.mjs";
import { runV1PlanChange } from "../scripts/lib/v1/cli-v1-r1-commands.mjs";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const fixturePath = path.join(repoRoot, "testdata/conformance/v1/fixtures/project/dependency-canonical.xml");
const requestTemplatePath = path.join(repoRoot, "testdata/conformance/v1/fixtures/change/set-task-2-percent-0-to-50.template.json");
const testRuntime = Object.freeze({
  binding_status: "verified",
  family: "node",
  version: "1.0.2",
  artifact_digest: digest("0".repeat(64)),
  manifest_digest: digest("1".repeat(64)),
  capability_profile: "miku-project-cli-core/v1",
  fixture_suite_version: "1"
});

describe("v1 artifact-set verifier", () => {
  it("recognizes only a canonical, fully bound three-member set as committed without modifying it", async () => {
    const artifact = await createCommittedArtifact();
    const projectBefore = await readFile(path.join(artifact.destination, "project.xml"));
    const provenanceBefore = await readFile(path.join(artifact.destination, "provenance.json"));

    const verified = await verifyV1ArtifactSet(artifact.destination, {
      expectedPlanResult: artifact.planResult
    });

    expect(verified.error).toBeNull();
    expect(verified.verification).toEqual({
      path: artifact.destination,
      publication_state: "committed",
      matches_expected_plan: true,
      bindings: {
        change_request_digest: sha256CanonicalJson(artifact.request),
        semantic_diff_digest: sha256CanonicalJson(artifact.planResult.data.semantic_diff),
        output_plan_digest: sha256CanonicalJson(artifact.planResult.data.output_plan)
      }
    });
    expect(verified.artifact_set).toMatchObject({
      kind: "miku_project_artifact_set",
      schema_version: "1",
      path: artifact.destination,
      publication_state: "committed",
      project_artifact_digest: artifact.provenance.provenance.output.artifact_digest
    });
    expect(validateV1ExpectedPlanResultBinding({
      provenance: verified.provenance,
      expectedPlanResult: artifact.planResult,
      artifactSetPath: artifact.destination
    })).toBe(true);
    expect(await readFile(path.join(artifact.destination, "project.xml"))).toEqual(projectBefore);
    expect(await readFile(path.join(artifact.destination, "provenance.json"))).toEqual(provenanceBefore);
  });

  it("classifies absent, marker-free, and known-invalid paths without repairing them", async () => {
    const directory = await mkdtemp(path.join(os.tmpdir(), "miku-project-v1-artifact-state-"));
    const absentPath = path.join(directory, "absent");
    const incompletePath = path.join(directory, "incomplete");
    await mkdir(incompletePath);
    await writeFile(path.join(incompletePath, "unexpected.tmp"), "partial", "utf8");

    const absent = await verifyV1ArtifactSet(absentPath);
    const incomplete = await verifyV1ArtifactSet(incompletePath);
    expect(absent.verification).toMatchObject({ publication_state: "absent", matches_expected_plan: null, bindings: null });
    expect(absent.error).toMatchObject({ code: "publication.artifact-absent" });
    expect(incomplete.verification).toMatchObject({ publication_state: "incomplete", matches_expected_plan: null, bindings: null });
    expect(incomplete.error).toMatchObject({ code: "publication.artifact-incomplete" });
    expect(await readFile(path.join(incompletePath, "unexpected.tmp"), "utf8")).toBe("partial");
    const incompleteProject = await readV1CommittedArtifactSetProject(incompletePath);
    expect(incompleteProject).toMatchObject({
      input: {
        role: "project",
        option: "--project",
        source: "directory",
        path: await realpath(incompletePath),
        digest: null
      },
      error: { code: "publication.artifact-incomplete" }
    });
    expect(incompleteProject).not.toHaveProperty("bytes");

    const malformed = await createCommittedArtifact();
    await writeFile(path.join(malformed.destination, "unexpected.tmp"), "do not repair", "utf8");
    const corrupt = await verifyV1ArtifactSet(malformed.destination);
    expect(corrupt.verification).toMatchObject({ publication_state: "corrupt", matches_expected_plan: null, bindings: null });
    expect(corrupt.error).toMatchObject({ code: "publication.artifact-corrupt" });
    expect(await readFile(path.join(malformed.destination, "unexpected.tmp"), "utf8")).toBe("do not repair");
    const corruptProject = await readV1CommittedArtifactSetProject(malformed.destination);
    expect(corruptProject).toMatchObject({
      input: {
        role: "project",
        option: "--project",
        source: "directory",
        path: malformed.destination,
        digest: null
      },
      error: { code: "publication.artifact-corrupt" }
    });
    expect(corruptProject).not.toHaveProperty("bytes");
  });

  it("rejects noncanonical/tampered members, distinguishes an indeterminate read, and detects RB-008 mismatch", async () => {
    const tampered = await createCommittedArtifact();
    const projectPath = path.join(tampered.destination, "project.xml");
    const originalXml = await readFile(projectPath);
    await writeFile(projectPath, Buffer.concat([originalXml, Buffer.from(" ", "utf8")]));
    const corrupt = await verifyV1ArtifactSet(tampered.destination);
    expect(corrupt.verification).toMatchObject({ publication_state: "corrupt", matches_expected_plan: null, bindings: null });
    expect(corrupt.error).toMatchObject({ code: "publication.artifact-corrupt" });

    const complete = await createCommittedArtifact();
    const otherPlanResult = await planResultForDestination(complete.directory, path.join(complete.directory, "other-destination"));
    const mismatch = await verifyV1ArtifactSet(complete.destination, { expectedPlanResult: otherPlanResult });
    expect(mismatch.verification).toMatchObject({ publication_state: "committed", matches_expected_plan: false });
    expect(mismatch.error).toMatchObject({
      code: "publication.expected-plan-mismatch",
      location: { rule_id: "RB-008" }
    });

    const indeterminate = await verifyV1ArtifactSet(path.join(complete.directory, "unreadable"), {
      fileSystem: {
        async lstat() {
          const error = new Error("permission denied");
          error.code = "EACCES";
          throw error;
        }
      }
    });
    expect(indeterminate.verification).toMatchObject({ publication_state: null, matches_expected_plan: null, bindings: null });
    expect(indeterminate.error).toMatchObject({ code: "io.input-read-failed", details: { phase: "artifact-set-root", error_code: "EACCES" } });
  });
});

async function createCommittedArtifact() {
  const directory = await mkdtemp(path.join(os.tmpdir(), "miku-project-v1-artifact-"));
  const requestedDestination = path.join(directory, "next-project");
  const request = await requestArtifact();
  const planResult = await planResultForDestination(directory, requestedDestination, request);
  const destination = planResult.data.output_plan.output.destination.path;
  const approvalPath = path.join(directory, "approval.json");
  const requestPath = path.join(directory, "request.json");
  const planResultPath = path.join(directory, "plan.result.json");
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
  await writeFile(requestPath, `${canonicalJsonText(request)}\n`, "utf8");
  await writeFile(planResultPath, `${canonicalJsonText(planResult)}\n`, "utf8");
  await writeFile(approvalPath, `${canonicalJsonText(approval)}\n`, "utf8");
  const applyInvocation = parseV1Invocation([
    "apply-change",
    "--project", fixturePath,
    "--request", requestPath,
    "--plan-result", planResultPath,
    "--approval", approvalPath
  ]);
  const applyPreparation = await prepareV1ApplyChange({
    invocation: applyInvocation,
    runtime: testRuntime,
    cwd: directory,
    stdin: Buffer.alloc(0)
  });
  expect(applyPreparation.error).toBeUndefined();
  const provenance = createV1C1Provenance({ applyPreparation });
  await mkdir(destination);
  await writeFile(path.join(destination, "project.xml"), applyPreparation.prepared.preflight_project_xml);
  await writeFile(path.join(destination, "provenance.json"), provenance.bytes);
  await writeFile(path.join(destination, "COMMITTED"), Buffer.alloc(0));
  return { directory, destination, request, planResult, provenance };
}

async function planResultForDestination(directory, destination, request = undefined) {
  const requestArtifactValue = request ?? await requestArtifact();
  const requestPath = path.join(directory, `request-${path.basename(destination)}.json`);
  const resultPath = path.join(directory, `plan-${path.basename(destination)}.result.json`);
  await writeFile(requestPath, `${canonicalJsonText(requestArtifactValue)}\n`, "utf8");
  const invocation = parseV1Invocation([
    "plan-change", "--project", fixturePath, "--request", requestPath, "--destination", destination, "--result", resultPath
  ]);
  const resultTransport = await reserveV1ResultTransport(invocation.options.result, { cwd: directory });
  return runV1PlanChange({
    invocation,
    resultTransport,
    runtime: testRuntime,
    cwd: directory,
    stdin: Buffer.alloc(0)
  });
}

async function requestArtifact() {
  const template = await readFile(requestTemplatePath, "utf8");
  return JSON.parse(template.replace("${BASE_STATE_DIGEST}", "a98f0c8b560382234572a61f360c9e96911bc75fc1d57b79968d6e60b5d751d0"));
}

function digest(value) {
  return { algorithm: "sha-256", value };
}
