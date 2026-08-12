import fsPromises, { mkdtemp, mkdir, readFile, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { prepareV1ApplyChange } from "../scripts/lib/v1/cli-v1-apply.mjs";
import { verifyV1ArtifactSet } from "../scripts/lib/v1/cli-v1-artifact-verifier.mjs";
import { parseV1Invocation } from "../scripts/lib/v1/cli-v1-argv.mjs";
import { canonicalJsonText, sha256CanonicalJson } from "../scripts/lib/v1/cli-v1-canonical-json.mjs";
import { reserveV1ResultTransport } from "../scripts/lib/v1/cli-v1-io.mjs";
import { publishV1ArtifactSet } from "../scripts/lib/v1/cli-v1-publisher.mjs";
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

describe("v1 exclusive artifact publisher", () => {
  it("creates exactly the committed C1 set through exclusive directory/member/marker operations", async () => {
    const material = await createPublicationMaterial();
    const published = await publishMaterial(material);

    expect(published).toMatchObject({
      destination: { path: material.destination },
      created_by_invocation: true,
      publication_state: "committed",
      cleanup: { status: "prohibited-after-commit", path: null },
      errors: []
    });
    expect(published.artifact_set).toMatchObject({
      kind: "miku_project_artifact_set",
      path: material.destination,
      project_artifact_digest: material.provenance.provenance.output.artifact_digest
    });
    expect((await verifyV1ArtifactSet(material.destination)).verification).toMatchObject({ publication_state: "committed" });
    expect(await readFile(path.join(material.destination, "COMMITTED"))).toEqual(Buffer.alloc(0));
  });

  it("does not overwrite a raced destination, and cleans only its own marker-free output after member or marker failure", async () => {
    const raced = await createPublicationMaterial();
    await mkdir(raced.destination);
    await writeFile(path.join(raced.destination, "caller-owned.txt"), "preserve", "utf8");
    const conflict = await publishMaterial(raced);
    expect(conflict).toMatchObject({
      created_by_invocation: false,
      publication_state: null,
      cleanup: { status: "not-needed", path: null },
      error: { code: "publication.reservation-conflict" }
    });
    expect(await readFile(path.join(raced.destination, "caller-owned.txt"), "utf8")).toBe("preserve");

    const unsupportedFilesystem = await createPublicationMaterial();
    const unsupported = await publishMaterial(unsupportedFilesystem, { fileSystem: {} });
    expect(unsupported).toMatchObject({
      created_by_invocation: false,
      publication_state: null,
      cleanup: { status: "not-needed", path: null },
      error: { code: "publication.capability-unsupported", details: { missing_operation: "mkdir" } }
    });
    await expect(fsPromises.lstat(unsupportedFilesystem.destination)).rejects.toMatchObject({ code: "ENOENT" });

    const writeFailure = await createPublicationMaterial();
    const failedWrite = await publishMaterial(writeFailure, {
      fileSystem: withFilesystemFault({
        open: async (memberPath, flags) => {
          if (path.basename(memberPath) === "project.xml") throw codedError("EIO");
          return fsPromises.open(memberPath, flags);
        }
      })
    });
    expect(failedWrite).toMatchObject({
      created_by_invocation: true,
      publication_state: "absent",
      cleanup: { status: "succeeded", path: writeFailure.destination },
      error: { code: "publication.write-failed" }
    });
    await expect(fsPromises.lstat(writeFailure.destination)).rejects.toMatchObject({ code: "ENOENT" });

    const markerFailure = await createPublicationMaterial();
    const failedMarker = await publishMaterial(markerFailure, {
      fileSystem: withFilesystemFault({
        open: async (memberPath, flags) => {
          if (path.basename(memberPath) === "COMMITTED") throw codedError("EIO");
          return fsPromises.open(memberPath, flags);
        }
      })
    });
    expect(failedMarker).toMatchObject({
      publication_state: "absent",
      cleanup: { status: "succeeded", path: markerFailure.destination },
      error: { code: "publication.write-failed" }
    });
    await expect(fsPromises.lstat(markerFailure.destination)).rejects.toMatchObject({ code: "ENOENT" });
  });

  it("leaves a marker-free incomplete set when cleanup fails and never cleans after the marker boundary", async () => {
    const cleanupFailure = await createPublicationMaterial();
    const failedCleanup = await publishMaterial(cleanupFailure, {
      fileSystem: withFilesystemFault({
        open: async (memberPath, flags) => {
          if (path.basename(memberPath) === "project.xml") throw codedError("EIO");
          return fsPromises.open(memberPath, flags);
        },
        rmdir: async () => { throw codedError("EPERM"); }
      })
    });
    expect(failedCleanup).toMatchObject({
      publication_state: "incomplete",
      cleanup: { status: "failed", path: cleanupFailure.destination },
      errors: [
        { code: "publication.write-failed" },
        { code: "publication.cleanup-failed" }
      ]
    });
    expect((await verifyV1ArtifactSet(cleanupFailure.destination)).verification).toMatchObject({ publication_state: "incomplete" });

    const postMarkerFailure = await createPublicationMaterial();
    let readdirCalls = 0;
    let rmdirCalls = 0;
    const postMarker = await publishMaterial(postMarkerFailure, {
      fileSystem: withFilesystemFault({
        readdir: async (directoryPath) => {
          readdirCalls += 1;
          if (readdirCalls === 2) throw codedError("EIO");
          return fsPromises.readdir(directoryPath);
        },
        rmdir: async (directoryPath) => {
          rmdirCalls += 1;
          return fsPromises.rmdir(directoryPath);
        }
      })
    });
    expect(postMarker).toMatchObject({
      publication_state: "committed",
      artifact_set: null,
      cleanup: { status: "prohibited-after-commit", path: null },
      error: { code: "publication.postwrite-verification-failed" }
    });
    expect(rmdirCalls).toBe(0);
    expect(await readFile(path.join(postMarkerFailure.destination, "COMMITTED"))).toEqual(Buffer.alloc(0));
  });
});

async function createPublicationMaterial() {
  const directory = await mkdtemp(path.join(os.tmpdir(), "miku-project-v1-publisher-"));
  const requestedDestination = path.join(directory, "next-project");
  const request = await requestArtifact();
  const planResult = await planResultForDestination(directory, requestedDestination, request);
  const destination = planResult.data.output_plan.output.destination.path;
  const requestPath = path.join(directory, "request.json");
  const planResultPath = path.join(directory, "plan.result.json");
  const approvalPath = path.join(directory, "approval.json");
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
  const applyPreparation = await prepareV1ApplyChange({
    invocation: parseV1Invocation([
      "apply-change", "--project", fixturePath, "--request", requestPath, "--plan-result", planResultPath, "--approval", approvalPath
    ]),
    runtime: testRuntime,
    cwd: directory,
    stdin: Buffer.alloc(0)
  });
  expect(applyPreparation.error).toBeUndefined();
  const provenance = createV1C1Provenance({ applyPreparation });
  return {
    directory,
    destination,
    projectBytes: applyPreparation.prepared.preflight_project_xml,
    provenanceBytes: provenance.bytes,
    provenance
  };
}

function publishMaterial(material, { fileSystem = fsPromises } = {}) {
  return publishV1ArtifactSet({
    destination: { path: material.destination },
    runtime: testRuntime,
    projectBytes: material.projectBytes,
    provenanceBytes: material.provenanceBytes,
    cwd: material.directory,
    fileSystem
  });
}

async function planResultForDestination(directory, destination, request) {
  const requestPath = path.join(directory, `request-${path.basename(destination)}.json`);
  const resultPath = path.join(directory, `plan-${path.basename(destination)}.result.json`);
  await writeFile(requestPath, `${canonicalJsonText(request)}\n`, "utf8");
  const invocation = parseV1Invocation([
    "plan-change", "--project", fixturePath, "--request", requestPath, "--destination", destination, "--result", resultPath
  ]);
  const resultTransport = await reserveV1ResultTransport(invocation.options.result, { cwd: directory });
  return runV1PlanChange({ invocation, resultTransport, runtime: testRuntime, cwd: directory, stdin: Buffer.alloc(0) });
}

async function requestArtifact() {
  const template = await readFile(requestTemplatePath, "utf8");
  return JSON.parse(template.replace("${BASE_STATE_DIGEST}", "a98f0c8b560382234572a61f360c9e96911bc75fc1d57b79968d6e60b5d751d0"));
}

function withFilesystemFault(overrides) {
  return Object.assign(Object.create(fsPromises), overrides);
}

function codedError(code) {
  const error = new Error(code);
  error.code = code;
  return error;
}

function digest(value) {
  return { algorithm: "sha-256", value };
}
