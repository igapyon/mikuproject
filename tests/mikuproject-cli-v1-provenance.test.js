import { mkdtemp, readFile, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { prepareV1ApplyChange } from "../scripts/lib/v1/cli-v1-apply.mjs";
import { parseV1Invocation } from "../scripts/lib/v1/cli-v1-argv.mjs";
import { canonicalJsonText, sha256CanonicalJson, sha256RawBytes } from "../scripts/lib/v1/cli-v1-canonical-json.mjs";
import { reserveV1ResultTransport } from "../scripts/lib/v1/cli-v1-io.mjs";
import {
  createV1C1Provenance,
  createV1StructuredObservations,
  V1_C1_PROVENANCE_TRANSFORMATIONS,
  validateV1C1ProvenanceBindings
} from "../scripts/lib/v1/cli-v1-provenance.mjs";
import { runV1PlanChange } from "../scripts/lib/v1/cli-v1-r1-commands.mjs";
import { validateArtifact } from "../scripts/generated/cli-v1-schema-validators.mjs";

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

describe("v1 C1 provenance and structured observations", () => {
  it("creates schema-valid deterministic provenance bytes bound to actual redecoded XML", async () => {
    const { applyPreparation } = await createPreparedApply();
    const first = createV1C1Provenance({ applyPreparation });
    const repeated = createV1C1Provenance({ applyPreparation });

    expect(validateArtifact(first.provenance)).toBe(true);
    expect(first).toEqual(repeated);
    expect(first.bytes.toString("utf8")).toBe(`${canonicalJsonText(first.provenance)}\n`);
    expect(first.raw_digest).toEqual(sha256RawBytes(first.bytes));
    expect(first.provenance).toMatchObject({
      kind: "miku_project_provenance",
      schema_version: "1",
      semantic_contract_version: "1",
      runtime: applyPreparation.prepared.runtime,
      input: {
        artifact_digest: applyPreparation.inputs[0].digest,
        state_digest: applyPreparation.prepared.semantic_diff.base_state_digest
      },
      change: {
        change_request_digest: sha256CanonicalJson(applyPreparation.prepared.change_request),
        semantic_diff_digest: sha256CanonicalJson(applyPreparation.prepared.semantic_diff),
        output_plan_digest: sha256CanonicalJson(applyPreparation.prepared.output_plan),
        target_task_uid: "2",
        before_percent_complete: 0,
        after_percent_complete: 50
      },
      output: {
        path: "project.xml",
        artifact_digest: applyPreparation.prepared.output_plan.preflight.project_artifact_digest,
        state_digest: applyPreparation.prepared.output_plan.preflight.proposed_state_digest
      },
      transformations: V1_C1_PROVENANCE_TRANSFORMATIONS,
      normalizations: [],
      losses: [],
      unsupported: []
    });
    expect(validateV1C1ProvenanceBindings({
      provenance: first.provenance,
      applyPreparation,
      outputBytes: applyPreparation.prepared.preflight_project_xml,
      outputState: applyPreparation.prepared.planned_state,
      outputNormalizations: []
    })).toBe(true);
  });

  it("rejects output bytes/state/observations that diverge from the approved output plan", async () => {
    const { applyPreparation } = await createPreparedApply();
    const bytesChanged = Buffer.concat([applyPreparation.prepared.preflight_project_xml, Buffer.from(" ", "utf8")]);
    expect(captureError(() => createV1C1Provenance({
      applyPreparation,
      output: { project_bytes: bytesChanged }
    }))).toMatchObject({ code: "change.binding-mismatch", location: { rule_id: "RB-007" } });

    const stateChanged = structuredClone(applyPreparation.prepared.planned_state);
    stateChanged.tasks.find((task) => task.uid === "2").percent_complete = 60;
    expect(captureError(() => createV1C1Provenance({
      applyPreparation,
      output: { state: stateChanged }
    }))).toMatchObject({ code: "change.binding-mismatch", location: { rule_id: "RB-007" } });

    expect(captureError(() => createV1C1Provenance({
      applyPreparation,
      output: {
        losses: [{ code: "xml.loss", path: "Project", description: "must not publish" }]
      }
    }))).toMatchObject({ code: "change.binding-mismatch", location: { rule_id: "RB-007" } });

    expect(captureError(() => createV1C1Provenance({
      applyPreparation,
      output: {
        normalizations: [{ code: "xml.changed", path: "Project", before: "a", after: "b" }]
      }
    }))).toMatchObject({ code: "change.binding-mismatch", location: { rule_id: "RB-007" } });
  });

  it("sorts and deduplicates structured observations without parsing messages", () => {
    const observations = createV1StructuredObservations({
      inputNormalizations: [
        { code: "z-code", path: "b", before: 1, after: 2 },
        { code: "a-code", path: "z", before: [1], after: [2] }
      ],
      outputNormalizations: [
        { code: "z-code", path: "b", before: 1, after: 2 },
        { code: "a-code", path: "a", before: null, after: true }
      ],
      losses: [
        { code: "loss-z", path: "b", description: "z" },
        { code: "loss-a", path: "a", description: "a" }
      ],
      unsupported: [
        { code: "unsupported-a", path: "a", description: "a" }
      ]
    });
    expect(observations).toEqual({
      normalizations: [
        { code: "a-code", path: "a", before: null, after: true },
        { code: "a-code", path: "z", before: [1], after: [2] },
        { code: "z-code", path: "b", before: 1, after: 2 }
      ],
      losses: [
        { code: "loss-a", path: "a", description: "a" },
        { code: "loss-z", path: "b", description: "z" }
      ],
      unsupported: [{ code: "unsupported-a", path: "a", description: "a" }]
    });
    expect(() => createV1StructuredObservations({
      inputNormalizations: [{ code: "same", path: "a", before: 1, after: 2 }],
      outputNormalizations: [{ code: "same", path: "a", before: 1, after: 3 }]
    })).toThrow("must not conflict");
  });
});

async function createPreparedApply() {
  const directory = await mkdtemp(path.join(os.tmpdir(), "miku-project-v1-provenance-"));
  const requestPath = path.join(directory, "request.json");
  const planResultPath = path.join(directory, "plan.result.json");
  const approvalPath = path.join(directory, "approval.json");
  const destination = path.join(directory, "next-project");
  const request = await requestArtifact();
  await writeFile(requestPath, `${canonicalJsonText(request)}\n`, "utf8");

  const planInvocation = parseV1Invocation([
    "plan-change", "--project", fixturePath, "--request", requestPath, "--destination", destination, "--result", planResultPath
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
  const invocation = parseV1Invocation([
    "apply-change",
    "--project", fixturePath,
    "--request", requestPath,
    "--plan-result", planResultPath,
    "--approval", approvalPath
  ]);
  const applyPreparation = await prepareV1ApplyChange({ invocation, runtime: testRuntime, cwd: directory, stdin: Buffer.alloc(0) });
  expect(applyPreparation.error).toBeUndefined();
  return { applyPreparation };
}

async function requestArtifact() {
  const template = await readFile(requestTemplatePath, "utf8");
  return JSON.parse(template.replace("${BASE_STATE_DIGEST}", "a98f0c8b560382234572a61f360c9e96911bc75fc1d57b79968d6e60b5d751d0"));
}

function digest(value) {
  return { algorithm: "sha-256", value };
}

function captureError(action) {
  try {
    action();
  } catch (error) {
    return error;
  }
  throw new Error("Expected action to throw");
}
