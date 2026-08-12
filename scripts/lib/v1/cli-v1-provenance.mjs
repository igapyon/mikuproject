import {
  canonicalJsonText,
  canonicalizeSemanticState,
  compareUnicodeScalars,
  sha256CanonicalJson,
  sha256RawBytes,
  sha256SemanticState
} from "./cli-v1-canonical-json.mjs";
import { createV1RejectedError } from "./cli-v1-errors.mjs";
import { decodeMsProjectXmlSubset } from "./cli-v1-xml-adapter.mjs";
import { validateV1SemanticState } from "./cli-v1-semantic-validator.mjs";
import { validateArtifact } from "../../generated/cli-v1-schema-validators.mjs";

export const V1_C1_PROVENANCE_TRANSFORMATIONS = Object.freeze([
  "decode",
  "validate",
  "dry-run-apply",
  "diff",
  "preflight-encode",
  "preflight-redecode",
  "approval-check",
  "apply",
  "reserve-output-directory",
  "encode",
  "redecode-validate",
  "write-provenance",
  "commit-marker"
]);

/**
 * Builds the `provenance.json` bytes that a later publisher will exclusively
 * write. This function is pure: it does not reserve or inspect a destination
 * directory. Its output state/XML can be supplied by P4.6.5; in P4.6.2 the
 * regenerated preflight bytes are used as the concrete output material.
 */
export function createV1C1Provenance({ applyPreparation, output = {} } = {}) {
  const context = normalizeV1C1ProvenanceContext({ applyPreparation, output });
  const provenance = createV1C1ProvenanceRecord(context);
  if (!validateArtifact(provenance)) {
    throw new TypeError("v1 C1 provenance builder produced an artifact-schema-invalid record");
  }
  if (!validateV1C1ProvenanceBindings({ provenance, ...context })) {
    throw provenanceBindingMismatch("The generated provenance does not match the approved plan and actual output material.");
  }
  const bytes = Buffer.from(`${canonicalJsonText(provenance)}\n`, "utf8");
  return Object.freeze({
    provenance: freezeJson(provenance),
    bytes,
    raw_digest: sha256RawBytes(bytes),
    observations: createV1StructuredObservations({
      inputNormalizations: applyPreparation.observations?.normalizations ?? [],
      outputNormalizations: provenance.normalizations,
      losses: provenance.losses,
      unsupported: provenance.unsupported
    })
  });
}

/**
 * Implements RB-007 for the material available at the C1 output boundary.
 * It deliberately requires all expected artifacts rather than trusting a
 * self-consistent provenance record by itself.
 */
export function validateV1C1ProvenanceBindings({
  provenance,
  applyPreparation,
  outputBytes,
  outputState,
  outputNormalizations = []
} = {}) {
  try {
    if (!validateArtifact(provenance)) return false;
    const prepared = applyPreparation?.prepared;
    const projectInput = applyPreparation?.inputs?.find((input) => input.role === "project");
    if (!prepared || !projectInput?.digest || !Buffer.isBuffer(outputBytes)) return false;
    if (!sameJson(provenance.runtime, prepared.runtime)) return false;
    if (!sameDigest(provenance.input.artifact_digest, projectInput.digest)
      || !sameDigest(provenance.input.state_digest, sha256SemanticState(prepared.base_state))) return false;
    if (!sameDigest(provenance.change.change_request_digest, sha256CanonicalJson(prepared.change_request))
      || !sameDigest(provenance.change.semantic_diff_digest, sha256CanonicalJson(prepared.semantic_diff))
      || !sameDigest(provenance.change.output_plan_digest, sha256CanonicalJson(prepared.output_plan))) return false;
    const change = prepared.semantic_diff.changes?.[0];
    if (!change
      || provenance.change.target_task_uid !== change.task_uid
      || provenance.change.before_percent_complete !== change.before
      || provenance.change.after_percent_complete !== change.after) return false;
    if (!sameDigest(provenance.output.artifact_digest, sha256RawBytes(outputBytes))
      || !sameDigest(provenance.output.state_digest, sha256SemanticState(outputState))) return false;
    if (!sameDigest(provenance.input.state_digest, prepared.semantic_diff.base_state_digest)
      || !sameDigest(provenance.output.state_digest, prepared.semantic_diff.proposed_state_digest)
      || !sameDigest(provenance.output.state_digest, prepared.output_plan.preflight.proposed_state_digest)
      || !sameDigest(provenance.output.artifact_digest, prepared.output_plan.preflight.project_artifact_digest)) return false;
    if (!sameJson(provenance.normalizations, canonicalizeV1ObservationItems(outputNormalizations, "normalization"))
      || !sameJson(provenance.normalizations, canonicalizeV1ObservationItems(prepared.output_plan.preflight.normalizations, "normalization"))) return false;
    if (provenance.losses.length !== 0 || provenance.unsupported.length !== 0) return false;
    const decoded = decodeMsProjectXmlSubset(outputBytes);
    const decodedValidation = validateV1SemanticState(decoded.state, { adapterIssues: decoded.adapter_issues });
    return decodedValidation.valid
      && sameJson(canonicalizeSemanticState(decoded.state), canonicalizeSemanticState(outputState))
      && sameJson(canonicalizeSemanticState(outputState), canonicalizeSemanticState(prepared.planned_state));
  } catch {
    return false;
  }
}

/**
 * Produces deterministic result observations without using message text. Exact
 * duplicate `code + path` entries collapse; conflicting entries are rejected
 * rather than choosing an arbitrary before/after or description.
 */
export function createV1StructuredObservations({
  inputNormalizations = [],
  outputNormalizations = [],
  losses = [],
  unsupported = []
} = {}) {
  return Object.freeze({
    normalizations: freezeJson(canonicalizeV1ObservationItems(
      [...inputNormalizations, ...outputNormalizations],
      "normalization"
    )),
    losses: freezeJson(canonicalizeV1ObservationItems(losses, "loss-or-unsupported")),
    unsupported: freezeJson(canonicalizeV1ObservationItems(unsupported, "loss-or-unsupported"))
  });
}

export function canonicalizeV1ObservationItems(items, kind) {
  if (!Array.isArray(items)) throw new TypeError(`${kind} observations must be arrays`);
  const sorted = items.map((item) => {
    if (!item || typeof item !== "object" || Array.isArray(item)
      || typeof item.code !== "string" || item.code.length === 0
      || typeof item.path !== "string") {
      throw new TypeError(`${kind} observation requires non-empty code and string path`);
    }
    if (kind === "normalization") {
      if (!Object.hasOwn(item, "before") || !Object.hasOwn(item, "after")) {
        throw new TypeError("normalization observation requires before and after values");
      }
      return { code: item.code, path: item.path, before: structuredClone(item.before), after: structuredClone(item.after) };
    }
    if (typeof item.description !== "string" || item.description.length === 0) {
      throw new TypeError("loss/unsupported observation requires a description");
    }
    return { code: item.code, path: item.path, description: item.description };
  }).sort(compareObservation);
  const canonical = [];
  for (const item of sorted) {
    const previous = canonical.at(-1);
    if (previous && previous.code === item.code && previous.path === item.path) {
      if (!sameJson(previous, item)) {
        throw new TypeError(`${kind} observations must not conflict for the same code and path`);
      }
      continue;
    }
    canonical.push(item);
  }
  return canonical;
}

function normalizeV1C1ProvenanceContext({ applyPreparation, output }) {
  const prepared = applyPreparation?.prepared;
  const projectInput = applyPreparation?.inputs?.find((input) => input.role === "project");
  if (!prepared || !projectInput?.digest) {
    throw new TypeError("v1 provenance requires successful P4.6.1 apply preparation with a read project input");
  }
  const outputBytes = Buffer.from(output.project_bytes ?? prepared.preflight_project_xml);
  const outputState = canonicalizeSemanticState(output.state ?? prepared.planned_state);
  const outputNormalizations = canonicalizeV1ObservationItems(
    output.normalizations ?? prepared.output_plan.preflight.normalizations,
    "normalization"
  );
  const losses = canonicalizeV1ObservationItems(output.losses ?? [], "loss-or-unsupported");
  const unsupported = canonicalizeV1ObservationItems(output.unsupported ?? [], "loss-or-unsupported");
  if (losses.length !== 0 || unsupported.length !== 0) {
    throw provenanceBindingMismatch("C1 provenance cannot be created for output with loss or unsupported observations.");
  }
  if (!sameJson(outputNormalizations, canonicalizeV1ObservationItems(prepared.output_plan.preflight.normalizations, "normalization"))) {
    throw provenanceBindingMismatch("Output normalizations differ from the approved output plan.");
  }
  return { applyPreparation, prepared, projectInput, outputBytes, outputState, outputNormalizations };
}

function createV1C1ProvenanceRecord({ prepared, projectInput, outputBytes, outputState, outputNormalizations }) {
  const change = prepared.semantic_diff.changes?.[0];
  if (!change) throw provenanceBindingMismatch("The approved semantic diff does not contain the required C1 change.");
  return {
    kind: "miku_project_provenance",
    schema_version: "1",
    semantic_contract_version: "1",
    format_profile: "miku-project-ms-project-xml-subset/v1",
    adapter: "ms-project-xml-adapter/v1",
    artifact_set: {
      kind: "miku_project_artifact_set",
      schema_version: "1",
      publication_protocol: "exclusive-directory-commit-marker/v1",
      commit_marker: "COMMITTED"
    },
    runtime: structuredClone(prepared.runtime),
    input: {
      format: "ms-project-xml",
      format_profile: "miku-project-ms-project-xml-subset/v1",
      artifact_digest: structuredClone(projectInput.digest),
      state_digest: sha256SemanticState(prepared.base_state)
    },
    change: {
      change_request_digest: sha256CanonicalJson(prepared.change_request),
      semantic_diff_digest: sha256CanonicalJson(prepared.semantic_diff),
      output_plan_digest: sha256CanonicalJson(prepared.output_plan),
      target_task_uid: change.task_uid,
      before_percent_complete: change.before,
      after_percent_complete: change.after
    },
    output: {
      format: "ms-project-xml",
      format_profile: "miku-project-ms-project-xml-subset/v1",
      path: "project.xml",
      artifact_digest: sha256RawBytes(outputBytes),
      state_digest: sha256SemanticState(outputState)
    },
    transformations: [...V1_C1_PROVENANCE_TRANSFORMATIONS],
    normalizations: outputNormalizations,
    losses: [],
    unsupported: []
  };
}

function provenanceBindingMismatch(message) {
  return createV1RejectedError({
    code: "change.binding-mismatch",
    message,
    scope: "artifact",
    path: "/",
    option: "--plan-result",
    artifactRole: "plan_result",
    ruleId: "RB-007",
    details: { rule_id: "RB-007" }
  });
}

function compareObservation(left, right) {
  return compareUnicodeScalars(left.code, right.code) || compareUnicodeScalars(left.path, right.path);
}

function sameDigest(left, right) {
  return Boolean(left && right && left.algorithm === "sha-256" && right.algorithm === "sha-256" && left.value === right.value);
}

function sameJson(left, right) {
  return canonicalJsonText(left) === canonicalJsonText(right);
}

function freezeJson(value) {
  if (Array.isArray(value)) return Object.freeze(value.map((item) => freezeJson(item)));
  if (value && typeof value === "object") {
    for (const key of Object.keys(value)) freezeJson(value[key]);
    return Object.freeze(value);
  }
  return value;
}
