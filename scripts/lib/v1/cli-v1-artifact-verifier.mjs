import fsPromises from "node:fs/promises";
import path from "node:path";

import {
  canonicalJsonText,
  sha256CanonicalJson,
  sha256RawBytes,
  sha256SemanticState
} from "./cli-v1-canonical-json.mjs";
import { createV1RejectedError, createV1RuntimeError } from "./cli-v1-errors.mjs";
import { parseV1JsonDocument } from "./cli-v1-json-artifact.mjs";
import { canonicalizeV1ObservationItems } from "./cli-v1-provenance.mjs";
import { validateV1SemanticState } from "./cli-v1-semantic-validator.mjs";
import { encodeMsProjectXmlSubset } from "./cli-v1-xml-encoder.mjs";
import { decodeMsProjectXmlSubset } from "./cli-v1-xml-adapter.mjs";
import { validateArtifact, validateCliResult } from "../../generated/cli-v1-schema-validators.mjs";

const ARTIFACT_MEMBER_NAMES = Object.freeze(["COMMITTED", "project.xml", "provenance.json"]);

/**
 * Inspects one C1 artifact-set path without creating, removing, repairing, or
 * opening any member for writing.  `committed` is returned only after every
 * fixed member, the canonical XML/JSON encodings, provenance schema, and
 * output digest/state bindings have been independently rechecked.
 */
export async function verifyV1ArtifactSet(artifactSetPath, {
  cwd = process.cwd(),
  fileSystem = fsPromises,
  expectedPlanResult = null
} = {}) {
  const candidatePath = resolveArtifactSetPath(artifactSetPath, cwd);
  let rootEntry;
  try {
    rootEntry = await fileSystem.lstat(candidatePath);
  } catch (error) {
    if (isNotFound(error)) return absentArtifactSet(candidatePath);
    return indeterminateArtifactSet(candidatePath, error, "artifact-set-root");
  }
  if (rootEntry.isSymbolicLink() || !rootEntry.isDirectory()) {
    return corruptArtifactSet(candidatePath, "artifact-set-root-type");
  }

  let artifactSetPathCanonical;
  try {
    artifactSetPathCanonical = await fileSystem.realpath(candidatePath);
  } catch (error) {
    return indeterminateArtifactSet(candidatePath, error, "artifact-set-root-canonicalization");
  }

  let memberNames;
  try {
    memberNames = await fileSystem.readdir(artifactSetPathCanonical);
  } catch (error) {
    return indeterminateArtifactSet(artifactSetPathCanonical, error, "artifact-set-members");
  }
  if (!memberNames.includes("COMMITTED")) {
    // A marker-free directory is always an interrupted/non-published set. It
    // must not be promoted to corrupt merely because its partial members are
    // odd: P4.6.4 is allowed to leave exactly such a directory after cleanup.
    return incompleteArtifactSet(artifactSetPathCanonical);
  }
  if (!hasExactArtifactMembers(memberNames)) {
    return corruptArtifactSet(artifactSetPathCanonical, "artifact-set-members");
  }

  const memberEntries = await inspectArtifactMembers(artifactSetPathCanonical, fileSystem);
  if (memberEntries.error) return memberEntries.error;
  const marker = memberEntries.entries.get("COMMITTED");
  if (marker.size !== 0) return corruptArtifactSet(artifactSetPathCanonical, "commit-marker-size");

  const projectPath = path.join(artifactSetPathCanonical, "project.xml");
  const provenancePath = path.join(artifactSetPathCanonical, "provenance.json");
  const projectBytes = await readArtifactMember(projectPath, "project.xml", artifactSetPathCanonical, fileSystem);
  if (projectBytes.error) return projectBytes.error;
  const provenanceBytes = await readArtifactMember(provenancePath, "provenance.json", artifactSetPathCanonical, fileSystem);
  if (provenanceBytes.error) return provenanceBytes.error;

  const content = inspectV1ArtifactSetContent({ projectBytes: projectBytes.bytes, provenanceBytes: provenanceBytes.bytes });
  if (!content.valid) return corruptArtifactSet(artifactSetPathCanonical, content.reason);
  const descriptor = artifactSetDescriptor({
    artifactSetPath: artifactSetPathCanonical,
    projectBytes: projectBytes.bytes,
    provenanceBytes: provenanceBytes.bytes
  });
  const matchesExpectedPlan = expectedPlanResult === null
    ? null
    : validateV1ExpectedPlanResultBinding({ provenance: content.provenance, expectedPlanResult, artifactSetPath: artifactSetPathCanonical });
  if (matchesExpectedPlan === false) {
    return committedArtifactSet({
      artifactSetPath: artifactSetPathCanonical,
      bindings: content.bindings,
      descriptor,
      provenance: content.provenance,
      matchesExpectedPlan,
      error: expectedPlanMismatch(artifactSetPathCanonical)
    });
  }
  return committedArtifactSet({
    artifactSetPath: artifactSetPathCanonical,
    bindings: content.bindings,
    descriptor,
    provenance: content.provenance,
    matchesExpectedPlan,
    error: null
  });
}

/**
 * Checks the two non-marker C1 member byte streams without consulting a
 * directory. P4.6.4 uses this before creating COMMITTED; P4.6.3 uses the same
 * primitive after observing the committed topology.
 */
export function inspectV1ArtifactSetContent({ projectBytes, provenanceBytes } = {}) {
  let project;
  let provenance;
  try {
    project = asArtifactBytes(projectBytes);
    const provenanceRaw = asArtifactBytes(provenanceBytes);
    try {
      provenance = parseV1JsonDocument(provenanceRaw, {
        option: "--artifact-set",
        role: "provenance"
      });
    } catch (error) {
      return invalidArtifactContent(`provenance-${error?.code ?? "invalid"}`);
    }
    if (!isCanonicalProvenanceJson(provenanceRaw, provenance)
      || !validateArtifact(provenance)
      || !hasCanonicalProvenanceObservations(provenance)) {
      return invalidArtifactContent("provenance-schema-or-canonicalization");
    }

    let decoded;
    try {
      decoded = decodeMsProjectXmlSubset(project);
    } catch (error) {
      return invalidArtifactContent(`project-xml-${error?.code ?? "invalid"}`);
    }
    const semanticValidation = validateV1SemanticState(decoded.state, { adapterIssues: decoded.adapter_issues });
    if (!semanticValidation.valid) return invalidArtifactContent("project-semantic-validation");
    let canonicalXml;
    try {
      canonicalXml = encodeMsProjectXmlSubset(decoded.state);
    } catch {
      return invalidArtifactContent("project-xml-canonicalization");
    }
    if (!project.equals(canonicalXml.bytes)) return invalidArtifactContent("project-xml-not-canonical");
    if (!sameDigest(provenance.output.artifact_digest, sha256RawBytes(project))
      || !sameDigest(provenance.output.state_digest, sha256SemanticState(decoded.state))
      || !sameJson(provenance.normalizations, canonicalXml.normalizations)) {
      return invalidArtifactContent("provenance-output-binding");
    }
    return Object.freeze({
      valid: true,
      provenance,
      bindings: provenanceBindings(provenance),
      project_artifact_digest: sha256RawBytes(project),
      provenance_digest: sha256RawBytes(provenanceRaw)
    });
  } catch {
    return invalidArtifactContent("artifact-content-unreadable");
  }
}

/**
 * Implements RB-008 without requiring an output directory to be writable.
 * The expected plan must bind every published output fact that it can know;
 * its input raw bytes are intentionally not materialized in an artifact set.
 */
export function validateV1ExpectedPlanResultBinding({ provenance, expectedPlanResult, artifactSetPath } = {}) {
  try {
    if (!validateArtifact(provenance)
      || !expectedPlanResult
      || !validateCliResult(expectedPlanResult)
      || expectedPlanResult.command !== "plan-change"
      || expectedPlanResult.status !== "succeeded") return false;
    const semanticDiff = expectedPlanResult.data?.semantic_diff;
    const outputPlan = expectedPlanResult.data?.output_plan;
    if (!validateArtifact(semanticDiff) || !validateArtifact(outputPlan)) return false;
    const change = semanticDiff.changes?.[0];
    if (!change || expectedPlanResult.io?.destination?.path !== artifactSetPath
      || outputPlan.output?.destination?.path !== artifactSetPath) return false;
    if (!sameDigest(semanticDiff.base_state_digest, outputPlan.base_state_digest)
      || !sameDigest(semanticDiff.change_request_digest, outputPlan.change_request_digest)
      || !sameDigest(outputPlan.semantic_diff_digest, sha256CanonicalJson(semanticDiff))
      || !sameDigest(outputPlan.preflight?.proposed_state_digest, semanticDiff.proposed_state_digest)
      || !sameJson(outputPlan.runtime, runtimeForOutputPlan(expectedPlanResult.runtime))) return false;
    if (!sameJson(provenance.runtime, outputPlan.runtime)
      || !sameDigest(provenance.input.state_digest, semanticDiff.base_state_digest)
      || !sameDigest(provenance.change.change_request_digest, outputPlan.change_request_digest)
      || !sameDigest(provenance.change.semantic_diff_digest, sha256CanonicalJson(semanticDiff))
      || !sameDigest(provenance.change.output_plan_digest, sha256CanonicalJson(outputPlan))
      || provenance.change.target_task_uid !== change.task_uid
      || provenance.change.before_percent_complete !== change.before
      || provenance.change.after_percent_complete !== change.after
      || !sameDigest(provenance.output.artifact_digest, outputPlan.preflight.project_artifact_digest)
      || !sameDigest(provenance.output.state_digest, semanticDiff.proposed_state_digest)
      || !sameJson(provenance.normalizations, outputPlan.preflight.normalizations)) return false;
    return true;
  } catch {
    return false;
  }
}

function resolveArtifactSetPath(artifactSetPath, cwd) {
  if (typeof artifactSetPath !== "string" || artifactSetPath.length === 0 || artifactSetPath.includes("\0")) {
    throw new TypeError("v1 artifact verification requires a non-empty artifact-set path");
  }
  return path.resolve(cwd, artifactSetPath);
}

async function inspectArtifactMembers(artifactSetPath, fileSystem) {
  const entries = new Map();
  for (const memberName of ARTIFACT_MEMBER_NAMES) {
    const memberPath = path.join(artifactSetPath, memberName);
    let entry;
    try {
      entry = await fileSystem.lstat(memberPath);
    } catch (error) {
      if (isNotFound(error)) return { error: corruptArtifactSet(artifactSetPath, `member-missing-${memberName}`) };
      return { error: indeterminateArtifactSet(artifactSetPath, error, `member-inspection-${memberName}`) };
    }
    if (entry.isSymbolicLink() || !entry.isFile()) {
      return { error: corruptArtifactSet(artifactSetPath, `member-type-${memberName}`) };
    }
    entries.set(memberName, entry);
  }
  return { entries };
}

async function readArtifactMember(memberPath, memberName, artifactSetPath, fileSystem) {
  try {
    return { bytes: Buffer.from(await fileSystem.readFile(memberPath)) };
  } catch (error) {
    if (isNotFound(error)) return { error: corruptArtifactSet(artifactSetPath, `member-disappeared-${memberName}`) };
    return { error: indeterminateArtifactSet(artifactSetPath, error, `member-read-${memberName}`) };
  }
}

function hasExactArtifactMembers(memberNames) {
  if (!Array.isArray(memberNames) || memberNames.length !== ARTIFACT_MEMBER_NAMES.length) return false;
  return memberNames.every((memberName) => ARTIFACT_MEMBER_NAMES.includes(memberName));
}

function isCanonicalProvenanceJson(bytes, provenance) {
  return bytes.equals(Buffer.from(`${canonicalJsonText(provenance)}\n`, "utf8"));
}

function asArtifactBytes(value) {
  if (Buffer.isBuffer(value) || value instanceof Uint8Array) return Buffer.from(value);
  throw new TypeError("artifact-set members must be raw bytes");
}

function invalidArtifactContent(reason) {
  return Object.freeze({ valid: false, reason });
}

function hasCanonicalProvenanceObservations(provenance) {
  try {
    return sameJson(provenance.normalizations, canonicalizeV1ObservationItems(provenance.normalizations, "normalization"))
      && sameJson(provenance.losses, canonicalizeV1ObservationItems(provenance.losses, "loss-or-unsupported"))
      && sameJson(provenance.unsupported, canonicalizeV1ObservationItems(provenance.unsupported, "loss-or-unsupported"));
  } catch {
    return false;
  }
}

function absentArtifactSet(artifactSetPath) {
  return classifiedArtifactSet({
    artifactSetPath,
    publicationState: "absent",
    error: artifactStateError("publication.artifact-absent", "The artifact-set path does not exist.", artifactSetPath)
  });
}

function incompleteArtifactSet(artifactSetPath) {
  return classifiedArtifactSet({
    artifactSetPath,
    publicationState: "incomplete",
    error: artifactStateError("publication.artifact-incomplete", "The artifact-set directory has no COMMITTED marker.", artifactSetPath)
  });
}

function corruptArtifactSet(artifactSetPath, reason) {
  return classifiedArtifactSet({
    artifactSetPath,
    publicationState: "corrupt",
    error: artifactStateError("publication.artifact-corrupt", "The artifact set does not satisfy the v1 committed-set contract.", artifactSetPath, { reason })
  });
}

function indeterminateArtifactSet(artifactSetPath, error, phase) {
  return classifiedArtifactSet({
    artifactSetPath,
    publicationState: null,
    error: createV1RuntimeError({
      code: "io.input-read-failed",
      message: "The artifact set could not be inspected completely.",
      scope: "filesystem",
      path: artifactSetPath,
      option: "--artifact-set",
      artifactRole: "project_artifact_set",
      details: { phase, error_code: error?.code ?? null }
    })
  });
}

function committedArtifactSet({ artifactSetPath, bindings, descriptor, provenance, matchesExpectedPlan, error }) {
  return Object.freeze({
    verification: Object.freeze({
      path: artifactSetPath,
      publication_state: "committed",
      matches_expected_plan: matchesExpectedPlan,
      bindings: freezeJson(bindings)
    }),
    artifact_set: freezeJson(descriptor),
    provenance: freezeJson(structuredClone(provenance)),
    error
  });
}

function classifiedArtifactSet({ artifactSetPath, publicationState, error }) {
  return Object.freeze({
    verification: Object.freeze({
      path: artifactSetPath,
      publication_state: publicationState,
      matches_expected_plan: null,
      bindings: null
    }),
    artifact_set: null,
    provenance: null,
    error
  });
}

function artifactSetDescriptor({ artifactSetPath, projectBytes, provenanceBytes }) {
  return {
    kind: "miku_project_artifact_set",
    schema_version: "1",
    path: artifactSetPath,
    publication_state: "committed",
    project_artifact_digest: sha256RawBytes(projectBytes),
    provenance_digest: sha256RawBytes(provenanceBytes)
  };
}

function provenanceBindings(provenance) {
  return {
    change_request_digest: { ...provenance.change.change_request_digest },
    semantic_diff_digest: { ...provenance.change.semantic_diff_digest },
    output_plan_digest: { ...provenance.change.output_plan_digest }
  };
}

function expectedPlanMismatch(artifactSetPath) {
  return createV1RejectedError({
    code: "publication.expected-plan-mismatch",
    message: "The committed artifact does not match the expected output plan.",
    scope: "artifact",
    path: path.join(artifactSetPath, "provenance.json"),
    option: "--expect-plan-result",
    artifactRole: "provenance",
    ruleId: "RB-008",
    details: { rule_id: "RB-008" }
  });
}

function artifactStateError(code, message, artifactSetPath, details = {}) {
  return createV1RejectedError({
    code,
    message,
    scope: "filesystem",
    path: artifactSetPath,
    option: "--artifact-set",
    artifactRole: "project_artifact_set",
    details
  });
}

function sameDigest(left, right) {
  return Boolean(left && right && left.algorithm === "sha-256" && right.algorithm === "sha-256" && left.value === right.value);
}

function runtimeForOutputPlan(runtime) {
  return {
    family: runtime?.family,
    version: runtime?.version,
    artifact_digest: runtime?.artifact_digest,
    manifest_digest: runtime?.manifest_digest,
    capability_profile: runtime?.capability_profile,
    fixture_suite_version: runtime?.fixture_suite_version
  };
}

function sameJson(left, right) {
  return canonicalJsonText(left) === canonicalJsonText(right);
}

function isNotFound(error) {
  return error?.code === "ENOENT";
}

function freezeJson(value) {
  if (Array.isArray(value)) return Object.freeze(value.map((item) => freezeJson(item)));
  if (value && typeof value === "object") {
    for (const key of Object.keys(value)) freezeJson(value[key]);
    return Object.freeze(value);
  }
  return value;
}
