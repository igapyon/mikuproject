import fsPromises from "node:fs/promises";
import path from "node:path";

import { preflightV1NewDestination } from "./cli-v1-destination.mjs";
import { createV1RejectedError, createV1RuntimeError, isCliV1Error } from "./cli-v1-errors.mjs";
import { inspectV1ArtifactSetContent, verifyV1ArtifactSet } from "./cli-v1-artifact-verifier.mjs";

const PRECOMMIT_MEMBER_NAMES = Object.freeze(["project.xml", "provenance.json"]);

/**
 * Exclusively publishes already-revalidated C1 output material.  The caller
 * owns apply/approval semantics; this state machine owns only directory
 * reservation, exact member creation, marker publication, and cleanup of its
 * own pre-marker files. It deliberately does not construct a CLI result.
 */
export async function publishV1ArtifactSet({
  destination,
  runtime,
  projectBytes,
  provenanceBytes,
  cwd = process.cwd(),
  fileSystem = fsPromises
} = {}) {
  const destinationPath = assertPublisherInputs({ destination, runtime, projectBytes, provenanceBytes, fileSystem, cwd });
  const capabilityError = publisherCapabilityError({ destinationPath, runtime, fileSystem });
  if (capabilityError) return preReservationFailure(destinationPath, capabilityError);
  const material = inspectV1ArtifactSetContent({ projectBytes, provenanceBytes });
  if (!material.valid) {
    return preReservationFailure(destinationPath, createV1RuntimeError({
      code: "internal.unexpected-error",
      message: "The publisher was given C1 artifact material that failed its required pre-publication verification.",
      scope: "artifact",
      path: destinationPath,
      artifactRole: "project_artifact_set",
      details: { reason: material.reason }
    }));
  }

  let approvedDestination;
  try {
    approvedDestination = await preflightV1NewDestination(destinationPath, { cwd, fileSystem });
  } catch (error) {
    return preReservationFailure(destinationPath, toReservationError(error, destinationPath));
  }
  if (approvedDestination.path !== destinationPath) {
    return preReservationFailure(destinationPath, createV1RejectedError({
      code: "change.binding-mismatch",
      message: "The publisher destination does not equal the approved canonical destination path.",
      scope: "filesystem",
      path: approvedDestination.path,
      option: "--plan-result",
      artifactRole: "plan_result",
      ruleId: "RB-005",
      details: { rule_id: "RB-005", approved_path: destinationPath, observed_path: approvedDestination.path }
    }));
  }

  const state = { destinationPath, createdMembers: [], markerBoundaryObserved: false };
  try {
    await fileSystem.mkdir(destinationPath);
  } catch (error) {
    return preReservationFailure(destinationPath, reservationConflictOrFailure(error, destinationPath));
  }

  try {
    await writeOwnedMember({
      fileSystem,
      memberPath: path.join(destinationPath, "project.xml"),
      memberName: "project.xml",
      bytes: projectBytes,
      state
    });
    await writeOwnedMember({
      fileSystem,
      memberPath: path.join(destinationPath, "provenance.json"),
      memberName: "provenance.json",
      bytes: provenanceBytes,
      state
    });
    const precommit = await inspectV1PrecommitArtifactSet(destinationPath, { fileSystem });
    if (!precommit.valid) {
      throw createV1RuntimeError({
        code: "publication.postwrite-verification-failed",
        message: "The uncommitted artifact members did not pass the required pre-marker verification.",
        scope: "filesystem",
        path: destinationPath,
        artifactRole: "project_artifact_set",
        details: { phase: "pre-marker", reason: precommit.reason }
      });
    }

    await createCommitMarker({ fileSystem, markerPath: path.join(destinationPath, "COMMITTED"), state });
    const postcommit = await verifyV1ArtifactSet(destinationPath, { fileSystem });
    if (postcommit.verification.publication_state !== "committed" || postcommit.error) {
      throw createV1RuntimeError({
        code: "publication.postwrite-verification-failed",
        message: "The artifact set could not be verified as committed after the COMMITTED marker was created.",
        scope: "filesystem",
        path: destinationPath,
        artifactRole: "project_artifact_set",
        details: { phase: "post-marker", observed_publication_state: postcommit.verification.publication_state }
      });
    }
    return freezePublishResult({
      destinationPath,
      createdByInvocation: true,
      publicationState: "committed",
      artifactSet: postcommit.artifact_set,
      cleanup: { status: "prohibited-after-commit", path: null },
      errors: []
    });
  } catch (error) {
    const primaryError = toPublisherRuntimeError(error, destinationPath);
    if (state.markerBoundaryObserved) {
      const observed = await safePostMarkerObservation(destinationPath, fileSystem);
      return freezePublishResult({
        destinationPath,
        createdByInvocation: true,
        publicationState: observed.verification.publication_state,
        artifactSet: null,
        cleanup: { status: "prohibited-after-commit", path: null },
        errors: [primaryError]
      });
    }
    const cleanup = await cleanupOwnedPrecommitFiles(state, fileSystem);
    if (cleanup.succeeded) {
      return freezePublishResult({
        destinationPath,
        createdByInvocation: true,
        publicationState: "absent",
        artifactSet: null,
        cleanup: { status: "succeeded", path: destinationPath },
        errors: [primaryError]
      });
    }
    const cleanupError = createV1RuntimeError({
      code: "publication.cleanup-failed",
      message: "The publisher could not safely clean up its pre-marker artifact directory.",
      scope: "filesystem",
      path: destinationPath,
      artifactRole: "project_artifact_set",
      details: { reason: cleanup.reason }
    });
    return freezePublishResult({
      destinationPath,
      createdByInvocation: true,
      publicationState: "incomplete",
      artifactSet: null,
      cleanup: { status: "failed", path: destinationPath },
      errors: [primaryError, cleanupError]
    });
  }
}

/**
 * Applies the same member-content verification used by the committed verifier
 * to the marker-free topology required immediately before publication.
 */
export async function inspectV1PrecommitArtifactSet(artifactSetPath, { fileSystem = fsPromises } = {}) {
  try {
    const root = await fileSystem.lstat(artifactSetPath);
    if (root.isSymbolicLink() || !root.isDirectory()) return { valid: false, reason: "artifact-set-root-type" };
    const names = await fileSystem.readdir(artifactSetPath);
    if (!hasExactNames(names, PRECOMMIT_MEMBER_NAMES)) return { valid: false, reason: "precommit-member-topology" };
    const files = {};
    for (const memberName of PRECOMMIT_MEMBER_NAMES) {
      const memberPath = path.join(artifactSetPath, memberName);
      const entry = await fileSystem.lstat(memberPath);
      if (entry.isSymbolicLink() || !entry.isFile()) return { valid: false, reason: `member-type-${memberName}` };
      files[memberName] = Buffer.from(await fileSystem.readFile(memberPath));
    }
    const content = inspectV1ArtifactSetContent({
      projectBytes: files["project.xml"],
      provenanceBytes: files["provenance.json"]
    });
    return content.valid
      ? { valid: true, content }
      : { valid: false, reason: content.reason };
  } catch (error) {
    return { valid: false, reason: `filesystem-${error?.code ?? "unknown"}` };
  }
}

function assertPublisherInputs({ destination, runtime, projectBytes, provenanceBytes, fileSystem, cwd }) {
  if (!destination || typeof destination.path !== "string" || destination.path.length === 0 || !path.isAbsolute(destination.path)) {
    throw new TypeError("v1 publisher requires an approved absolute destination path");
  }
  if (!(Buffer.isBuffer(projectBytes) || projectBytes instanceof Uint8Array)
    || !(Buffer.isBuffer(provenanceBytes) || provenanceBytes instanceof Uint8Array)) {
    throw new TypeError("v1 publisher requires raw project.xml and provenance.json bytes");
  }
  if (!cwd || typeof cwd !== "string") throw new TypeError("v1 publisher requires a cwd string");
  return destination.path;
}

function publisherCapabilityError({ destinationPath, runtime, fileSystem }) {
  if (!runtime || runtime.binding_status !== "verified" || runtime.capability_profile !== "miku-project-cli-core/v1") {
    return createV1RuntimeError({
      code: "runtime.capability-missing",
      message: "Artifact publication requires a verified miku-project-cli-core/v1 runtime binding.",
      scope: "runtime",
      path: destinationPath,
      artifactRole: "project_artifact_set",
      details: { required_capability_profile: "miku-project-cli-core/v1" }
    });
  }
  const missingMethod = ["mkdir", "open", "lstat", "readdir", "readFile", "realpath", "unlink", "rmdir"]
    .find((methodName) => typeof fileSystem?.[methodName] !== "function");
  if (!missingMethod) return null;
  return createV1RuntimeError({
    code: "publication.capability-unsupported",
    message: "The selected filesystem adapter cannot provide the exclusive publication operations required by v1.",
    scope: "filesystem",
    path: destinationPath,
    artifactRole: "project_artifact_set",
    details: { missing_operation: missingMethod }
  });
}

async function writeOwnedMember({ fileSystem, memberPath, memberName, bytes, state }) {
  let handle = null;
  let closed = false;
  try {
    handle = await fileSystem.open(memberPath, "wx");
    state.createdMembers.push(memberName);
    await handle.writeFile(bytes);
    await handle.close();
    closed = true;
  } catch (error) {
    if (handle && !closed) {
      try { await handle.close(); } catch { /* preserve the first failure */ }
    }
    throw createV1RuntimeError({
      code: "publication.write-failed",
      message: "A project artifact member could not be written and closed exclusively.",
      scope: "filesystem",
      path: memberPath,
      artifactRole: memberName === "project.xml" ? "external_project" : "provenance",
      details: { member: memberName, error_code: error?.code ?? null }
    });
  }
}

async function createCommitMarker({ fileSystem, markerPath, state }) {
  let handle = null;
  let closed = false;
  try {
    handle = await fileSystem.open(markerPath, "wx");
    // From successful exclusive creation onward, a marker may be visible. No
    // cleanup is safe even if close or later verification fails.
    state.markerBoundaryObserved = true;
    await handle.close();
    closed = true;
  } catch (error) {
    if (error?.code === "EEXIST") state.markerBoundaryObserved = true;
    if (handle && !closed) {
      try { await handle.close(); } catch { /* the marker boundary remains observed */ }
    }
    throw createV1RuntimeError({
      code: "publication.write-failed",
      message: "The empty COMMITTED marker could not be created and closed exclusively.",
      scope: "filesystem",
      path: markerPath,
      artifactRole: "commit_marker",
      details: { member: "COMMITTED", error_code: error?.code ?? null }
    });
  }
}

async function cleanupOwnedPrecommitFiles(state, fileSystem) {
  try {
    for (const memberName of [...state.createdMembers].reverse()) {
      const memberPath = path.join(state.destinationPath, memberName);
      let entry;
      try {
        entry = await fileSystem.lstat(memberPath);
      } catch (error) {
        if (error?.code === "ENOENT") continue;
        return { succeeded: false, reason: `member-inspection-${memberName}-${error?.code ?? "unknown"}` };
      }
      if (entry.isSymbolicLink() || !entry.isFile()) return { succeeded: false, reason: `member-ownership-${memberName}` };
      await fileSystem.unlink(memberPath);
    }
    const remaining = await fileSystem.readdir(state.destinationPath);
    if (remaining.length !== 0) return { succeeded: false, reason: "destination-not-empty" };
    const root = await fileSystem.lstat(state.destinationPath);
    if (root.isSymbolicLink() || !root.isDirectory()) return { succeeded: false, reason: "destination-ownership" };
    await fileSystem.rmdir(state.destinationPath);
    return { succeeded: true };
  } catch (error) {
    return { succeeded: false, reason: `cleanup-${error?.code ?? "unknown"}` };
  }
}

async function safePostMarkerObservation(destinationPath, fileSystem) {
  try {
    return await verifyV1ArtifactSet(destinationPath, { fileSystem });
  } catch {
    return { verification: { publication_state: null } };
  }
}

function hasExactNames(names, expectedNames) {
  return Array.isArray(names)
    && names.length === expectedNames.length
    && names.every((name) => expectedNames.includes(name));
}

function preReservationFailure(destinationPath, error) {
  return freezePublishResult({
    destinationPath,
    createdByInvocation: false,
    publicationState: null,
    artifactSet: null,
    cleanup: { status: "not-needed", path: null },
    errors: [error]
  });
}

function toReservationError(error, destinationPath) {
  if (isCliV1Error(error) && error.code === "publication.destination-exists") {
    return createV1RejectedError({
      code: "publication.reservation-conflict",
      message: "The approved destination is no longer available for exclusive publication.",
      scope: "filesystem",
      path: destinationPath,
      artifactRole: "project_artifact_set",
      details: { approved_path: destinationPath }
    });
  }
  return toPublisherRuntimeError(error, destinationPath);
}

function reservationConflictOrFailure(error, destinationPath) {
  if (error?.code === "EEXIST") {
    return createV1RejectedError({
      code: "publication.reservation-conflict",
      message: "The approved destination was created by another process before exclusive reservation completed.",
      scope: "filesystem",
      path: destinationPath,
      artifactRole: "project_artifact_set",
      details: { approved_path: destinationPath }
    });
  }
  return createV1RuntimeError({
    code: "publication.write-failed",
    message: "The approved destination directory could not be created exclusively.",
    scope: "filesystem",
    path: destinationPath,
    artifactRole: "project_artifact_set",
    details: { phase: "directory-reservation", error_code: error?.code ?? null }
  });
}

function toPublisherRuntimeError(error, destinationPath) {
  if (isCliV1Error(error)) return error;
  return createV1RuntimeError({
    code: "internal.unexpected-error",
    message: "The v1 publisher encountered an unexpected failure.",
    scope: "filesystem",
    path: destinationPath,
    artifactRole: "project_artifact_set",
    details: { error_name: error instanceof Error ? error.name : typeof error }
  });
}

function freezePublishResult({ destinationPath, createdByInvocation, publicationState, artifactSet, cleanup, errors }) {
  return Object.freeze({
    destination: Object.freeze({ path: destinationPath }),
    created_by_invocation: createdByInvocation,
    publication_state: publicationState,
    artifact_set: artifactSet ? freezeJson(artifactSet) : null,
    cleanup: Object.freeze({ ...cleanup }),
    error: errors[0] ?? null,
    errors: Object.freeze([...errors])
  });
}

function freezeJson(value) {
  if (Array.isArray(value)) return Object.freeze(value.map((item) => freezeJson(item)));
  if (value && typeof value === "object") {
    for (const key of Object.keys(value)) freezeJson(value[key]);
    return Object.freeze(value);
  }
  return value;
}
