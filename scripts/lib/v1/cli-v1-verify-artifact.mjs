import fsPromises from "node:fs/promises";
import path from "node:path";

import { verifyV1ArtifactSet } from "./cli-v1-artifact-verifier.mjs";
import { createV1RejectedError, createV1RuntimeError, isCliV1Error } from "./cli-v1-errors.mjs";
import { readV1JsonArtifact } from "./cli-v1-json-artifact.mjs";
import { createV1DiagnosticFromError, createV1Result } from "./cli-v1-result.mjs";
import { validateCliResult } from "../../generated/cli-v1-schema-validators.mjs";

/**
 * Executes the read-only `verify-artifact` service.  Publication state is
 * always observed through the verifier; this layer only turns that observation
 * and an optional, strict plan-result input into the command result contract.
 */
export async function runV1VerifyArtifact({
  invocation,
  resultTransport,
  runtime,
  cwd = process.cwd(),
  stdin = process.stdin,
  fileSystem = fsPromises
} = {}) {
  assertVerifyArtifactInvocation(invocation);
  assertResultTransport(resultTransport);

  let artifactInput = unreadArtifactSetInput(invocation.options["artifact-set"], cwd);
  const expectedRequested = Object.hasOwn(invocation.options, "expect-plan-result");
  let expectedPlanInput = expectedRequested
    ? unreadExpectedPlanInput(invocation.options["expect-plan-result"], cwd)
    : null;

  let verificationResult;
  try {
    verificationResult = await verifyV1ArtifactSet(invocation.options["artifact-set"], { cwd, fileSystem });
  } catch (error) {
    const result = verificationResultFailure({
      error: toUnexpectedV1RuntimeError(error),
      runtime,
      artifactInput,
      expectedPlanInput,
      resultTarget: resultTransport.target,
      verification: null
    });
    await resultTransport.writeResult(result);
    return result;
  }
  artifactInput = artifactSetInputMetadata(verificationResult.verification.path);

  if (verificationResult.error) {
    const result = verificationResultFailure({
      error: verificationResult.error,
      runtime,
      artifactInput,
      expectedPlanInput,
      resultTarget: resultTransport.target,
      verification: verificationResult.verification
    });
    await resultTransport.writeResult(result);
    return result;
  }

  if (expectedRequested) {
    const expectedRead = await readV1JsonArtifact(invocation.options["expect-plan-result"], {
      role: "expected_plan_result",
      option: "--expect-plan-result",
      cwd,
      stdin,
      fileSystem
    });
    expectedPlanInput = expectedRead.input ?? expectedPlanInput;
    if (expectedRead.error) {
      const result = verificationResultFailure({
        error: expectedRead.error,
        runtime,
        artifactInput,
        expectedPlanInput,
        resultTarget: resultTransport.target,
        verification: verificationResult.verification
      });
      await resultTransport.writeResult(result);
      return result;
    }
    if (!isSucceededPlanResult(expectedRead.value)) {
      const result = verificationResultFailure({
        error: invalidExpectedPlanResultEnvelope(expectedPlanInput),
        runtime,
        artifactInput,
        expectedPlanInput,
        resultTarget: resultTransport.target,
        // RB-008 compares a committed artifact with a schema-valid successful
        // plan result.  An invalid envelope never reaches that comparison, so
        // keep the committed observation while leaving its match unknown.
        verification: verificationResult.verification
      });
      await resultTransport.writeResult(result);
      return result;
    }

    // Re-run the read-only verifier with the now-strict expected result. This
    // keeps the final reported state and RB-008 binding in one observation.
    try {
      verificationResult = await verifyV1ArtifactSet(invocation.options["artifact-set"], {
        cwd,
        fileSystem,
        expectedPlanResult: expectedRead.value
      });
    } catch (error) {
      const result = verificationResultFailure({
        error: toUnexpectedV1RuntimeError(error),
        runtime,
        artifactInput,
        expectedPlanInput,
        resultTarget: resultTransport.target,
        verification: null
      });
      await resultTransport.writeResult(result);
      return result;
    }
    artifactInput = artifactSetInputMetadata(verificationResult.verification.path);
    if (verificationResult.error) {
      const result = verificationResultFailure({
        error: verificationResult.error,
        runtime,
        artifactInput,
        expectedPlanInput,
        resultTarget: resultTransport.target,
        verification: verificationResult.verification
      });
      await resultTransport.writeResult(result);
      return result;
    }
  }

  const result = createV1Result({
    command: "verify-artifact",
    runtime,
    status: "succeeded",
    io: verifyArtifactIo({ artifactInput, expectedPlanInput, resultTarget: resultTransport.target }),
    effects: artifactObservationEffects(verificationResult.verification),
    data: { verification: verificationResult.verification }
  });
  await resultTransport.writeResult(result);
  return result;
}

/** Implements the result-level RB-011 verifier-path/state checks. */
export function validateV1VerifyArtifactResultBindings({ result } = {}) {
  try {
    if (!validateCliResult(result)
      || result.command !== "verify-artifact"
      || !["succeeded", "rejected"].includes(result.status)
      || !result.data?.verification) return false;
    const artifactInput = result.io?.inputs?.[0];
    const verification = result.data.verification;
    const effect = result.effects?.project_artifact;
    if (artifactInput?.role !== "artifact_set"
      || artifactInput.option !== "--artifact-set"
      || artifactInput.source !== "filesystem-path"
      || artifactInput.digest !== null
      || artifactInput.path !== verification.path
      || effect?.path !== verification.path
      || effect.publication_state !== verification.publication_state
      || effect.created_by_invocation !== false
      || result.effects?.cleanup?.status !== "not-needed"
      || result.effects?.cleanup?.path !== null) return false;
    if (result.status === "succeeded") {
      return verification.publication_state === "committed"
        && verification.bindings !== null
        && (hasExpectedPlanInput(result)
          ? verification.matches_expected_plan === true
          : verification.matches_expected_plan === null);
    }
    return (verification.publication_state === "committed"
      && verification.matches_expected_plan === false
      && verification.bindings !== null)
      || (verification.publication_state === "committed"
        && verification.matches_expected_plan === null
        && verification.bindings !== null
        && hasExpectedPlanInput(result))
      || (["absent", "incomplete", "corrupt"].includes(verification.publication_state)
        && verification.matches_expected_plan === null
        && verification.bindings === null);
  } catch {
    return false;
  }
}

function verificationResultFailure({ error, runtime, artifactInput, expectedPlanInput, resultTarget, verification }) {
  return createV1Result({
    command: "verify-artifact",
    runtime,
    status: error.status,
    io: verifyArtifactIo({ artifactInput, expectedPlanInput, resultTarget }),
    diagnostics: [createV1DiagnosticFromError(error)],
    effects: verification ? artifactObservationEffects(verification) : noArtifactObservationEffects(),
    data: verification ? { verification } : null
  });
}

function verifyArtifactIo({ artifactInput, expectedPlanInput, resultTarget }) {
  const inputs = [copyInput(artifactInput)];
  if (expectedPlanInput) inputs.push(copyInput(expectedPlanInput));
  return {
    stdin_option: expectedPlanInput?.source === "stdin" ? "--expect-plan-result" : null,
    inputs,
    result: { target: resultTarget.target, path: resultTarget.path },
    destination: null
  };
}

function artifactObservationEffects(verification) {
  return {
    project_input_modified: false,
    project_artifact: {
      path: verification.path,
      publication_state: verification.publication_state,
      created_by_invocation: false
    },
    cleanup: { status: "not-needed", path: null }
  };
}

function noArtifactObservationEffects() {
  return {
    project_input_modified: false,
    project_artifact: null,
    cleanup: { status: "not-needed", path: null }
  };
}

function artifactSetInputMetadata(inputPath) {
  return {
    role: "artifact_set",
    option: "--artifact-set",
    source: "filesystem-path",
    path: inputPath,
    digest: null
  };
}

function unreadArtifactSetInput(value, cwd) {
  return artifactSetInputMetadata(path.resolve(cwd, value));
}

function unreadExpectedPlanInput(value, cwd) {
  return {
    role: "expected_plan_result",
    option: "--expect-plan-result",
    source: value === "-" ? "stdin" : "file",
    path: value === "-" ? null : path.resolve(cwd, value),
    digest: null
  };
}

function invalidExpectedPlanResultEnvelope(input) {
  return createV1RejectedError({
    code: "change.binding-mismatch",
    message: "The --expect-plan-result input must be a schema-valid successful v1 plan-change result envelope.",
    scope: "artifact",
    path: input.path,
    option: "--expect-plan-result",
    artifactRole: "expected_plan_result",
    details: {}
  });
}

function isSucceededPlanResult(value) {
  return Boolean(value
    && validateCliResult(value)
    && value.command === "plan-change"
    && value.status === "succeeded");
}

function hasExpectedPlanInput(result) {
  return result.io.inputs.some((input) => input.role === "expected_plan_result");
}

function toUnexpectedV1RuntimeError(error) {
  if (isCliV1Error(error)) return error;
  return createV1RuntimeError({
    message: "The v1 artifact verification service encountered an unexpected internal error.",
    scope: "filesystem",
    option: "--artifact-set",
    artifactRole: "project_artifact_set",
    details: { error_name: error instanceof Error ? error.name : typeof error }
  });
}

function copyInput(input) {
  return {
    role: input.role,
    option: input.option,
    source: input.source,
    path: input.path,
    digest: input.digest ? { ...input.digest } : null
  };
}

function assertVerifyArtifactInvocation(invocation) {
  if (!invocation || invocation.kind !== "workflow" || invocation.command !== "verify-artifact") {
    throw new TypeError("runV1VerifyArtifact requires a parsed v1 verify-artifact workflow invocation");
  }
}

function assertResultTransport(resultTransport) {
  if (!resultTransport || !resultTransport.target || typeof resultTransport.writeResult !== "function") {
    throw new TypeError("runV1VerifyArtifact requires a reserved v1 result transport");
  }
}
