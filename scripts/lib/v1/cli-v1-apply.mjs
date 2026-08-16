import fsPromises from "node:fs/promises";
import path from "node:path";

import {
  getV1ApprovedDestinationFromPlanResult,
  planV1SetTaskPercentComplete,
  prepareV1ApprovedChange
} from "./cli-v1-change.mjs";
import { canonicalJsonText, sha256RawBytes, sha256SemanticState } from "./cli-v1-canonical-json.mjs";
import { preflightV1NewDestination } from "./cli-v1-destination.mjs";
import { createV1RejectedError, createV1RuntimeError, isCliV1Error } from "./cli-v1-errors.mjs";
import { readV1JsonArtifact } from "./cli-v1-json-artifact.mjs";
import { publishV1ArtifactSet } from "./cli-v1-publisher.mjs";
import { createV1C1Provenance } from "./cli-v1-provenance.mjs";
import { prepareV1ExternalProjectInput } from "./cli-v1-r1-commands.mjs";
import { createV1DiagnosticFromError, createV1Result } from "./cli-v1-result.mjs";
import { semanticIssuesToV1Errors, validateV1SemanticState } from "./cli-v1-semantic-validator.mjs";
import { validateCliResult } from "../../generated/cli-v1-schema-validators.mjs";

/**
 * Reads and revalidates every apply-change input, then recomputes the C1 plan.
 * This is the P4.6.1 boundary: it returns internal material for a later
 * publisher but never creates, reserves, repairs, or removes a destination.
 */
export async function prepareV1ApplyChange({
  invocation,
  runtime,
  cwd = process.cwd(),
  stdin = process.stdin,
  fileSystem = fsPromises
} = {}) {
  assertV1ApplyPreparationInvocation(invocation);
  const inputs = [
    unreadV1ApplyPreparationInput(invocation.options.project, "project", "--project", cwd),
    unreadV1ApplyPreparationInput(invocation.options.request, "change_request", "--request", cwd),
    unreadV1ApplyPreparationInput(invocation.options["plan-result"], "plan_result", "--plan-result", cwd),
    unreadV1ApplyPreparationInput(invocation.options.approval, "approval", "--approval", cwd)
  ];

  const projectRead = await prepareV1ExternalProjectInput(invocation.options.project, { cwd, stdin, fileSystem });
  inputs[0] = projectRead.input ?? inputs[0];
  if (projectRead.error) return failedV1ApplyPreparation({ inputs, error: projectRead.error });
  if (!projectRead.validation.valid) {
    const errors = semanticIssuesToV1Errors(projectRead.validation);
    return failedV1ApplyPreparation({
      inputs,
      error: errors[0],
      errors,
      observations: v1ApplyPreparationSemanticObservations(projectRead.decoded, projectRead.validation)
    });
  }

  const requestRead = await readV1JsonArtifact(invocation.options.request, {
    role: "change_request",
    option: "--request",
    cwd,
    stdin,
    fileSystem
  });
  inputs[1] = requestRead.input ?? inputs[1];
  if (requestRead.error) {
    return failedV1ApplyPreparation({ inputs, error: requestRead.error, observations: v1ApplyPreparationInputObservations(projectRead.decoded) });
  }

  const planResultRead = await readV1JsonArtifact(invocation.options["plan-result"], {
    role: "plan_result",
    option: "--plan-result",
    cwd,
    stdin,
    fileSystem
  });
  inputs[2] = planResultRead.input ?? inputs[2];
  if (planResultRead.error) {
    return failedV1ApplyPreparation({ inputs, error: planResultRead.error, observations: v1ApplyPreparationInputObservations(projectRead.decoded) });
  }

  let approvedDestination;
  try {
    approvedDestination = getV1ApprovedDestinationFromPlanResult(planResultRead.value, runtime);
  } catch (error) {
    return failedV1ApplyPreparation({
      inputs,
      error: toV1ApplyPreparationError(error),
      observations: v1ApplyPreparationInputObservations(projectRead.decoded)
    });
  }

  const approvalRead = await readV1JsonArtifact(invocation.options.approval, {
    role: "approval",
    option: "--approval",
    cwd,
    stdin,
    fileSystem
  });
  inputs[3] = approvalRead.input ?? inputs[3];
  if (approvalRead.error) {
    return failedV1ApplyPreparation({
      inputs,
      destination: approvedDestination,
      error: approvalRead.error,
      observations: v1ApplyPreparationInputObservations(projectRead.decoded)
    });
  }

  let prepared;
  try {
    prepared = prepareV1ApprovedChange({
      state: projectRead.decoded.state,
      changeRequest: requestRead.value,
      planResult: planResultRead.value,
      approval: approvalRead.value,
      runtime,
      destination: approvedDestination
    });
  } catch (error) {
    return failedV1ApplyPreparation({
      inputs,
      destination: approvedDestination,
      error: toV1ApplyPreparationError(error),
      observations: v1ApplyPreparationInputObservations(projectRead.decoded)
    });
  }

  let recheckedDestination;
  try {
    // The approved canonical absolute path, not the caller's current cwd, is
    // authoritative at apply time. This check remains read-only.
    recheckedDestination = await preflightV1NewDestination(approvedDestination.path, {
      cwd,
      fileSystem,
      projectInput: inputs[0]
    });
  } catch (error) {
    return failedV1ApplyPreparation({
      inputs,
      destination: approvedDestination,
      error: toV1ApplyDestinationError(error, approvedDestination),
      observations: v1ApplyPreparationObservations(projectRead.decoded, prepared)
    });
  }
  if (recheckedDestination.path !== approvedDestination.path) {
    return failedV1ApplyPreparation({
      inputs,
      destination: approvedDestination,
      error: createV1RejectedError({
        code: "change.binding-mismatch",
        message: "The approved destination parent no longer resolves to the planning-time canonical path.",
        scope: "filesystem",
        path: recheckedDestination.path,
        option: "--plan-result",
        artifactRole: "plan_result",
        ruleId: "RB-005",
        details: {
          approved_path: approvedDestination.path,
          rechecked_path: recheckedDestination.path
        }
      }),
      observations: v1ApplyPreparationObservations(projectRead.decoded, prepared)
    });
  }

  return Object.freeze({
    inputs: copyV1ApplyPreparationInputs(inputs),
    destination: Object.freeze({ ...approvedDestination }),
    observations: v1ApplyPreparationObservations(projectRead.decoded, prepared),
    prepared
  });
}

/**
 * Executes the approved C1 change exactly through the P4.6 sequence after a
 * caller has reserved its result channel: revalidate inputs/bindings, apply
 * and validate the planned semantic state, build provenance, then publish one
 * new artifact set.  It never reuses a destination and only reports success
 * after the publisher's post-marker verifier has returned its descriptor.
 *
 * A result-channel write error is deliberately allowed to escape this service.
 * At that point the artifact may already be committed, so replacing it with a
 * synthetic failure result would hide an unknown outcome and invite re-apply.
 */
export async function runV1ApplyChange({
  invocation,
  resultTransport,
  runtime,
  cwd = process.cwd(),
  stdin = process.stdin,
  fileSystem = fsPromises
} = {}) {
  assertV1ApplyPreparationInvocation(invocation);
  assertV1ApplyResultTransport(resultTransport);

  const preparation = await prepareV1ApplyChange({ invocation, runtime, cwd, stdin, fileSystem });
  if (preparation.error) {
    const result = createV1ApplyFailureResult({
      errors: preparation.errors,
      runtime,
      io: v1ApplyIo(preparation.inputs, preparation.destination, resultTransport.target),
      observations: preparation.observations,
      effects: noV1ApplyArtifactEffects()
    });
    await resultTransport.writeResult(result);
    return result;
  }

  let applied;
  try {
    applied = materializeV1ApprovedApply({ preparation, runtime });
  } catch (error) {
    const errors = materializationErrors(error);
    const result = createV1ApplyFailureResult({
      errors,
      runtime,
      io: v1ApplyIo(preparation.inputs, preparation.destination, resultTransport.target),
      observations: preparation.observations,
      effects: noV1ApplyArtifactEffects()
    });
    await resultTransport.writeResult(result);
    return result;
  }

  let provenance;
  try {
    provenance = createV1C1Provenance({
      applyPreparation: preparation,
      output: {
        project_bytes: applied.projectBytes,
        state: applied.state,
        normalizations: applied.normalizations,
        losses: [],
        unsupported: []
      }
    });
  } catch (error) {
    const result = createV1ApplyFailureResult({
      errors: materializationErrors(error),
      runtime,
      io: v1ApplyIo(preparation.inputs, preparation.destination, resultTransport.target),
      observations: preparation.observations,
      effects: noV1ApplyArtifactEffects()
    });
    await resultTransport.writeResult(result);
    return result;
  }

  const published = await publishV1ArtifactSet({
    destination: preparation.destination,
    runtime,
    projectBytes: applied.projectBytes,
    provenanceBytes: provenance.bytes,
    cwd,
    fileSystem
  });
  const io = v1ApplyIo(preparation.inputs, preparation.destination, resultTransport.target);
  if (published.error) {
    const result = createV1ApplyFailureResult({
      errors: published.errors,
      runtime,
      io,
      observations: provenance.observations,
      effects: v1ApplyPublicationEffects(published)
    });
    await resultTransport.writeResult(result);
    return result;
  }

  try {
    assertV1ApplySuccessBindings({ preparation, published, io });
  } catch (error) {
    // The marker was already created. Do not translate this into a rejected
    // apply result or cleanup attempt: verification is the recovery path.
    const result = createV1ApplyFailureResult({
      errors: materializationErrors(error),
      runtime,
      io,
      observations: provenance.observations,
      effects: v1ApplyPublicationEffects(published, { hideCommittedState: true })
    });
    await resultTransport.writeResult(result);
    return result;
  }

  const result = createV1Result({
    command: "apply-change",
    runtime,
    status: "succeeded",
    io,
    observations: provenance.observations,
    effects: v1ApplyPublicationEffects(published),
    data: { artifact_set: published.artifact_set }
  });
  await resultTransport.writeResult(result);
  return result;
}

/**
 * Implements the result-level RB-011 check shared by conformance runners.
 * Publication already enforces these facts before a success result is emitted;
 * keeping the pure checker public lets Node and downstream ports validate a
 * received plan/result pair without rerunning or modifying an apply.
 */
export function validateV1ApplyResultBindings({ applyResult, planResult } = {}) {
  try {
    if (!validateCliResult(applyResult)
      || !validateCliResult(planResult)
      || applyResult.command !== "apply-change"
      || applyResult.status !== "succeeded"
      || planResult.command !== "plan-change"
      || planResult.status !== "succeeded") return false;
    const destinationPath = planResult.data?.output_plan?.output?.destination?.path;
    const effect = applyResult.effects?.project_artifact;
    const descriptor = applyResult.data?.artifact_set;
    return Boolean(destinationPath
      && applyResult.io?.destination?.path === destinationPath
      && planResult.io?.destination?.path === destinationPath
      && effect?.path === destinationPath
      && effect.publication_state === "committed"
      && effect.created_by_invocation === true
      && applyResult.effects?.cleanup?.status === "prohibited-after-commit"
      && applyResult.effects?.cleanup?.path === null
      && descriptor?.path === destinationPath
      && descriptor.publication_state === "committed");
  } catch {
    return false;
  }
}

function materializeV1ApprovedApply({ preparation, runtime }) {
  const prepared = preparation.prepared;
  let actual;
  try {
    // This is intentionally distinct from P4.6.1's planning revalidation:
    // the apply stage recomputes the allowed mutation from the revalidated
    // base state and then checks the post-apply state before bytes are handed
    // to provenance or the publisher.
    actual = planV1SetTaskPercentComplete({
      state: prepared.base_state,
      changeRequest: prepared.change_request,
      runtime,
      destination: preparation.destination
    });
  } catch (error) {
    throw toV1ApplyMaterializationError(error);
  }

  const validation = validateV1SemanticState(actual.planned_state);
  if (!validation.valid) {
    const errors = semanticIssuesToV1Errors(validation);
    throw errors.length === 1
      ? errors[0]
      : createV1RuntimeError({
        message: "The approved C1 apply stage produced more than one post-apply semantic validation issue.",
        scope: "semantic",
        path: "state",
        option: "--request",
        artifactRole: "change_request",
        details: { issue_count: errors.length }
      });
  }

  const projectBytes = Buffer.from(actual.preflight_project_xml);
  const stateDigest = sha256SemanticState(actual.planned_state);
  const projectDigest = sha256RawBytes(projectBytes);
  if (canonicalJsonText(actual.semantic_diff) !== canonicalJsonText(prepared.semantic_diff)
    || canonicalJsonText(actual.output_plan) !== canonicalJsonText(prepared.output_plan)
    || canonicalJsonText(actual.planned_state) !== canonicalJsonText(prepared.planned_state)
    || !sameV1ApplyDigest(stateDigest, prepared.output_plan.preflight.proposed_state_digest)
    || !sameV1ApplyDigest(projectDigest, prepared.output_plan.preflight.project_artifact_digest)
    || !projectBytes.equals(Buffer.from(prepared.preflight_project_xml))) {
    throw createV1RuntimeError({
      message: "The actual C1 apply output did not match the approved output-plan preflight material.",
      scope: "artifact",
      path: preparation.destination.path,
      option: "--plan-result",
      artifactRole: "plan_result",
      details: { required_rule: "RB-007" }
    });
  }

  return Object.freeze({
    state: actual.planned_state,
    projectBytes,
    normalizations: actual.output_plan.preflight.normalizations
  });
}

function createV1ApplyFailureResult({ errors, runtime, io, observations, effects }) {
  const normalizedErrors = Array.isArray(errors) ? errors : [errors];
  if (normalizedErrors.length === 0 || normalizedErrors.some((error) => !isCliV1Error(error))) {
    throw new TypeError("v1 apply failure results require one or more CliV1Error values");
  }
  const status = normalizedErrors.some((error) => error.status === "runtime-error")
    ? "runtime-error"
    : normalizedErrors[0].status;
  if (normalizedErrors.some((error) => error.status !== status && error.status !== "runtime-error")) {
    throw new TypeError("v1 apply failure diagnostics must have compatible result statuses");
  }
  return createV1Result({
    command: "apply-change",
    runtime,
    status,
    io,
    diagnostics: normalizedErrors.map(createV1DiagnosticFromError),
    observations,
    effects,
    data: null
  });
}

function materializationErrors(error) {
  return [toV1ApplyMaterializationError(error)];
}

function toV1ApplyMaterializationError(error) {
  if (isCliV1Error(error)) return error;
  return createV1RuntimeError({
    message: "The v1 apply-change service encountered an unexpected error before publication.",
    scope: "internal",
    details: { error_name: error instanceof Error ? error.name : typeof error }
  });
}

function v1ApplyIo(inputs, destination, resultTarget) {
  const copiedInputs = copyV1ApplyPreparationInputs(inputs);
  const stdinInput = copiedInputs.find((input) => input.source === "stdin");
  return {
    stdin_option: stdinInput ? stdinInput.option : null,
    inputs: copiedInputs,
    result: { target: resultTarget.target, path: resultTarget.path },
    destination: destination ? { ...destination } : null
  };
}

function noV1ApplyArtifactEffects() {
  return {
    project_input_modified: false,
    project_artifact: null,
    cleanup: { status: "not-needed", path: null }
  };
}

function v1ApplyPublicationEffects(published, { hideCommittedState = false } = {}) {
  const state = hideCommittedState && published.publication_state === "committed"
    ? null
    : published.publication_state;
  return {
    project_input_modified: false,
    project_artifact: published.created_by_invocation
      ? {
        path: published.destination.path,
        publication_state: state,
        created_by_invocation: true
      }
      : null,
    cleanup: { ...published.cleanup }
  };
}

function assertV1ApplySuccessBindings({ preparation, published, io }) {
  const destinationPath = preparation.destination?.path;
  const descriptorPath = published.artifact_set?.path;
  if (published.publication_state !== "committed"
    || !published.artifact_set
    || !published.created_by_invocation
    || published.cleanup?.status !== "prohibited-after-commit"
    || io.destination?.path !== destinationPath
    || published.destination?.path !== destinationPath
    || descriptorPath !== destinationPath) {
    throw createV1RuntimeError({
      message: "The committed C1 publication result violated the required RB-011 destination/effect binding.",
      scope: "artifact",
      path: destinationPath ?? null,
      option: "--plan-result",
      artifactRole: "project_artifact_set",
      details: { rule_id: "RB-011" }
    });
  }
}

function sameV1ApplyDigest(left, right) {
  return Boolean(left && right && left.algorithm === "sha-256" && right.algorithm === "sha-256" && left.value === right.value);
}

function assertV1ApplyResultTransport(resultTransport) {
  if (!resultTransport || !resultTransport.target || typeof resultTransport.writeResult !== "function") {
    throw new TypeError("runV1ApplyChange requires a reserved v1 result transport");
  }
}

function failedV1ApplyPreparation({ inputs, destination = null, error, errors = undefined, observations = undefined }) {
  if (!isCliV1Error(error)) throw new TypeError("failed v1 apply preparation requires a CliV1Error");
  return Object.freeze({
    inputs: copyV1ApplyPreparationInputs(inputs),
    destination: destination ? Object.freeze({ ...destination }) : null,
    observations: observations ?? { normalizations: [], losses: [], unsupported: [] },
    error,
    errors: errors ? Object.freeze([...errors]) : Object.freeze([error]),
    prepared: null
  });
}

function toV1ApplyDestinationError(error, approvedDestination) {
  if (isCliV1Error(error) && error.code === "publication.destination-exists") {
    return createV1RejectedError({
      code: "publication.reservation-conflict",
      message: "The approved destination is no longer available for exclusive creation.",
      scope: "filesystem",
      path: approvedDestination.path,
      option: "--plan-result",
      artifactRole: "project_artifact_set",
      details: { approved_path: approvedDestination.path }
    });
  }
  if (isCliV1Error(error) && error.code === "publication.destination-unsafe") {
    return createV1RejectedError({
      code: "change.binding-mismatch",
      message: "The approved destination parent no longer satisfies the planning-time path binding.",
      scope: "filesystem",
      path: error.location.path,
      option: "--plan-result",
      artifactRole: "plan_result",
      ruleId: "RB-005",
      details: { ...error.details, approved_path: approvedDestination.path }
    });
  }
  if (isCliV1Error(error) && error.status === "rejected") {
    return createV1RejectedError({
      code: error.code,
      message: error.message,
      scope: error.location.scope,
      path: error.location.path,
      option: "--plan-result",
      artifactRole: "project_artifact_set",
      ruleId: error.location.rule_id,
      details: { ...error.details, approved_path: approvedDestination.path }
    });
  }
  if (isCliV1Error(error) && error.status === "runtime-error") {
    return createV1RuntimeError({
      code: error.code,
      message: error.message,
      scope: error.location.scope,
      path: error.location.path,
      option: "--plan-result",
      artifactRole: "project_artifact_set",
      details: { ...error.details, approved_path: approvedDestination.path }
    });
  }
  return toV1ApplyPreparationError(error);
}

function toV1ApplyPreparationError(error) {
  if (isCliV1Error(error)) return error;
  return createV1RuntimeError({
    message: "The v1 apply preparation service encountered an unexpected internal error.",
    scope: "internal",
    details: { error_name: error instanceof Error ? error.name : typeof error }
  });
}

function v1ApplyPreparationInputObservations(decoded) {
  return { normalizations: [...decoded.normalizations], losses: [], unsupported: [] };
}

function v1ApplyPreparationSemanticObservations(decoded, validation) {
  return {
    normalizations: [...decoded.normalizations],
    losses: [],
    unsupported: validation.issues
      .filter((issue) => issue.code === "semantic.unsupported")
      .map((issue) => ({ code: issue.code, path: issue.path, description: issue.message }))
  };
}

function v1ApplyPreparationObservations(decoded, prepared) {
  return {
    normalizations: [...decoded.normalizations, ...prepared.output_plan.preflight.normalizations],
    losses: [...prepared.output_plan.preflight.losses],
    unsupported: [...prepared.output_plan.preflight.unsupported]
  };
}

function unreadV1ApplyPreparationInput(value, role, option, cwd) {
  return {
    role,
    option,
    source: value === "-" ? "stdin" : "file",
    path: value === "-" ? null : path.resolve(cwd, value),
    digest: null
  };
}

function copyV1ApplyPreparationInputs(inputs) {
  return inputs.map((input) => ({
    role: input.role,
    option: input.option,
    source: input.source,
    path: input.path,
    digest: input.digest ? { ...input.digest } : null
  }));
}

function assertV1ApplyPreparationInvocation(invocation) {
  if (!invocation || invocation.kind !== "workflow" || invocation.command !== "apply-change") {
    throw new TypeError("prepareV1ApplyChange requires a parsed v1 apply-change workflow invocation");
  }
}
