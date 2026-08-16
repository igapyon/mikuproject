import { canonicalJsonText, canonicalizeSemanticState, sha256CanonicalJson, sha256SemanticState } from "./cli-v1-canonical-json.mjs";
import { createV1RejectedError } from "./cli-v1-errors.mjs";
import { encodeMsProjectXmlSubset, isV1XmlSemanticRoundTripEquivalent } from "./cli-v1-xml-encoder.mjs";
import { validateV1SemanticState } from "./cli-v1-semantic-validator.mjs";
import { validateArtifact, validateCliResult } from "../../generated/cli-v1-schema-validators.mjs";

/**
 * Validates and dry-runs the only C1 operation.  Returned planned state is
 * internal-only; the caller may expose the resulting diff and output plan but
 * never this whole state through the normal CLI result surface.
 */
export function planV1SetTaskPercentComplete({ state, changeRequest, runtime, destination } = {}) {
  assertVerifiedRuntime(runtime);
  assertChangeRequestArtifact(changeRequest);
  const baseState = canonicalizeSemanticState(state);
  const preValidation = validateV1SemanticState(baseState);
  if (!preValidation.valid) {
    throw createV1RejectedError({
      code: preValidation.status === "unsupported" ? "semantic.unsupported" : "semantic.invalid",
      message: "A change request cannot be planned from an invalid or unsupported semantic state.",
      scope: "semantic",
      path: "state",
      option: "--project",
      artifactRole: "external_project",
      details: { issue_count: preValidation.issues.length }
    });
  }
  const baseStateDigest = sha256SemanticState(baseState);
  const operation = changeRequest.operations[0];
  assertRequestPreconditions({ baseState, baseStateDigest, changeRequest, operation });

  const targetIndex = baseState.tasks.findIndex((task) => task.uid === operation.target.task_uid);
  if (targetIndex === -1) {
    throw changeRequestInvalid("Target task UID does not exist in the current semantic state.", {
      target_task_uid: operation.target.task_uid
    });
  }
  const target = baseState.tasks[targetIndex];
  if (target.summary !== false) {
    throw createV1RejectedError({
      code: "change.operation-unsupported",
      message: "set_task_percent_complete is supported only for a leaf task.",
      scope: "semantic",
      path: `tasks[uid=${target.uid}].summary`,
      option: "--request",
      artifactRole: "change_request",
      ruleId: "S-I022",
      details: { target_task_uid: target.uid }
    });
  }
  if (target.percent_complete !== operation.preconditions.expected_percent_complete) {
    throw createV1RejectedError({
      code: "change.precondition-failed",
      message: "Change request expected_percent_complete does not match the current target task value.",
      scope: "semantic",
      path: `tasks[uid=${target.uid}].percent_complete`,
      option: "--request",
      artifactRole: "change_request",
      details: {
        target_task_uid: target.uid,
        expected_percent_complete: operation.preconditions.expected_percent_complete,
        actual_percent_complete: target.percent_complete
      }
    });
  }
  if (target.percent_complete === operation.value.percent_complete) {
    throw createV1RejectedError({
      code: "change.no-op",
      message: "Change request must set percent_complete to a value different from the current value.",
      scope: "semantic",
      path: `tasks[uid=${target.uid}].percent_complete`,
      option: "--request",
      artifactRole: "change_request",
      details: { target_task_uid: target.uid, percent_complete: target.percent_complete }
    });
  }

  const plannedState = structuredClone(baseState);
  plannedState.tasks[targetIndex].percent_complete = operation.value.percent_complete;
  const postValidation = validateV1SemanticState(plannedState);
  if (!postValidation.valid) {
    throw changeRequestInvalid("Dry-run operation did not produce a valid semantic state.", {
      phase: "post-apply-validation",
      issue_count: postValidation.issues.length
    });
  }
  if (!isOnlyApprovedPercentChange(baseState, plannedState, target.uid, target.percent_complete, operation.value.percent_complete)) {
    throw new TypeError("C1 dry-run changed semantic data outside the approved target percent_complete field");
  }

  const changeRequestDigest = sha256CanonicalJson(changeRequest);
  const semanticDiff = {
    kind: "miku_project_semantic_diff",
    schema_version: "1",
    semantic_contract_version: "1",
    base_state_digest: baseStateDigest,
    proposed_state_digest: sha256SemanticState(plannedState),
    change_request_digest: changeRequestDigest,
    changes: [{
      kind: "set_task_percent_complete",
      task_uid: target.uid,
      before: target.percent_complete,
      after: operation.value.percent_complete
    }],
    preservation: { semantic_equivalent_except_changes: true },
    provenance: { losses: [], normalizations: [], unsupported: [] }
  };
  assertArtifact(semanticDiff, "semantic diff");

  const encoded = encodeMsProjectXmlSubset(plannedState);
  if (!isV1XmlSemanticRoundTripEquivalent(plannedState, encoded.bytes)) {
    throw new TypeError("C1 XML preflight encode/redecode did not preserve the planned semantic state");
  }
  const outputPlan = {
    kind: "miku_project_output_plan",
    schema_version: "1",
    semantic_contract_version: "1",
    base_state_digest: baseStateDigest,
    change_request_digest: changeRequestDigest,
    semantic_diff_digest: sha256CanonicalJson(semanticDiff),
    runtime: runtimeForOutputPlan(runtime),
    output: {
      format: "ms-project-xml",
      format_profile: "miku-project-ms-project-xml-subset/v1",
      adapter: "ms-project-xml-adapter/v1",
      artifact_set: "miku_project_artifact_set/v1",
      destination: { path: destination.path, write_mode: "create-new-directory" },
      publication: {
        strategy: "exclusive-directory-commit-marker/v1",
        directory_create: "exclusive",
        commit_marker: { path: "COMMITTED", create_mode: "exclusive-empty-file" },
        runtime_filesystem_supported: true
      },
      members: [
        { role: "external_project", path: "project.xml" },
        { role: "provenance", path: "provenance.json" },
        { role: "commit_marker", path: "COMMITTED", size: 0 }
      ]
    },
    preflight: {
      proposed_state_digest: sha256SemanticState(plannedState),
      project_artifact_digest: encoded.raw_digest,
      normalizations: [...encoded.normalizations],
      losses: [],
      unsupported: []
    }
  };
  assertArtifact(outputPlan, "output plan");
  if (!validateV1PlanChangeBindings({ changeRequest, semanticDiff, outputPlan, runtime, destination })) {
    throw new TypeError("C1 planner produced cross-artifact binding divergence");
  }
  return Object.freeze({
    semantic_diff: semanticDiff,
    output_plan: outputPlan,
    planned_state: canonicalizeSemanticState(plannedState),
    preflight_project_xml: Buffer.from(encoded.bytes)
  });
}

/** Implements the planning subset of RB-001 through RB-005. */
export function validateV1PlanChangeBindings({ changeRequest, semanticDiff, outputPlan, runtime, destination } = {}) {
  return findV1PlanChangeBindingMismatchRuleId({
    changeRequest,
    semanticDiff,
    outputPlan,
    runtime,
    destination
  }) === null;
}

/** Returns the first failed RB-001..RB-005 rule, or null when all pass. */
export function findV1PlanChangeBindingMismatchRuleId({ changeRequest, semanticDiff, outputPlan, runtime, destination } = {}) {
  try {
    if (!validateArtifact(changeRequest) || !validateArtifact(semanticDiff) || !validateArtifact(outputPlan)) return "artifact-schema";
    const requestDigest = sha256CanonicalJson(changeRequest);
    const diffDigest = sha256CanonicalJson(semanticDiff);
    if (!sameDigest(semanticDiff.base_state_digest, changeRequest.base.state_digest)
      || !sameDigest(outputPlan.base_state_digest, changeRequest.base.state_digest)) return "RB-001";
    if (!sameDigest(semanticDiff.change_request_digest, requestDigest)
      || !sameDigest(outputPlan.change_request_digest, requestDigest)) return "RB-002";
    if (!sameDigest(outputPlan.semantic_diff_digest, diffDigest)
      || !sameDigest(outputPlan.preflight.proposed_state_digest, semanticDiff.proposed_state_digest)) return "RB-003";
    if (runtime?.binding_status !== "verified"
      || canonicalJsonText(outputPlan.runtime) !== canonicalJsonText(runtimeForOutputPlan(runtime))) return "RB-004";
    if (!destination || outputPlan.output.destination.path !== destination.path) return "RB-005";
    return null;
  } catch {
    return "artifact-schema";
  }
}

/**
 * Returns the destination carried by a successful planning result only after
 * validating the result envelope and its runtime/path bindings. apply-change
 * deliberately has no independent destination option.
 */
export function getV1ApprovedDestinationFromPlanResult(planResult, runtime) {
  assertV1PlanChangeResultEnvelope(planResult, runtime);
  return Object.freeze({
    requested_path: planResult.io.destination.requested_path,
    path: planResult.data.output_plan.output.destination.path
  });
}

/** Implements the pure cross-artifact part of RB-001 through RB-006. */
export function validateV1ApplyPreparationBindings({
  changeRequest,
  planResult,
  approval,
  runtime,
  destination
} = {}) {
  try {
    assertV1PlanChangeResultEnvelope(planResult, runtime);
    assertV1ChangeApprovalArtifact(approval);
    assertV1ApplyPreparationBindings({ changeRequest, planResult, approval, runtime, destination });
    return true;
  } catch {
    return false;
  }
}

/**
 * Recomputes the approved C1 plan from the current state and returns internal
 * apply material only when the request, result, approval, runtime and
 * destination remain exactly bound. No filesystem publication occurs here.
 */
export function prepareV1ApprovedChange({
  state,
  changeRequest,
  planResult,
  approval,
  runtime,
  destination
} = {}) {
  assertV1PlanChangeResultEnvelope(planResult, runtime);
  assertV1ChangeApprovalArtifact(approval);
  assertV1ApplyPreparationBindings({ changeRequest, planResult, approval, runtime, destination });

  const regenerated = planV1SetTaskPercentComplete({ state, changeRequest, runtime, destination });
  if (canonicalJsonText(regenerated.semantic_diff) !== canonicalJsonText(planResult.data.semantic_diff)
    || canonicalJsonText(regenerated.output_plan) !== canonicalJsonText(planResult.data.output_plan)) {
    throw changeBindingMismatch("The planning result does not match a fresh plan from the current project and request.", {
      rule_id: "RB-001..RB-006"
    });
  }
  if (regenerated.output_plan.preflight.losses.length !== 0
    || regenerated.output_plan.preflight.unsupported.length !== 0) {
    throw changeBindingMismatch("An approved v1 change must not contain loss or unsupported output observations.", {
      rule_id: "RB-006"
    });
  }
  return Object.freeze({
    semantic_diff: regenerated.semantic_diff,
    output_plan: regenerated.output_plan,
    approval: structuredClone(approval),
    base_state: canonicalizeSemanticState(state),
    change_request: structuredClone(changeRequest),
    runtime: runtimeForOutputPlan(runtime),
    planned_state: regenerated.planned_state,
    preflight_project_xml: Buffer.from(regenerated.preflight_project_xml)
  });
}

function assertRequestPreconditions({ baseStateDigest, changeRequest }) {
  if (!sameDigest(changeRequest.base.state_digest, baseStateDigest)) {
    throw createV1RejectedError({
      code: "change.precondition-failed",
      message: "Change request base state digest does not match the current semantic state.",
      scope: "semantic",
      path: "base.state_digest",
      option: "--request",
      artifactRole: "change_request",
      details: {
        expected_state_digest: baseStateDigest.value,
        request_state_digest: changeRequest.base.state_digest?.value ?? null
      }
    });
  }
}

function assertChangeRequestArtifact(changeRequest) {
  if (!changeRequest || typeof changeRequest !== "object") {
    throw changeRequestInvalid("Change request must be a JSON object.");
  }
  if (changeRequest.kind !== "miku_project_change_request") {
    throw createV1RejectedError({
      code: "artifact.kind-unsupported",
      message: "The --request artifact kind is not miku_project_change_request.",
      scope: "artifact",
      path: "/kind",
      option: "--request",
      artifactRole: "change_request",
      details: { received_kind: changeRequest.kind ?? null }
    });
  }
  if (changeRequest.schema_version !== "1" || changeRequest.semantic_contract_version !== "1") {
    throw createV1RejectedError({
      code: "artifact.schema-version-unsupported",
      message: "The --request artifact does not use schema and semantic contract version 1.",
      scope: "artifact",
      path: "/schema_version",
      option: "--request",
      artifactRole: "change_request",
      details: {
        schema_version: changeRequest.schema_version ?? null,
        semantic_contract_version: changeRequest.semantic_contract_version ?? null
      }
    });
  }
  const requestedOperation = changeRequest.operations?.[0];
  if (requestedOperation?.kind !== undefined && requestedOperation.kind !== "set_task_percent_complete") {
    throw createV1RejectedError({
      code: "change.operation-unsupported",
      message: "The requested change operation is outside the C1 allowlist.",
      scope: "artifact",
      path: "/operations/0/kind",
      option: "--request",
      artifactRole: "change_request",
      details: { received_kind: requestedOperation.kind }
    });
  }
  if (!validateArtifact(changeRequest)) {
    throw changeRequestInvalid("The --request artifact does not satisfy miku_project_change_request/v1.");
  }
}

function assertV1PlanChangeResultEnvelope(planResult, runtime) {
  if (!planResult || typeof planResult !== "object" || !validateCliResult(planResult)
    || planResult.command !== "plan-change" || planResult.status !== "succeeded"
    || !planResult.data?.semantic_diff || !planResult.data?.output_plan) {
    throw changeBindingMismatch("The --plan-result input must be a successful v1 plan-change result envelope.", {
      rule_id: null
    }, "/", null);
  }
  if (!runtime || typeof runtime !== "object"
    || canonicalJsonText(planResult.runtime) !== canonicalJsonText(runtime)) {
    throw changeBindingMismatch("The --plan-result runtime binding does not match the current runtime.", {
      rule_id: "RB-004"
    }, "/runtime", "RB-004");
  }
  if (planResult.io?.destination?.path !== planResult.data.output_plan.output?.destination?.path) {
    throw changeBindingMismatch("The --plan-result destination binding is inconsistent.", {
      rule_id: "RB-005"
    }, "/io/destination/path", "RB-005");
  }
  if (planResult.observations.losses.length !== 0
    || planResult.observations.unsupported.length !== 0
    || planResult.data.output_plan.preflight.losses.length !== 0
    || planResult.data.output_plan.preflight.unsupported.length !== 0) {
    throw changeBindingMismatch("The --plan-result input contains loss or unsupported observations.", {
      rule_id: "RB-006"
    }, "/observations", "RB-006");
  }
}

function assertV1ChangeApprovalArtifact(approval) {
  if (!approval || typeof approval !== "object"
    || approval.kind !== "miku_project_change_approval"
    || !validateArtifact(approval)) {
    throw createV1RejectedError({
      code: "change.approval-invalid",
      message: "The --approval input must satisfy miku_project_change_approval/v1 and explicitly approve the change.",
      scope: "artifact",
      path: "/",
      option: "--approval",
      artifactRole: "approval",
      details: {}
    });
  }
}

function assertV1ApplyPreparationBindings({ changeRequest, planResult, approval, runtime, destination }) {
  const semanticDiff = planResult.data.semantic_diff;
  const outputPlan = planResult.data.output_plan;
  const mismatchRuleId = findV1PlanChangeBindingMismatchRuleId({ changeRequest, semanticDiff, outputPlan, runtime, destination });
  if (mismatchRuleId !== null) {
    throw changeBindingMismatch("The request, planning result, runtime, or destination binding does not match.", {
      rule_id: mismatchRuleId
    }, null, mismatchRuleId === "artifact-schema" ? null : mismatchRuleId);
  }
  if (!sameDigest(approval.base_state_digest, semanticDiff.base_state_digest)
    || !sameDigest(approval.change_request_digest, sha256CanonicalJson(changeRequest))
    || !sameDigest(approval.semantic_diff_digest, sha256CanonicalJson(semanticDiff))
    || !sameDigest(approval.output_plan_digest, sha256CanonicalJson(outputPlan))) {
    throw createV1RejectedError({
      code: "change.binding-mismatch",
      message: "The --approval digests do not match the current request and approved planning result.",
      scope: "artifact",
      path: "/",
      option: "--approval",
      artifactRole: "approval",
      ruleId: "RB-006",
      details: { rule_id: "RB-006" }
    });
  }
}

function assertVerifiedRuntime(runtime) {
  if (!runtime || runtime.binding_status !== "verified" || runtime.capability_profile !== "miku-project-cli-core/v1") {
    throw new TypeError("C1 planning requires a verified miku-project-cli-core/v1 runtime binding");
  }
}

function runtimeForOutputPlan(runtime) {
  return {
    family: runtime.family,
    version: runtime.version,
    artifact_digest: { ...runtime.artifact_digest },
    manifest_digest: { ...runtime.manifest_digest },
    capability_profile: runtime.capability_profile,
    fixture_suite_version: runtime.fixture_suite_version
  };
}

function isOnlyApprovedPercentChange(beforeState, afterState, targetUid, beforeValue, afterValue) {
  const maskedAfter = canonicalizeSemanticState(afterState);
  const target = maskedAfter.tasks.find((task) => task.uid === targetUid);
  if (!target || target.percent_complete !== afterValue) return false;
  target.percent_complete = beforeValue;
  return canonicalJsonText(canonicalizeSemanticState(beforeState)) === canonicalJsonText(maskedAfter);
}

function sameDigest(left, right) {
  return Boolean(left && right && left.algorithm === "sha-256" && right.algorithm === "sha-256" && left.value === right.value);
}

function assertArtifact(value, name) {
  if (!validateArtifact(value)) throw new TypeError(`C1 ${name} does not satisfy miku_project_artifacts/v1`);
}

function changeRequestInvalid(message, details = {}) {
  return createV1RejectedError({
    code: "change.request-invalid",
    message,
    scope: "artifact",
    path: null,
    option: "--request",
    artifactRole: "change_request",
    details
  });
}

function changeBindingMismatch(message, details = {}, bindingPath = null, ruleId = "RB-006") {
  return createV1RejectedError({
    code: "change.binding-mismatch",
    message,
    scope: "artifact",
    path: bindingPath,
    option: "--plan-result",
    artifactRole: "plan_result",
    ruleId,
    details
  });
}
