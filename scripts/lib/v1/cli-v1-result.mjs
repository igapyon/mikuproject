import { canonicalJsonText, compareUnicodeScalars } from "./cli-v1-canonical-json.mjs";
import { isCliV1Error } from "./cli-v1-errors.mjs";
import { validateCliDiagnostic, validateCliResult } from "../../generated/cli-v1-schema-validators.mjs";

const STATUS_EXIT_CODES = Object.freeze({
  succeeded: 0,
  rejected: 1,
  "usage-error": 2,
  "runtime-error": 3
});

const COMMAND_SIDE_EFFECT_CLASSES = Object.freeze({
  cli: "none",
  inspect: "read-only",
  validate: "read-only",
  "plan-change": "exchange-artifact-generation",
  "apply-change": "meaning-change-and-project-artifact-generation",
  "verify-artifact": "read-only"
});

const ERROR_METADATA = Object.freeze({
  "cli.unknown-command": Object.freeze({ category: "usage", retryability: "after-input-change" }),
  "cli.unknown-option": Object.freeze({ category: "usage", retryability: "after-input-change" }),
  "cli.missing-option": Object.freeze({ category: "usage", retryability: "after-input-change" }),
  "cli.duplicate-option": Object.freeze({ category: "usage", retryability: "after-input-change" }),
  "cli.unexpected-argument": Object.freeze({ category: "usage", retryability: "after-input-change" }),
  "cli.invalid-option-value": Object.freeze({ category: "usage", retryability: "after-input-change" }),
  "cli.multiple-stdin-sources": Object.freeze({ category: "usage", retryability: "after-input-change" }),
  "io.input-not-found": Object.freeze({ category: "io", retryability: "after-input-change" }),
  "io.input-type-invalid": Object.freeze({ category: "io", retryability: "after-input-change" }),
  "io.input-symlink-rejected": Object.freeze({ category: "io", retryability: "after-input-change" }),
  "io.input-read-failed": Object.freeze({ category: "io", retryability: "after-environment-change" }),
  "io.result-path-exists": Object.freeze({ category: "io", retryability: "after-input-change" }),
  "io.result-path-unsafe": Object.freeze({ category: "io", retryability: "after-input-change" }),
  "io.result-reservation-failed": Object.freeze({ category: "io", retryability: "after-environment-change" }),
  "text.invalid-utf8": Object.freeze({ category: "encoding", retryability: "after-input-change" }),
  "json.invalid": Object.freeze({ category: "json", retryability: "after-input-change" }),
  "json.bom-not-allowed": Object.freeze({ category: "json", retryability: "after-input-change" }),
  "json.duplicate-key": Object.freeze({ category: "json", retryability: "after-input-change" }),
  "artifact.kind-unsupported": Object.freeze({ category: "artifact", retryability: "after-input-change" }),
  "artifact.schema-version-unsupported": Object.freeze({ category: "artifact", retryability: "after-input-change" }),
  "xml.invalid": Object.freeze({ category: "xml", retryability: "after-input-change" }),
  "xml.encoding-unsupported": Object.freeze({ category: "xml", retryability: "after-input-change" }),
  "xml.profile-unsupported": Object.freeze({ category: "xml", retryability: "after-input-change" }),
  "semantic.invalid": Object.freeze({ category: "semantic", retryability: "after-input-change" }),
  "semantic.unsupported": Object.freeze({ category: "semantic", retryability: "after-input-change" }),
  "change.request-invalid": Object.freeze({ category: "change", retryability: "after-input-change" }),
  "change.operation-unsupported": Object.freeze({ category: "change", retryability: "after-input-change" }),
  "change.precondition-failed": Object.freeze({ category: "change", retryability: "after-replan-and-approval" }),
  "change.no-op": Object.freeze({ category: "change", retryability: "after-input-change" }),
  "change.binding-mismatch": Object.freeze({ category: "change", retryability: "after-replan-and-approval" }),
  "change.approval-invalid": Object.freeze({ category: "change", retryability: "after-replan-and-approval" }),
  "publication.destination-exists": Object.freeze({ category: "publication", retryability: "after-replan-and-approval" }),
  "publication.destination-unsafe": Object.freeze({ category: "publication", retryability: "after-input-change" }),
  "publication.capability-unsupported": Object.freeze({ category: "publication", retryability: "after-environment-change" }),
  "publication.reservation-conflict": Object.freeze({ category: "publication", retryability: "after-replan-and-approval" }),
  "publication.write-failed": Object.freeze({ category: "publication", retryability: "after-environment-change" }),
  "publication.postwrite-verification-failed": Object.freeze({ category: "publication", retryability: "not-retryable" }),
  "publication.cleanup-failed": Object.freeze({ category: "publication", retryability: "not-retryable" }),
  "publication.artifact-absent": Object.freeze({ category: "publication", retryability: "not-retryable" }),
  "publication.artifact-incomplete": Object.freeze({ category: "publication", retryability: "not-retryable" }),
  "publication.artifact-corrupt": Object.freeze({ category: "publication", retryability: "not-retryable" }),
  "publication.expected-plan-mismatch": Object.freeze({ category: "publication", retryability: "after-replan-and-approval" }),
  "runtime.manifest-invalid": Object.freeze({ category: "runtime", retryability: "not-retryable" }),
  "runtime.artifact-digest-mismatch": Object.freeze({ category: "runtime", retryability: "not-retryable" }),
  "runtime.capability-missing": Object.freeze({ category: "runtime", retryability: "after-environment-change" }),
  "internal.unexpected-error": Object.freeze({ category: "internal", retryability: "not-retryable" })
});

export const V1_CONTRACT_DESCRIPTOR = Object.freeze({
  product: "miku-project",
  product_contract_version: "1",
  artifact_schema: "miku_project_artifacts/v1",
  result_schema: "miku_project_cli_result/v1",
  diagnostic_schema: "miku_project_cli_diagnostic/v1",
  diagnostic_catalog_version: "1"
});

export function createUnverifiedRuntimeBinding({ family = "node", version } = {}) {
  if ((family !== "node" && family !== "java") || typeof version !== "string" || version.length === 0) {
    throw new TypeError("an unverified v1 runtime binding requires family node/java and a version");
  }
  return {
    binding_status: "unverified",
    family,
    version,
    artifact_digest: null,
    manifest_digest: null,
    capability_profile: null,
    fixture_suite_version: null
  };
}

export function createV1DiagnosticFromError(error) {
  if (!isCliV1Error(error)) {
    throw new TypeError("v1 diagnostics require a CliV1Error with an explicit stable code");
  }
  const metadata = ERROR_METADATA[error.code];
  if (!metadata) {
    throw new TypeError(`v1 diagnostic metadata is missing for ${error.code}`);
  }
  const diagnostic = {
    kind: "miku_project_cli_diagnostic",
    schema_version: "1",
    code: error.code,
    severity: "error",
    category: metadata.category,
    message: error.message,
    location: { ...error.location },
    retryability: metadata.retryability,
    details: { ...error.details }
  };
  assertDiagnostic(diagnostic);
  return diagnostic;
}

export function createV1ErrorResult({
  error,
  command = "cli",
  runtime,
  resultTarget = { target: "stdout", path: null },
  io = undefined,
  data = null
} = {}) {
  if (!isCliV1Error(error)) {
    throw new TypeError("v1 error results require a CliV1Error");
  }
  return createV1Result({
    command,
    runtime,
    status: error.status,
    io: io ?? createEmptyIo(resultTarget),
    diagnostics: [createV1DiagnosticFromError(error)],
    data
  });
}

export function createV1Result({
  command,
  runtime,
  status,
  io,
  diagnostics = [],
  data = null,
  effects = undefined,
  observations = undefined
} = {}) {
  if (!Object.hasOwn(COMMAND_SIDE_EFFECT_CLASSES, command)) {
    throw new TypeError(`unsupported v1 result command: ${String(command)}`);
  }
  if (!Object.hasOwn(STATUS_EXIT_CODES, status)) {
    throw new TypeError(`unsupported v1 result status: ${String(status)}`);
  }
  if (!runtime || typeof runtime !== "object") {
    throw new TypeError("v1 result runtime binding is required");
  }
  if (!io || typeof io !== "object") {
    throw new TypeError("v1 result I/O metadata is required");
  }

  const sortedDiagnostics = sortDiagnostics(diagnostics);
  if (status === "succeeded" && sortedDiagnostics.length !== 0) {
    throw new TypeError("a succeeded v1 result must not contain diagnostics");
  }
  if (status !== "succeeded" && sortedDiagnostics.length === 0) {
    throw new TypeError("a non-success v1 result must contain at least one diagnostic");
  }

  const result = {
    kind: "miku_project_cli_result",
    schema_version: "1",
    contract: { ...V1_CONTRACT_DESCRIPTOR },
    runtime: { ...runtime },
    command,
    side_effect_class: COMMAND_SIDE_EFFECT_CLASSES[command],
    status,
    exit_code: STATUS_EXIT_CODES[status],
    io: copyIo(io),
    effects: effects ? copyEffects(effects) : defaultEffects(),
    observations: observations ? copyObservations(observations) : defaultObservations(),
    next_action: deriveV1NextAction({ command, status, diagnostics: sortedDiagnostics }),
    diagnostics: sortedDiagnostics,
    data
  };
  assertResult(result);
  return result;
}

export function deriveV1NextAction({ command, status, diagnostics = [] }) {
  if (status === "succeeded") {
    if (command === "plan-change") {
      return { action: "request-human-approval", command: null, source_retryability: null };
    }
    if (command === "apply-change") {
      return { action: "verify-artifact", command: "verify-artifact", source_retryability: null };
    }
    return { action: "complete", command: null, source_retryability: null };
  }

  const retryability = mostConservativeRetryability(diagnostics);
  switch (retryability) {
    case "after-input-change":
      return { action: "revise-invocation-or-input", command: null, source_retryability: retryability };
    case "after-environment-change":
      return { action: "repair-environment", command: null, source_retryability: retryability };
    case "after-replan-and-approval":
      return { action: "replan-and-request-human-approval", command: "plan-change", source_retryability: retryability };
    case "not-retryable":
      return { action: "abort-and-investigate", command: null, source_retryability: retryability };
    default:
      throw new TypeError("a non-success v1 result requires diagnostics with a supported retryability");
  }
}

export function serializeV1Result(result) {
  assertResult(result);
  return `${canonicalJsonText(result)}\n`;
}

function createEmptyIo(resultTarget) {
  return {
    stdin_option: null,
    inputs: [],
    result: { target: resultTarget.target, path: resultTarget.path },
    destination: null
  };
}

function defaultEffects() {
  return {
    project_input_modified: false,
    project_artifact: null,
    cleanup: { status: "not-needed", path: null }
  };
}

function defaultObservations() {
  return { normalizations: [], losses: [], unsupported: [] };
}

function copyIo(io) {
  return {
    stdin_option: io.stdin_option ?? null,
    inputs: Array.isArray(io.inputs) ? io.inputs.map((input) => ({
      ...input,
      digest: input.digest ? { ...input.digest } : null
    })) : [],
    result: { ...io.result },
    destination: io.destination ? { ...io.destination } : null
  };
}

function copyEffects(effects) {
  return {
    project_input_modified: effects.project_input_modified,
    project_artifact: effects.project_artifact ? { ...effects.project_artifact } : null,
    cleanup: { ...effects.cleanup }
  };
}

function copyObservations(observations) {
  return {
    normalizations: sortObservations(observations.normalizations ?? []).map((item) => ({ ...item })),
    losses: sortObservations(observations.losses ?? []).map((item) => ({ ...item })),
    unsupported: sortObservations(observations.unsupported ?? []).map((item) => ({ ...item }))
  };
}

function sortDiagnostics(diagnostics) {
  if (!Array.isArray(diagnostics)) {
    throw new TypeError("v1 diagnostics must be an array");
  }
  return diagnostics.map((diagnostic) => {
    assertDiagnostic(diagnostic);
    return {
      ...diagnostic,
      location: { ...diagnostic.location },
      details: { ...diagnostic.details }
    };
  }).sort((left, right) => compareDiagnostic(left, right));
}

function sortObservations(items) {
  if (!Array.isArray(items)) {
    throw new TypeError("v1 observations must be arrays");
  }
  return [...items].sort((left, right) => compareUnicodeScalars(left.code, right.code)
    || compareUnicodeScalars(left.path, right.path));
}

function compareDiagnostic(left, right) {
  return compareUnicodeScalars(left.code, right.code)
    || compareUnicodeScalars(left.location.scope, right.location.scope)
    || compareNullableUnicodeScalars(left.location.path, right.location.path)
    || compareNullableUnicodeScalars(left.location.option, right.location.option);
}

function compareNullableUnicodeScalars(left, right) {
  if (left === null && right === null) {
    return 0;
  }
  if (left === null) {
    return -1;
  }
  if (right === null) {
    return 1;
  }
  return compareUnicodeScalars(left, right);
}

function mostConservativeRetryability(diagnostics) {
  const priority = Object.freeze({
    "after-input-change": 1,
    "after-environment-change": 2,
    "after-replan-and-approval": 3,
    "not-retryable": 4
  });
  let selected = null;
  for (const diagnostic of diagnostics) {
    if (!Object.hasOwn(priority, diagnostic.retryability)) {
      throw new TypeError(`unsupported v1 retryability: ${String(diagnostic.retryability)}`);
    }
    if (selected === null || priority[diagnostic.retryability] > priority[selected]) {
      selected = diagnostic.retryability;
    }
  }
  return selected;
}

function assertDiagnostic(diagnostic) {
  if (!validateCliDiagnostic(diagnostic)) {
    throw new TypeError(`v1 diagnostic does not satisfy its schema: ${formatSchemaErrors(validateCliDiagnostic.errors)}`);
  }
}

function assertResult(result) {
  if (!validateCliResult(result)) {
    throw new TypeError(`v1 result does not satisfy its schema: ${formatSchemaErrors(validateCliResult.errors)}`);
  }
}

function formatSchemaErrors(errors) {
  return Array.isArray(errors)
    ? errors.map((error) => `${error.instancePath || "/"} ${error.keyword}`).join(", ")
    : "unknown schema error";
}
