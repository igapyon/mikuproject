export class CliV1Error extends Error {
  constructor({ code, status, message, location, details = {} }) {
    super(message);
    this.name = "CliV1Error";
    this.code = code;
    this.status = status;
    this.location = Object.freeze({
      scope: location?.scope ?? "internal",
      path: location?.path ?? null,
      option: location?.option ?? null,
      artifact_role: location?.artifact_role ?? null,
      rule_id: location?.rule_id ?? null
    });
    this.details = Object.freeze({ ...details });
  }
}

export function createV1UsageError({ code, message, option = null, details = {}, command = null }) {
  return new CliV1Error({
    code,
    status: "usage-error",
    message,
    location: {
      scope: option ? "option" : "command",
      path: null,
      option,
      artifact_role: null,
      rule_id: null
    },
    details: command ? { command, ...details } : details
  });
}

export function createV1IoError({ code, status, message, path = null, option = null, details = {} }) {
  return new CliV1Error({
    code,
    status,
    message,
    location: {
      scope: "filesystem",
      path,
      option,
      artifact_role: null,
      rule_id: null
    },
    details
  });
}

/**
 * Creates a rejected v1 error for a checked input/domain condition.  This is
 * deliberately separate from usage and filesystem failures: XML/profile and
 * semantic validation have a stable diagnostic code and a semantic rule/path
 * even before a public workflow command is wired to the legacy entrypoint.
 */
export function createV1RejectedError({
  code,
  message,
  scope = "input",
  path = null,
  option = "--project",
  artifactRole = "external_project",
  ruleId = null,
  details = {}
}) {
  return new CliV1Error({
    code,
    status: "rejected",
    message,
    location: {
      scope,
      path,
      option,
      artifact_role: artifactRole,
      rule_id: ruleId
    },
    details
  });
}

/** Creates a schema-addressable runtime failure without exposing raw errors. */
export function createV1RuntimeError({
  code = "internal.unexpected-error",
  message,
  scope = "internal",
  path = null,
  option = null,
  artifactRole = null,
  details = {}
}) {
  return new CliV1Error({
    code,
    status: "runtime-error",
    message,
    location: {
      scope,
      path,
      option,
      artifact_role: artifactRole,
      rule_id: null
    },
    details
  });
}

export function isCliV1Error(error) {
  return error instanceof CliV1Error;
}
