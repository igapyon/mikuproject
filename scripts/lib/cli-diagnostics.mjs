import { summarizeCommandFromArgv } from "./cli-argv.mjs";
import {
  CliUsageError,
  extractErrorDetails,
  inferErrorCode
} from "./cli-errors.mjs";
import { buildIoDiagnosticsFromArgv } from "./cli-io.mjs";

export const DIAGNOSTICS_VERSION = 1;

export function parseDiagnosticsFormat(value) {
  if (value === undefined) {
    return "text";
  }
  if (value === "text" || value === "json") {
    return value;
  }
  throw new CliUsageError("--diagnostics には text または json を指定してください", "invalid_diagnostics_option", {
    option: "--diagnostics",
    expected_values: ["text", "json"]
  });
}

export function buildPatchDiagnosticsText(result, contextLabel) {
  const lines = [];
  const warningCount = Array.isArray(result.warnings) ? result.warnings.length : 0;
  const changeCount = Array.isArray(result.changes) ? result.changes.length : 0;
  const status = determineStatus({
    ok: true,
    warnings: result.warnings || [],
    errors: [],
    changes_summary: summarizeChanges(result.changes || [])
  });

  lines.push(`[miku-project-cli] ${contextLabel} patch_json status=${status} changes=${changeCount} warnings=${warningCount}`);
  for (const warning of result.warnings || []) {
    lines.push(`[warning] ${formatWarning(warning)}`);
  }
  return lines;
}

export function buildPatchDiagnosticsJson(result, contextLabel, io) {
  return buildCommandDiagnostics(contextLabel, {
    io,
    warnings: result.warnings || [],
    errors: [],
    changes_summary: summarizeChanges(result.changes || [])
  });
}

export function summarizeChanges(changes) {
  const byScope = {
    project: 0,
    tasks: 0,
    resources: 0,
    assignments: 0,
    calendars: 0
  };
  const affectedItems = {
    project: new Set(),
    tasks: new Set(),
    resources: new Set(),
    assignments: new Set(),
    calendars: new Set()
  };

  for (const change of changes) {
    if (!Object.hasOwn(byScope, change.scope)) {
      continue;
    }
    byScope[change.scope] += 1;
    affectedItems[change.scope].add(change.uid);
  }

  return {
    total_changes: changes.length,
    by_scope: byScope,
    affected_items: {
      project: affectedItems.project.size,
      tasks: affectedItems.tasks.size,
      resources: affectedItems.resources.size,
      assignments: affectedItems.assignments.size,
      calendars: affectedItems.calendars.size
    }
  };
}

export function buildCommandDiagnostics(context, extra = {}) {
  const warnings = Array.isArray(extra.warnings) ? extra.warnings : [];
  const errors = Array.isArray(extra.errors) ? extra.errors : [];
  const status = typeof extra.status === "string"
    ? extra.status
    : determineStatus({
      ok: extra.ok !== false,
      warnings,
      errors,
      changes_summary: extra.changes_summary
    });
  const exitCode = typeof extra.exit_code === "number"
    ? extra.exit_code
    : determineExitCodeFromStatus(status);
  return {
    ok: extra.ok !== false,
    diagnostics_version: DIAGNOSTICS_VERSION,
    command: context,
    context,
    status,
    exit_code: exitCode,
    warning_count: warnings.length,
    error_count: errors.length,
    warnings,
    errors,
    ...extra
  };
}

export function determineExitCodeFromStatus(status) {
  if (status === "error") {
    return 1;
  }
  return 0;
}

export function determineStatus(input) {
  if (!input.ok || (input.errors || []).length > 0) {
    return "error";
  }
  if ((input.warnings || []).length > 0) {
    return "warning";
  }
  if (input.changes_summary && input.changes_summary.total_changes === 0) {
    return "noop";
  }
  return "success";
}

export function buildChangedItems(changes) {
  const grouped = {
    project: [],
    tasks: [],
    resources: [],
    assignments: [],
    calendars: []
  };

  for (const change of changes) {
    if (!Object.hasOwn(grouped, change.scope)) {
      continue;
    }
    grouped[change.scope].push({
      uid: change.uid,
      label: change.label,
      field: change.field,
      before: change.before,
      after: change.after
    });
  }

  return grouped;
}

export function formatValidationOutput(report, diagnosticsFormat) {
  if (diagnosticsFormat === "json") {
    return `${JSON.stringify(report, null, 2)}\n`;
  }

  const lines = [
    `[miku-project-cli] validate-patch ok=${report.ok ? "true" : "false"} status=${report.status} warnings=${report.warnings.length} errors=${report.errors.length} changes=${report.changes_summary.total_changes}`
  ];
  for (const error of report.errors) {
    lines.push(`[error] ${error.message}`);
  }
  for (const warning of report.warnings) {
    lines.push(`[warning] ${formatWarning(warning)}`);
  }
  lines.push(`[changes] project=${report.changes_summary.by_scope.project} tasks=${report.changes_summary.by_scope.tasks} resources=${report.changes_summary.by_scope.resources} assignments=${report.changes_summary.by_scope.assignments} calendars=${report.changes_summary.by_scope.calendars}`);
  return `${lines.join("\n")}\n`;
}

export function formatWarning(warning) {
  const fragments = [warning.message];
  if (warning.scope) {
    fragments.push(`scope=${warning.scope}`);
  }
  if (warning.uid) {
    fragments.push(`uid=${warning.uid}`);
  }
  if (warning.label) {
    fragments.push(`label=${warning.label}`);
  }
  return fragments.join(" ");
}

export function writeDiagnostics(stream, diagnostics) {
  if (!Array.isArray(diagnostics)) {
    stream.write(`${JSON.stringify(diagnostics, null, 2)}\n`);
    return;
  }
  for (const line of diagnostics) {
    stream.write(`${line}\n`);
  }
}

export function buildErrorDiagnostics(argv, error, exitCode) {
  const context = summarizeCommandFromArgv(argv);
  const errorCode = inferErrorCode(error, context);
  return {
    ok: false,
    diagnostics_version: DIAGNOSTICS_VERSION,
    command: context,
    context,
    status: "error",
    exit_code: exitCode,
    warning_count: 0,
    error_count: 1,
    io: buildIoDiagnosticsFromArgv(argv),
    error_type: error instanceof CliUsageError ? "usage_error" : "processing_error",
    error_code: errorCode,
    error_details: extractErrorDetails(error),
    warnings: [],
    errors: [buildErrorItem(error, context)]
  };
}

export function buildErrorItem(error, context) {
  const code = inferErrorCode(error, context);
  const item = {
    code,
    message: error instanceof Error ? error.message : String(error)
  };
  const details = extractErrorDetails(error);
  if (details) {
    item.details = details;
  }
  return item;
}
