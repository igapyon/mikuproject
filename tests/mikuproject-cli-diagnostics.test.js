import { describe, expect, it } from "vitest";

import { CliUsageError } from "../scripts/lib/cli-errors.mjs";
import {
  buildCommandDiagnostics,
  buildErrorDiagnostics,
  determineStatus,
  formatValidationOutput,
  parseDiagnosticsFormat,
  summarizeChanges
} from "../scripts/lib/cli-diagnostics.mjs";

describe("legacy CLI diagnostics boundary", () => {
  it("keeps diagnostics option parsing and status precedence", () => {
    expect(parseDiagnosticsFormat()).toBe("text");
    expect(parseDiagnosticsFormat("json")).toBe("json");
    expect(() => parseDiagnosticsFormat("yaml")).toThrow(CliUsageError);
    expect(determineStatus({ ok: true, warnings: [], errors: [], changes_summary: { total_changes: 0 } })).toBe("noop");
    expect(determineStatus({ ok: true, warnings: [{ message: "warning" }], errors: [] })).toBe("warning");
    expect(determineStatus({ ok: false, warnings: [], errors: [] })).toBe("error");
  });

  it("keeps structured command and error diagnostics", () => {
    const changesSummary = summarizeChanges([{ scope: "tasks", uid: "1" }]);
    expect(buildCommandDiagnostics("state apply-patch", { changes_summary: changesSummary })).toMatchObject({
      ok: true,
      diagnostics_version: 1,
      status: "success",
      exit_code: 0,
      changes_summary: { total_changes: 1 }
    });
    const error = new CliUsageError("未対応のコマンドです: unknown", "unsupported_command");
    expect(buildErrorDiagnostics(["unknown", "--diagnostics", "json"], error, 2)).toMatchObject({
      ok: false,
      command: "unknown",
      error_type: "usage_error",
      error_code: "unsupported_command",
      io: { inputs: [], output: { mode: "stdout" } },
      errors: [{ code: "unsupported_command" }]
    });
  });

  it("keeps text and JSON validation report formatting", () => {
    const report = {
      ok: false,
      status: "error",
      warnings: [],
      errors: [{ message: "invalid patch" }],
      changes_summary: {
        total_changes: 0,
        by_scope: { project: 0, tasks: 0, resources: 0, assignments: 0, calendars: 0 }
      }
    };
    expect(formatValidationOutput(report, "text")).toContain("validate-patch ok=false status=error");
    expect(JSON.parse(formatValidationOutput(report, "json"))).toEqual(report);
  });
});
