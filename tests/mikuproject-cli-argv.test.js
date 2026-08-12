import { describe, expect, it } from "vitest";

import { CliUsageError } from "../scripts/lib/cli-errors.mjs";
import {
  detectRequestedDiagnosticsFormat,
  parseArgs,
  summarizeCommandFromArgv
} from "../scripts/lib/cli-argv.mjs";

describe("legacy CLI argv boundary", () => {
  it("parses command tokens, flags, and the final repeated option value", () => {
    expect(parseArgs([
      "ai", "detect-kind", "--in", "input.json", "--diagnostics", "text", "--diagnostics", "json", "--help", "--version"
    ])).toEqual({
      command: ["ai", "detect-kind"],
      options: {
        in: "input.json",
        diagnostics: "json",
        help: true,
        version: true
      }
    });
  });

  it("keeps the missing option value error contract", () => {
    expect(() => parseArgs(["ai", "detect-kind", "--in"])).toThrow(CliUsageError);
    try {
      parseArgs(["ai", "detect-kind", "--in"]);
    } catch (error) {
      expect(error).toMatchObject({
        code: "missing_option_value",
        details: { option: "--in" }
      });
    }
  });

  it("detects requested JSON diagnostics and summarizes commands for error output", () => {
    expect(detectRequestedDiagnosticsFormat(["ai", "detect-kind", "--diagnostics", "json"])).toBe("json");
    expect(detectRequestedDiagnosticsFormat(["ai", "detect-kind", "--diagnostics", "text"])).toBe("text");
    expect(summarizeCommandFromArgv(["state", "diff", "--before", "before.json", "--after", "after.json"])).toBe("state diff");
    expect(summarizeCommandFromArgv(["--help"])).toBe("cli");
  });
});
