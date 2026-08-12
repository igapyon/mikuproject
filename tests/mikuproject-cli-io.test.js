import { describe, expect, it } from "vitest";

import { CliUsageError } from "../scripts/lib/cli-errors.mjs";
import {
  buildIoDiagnostics,
  buildIoDiagnosticsFromArgv,
  decodeBase64Input,
  ensureBinaryInputSource,
  ensureBinaryOutputTarget,
  ensureSingleStdinSource
} from "../scripts/lib/cli-io.mjs";

describe("legacy CLI I/O boundary", () => {
  it("keeps Base64 decoding and binary target validation", () => {
    expect(decodeBase64Input("aGVsbG8=\n", "test").toString("utf8")).toBe("hello");
    expect(() => decodeBase64Input("not base64", "test")).toThrow(CliUsageError);
    expect(() => ensureBinaryInputSource({ in: "input.xlsx", "in-base64": "-" }, "import xlsx")).toThrow(/同時に指定/);
    expect(() => ensureBinaryOutputTarget({}, "export xlsx")).toThrow(/--out/);
  });

  it("keeps stdin conflict detection and diagnostics I/O descriptions", () => {
    expect(() => ensureSingleStdinSource([
      { optionName: "--state", value: "-", allowImplicitStdin: false },
      { optionName: "--in", value: undefined, allowImplicitStdin: true }
    ])).toThrow(/1 つだけ/);
    expect(buildIoDiagnostics({
      inputs: [
        { optionName: "--in", value: "-", allowImplicitStdin: true },
        { optionName: "--state", value: "state.json", allowImplicitStdin: false }
      ],
      output: "output.json"
    })).toEqual({
      inputs: [
        { option: "--in", mode: "stdin" },
        { option: "--state", mode: "file", path: "state.json" }
      ],
      output: { mode: "file", path: "output.json" }
    });
  });

  it("keeps error diagnostics argv I/O and implicit stdin behavior", () => {
    expect(buildIoDiagnosticsFromArgv([
      "ai", "detect-kind", "--diagnostics", "json"
    ])).toEqual({
      inputs: [{ option: "--in", mode: "stdin_implicit" }],
      output: { mode: "stdout" }
    });
    expect(buildIoDiagnosticsFromArgv([
      "export", "xlsx", "--in", "state.json", "--out-base64", "-"
    ])).toEqual({
      inputs: [{ option: "--in", mode: "file", path: "state.json" }],
      output: { mode: "stdout_base64" }
    });
  });
});
