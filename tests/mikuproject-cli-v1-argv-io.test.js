import { existsSync } from "node:fs";
import { mkdtemp, readFile, realpath } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { describe, expect, it } from "vitest";

import {
  isV1ControlInvocation,
  isV1WorkflowCommand,
  parseV1Invocation
} from "../scripts/lib/v1/cli-v1-argv.mjs";
import { isCliV1Error } from "../scripts/lib/v1/cli-v1-errors.mjs";
import { reserveV1ResultTransport } from "../scripts/lib/v1/cli-v1-io.mjs";
import { prepareV1WorkflowInvocation, recognizesV1Workflow } from "../scripts/lib/v1/cli-v1-router.mjs";
import {
  createUnverifiedRuntimeBinding,
  createV1ErrorResult,
  createV1Result,
  serializeV1Result
} from "../scripts/lib/v1/cli-v1-result.mjs";
import { validateCliResult } from "../scripts/generated/cli-v1-schema-validators.mjs";

describe("v1 argv grammar", () => {
  it("recognizes only the five v1 workflow words and parses their fixed option grammar", () => {
    expect(isV1WorkflowCommand("inspect")).toBe(true);
    expect(isV1WorkflowCommand("validate")).toBe(true);
    expect(isV1WorkflowCommand("plan-change")).toBe(true);
    expect(isV1WorkflowCommand("apply-change")).toBe(true);
    expect(isV1WorkflowCommand("verify-artifact")).toBe(true);
    expect(isV1WorkflowCommand("ai")).toBe(false);
    expect(recognizesV1Workflow(["validate", "--project", "project.xml"])).toBe(true);
    expect(recognizesV1Workflow(["ai", "spec"])).toBe(false);

    expect(parseV1Invocation([
      "inspect", "--project", "project.xml", "--purpose", "project_overview"
    ])).toEqual({
      kind: "workflow",
      command: "inspect",
      sideEffectClass: "read-only",
      options: { project: "project.xml", purpose: "project_overview", result: "-" }
    });
    expect(parseV1Invocation([
      "inspect", "--project", "-", "--purpose", "task_change_context", "--task-uid", "2", "--result", "result.json"
    ])).toMatchObject({ kind: "workflow", command: "inspect", options: { project: "-", purpose: "task_change_context", "task-uid": "2", result: "result.json" } });
    expect(parseV1Invocation([
      "validate", "--project", "project.xml", "--result", "-"
    ])).toMatchObject({ kind: "workflow", command: "validate" });
    expect(parseV1Invocation([
      "plan-change", "--project", "project.xml", "--request", "request.json", "--destination", "next-project"
    ])).toMatchObject({ kind: "workflow", command: "plan-change", sideEffectClass: "exchange-artifact-generation" });
    expect(parseV1Invocation([
      "apply-change", "--project", "project.xml", "--request", "request.json", "--plan-result", "plan.json", "--approval", "approval.json"
    ])).toMatchObject({ kind: "workflow", command: "apply-change", sideEffectClass: "meaning-change-and-project-artifact-generation" });
    expect(parseV1Invocation([
      "verify-artifact", "--artifact-set", "project-result", "--expect-plan-result", "plan.json"
    ])).toMatchObject({ kind: "workflow", command: "verify-artifact", sideEffectClass: "read-only" });
  });

  it("keeps control operations distinct from workflow result transport", () => {
    expect(isV1ControlInvocation(["--help"])).toBe(true);
    expect(isV1ControlInvocation(["--version"])).toBe(true);
    expect(isV1ControlInvocation(["inspect", "--help"])).toBe(true);
    expect(parseV1Invocation(["--help"])).toEqual({ kind: "control", control: "help", command: null });
    expect(parseV1Invocation(["validate", "--help"])).toEqual({ kind: "control", control: "command-help", command: "validate" });
    expectError(["inspect", "--help", "--project", "project.xml"], "cli.unexpected-argument");
  });

  it("rejects invalid grammar before any project input is read", () => {
    expectError([], "cli.unknown-command");
    expectError(["legacy-command"], "cli.unknown-command");
    expectError(["validate", "project.xml"], "cli.unexpected-argument");
    expectError(["validate", "--unknown", "value"], "cli.unknown-option");
    expectError(["validate", "--project=value"], "cli.unknown-option");
    expectError(["validate", "--project"], "cli.missing-option");
    expectError(["validate", "--project", "one.xml", "--project", "two.xml"], "cli.duplicate-option");
    expectError(["inspect", "--project", "one.xml", "--purpose", "project_overview", "--task-uid", "2"], "cli.invalid-option-value");
    expectError(["inspect", "--project", "one.xml", "--purpose", "task_change_context"], "cli.missing-option");
    expectError(["apply-change", "--project", "-", "--request", "-", "--plan-result", "plan.json", "--approval", "approval.json"], "cli.multiple-stdin-sources");
    expectError(["verify-artifact", "--artifact-set", "-"], "cli.invalid-option-value");
    expectError(["validate", "--project", "\0"], "cli.invalid-option-value");
  });
});

describe("v1 result transport and result envelope", () => {
  it("uses stdout only for the default result channel and writes one caller-supplied JSON document", async () => {
    const written = [];
    const result = usageResult();
    const transport = await reserveV1ResultTransport(undefined, {
      stdout: { write(value) { written.push(value); } }
    });
    expect(transport.target).toEqual({ target: "stdout", path: null });
    await transport.writeResult(result);
    expect(written).toEqual([serializeV1Result(result)]);
    await expect(transport.writeResult(result)).rejects.toThrow("already completed");
  });

  it("reserves a new result file exclusively, canonicalizes its parent, and never overwrites an existing entry", async () => {
    const temporaryDirectory = await mkdtemp(path.join(os.tmpdir(), "miku-project-v1-result-"));
    const canonicalTemporaryDirectory = await realpath(temporaryDirectory);
    const result = usageResult();
    const transport = await reserveV1ResultTransport("result.json", { cwd: temporaryDirectory });
    expect(transport.target).toEqual({ target: "file", path: path.join(canonicalTemporaryDirectory, "result.json") });
    expect(await readFile(transport.target.path, "utf8")).toBe("");
    await transport.writeResult(result);
    expect(await readFile(transport.target.path, "utf8")).toBe(serializeV1Result(result));

    await expect(reserveV1ResultTransport("result.json", { cwd: temporaryDirectory })).rejects.toMatchObject({
      code: "io.result-path-exists",
      status: "rejected"
    });
    await expect(reserveV1ResultTransport("missing/result.json", { cwd: temporaryDirectory })).rejects.toMatchObject({
      code: "io.result-path-unsafe",
      status: "rejected"
    });
    await expect(reserveV1ResultTransport("result.json/", { cwd: temporaryDirectory })).rejects.toMatchObject({
      code: "io.result-path-unsafe",
      status: "rejected"
    });
    await expect(reserveV1ResultTransport(".", { cwd: temporaryDirectory })).rejects.toMatchObject({
      code: "io.result-path-unsafe",
      status: "rejected"
    });
  });

  it("removes only its own unfinished reserved result file and keeps router preparation independent from process.argv", async () => {
    const temporaryDirectory = await mkdtemp(path.join(os.tmpdir(), "miku-project-v1-result-abort-"));
    const canonicalTemporaryDirectory = await realpath(temporaryDirectory);
    const transport = await reserveV1ResultTransport("reserved.json", { cwd: temporaryDirectory });
    expect(existsSync(transport.target.path)).toBe(true);
    await transport.abort();
    expect(existsSync(transport.target.path)).toBe(false);

    const prepared = await prepareV1WorkflowInvocation([
      "validate", "--project", "project.xml", "--result", "prepared.json"
    ], { cwd: temporaryDirectory });
    expect(prepared.invocation).toMatchObject({ kind: "workflow", command: "validate" });
    expect(prepared.resultTransport.target).toEqual({ target: "file", path: path.join(canonicalTemporaryDirectory, "prepared.json") });
    await prepared.resultTransport.abort();
  });

  it("builds schema-valid deterministic envelopes and derives the most conservative next action", () => {
    const runtime = createUnverifiedRuntimeBinding({ version: "1.0.2" });
    let usageError;
    try {
      parseV1Invocation(["validate", "--unknown", "value"]);
    } catch (error) {
      usageError = error;
    }
    const usageResult = createV1ErrorResult({ error: usageError, runtime });
    expect(usageResult).toMatchObject({
      command: "cli",
      status: "usage-error",
      exit_code: 2,
      next_action: {
        action: "revise-invocation-or-input",
        command: null,
        source_retryability: "after-input-change"
      }
    });
    expect(validateCliResult(usageResult)).toBe(true);
    expect(serializeV1Result(usageResult)).toBe(`${serializeV1Result(usageResult).trimEnd()}\n`);

    const runtimeError = createV1Result({
      command: "cli",
      runtime,
      status: "runtime-error",
      io: {
        stdin_option: null,
        inputs: [],
        result: { target: "stdout", path: null },
        destination: null
      },
      diagnostics: [
        diagnostic("io.result-reservation-failed", "io", "after-environment-change"),
        diagnostic("runtime.manifest-invalid", "runtime", "not-retryable")
      ]
    });
    expect(runtimeError.diagnostics.map((item) => item.code)).toEqual([
      "io.result-reservation-failed",
      "runtime.manifest-invalid"
    ]);
    expect(runtimeError.next_action).toEqual({
      action: "abort-and-investigate",
      command: null,
      source_retryability: "not-retryable"
    });
    expect(validateCliResult(runtimeError)).toBe(true);

    const validationSuccess = createV1Result({
      command: "validate",
      runtime: verifiedRuntime(),
      status: "succeeded",
      io: projectIo(),
      data: {
        validation: {
          valid: true,
          format_profile: "miku-project-ms-project-xml-subset/v1",
          state_digest: digest("d")
        }
      }
    });
    expect(validationSuccess).toMatchObject({ status: "succeeded", exit_code: 0, next_action: { action: "complete" } });
    expect(validateCliResult(validationSuccess)).toBe(true);

    const validationRejection = createV1Result({
      command: "validate",
      runtime: verifiedRuntime(),
      status: "rejected",
      io: projectIo(),
      diagnostics: [
        {
          ...diagnostic("semantic.invalid", "semantic", "after-input-change"),
          location: {
            scope: "semantic",
            path: "tasks[uid=2].percent_complete",
            option: "--project",
            artifact_role: "external_project",
            rule_id: "S-I012"
          }
        }
      ],
      data: { validation: { valid: false, format_profile: null, state_digest: null } }
    });
    expect(validationRejection).toMatchObject({ status: "rejected", exit_code: 1, next_action: { action: "revise-invocation-or-input" } });
    expect(validateCliResult(validationRejection)).toBe(true);
  });
});

function expectError(argv, expectedCode) {
  try {
    parseV1Invocation(argv);
  } catch (error) {
    expect(isCliV1Error(error)).toBe(true);
    expect(error.code).toBe(expectedCode);
    return;
  }
  throw new Error(`expected ${expectedCode} for ${JSON.stringify(argv)}`);
}

function diagnostic(code, category, retryability) {
  return {
    kind: "miku_project_cli_diagnostic",
    schema_version: "1",
    code,
    severity: "error",
    category,
    message: code,
    location: {
      scope: "internal",
      path: null,
      option: null,
      artifact_role: null,
      rule_id: null
    },
    retryability,
    details: {}
  };
}

function usageResult() {
  let error;
  try {
    parseV1Invocation(["unknown-v1-command"]);
  } catch (caught) {
    error = caught;
  }
  return createV1ErrorResult({
    error,
    runtime: createUnverifiedRuntimeBinding({ version: "1.0.2" })
  });
}

function digest(character) {
  return { algorithm: "sha-256", value: character.repeat(64) };
}

function verifiedRuntime() {
  return {
    binding_status: "verified",
    family: "node",
    version: "1.0.2",
    artifact_digest: digest("a"),
    manifest_digest: digest("b"),
    capability_profile: "miku-project-cli-core/v1",
    fixture_suite_version: "1"
  };
}

function projectIo() {
  return {
    stdin_option: null,
    inputs: [{
      role: "project",
      option: "--project",
      source: "file",
      path: "/work/project.xml",
      digest: digest("c")
    }],
    result: { target: "stdout", path: null },
    destination: null
  };
}
