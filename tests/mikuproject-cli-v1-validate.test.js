import fs from "node:fs";
import { mkdtemp, readFile, realpath, symlink, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { parseV1Invocation } from "../scripts/lib/v1/cli-v1-argv.mjs";
import { sha256RawBytes } from "../scripts/lib/v1/cli-v1-canonical-json.mjs";
import { reserveV1ResultTransport } from "../scripts/lib/v1/cli-v1-io.mjs";
import { runV1Validate } from "../scripts/lib/v1/cli-v1-r1-commands.mjs";
import { serializeV1Result } from "../scripts/lib/v1/cli-v1-result.mjs";
import { validateCliResult } from "../scripts/generated/cli-v1-schema-validators.mjs";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const fixtureRoot = path.join(repoRoot, "testdata/conformance/v1/fixtures/project");
const suiteCases = new Map(readJson("testdata/conformance/v1/suite-index.json").cases.map((item) => [item.id, item]));
const testRuntime = Object.freeze({
  binding_status: "verified",
  family: "node",
  version: "1.0.2",
  artifact_digest: digest("0".repeat(64)),
  manifest_digest: digest("1".repeat(64)),
  capability_profile: "miku-project-cli-core/v1",
  fixture_suite_version: "1"
});

describe("v1 R1 validate service", () => {
  it("runs CV-VALID-001 from a regular file to a reserved result file without changing its input", async () => {
    const temporaryDirectory = await mkdtemp(path.join(os.tmpdir(), "miku-project-v1-validate-valid-"));
    const canonicalTemporaryDirectory = await realpath(temporaryDirectory);
    const fixturePath = path.join(fixtureRoot, "dependency-canonical.xml");
    const before = await readFile(fixturePath);
    const { result, output } = await invokeValidate([
      "validate", "--project", fixturePath, "--result", "validate.result.json"
    ], { cwd: temporaryDirectory });
    const expected = suiteCases.get("CV-VALID-001");

    expect(validateCliResult(result)).toBe(true);
    expect(result).toMatchObject({
      command: "validate",
      status: expected.expected_status,
      exit_code: expected.expected_exit_code,
      next_action: expected.expected_next_action,
      io: {
        stdin_option: null,
        inputs: [{
          role: "project",
          option: "--project",
          source: "file",
          path: await realpath(fixturePath),
          digest: sha256RawBytes(before)
        }],
        result: { target: "file", path: path.join(canonicalTemporaryDirectory, "validate.result.json") },
        destination: null
      },
      effects: {
        project_input_modified: false,
        project_artifact: null,
        cleanup: { status: "not-needed", path: null }
      },
      observations: { normalizations: [], losses: [], unsupported: [] },
      diagnostics: [],
      data: {
        validation: {
          valid: true,
          format_profile: "miku-project-ms-project-xml-subset/v1",
          state_digest: digest("a98f0c8b560382234572a61f360c9e96911bc75fc1d57b79968d6e60b5d751d0")
        }
      }
    });
    expect(output).toEqual([]);
    expect(await readFile(result.io.result.path, "utf8")).toBe(serializeV1Result(result));
    expect(await readFile(fixturePath)).toEqual(before);
  });

  it("runs CV-INVALID-001 from explicit stdin and returns its only result document to stdout", async () => {
    const input = readFixture("dependency-percent-101.xml");
    const { result, output } = await invokeValidate([
      "validate", "--project", "-"
    ], { stdin: input });
    const expected = suiteCases.get("CV-INVALID-001");

    expect(validateCliResult(result)).toBe(true);
    expect(result).toMatchObject({
      status: expected.expected_status,
      exit_code: expected.expected_exit_code,
      next_action: expected.expected_next_action,
      io: {
        stdin_option: "--project",
        inputs: [{ role: "project", option: "--project", source: "stdin", path: null, digest: sha256RawBytes(input) }],
        result: { target: "stdout", path: null },
        destination: null
      },
      effects: { project_input_modified: false, project_artifact: null, cleanup: { status: "not-needed", path: null } },
      observations: { normalizations: [], losses: [], unsupported: [] },
      data: { validation: { valid: false, format_profile: "miku-project-ms-project-xml-subset/v1", state_digest: null } }
    });
    expect(result.diagnostics.map((item) => item.code)).toEqual(expected.expected_diagnostic_codes);
    expect(result.diagnostics.map((item) => item.location.rule_id)).toEqual(expected.expected_rule_ids);
    expect(output).toEqual([serializeV1Result(result)]);
  });

  it("runs CV-UNSUPPORTED-001 as a rejected validation without inventing a state digest", async () => {
    const fixturePath = path.join(fixtureRoot, "dependency-unsupported-actual.xml");
    const { result, output } = await invokeValidate([
      "validate", "--project", fixturePath
    ]);
    const expected = suiteCases.get("CV-UNSUPPORTED-001");

    expect(validateCliResult(result)).toBe(true);
    expect(result).toMatchObject({
      status: expected.expected_status,
      exit_code: expected.expected_exit_code,
      next_action: expected.expected_next_action,
      data: { validation: { valid: false, format_profile: "miku-project-ms-project-xml-subset/v1", state_digest: null } },
      observations: {
        losses: [],
        unsupported: [{ code: "semantic.unsupported", path: "tasks[uid=2].actual_start" }]
      }
    });
    expect(result.diagnostics.map((item) => item.code)).toEqual(expected.expected_diagnostic_codes);
    expect(result.diagnostics.map((item) => item.location.rule_id)).toEqual(expected.expected_rule_ids);
    expect(output).toEqual([serializeV1Result(result)]);
  });

  it("runs the hierarchy forest and summary negative cases with their stable rule IDs and no side effects", async () => {
    for (const { caseId, fixtureName } of [
      { caseId: "CV-HIERARCHY-INVALID-PREORDER-001", fixtureName: "hierarchy-invalid-preorder.xml" },
      { caseId: "CV-HIERARCHY-INVALID-SUMMARY-001", fixtureName: "hierarchy-invalid-summary.xml" }
    ]) {
      const expected = suiteCases.get(caseId);
      const fixturePath = path.join(fixtureRoot, fixtureName);
      const before = await readFile(fixturePath);
      const { result, output } = await invokeValidate(["validate", "--project", fixturePath]);

      expect(validateCliResult(result)).toBe(true);
      expect(result).toMatchObject({
        command: "validate",
        status: expected.expected_status,
        exit_code: expected.expected_exit_code,
        next_action: expected.expected_next_action,
        effects: { project_input_modified: false, project_artifact: null, cleanup: { status: "not-needed", path: null } },
        data: { validation: { valid: false, format_profile: "miku-project-ms-project-xml-subset/v1", state_digest: null } }
      });
      expect(result.diagnostics.map((item) => item.code)).toEqual(expected.expected_diagnostic_codes);
      expect(result.diagnostics.map((item) => item.location.rule_id)).toEqual(expected.expected_rule_ids);
      expect(result.diagnostics.map((item) => item.location.path)).toEqual(expected.expected_diagnostic_paths);
      expect(output).toEqual([serializeV1Result(result)]);
      expect(await readFile(fixturePath)).toEqual(before);
    }
  });

  it("produces byte-identical stdout result documents for the same runtime, file input, and path token", async () => {
    const fixturePath = path.join(fixtureRoot, "dependency-canonical.xml");
    const first = await invokeValidate(["validate", "--project", fixturePath]);
    const second = await invokeValidate(["validate", "--project", fixturePath]);

    expect(first.result).toEqual(second.result);
    expect(first.output).toEqual(second.output);
  });

  it("rejects missing, symlink, and incomplete artifact-set directory entries before XML decoding", async () => {
    const temporaryDirectory = await mkdtemp(path.join(os.tmpdir(), "miku-project-v1-validate-input-"));
    const missing = await invokeValidate([
      "validate", "--project", "missing.xml"
    ], { cwd: temporaryDirectory });
    expect(missing.result).toMatchObject({
      status: "rejected",
      exit_code: 1,
      diagnostics: [{ code: "io.input-not-found" }],
      data: { validation: { valid: false, format_profile: null, state_digest: null } }
    });

    const targetPath = path.join(temporaryDirectory, "target.xml");
    const linkPath = path.join(temporaryDirectory, "input-link.xml");
    await writeFile(targetPath, readFixture("dependency-canonical.xml"));
    await symlink(targetPath, linkPath);
    const linked = await invokeValidate([
      "validate", "--project", linkPath
    ], { cwd: temporaryDirectory });
    expect(linked.result).toMatchObject({
      status: "rejected",
      diagnostics: [{ code: "io.input-symlink-rejected" }],
      data: { validation: { valid: false, format_profile: null, state_digest: null } }
    });

    const directory = await invokeValidate([
      "validate", "--project", temporaryDirectory
    ], { cwd: temporaryDirectory });
    expect(validateCliResult(directory.result)).toBe(true);
    expect(directory.result).toMatchObject({
      status: "rejected",
      diagnostics: [{ code: "publication.artifact-incomplete" }],
      data: { validation: { valid: false, format_profile: null, state_digest: null } }
    });
  });
});

async function invokeValidate(argv, { cwd = repoRoot, stdin = Buffer.alloc(0) } = {}) {
  const output = [];
  const invocation = parseV1Invocation(argv);
  const resultTransport = await reserveV1ResultTransport(invocation.options.result, {
    cwd,
    stdout: { write(value) { output.push(value); } }
  });
  const result = await runV1Validate({ invocation, resultTransport, runtime: testRuntime, cwd, stdin });
  return { result, output };
}

function readFixture(name) {
  return fs.readFileSync(path.join(fixtureRoot, name));
}

function readJson(relativePath) {
  return JSON.parse(fs.readFileSync(path.join(repoRoot, relativePath), "utf8"));
}

function digest(value) {
  return { algorithm: "sha-256", value };
}
