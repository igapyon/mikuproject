import fs from "node:fs";
import { copyFile, mkdir, mkdtemp, readFile, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { spawnSync } from "node:child_process";

import { describe, expect, it } from "vitest";

import {
  CanonicalJsonError,
  canonicalizeSemanticState,
  canonicalJsonText,
  compareUnicodeScalars,
  sha256CanonicalJson,
  sha256RawBytes,
  sha256SemanticState
} from "../scripts/lib/v1/cli-v1-canonical-json.mjs";
import {
  checkCliV1SchemaValidators,
  writeCliV1SchemaValidators
} from "../scripts/lib/v1/cli-v1-schema-validator-generator.mjs";
import {
  validateArtifact,
  validateCliDiagnostic,
  validateCliResult,
  validateRuntimeManifest
} from "../scripts/generated/cli-v1-schema-validators.mjs";

const repoRoot = path.resolve(import.meta.dirname, "..");
const contractCasesPath = path.join(repoRoot, "testdata/conformance/v1/contract-cases.json");
const generatedValidatorPath = path.join(repoRoot, "scripts/generated/cli-v1-schema-validators.mjs");
const goldenSemanticPaths = [
  [
    "testdata/conformance/v1/golden/semantic/dependency.state.json",
    "a98f0c8b560382234572a61f360c9e96911bc75fc1d57b79968d6e60b5d751d0"
  ],
  [
    "testdata/conformance/v1/golden/semantic/dependency-percent-50.state.json",
    "1c72d70cc114853b2a61f1c4798794093e46419f16b6cc49819a9d050cb67a08"
  ]
];

describe("v1 generated schema validators", () => {
  it("validates the checked-in official positive examples", () => {
    const artifactPaths = [
      "testdata/conformance/v1/golden/semantic/dependency.state.json",
      "testdata/conformance/v1/golden/semantic/dependency-percent-50.state.json",
      "testdata/conformance/v1/golden/projection/dependency.project-overview.json",
      "docs/examples/artifacts-v1/change-approval.example.json",
      "docs/examples/artifacts-v1/change-request.example.json",
      "docs/examples/artifacts-v1/provenance.example.json",
      "docs/examples/artifacts-v1/task-change-context-projection.example.json"
    ];
    const resultPaths = [
      "docs/examples/cli-v1/apply-change-incomplete.result.json",
      "docs/examples/cli-v1/apply-change-succeeded.result.json",
      "docs/examples/cli-v1/inspect-succeeded.result.json",
      "docs/examples/cli-v1/plan-change-succeeded.result.json",
      "docs/examples/cli-v1/runtime-manifest-invalid.result.json",
      "docs/examples/cli-v1/usage-error.result.json",
      "docs/examples/cli-v1/validate-rejected.result.json",
      "docs/examples/cli-v1/verify-artifact-committed.result.json",
      "docs/examples/cli-v1/verify-artifact-expected-plan-mismatch.result.json"
    ];
    const manifestPaths = [
      "docs/examples/runtime-manifest-v1/java-runtime-manifest.example.json",
      "docs/examples/runtime-manifest-v1/node-runtime-manifest.example.json"
    ];

    for (const relativePath of artifactPaths) {
      expect(validateArtifact(readJson(relativePath)), relativePath).toBe(true);
    }
    for (const relativePath of resultPaths) {
      const result = readJson(relativePath);
      expect(validateCliResult(result), relativePath).toBe(true);
      for (const diagnostic of result.diagnostics) {
        expect(validateCliDiagnostic(diagnostic), `${relativePath} diagnostic`).toBe(true);
      }
    }
    for (const relativePath of manifestPaths) {
      expect(validateRuntimeManifest(readJson(relativePath)), relativePath).toBe(true);
    }

    const malformedRepositoryManifest = readJson("docs/examples/runtime-manifest-v1/node-runtime-manifest.example.json");
    malformedRepositoryManifest.source.contract.repository = "not a URI";
    expect(validateRuntimeManifest(malformedRepositoryManifest)).toBe(false);
  });

  it("applies all schema-layer contract mutations without making Ajv error text part of the contract", () => {
    const contractCases = readJson("testdata/conformance/v1/contract-cases.json");
    const schemaCases = contractCases.cases.filter((testCase) => testCase.validation_layer === "json-schema");
    expect(schemaCases).toHaveLength(18);

    for (const testCase of schemaCases) {
      const documents = new Map(testCase.inputs.map((input) => [
        input.role,
        readJsonFromPath(path.resolve(path.dirname(contractCasesPath), input.path))
      ]));
      for (const mutation of testCase.mutations) {
        applyJsonPointerMutation(documents.get(mutation.input_role), mutation);
      }
      const actualValid = validateDocumentByRole(documents.get("result"), "result");
      expect(actualValid, testCase.id).toBe(testCase.expected_valid);
    }
  });

  it("keeps generation deterministic, detects schema drift, and leaves a repository-independent module", async () => {
    await expect(checkCliV1SchemaValidators()).resolves.toBeDefined();
    const generatedSource = await readFile(generatedValidatorPath, "utf8");
    expect(generatedSource).not.toMatch(/^\s*import\s/m);
    expect(generatedSource).not.toMatch(/^\s*export\s+.+?\s+from\s/m);
    expect(generatedSource).not.toContain(repoRoot);

    const standaloneDirectory = await mkdtemp(path.join(os.tmpdir(), "miku-project-v1-validator-standalone-"));
    const standaloneValidatorPath = path.join(standaloneDirectory, "cli-v1-schema-validators.mjs");
    await copyFile(generatedValidatorPath, standaloneValidatorPath);
    const standaloneResult = spawnSync(process.execPath, [
      "--input-type=module",
      "--eval",
      `const validators = await import(${JSON.stringify(pathToFileURL(standaloneValidatorPath).href)}); if (typeof validators.validateArtifact !== "function" || typeof validators.validateCliResult !== "function" || typeof validators.validateCliDiagnostic !== "function" || typeof validators.validateRuntimeManifest !== "function") process.exit(1);`
    ], {
      cwd: standaloneDirectory,
      encoding: "utf8"
    });
    expect(standaloneResult.status, standaloneResult.stderr).toBe(0);

    const temporaryRoot = await mkdtemp(path.join(os.tmpdir(), "miku-project-v1-validator-drift-"));
    const sourceSchemaDirectory = path.join(repoRoot, "docs/schemas");
    const temporarySchemaDirectory = path.join(temporaryRoot, "docs/schemas");
    await mkdir(temporarySchemaDirectory, { recursive: true });
    for (const entry of fs.readdirSync(sourceSchemaDirectory)) {
      await copyFile(path.join(sourceSchemaDirectory, entry), path.join(temporarySchemaDirectory, entry));
    }
    await writeCliV1SchemaValidators({ rootDirectory: temporaryRoot });
    await expect(checkCliV1SchemaValidators({ rootDirectory: temporaryRoot })).resolves.toBeDefined();

    const artifactSchemaPath = path.join(temporarySchemaDirectory, "miku-project-artifacts-v1.schema.json");
    const originalSchemaSource = await readFile(artifactSchemaPath, "utf8");
    const changedSchemaSource = originalSchemaSource.replace(
      '"identity": { "type": "string", "minLength": 1 }',
      '"identity": { "type": "string", "minLength": 2 }'
    );
    expect(changedSchemaSource).not.toBe(originalSchemaSource);
    await writeFile(artifactSchemaPath, changedSchemaSource, "utf8");
    await expect(checkCliV1SchemaValidators({ rootDirectory: temporaryRoot })).rejects.toThrow("out of date");
    await writeCliV1SchemaValidators({ rootDirectory: temporaryRoot });
    await expect(checkCliV1SchemaValidators({ rootDirectory: temporaryRoot })).resolves.toBeDefined();
  });
});

describe("v1 canonical JSON and semantic digest", () => {
  it("uses the exact escape and Unicode scalar ordering rules without normalizing strings", () => {
    expect(canonicalJsonText({
      "😀": 1,
      "\ue000": 2,
      control: "\u0000\b\t\n\f\r\u001f\"\\/",
      composed: "é",
      decomposed: "e\u0301"
    })).toBe('{"composed":"é","control":"\\u0000\\b\\t\\n\\f\\r\\u001f\\"\\\\/","decomposed":"é","":2,"😀":1}');
    expect(compareUnicodeScalars("\ue000", "😀")).toBeLessThan(0);
    expect(compareUnicodeScalars("é", "e\u0301")).toBeGreaterThan(0);
  });

  it("rejects values outside the v1 JSON domain without mutating valid input", () => {
    const value = { b: [1, true, null], a: "safe" };
    expect(canonicalJsonText(value)).toBe('{"a":"safe","b":[1,true,null]}');
    expect(value).toEqual({ b: [1, true, null], a: "safe" });
    for (const invalid of [
      { value: undefined },
      [1, , 3],
      { value: Number.NaN },
      { value: Number.POSITIVE_INFINITY },
      { value: 1.5 },
      { value: -0 },
      { value: 9007199254740992 },
      { value: BigInt(1) },
      { value: "\ud800" },
      Object.assign([], { named: "not-json" })
    ]) {
      expect(() => canonicalJsonText(invalid)).toThrow(CanonicalJsonError);
    }
  });

  it("preserves task preorder, sorts the declared semantic collections, and matches the two fixed digests", () => {
    for (const [relativePath, expectedDigest] of goldenSemanticPaths) {
      const semanticState = readJson(relativePath);
      expect(sha256CanonicalJson(semanticState)).toEqual({ algorithm: "sha-256", value: expectedDigest });
      expect(sha256SemanticState(semanticState)).toEqual({ algorithm: "sha-256", value: expectedDigest });
    }

    const state = readJson("testdata/conformance/v1/golden/semantic/dependency.state.json");
    state.dependencies = [
      { predecessor_uid: "2", successor_uid: "3", type: "FS", lag: "PT0H0M0S" },
      { predecessor_uid: "1", successor_uid: "2", type: "FS", lag: "PT0H0M0S" }
    ];
    state.resources = [{ uid: "😀" }, { uid: "\ue000" }];
    state.assignments = [{ uid: "2", task_uid: "2" }, { uid: "1", task_uid: "1" }];
    state.calendars = [{ uid: "2" }, { uid: "1" }];
    const original = structuredClone(state);

    const canonical = canonicalizeSemanticState(state);
    expect(canonical.tasks).toEqual(original.tasks);
    expect(canonical.dependencies.map((dependency) => dependency.predecessor_uid)).toEqual(["1", "2"]);
    expect(canonical.resources.map((resource) => resource.uid)).toEqual(["\ue000", "😀"]);
    expect(canonical.assignments.map((assignment) => assignment.uid)).toEqual(["1", "2"]);
    expect(canonical.calendars.map((calendar) => calendar.uid)).toEqual(["1", "2"]);
    expect(state).toEqual(original);
    expect(sha256RawBytes(Buffer.from("abc"))).toEqual({
      algorithm: "sha-256",
      value: "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad"
    });
  });
});

function readJson(relativePath) {
  return readJsonFromPath(path.join(repoRoot, relativePath));
}

function readJsonFromPath(filePath) {
  return JSON.parse(fs.readFileSync(filePath, "utf8"));
}

function validateDocumentByRole(document, role) {
  switch (role) {
    case "result":
      return validateCliResult(document);
    case "diagnostic":
      return validateCliDiagnostic(document);
    case "runtime_manifest":
      return validateRuntimeManifest(document);
    default:
      return validateArtifact(document);
  }
}

function applyJsonPointerMutation(document, mutation) {
  const tokens = parseJsonPointer(mutation.pointer);
  if (tokens.length === 0) {
    throw new Error(`root mutation is not supported for ${mutation.operation}`);
  }
  let parent = document;
  for (const token of tokens.slice(0, -1)) {
    if (!parent || typeof parent !== "object" || !Object.hasOwn(parent, token)) {
      throw new Error(`mutation pointer does not exist: ${mutation.pointer}`);
    }
    parent = parent[token];
  }
  const key = tokens.at(-1);
  if (mutation.operation === "remove") {
    if (Array.isArray(parent)) {
      parent.splice(requireArrayIndex(key, mutation.pointer, parent.length), 1);
    } else if (Object.hasOwn(parent, key)) {
      delete parent[key];
    } else {
      throw new Error(`mutation pointer does not exist: ${mutation.pointer}`);
    }
    return;
  }
  if (mutation.operation === "replace") {
    if (Array.isArray(parent)) {
      parent[requireArrayIndex(key, mutation.pointer, parent.length)] = structuredClone(mutation.value);
    } else if (Object.hasOwn(parent, key)) {
      parent[key] = structuredClone(mutation.value);
    } else {
      throw new Error(`mutation pointer does not exist: ${mutation.pointer}`);
    }
    return;
  }
  if (mutation.operation === "add") {
    if (Array.isArray(parent)) {
      if (key === "-") {
        parent.push(structuredClone(mutation.value));
      } else {
        parent.splice(requireArrayIndex(key, mutation.pointer, parent.length, true), 0, structuredClone(mutation.value));
      }
    } else {
      parent[key] = structuredClone(mutation.value);
    }
    return;
  }
  throw new Error(`unsupported mutation operation: ${mutation.operation}`);
}

function parseJsonPointer(pointer) {
  if (typeof pointer !== "string" || !pointer.startsWith("/")) {
    throw new Error(`invalid JSON Pointer: ${String(pointer)}`);
  }
  return pointer.slice(1).split("/").map((token) => token.replaceAll("~1", "/").replaceAll("~0", "~"));
}

function requireArrayIndex(value, pointer, length, allowEnd = false) {
  if (!/^(0|[1-9][0-9]*)$/.test(value)) {
    throw new Error(`array JSON Pointer token is invalid: ${pointer}`);
  }
  const index = Number(value);
  if (!Number.isSafeInteger(index) || index < 0) {
    throw new Error(`array JSON Pointer token is invalid: ${pointer}`);
  }
  if (index > length || (!allowEnd && index === length)) {
    throw new Error(`array JSON Pointer token is out of range: ${pointer}`);
  }
  return index;
}
