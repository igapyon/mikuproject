import { readdirSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import {
  ALL_SUITE,
  FAST_SUITE,
  FULL_ONLY_SUITE,
  FULL_SUITE,
  SUITES
} from "../scripts/lib/test-suite-topology.mjs";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..");
const testDirectory = path.resolve(repoRoot, "tests");

describe("test suite topology", () => {
  it("keeps fast focused and reserves expensive integration tests for full", () => {
    expect(SUITES.fast).toEqual(FAST_SUITE);
    expect(SUITES.full).toEqual(FULL_SUITE);
    expect(SUITES.all).toEqual(ALL_SUITE);
    expect(FULL_SUITE).toEqual([...FAST_SUITE, ...FULL_ONLY_SUITE]);
    expect(FAST_SUITE).not.toContain("tests/mikuproject-cli.test.js");
    expect(FAST_SUITE).not.toContain("tests/miku-project-browser-runtime.test.js");
    expect(FULL_ONLY_SUITE).toEqual([
      "tests/miku-project-browser-runtime.test.js",
      "tests/mikuproject-cli-v1-r1-integration.test.js",
      "tests/mikuproject-cli.test.js"
    ]);
  });

  it("assigns every checked-in test file exactly once to the complete suite", () => {
    const discovered = collectTestPaths(testDirectory, "tests").sort();
    const assigned = [...ALL_SUITE].sort();

    expect(assigned).toEqual(discovered);
    expect(new Set(ALL_SUITE).size).toBe(ALL_SUITE.length);
    expect(new Set(FAST_SUITE).size).toBe(FAST_SUITE.length);
    expect(new Set(FULL_ONLY_SUITE).size).toBe(FULL_ONLY_SUITE.length);
  });
});

function collectTestPaths(directory, relativeDirectory) {
  return readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const relativePath = path.posix.join(relativeDirectory, entry.name);
    if (entry.isDirectory()) {
      return collectTestPaths(path.join(directory, entry.name), relativePath);
    }
    return entry.isFile() && entry.name.endsWith(".test.js") ? [relativePath] : [];
  });
}
