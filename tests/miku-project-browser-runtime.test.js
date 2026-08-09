import { readFileSync } from "node:fs";
import { appendFile, copyFile, mkdir, mkdtemp, readFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { spawnSync } from "node:child_process";

import { beforeAll, describe, expect, it } from "vitest";

import { WEB_SURFACE_MODULE_RELATIVE_PATHS } from "../scripts/lib/runtime-module-paths.mjs";

const repoRoot = path.resolve(import.meta.dirname, "..");
const builderPath = path.resolve(repoRoot, "scripts/build-browser-runtime.mjs");
const smokePath = path.resolve(repoRoot, "scripts/smoke-browser-runtime.mjs");
const createManifestPath = path.resolve(repoRoot, "scripts/create-browser-runtime-manifest.mjs");
const verifyManifestPath = path.resolve(repoRoot, "scripts/verify-browser-runtime-manifest.mjs");
const packageJson = JSON.parse(readFileSync(path.resolve(repoRoot, "package.json"), "utf8"));
const browserSmokeFixture = readFileSync(
  path.resolve(repoRoot, "tests/fixtures/browser-runtime-smoke.html"),
  "utf8"
);

let runtimePath;
let runtimeSource;

beforeAll(async () => {
  const tempDir = await mkdtemp(path.join(os.tmpdir(), "miku-project-browser-runtime-test-"));
  runtimePath = path.join(tempDir, "miku-project-runtime.mjs");
  const result = spawnSync(process.execPath, [builderPath, "--out", runtimePath], {
    cwd: repoRoot,
    encoding: "utf8"
  });
  if (result.status !== 0) {
    throw new Error(`browser runtime build failed:\n${result.stdout}\n${result.stderr}`);
  }
  runtimeSource = await readFile(runtimePath, "utf8");
});

describe("browser runtime bundle", () => {
  it("exports the version, embedded core paths, and loader contract", () => {
    expect(runtimeSource).toContain(`export const version = ${JSON.stringify(packageJson.version)};`);
    expect(runtimeSource).toContain("export const embeddedCorePaths = Object.freeze(");
    expect(runtimeSource).toContain("src/js/core-api.js");
    for (const relativePath of WEB_SURFACE_MODULE_RELATIVE_PATHS) {
      expect(runtimeSource).not.toContain(relativePath);
    }
    expect(runtimeSource).toContain("export function loadMikuProjectRuntime(options = {})");
    expect(runtimeSource).toContain("if (options.reuseExisting !== true)");
    expect(runtimeSource).toContain("export default loadMikuProjectRuntime;");
  });

  it("contains no Node, CLI, or UI-entrypoint runtime dependency", () => {
    expect(runtimeSource).not.toMatch(/node:/);
    expect(runtimeSource).not.toMatch(/(^|[^A-Za-z0-9_$])process([^A-Za-z0-9_$]|$)/m);
    expect(runtimeSource).not.toContain("miku-project-cli.mjs");
    for (const relativePath of WEB_SURFACE_MODULE_RELATIVE_PATHS) {
      expect(runtimeSource).not.toContain(relativePath);
    }
  });

  it("keeps a browser-loadable smoke fixture aligned with the runtime contract", () => {
    expect(browserSmokeFixture).toContain('from "/bundle/miku-project-runtime.mjs"');
    expect(browserSmokeFixture).toContain("loadMikuProjectRuntime({ expectedVersion: version })");
    expect(browserSmokeFixture).toContain('document.body.dataset.status = "ok"');
  });

  it("loads the canonical core API in an isolated browser smoke process", () => {
    const result = spawnSync(process.execPath, [
      smokePath,
      runtimePath,
      "--expected-version",
      packageJson.version
    ], {
      cwd: repoRoot,
      encoding: "utf8"
    });

    expect(result.status, result.stderr).toBe(0);
    expect(result.stdout).toContain(`[smoke:browser-runtime] ok ${packageJson.version}`);
  });

  it("rejects an unexpected runtime version", () => {
    const result = spawnSync(process.execPath, [
      smokePath,
      runtimePath,
      "--expected-version",
      "0.0.0"
    ], {
      cwd: repoRoot,
      encoding: "utf8"
    });

    expect(result.status).not.toBe(0);
    expect(result.stderr).toContain(`expected 0.0.0, actual ${packageJson.version}`);
  });

  it("creates a release lock and rejects a tampered runtime", async () => {
    const releaseAssetName = `miku-project-runtime-${packageJson.version}.mjs`;
    const releaseDir = await mkdtemp(path.join(os.tmpdir(), "miku-project-browser-runtime-lock-"));
    const releaseRuntimePath = path.join(releaseDir, releaseAssetName);
    const manifestPath = path.join(releaseDir, `miku-project-runtime-${packageJson.version}.json`);
    await copyFile(runtimePath, releaseRuntimePath);

    const createResult = spawnSync(process.execPath, [
      createManifestPath,
      releaseRuntimePath,
      "--release-tag",
      `v${packageJson.version}`,
      "--out",
      manifestPath
    ], {
      cwd: repoRoot,
      encoding: "utf8"
    });
    expect(createResult.status, createResult.stderr).toBe(0);

    const manifest = JSON.parse(await readFile(manifestPath, "utf8"));
    expect(manifest).toMatchObject({
      schema_version: "miku-project.browser-runtime-lock/v1",
      release_tag: `v${packageJson.version}`,
      package_version: packageJson.version,
      asset_name: releaseAssetName
    });
    expect(manifest.sha256).toMatch(/^[0-9a-f]{64}$/);

    const verifyResult = spawnSync(process.execPath, [
      verifyManifestPath,
      manifestPath,
      releaseRuntimePath,
      "--expected-release-tag",
      `v${packageJson.version}`
    ], {
      cwd: repoRoot,
      encoding: "utf8"
    });
    expect(verifyResult.status, verifyResult.stderr).toBe(0);

    const tamperedDir = path.join(releaseDir, "tampered");
    const tamperedRuntimePath = path.join(tamperedDir, releaseAssetName);
    await mkdir(tamperedDir);
    await copyFile(releaseRuntimePath, tamperedRuntimePath);
    await appendFile(tamperedRuntimePath, "\n// tampered\n", "utf8");
    const tamperedResult = spawnSync(process.execPath, [
      verifyManifestPath,
      manifestPath,
      tamperedRuntimePath
    ], {
      cwd: repoRoot,
      encoding: "utf8"
    });
    expect(tamperedResult.status).not.toBe(0);
    expect(tamperedResult.stderr).toContain("runtime sha256 mismatch");
  });
});
