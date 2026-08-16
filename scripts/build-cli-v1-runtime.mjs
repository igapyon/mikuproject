#!/usr/bin/env node

import fs from "node:fs";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

import { canonicalJsonText, compareUnicodeScalars, sha256RawBytes } from "./lib/v1/cli-v1-canonical-json.mjs";
import { V1_CORE_CAPABILITIES } from "./lib/v1/cli-v1-runtime-manifest.mjs";
import { validateRuntimeManifest } from "./generated/cli-v1-schema-validators.mjs";

const ROOT = process.cwd();
const PRODUCT_REPOSITORY = "https://github.com/igapyon/miku-project";

if (isMainModule()) {
  main();
}

/**
 * Builds a release-candidate Node runtime only from a clean, exact-tagged
 * source tree. The generated directory is deliberately fresh: it never
 * replaces a previous runtime bundle or reuses a stale manifest.
 */
export function buildReleaseNodeRuntime({ root = ROOT, outDir = undefined } = {}) {
  const resolvedRoot = path.resolve(root);
  const packageVersion = readPackageVersion(resolvedRoot);
  const sourceIdentity = assertCleanTaggedReleaseSource({ root: resolvedRoot, packageVersion });
  return buildNodeRuntime({
    root: resolvedRoot,
    outDir: outDir ?? path.join(resolvedRoot, "runtime", "node"),
    packageVersion,
    runtimeVersion: packageVersion,
    sourceIdentity
  });
}

/**
 * Test-only construction hook. It is intentionally not exposed through the
 * CLI; its caller must supply an explicit source identity and use a temporary
 * directory. These artifacts have no Release checksum or external manifest
 * pin and therefore are not release candidates.
 */
export function buildNodeRuntimeForTest({
  outDir,
  sourceRevision,
  sourceTag,
  root = ROOT,
  packageVersion = readPackageVersion(root),
  runtimeVersion = packageVersion
} = {}) {
  if (!isNonEmptyString(outDir) || !isRevision(sourceRevision) || !isReleaseTag(sourceTag)) {
    throw new TypeError("test runtime construction requires outDir, sourceRevision, and sourceTag");
  }
  return buildNodeRuntime({
    root,
    outDir,
    packageVersion,
    runtimeVersion,
    sourceIdentity: { revision: sourceRevision, tag: sourceTag }
  });
}

export function computeConformanceCorpusDigest({ root = ROOT } = {}) {
  const corpusRoot = path.join(root, "testdata", "conformance", "v1");
  const files = collectRegularFiles(corpusRoot, corpusRoot);
  const lines = files
    .sort(compareUnicodeScalars)
    .map((relativePath) => `${sha256RawBytes(fs.readFileSync(path.join(corpusRoot, relativePath))).value}  ${relativePath}\n`);
  return sha256RawBytes(Buffer.from(lines.join(""), "utf8"));
}

function main() {
  const args = parseArgs(process.argv.slice(2));
  const result = buildReleaseNodeRuntime({ outDir: args.outDir });
  process.stdout.write(`${canonicalJsonText({
    kind: "miku_project_runtime_build",
    schema_version: "1",
    runtime_directory: result.runtimeDirectory,
    manifest_path: result.manifestPath,
    manifest_digest: result.manifestDigest
  })}\n`);
}

function parseArgs(argv) {
  if (argv.length === 0) {
    return { outDir: path.join(ROOT, "runtime", "node") };
  }
  if (argv.length === 2 && argv[0] === "--out-dir" && isNonEmptyString(argv[1])) {
    return { outDir: path.resolve(argv[1]) };
  }
  throw new Error("usage: node scripts/build-cli-v1-runtime.mjs [--out-dir <fresh-runtime-directory>]");
}

function buildNodeRuntime({ root, outDir, packageVersion, runtimeVersion, sourceIdentity }) {
  if (!isSemver(packageVersion) || !isSemver(runtimeVersion) || !isRevision(sourceIdentity?.revision) || !isReleaseTag(sourceIdentity?.tag)) {
    throw new TypeError("Node runtime build received an invalid version or source identity");
  }
  const runtimeDirectory = path.resolve(outDir);
  assertFreshRuntimeDirectory(runtimeDirectory);

  const executableName = `miku-project-node-${runtimeVersion}.mjs`;
  const sourcesName = `miku-project-node-${runtimeVersion}-sources.tgz`;
  const executablePath = path.join(runtimeDirectory, executableName);
  const sourcesPath = path.join(runtimeDirectory, sourcesName);
  const manifestPath = path.join(runtimeDirectory, "runtime-manifest.json");

  fs.mkdirSync(runtimeDirectory, { recursive: false, mode: 0o755 });
  try {
    runBundleBuilder({ root, executablePath, sourcesPath, packageVersion, runtimeVersion });
    const manifest = createNodeRuntimeManifest({
      root,
      packageVersion,
      runtimeVersion,
      sourceIdentity,
      executableName,
      sourcesName,
      executablePath,
      sourcesPath
    });
    if (!validateRuntimeManifest(manifest)) {
      throw new Error("generated runtime manifest does not satisfy the v1 schema");
    }
    validateReleaseManifestConventions({ manifest, executableName, sourcesName, packageVersion, runtimeVersion });
    const manifestBytes = Buffer.from(`${canonicalJsonText(manifest)}\n`, "utf8");
    fs.writeFileSync(manifestPath, manifestBytes, { flag: "wx", mode: 0o644 });
    return Object.freeze({
      runtimeDirectory,
      manifestPath,
      manifest,
      manifestDigest: sha256RawBytes(manifestBytes)
    });
  } catch (error) {
    fs.rmSync(runtimeDirectory, { recursive: true, force: true });
    throw error;
  }
}

function createNodeRuntimeManifest({
  root,
  packageVersion,
  runtimeVersion,
  sourceIdentity,
  executableName,
  sourcesName,
  executablePath,
  sourcesPath
}) {
  return {
    kind: "miku_project_runtime_manifest",
    schema_version: "1",
    product: {
      name: "miku-project",
      release_version: packageVersion,
      product_contract_version: "1",
      semantic_contract_version: "1",
      format_contract_version: "1",
      change_contract_version: "1",
      cli_contract_version: "1",
      artifact_schema: "miku_project_artifacts/v1",
      result_schema: "miku_project_cli_result/v1",
      diagnostic_schema: "miku_project_cli_diagnostic/v1",
      diagnostic_catalog_version: "1"
    },
    runtime: {
      family: "node",
      role: "reference",
      version: runtimeVersion,
      launcher: "node"
    },
    compatibility: {
      capabilities: {
        catalog_version: "1",
        profiles: ["miku-project-cli-core/v1"],
        provided: [...V1_CORE_CAPABILITIES],
        extensions: []
      },
      conformance: {
        fixture_suite_version: "1",
        corpus_digest: computeConformanceCorpusDigest({ root })
      }
    },
    artifacts: {
      executable: describeArtifact(executablePath, executableName, "text/javascript"),
      sources: describeArtifact(sourcesPath, sourcesName, "application/gzip")
    },
    source: {
      contract: {
        repository: PRODUCT_REPOSITORY,
        revision: sourceIdentity.revision,
        tag: sourceIdentity.tag
      },
      runtime: {
        repository: PRODUCT_REPOSITORY,
        revision: sourceIdentity.revision,
        tag: sourceIdentity.tag
      }
    },
    reference_runtime: null
  };
}

function describeArtifact(artifactPath, artifactName, mediaType) {
  const bytes = fs.readFileSync(artifactPath);
  return {
    path: artifactName,
    media_type: mediaType,
    size_bytes: bytes.length,
    digest: sha256RawBytes(bytes)
  };
}

function runBundleBuilder({ root, executablePath, sourcesPath, packageVersion, runtimeVersion }) {
  const result = spawnSync(process.execPath, [
    path.join(root, "scripts", "build-cli-bundle.mjs"),
    "--out", executablePath,
    "--sources-out", sourcesPath,
    "--package-version", packageVersion,
    "--v1-runtime-version", runtimeVersion
  ], { cwd: root, encoding: "utf8" });
  if (result.status !== 0) {
    throw new Error(`versioned runtime bundle build failed: ${result.stderr || result.stdout || `exit ${String(result.status)}`}`);
  }
}

function assertCleanTaggedReleaseSource({ root, packageVersion }) {
  const status = runGit(root, ["status", "--porcelain"]);
  if (status.stdout.trim().length !== 0) {
    throw new Error("release runtime build requires a clean working tree");
  }
  const revision = runGit(root, ["rev-parse", "HEAD"]).stdout.trim();
  const tag = runGit(root, ["describe", "--tags", "--exact-match", "HEAD"]).stdout.trim();
  if (tag !== `v${packageVersion}`) {
    throw new Error(`release runtime build requires HEAD to be tagged v${packageVersion}`);
  }
  if (!isRevision(revision)) {
    throw new Error("release runtime build could not determine a full source revision");
  }
  return { revision, tag };
}

function runGit(root, args) {
  const result = spawnSync("git", args, { cwd: root, encoding: "utf8" });
  if (result.status !== 0) {
    throw new Error(`git ${args.join(" ")} failed: ${result.stderr.trim()}`);
  }
  return result;
}

function assertFreshRuntimeDirectory(runtimeDirectory) {
  try {
    fs.lstatSync(runtimeDirectory);
  } catch (error) {
    if (error?.code === "ENOENT") {
      const parentDirectory = path.dirname(runtimeDirectory);
      fs.mkdirSync(parentDirectory, { recursive: true, mode: 0o755 });
      const parent = fs.lstatSync(parentDirectory);
      if (parent.isSymbolicLink() || !parent.isDirectory()) {
        throw new Error("runtime output parent must be a regular directory");
      }
      return;
    }
    throw error;
  }
  throw new Error("runtime output directory must not already exist");
}

function collectRegularFiles(root, currentDirectory) {
  const files = [];
  const entries = fs.readdirSync(currentDirectory, { withFileTypes: true })
    .sort((left, right) => compareUnicodeScalars(left.name, right.name));
  for (const entry of entries) {
    const absolutePath = path.join(currentDirectory, entry.name);
    const relativePath = path.relative(root, absolutePath).split(path.sep).join("/");
    if (entry.isSymbolicLink()) {
      throw new Error(`conformance corpus must not contain a symbolic link: ${relativePath}`);
    }
    if (entry.isDirectory()) {
      files.push(...collectRegularFiles(root, absolutePath));
      continue;
    }
    if (!entry.isFile()) {
      throw new Error(`conformance corpus must contain regular files only: ${relativePath}`);
    }
    files.push(relativePath);
  }
  return files;
}

function validateReleaseManifestConventions({ manifest, executableName, sourcesName, packageVersion, runtimeVersion }) {
  if (manifest.artifacts.executable.path !== executableName
    || manifest.artifacts.sources.path !== sourcesName
    || manifest.product.release_version !== packageVersion
    || manifest.runtime.version !== runtimeVersion
    || manifest.source.contract.tag !== `v${packageVersion}`
    || manifest.source.runtime.tag !== `v${runtimeVersion}`
    || manifest.compatibility.capabilities.provided.join("\n") !== V1_CORE_CAPABILITIES.join("\n")) {
    throw new Error("generated manifest violates the fixed v1 release naming or capability conventions");
  }
}

function readPackageVersion(root) {
  const packageJson = JSON.parse(fs.readFileSync(path.join(root, "package.json"), "utf8"));
  if (!isSemver(packageJson.version)) {
    throw new Error("package.json version must be SemVer for a runtime build");
  }
  return packageJson.version;
}

function isMainModule() {
  return process.argv[1] && path.resolve(process.argv[1]) === path.resolve(fileURLToPath(import.meta.url));
}

function isNonEmptyString(value) {
  return typeof value === "string" && value.length > 0;
}

function isRevision(value) {
  return typeof value === "string" && /^[0-9a-f]{40}$/.test(value);
}

function isReleaseTag(value) {
  return typeof value === "string" && /^v.+$/.test(value);
}

function isSemver(value) {
  return typeof value === "string"
    && /^(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)(?:-[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*)?(?:\+[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*)?$/.test(value);
}
