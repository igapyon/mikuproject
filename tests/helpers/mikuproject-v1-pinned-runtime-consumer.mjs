import fsPromises from "node:fs/promises";
import path from "node:path";

import { canonicalJsonText } from "../../scripts/lib/v1/cli-v1-canonical-json.mjs";
import { parseV1JsonDocument } from "../../scripts/lib/v1/cli-v1-json-artifact.mjs";
import {
  matchesV1RuntimeFileDescriptor,
  readV1RegularRuntimeFile,
  V1_CORE_CAPABILITIES
} from "../../scripts/lib/v1/cli-v1-runtime-manifest.mjs";
import { validateCliResult, validateRuntimeManifest } from "../../scripts/generated/cli-v1-schema-validators.mjs";

const PRODUCT_REPOSITORY = "https://github.com/igapyon/miku-project";
const CAPABILITY_PROFILE = "miku-project-cli-core/v1";
const MANIFEST_FILENAME = "runtime-manifest.json";

/**
 * Test/audit-owned consumer boundary for the P4.10 distribution smoke and the
 * Gate G4 release-candidate verifier. This is not a product launcher:
 * production consumers will obtain the outer pin from a Release checksum or a
 * Skills lock. Keeping it here lets the checks prove that the pin is verified
 * before any manifest-derived launch decision.
 */
export class PinnedRuntimeConsumerError extends Error {
  constructor(code, message) {
    super(message);
    this.name = "PinnedRuntimeConsumerError";
    this.code = code;
  }
}

/**
 * Verifies a raw manifest SHA-256 supplied outside the distributed runtime,
 * then validates the pinned manifest enough to construct one exact Node launch
 * descriptor.  Do not move parsing or path resolution before the pin check.
 */
export async function preflightPinnedNodeRuntimeConsumer({
  runtimeDirectory,
  trustedManifestDigest,
  productVersion,
  runtimeVersion,
  conformanceCorpusDigest,
  fileSystem = fsPromises
} = {}) {
  assertNonEmptyString(runtimeDirectory, "runtimeDirectory");
  assertSemver(productVersion, "productVersion");
  assertSemver(runtimeVersion, "runtimeVersion");
  assertDigest(trustedManifestDigest, "trustedManifestDigest");
  assertDigest(conformanceCorpusDigest, "conformanceCorpusDigest");
  assertFileSystem(fileSystem);

  const directory = path.resolve(runtimeDirectory);
  const manifestPath = path.join(directory, MANIFEST_FILENAME);
  const rawManifest = await readPinnedManifest(manifestPath, { fileSystem });

  // This is deliberately the first check after the raw bytes are available.
  // A distributed manifest must never be allowed to supply its own expected
  // digest, nor influence JSON parsing, executable selection, or spawning.
  if (!sameDigest(rawManifest.digest, trustedManifestDigest)) {
    throw consumerFailure(
      "consumer.manifest-pin-mismatch",
      "The distributed runtime manifest does not match the external trust anchor."
    );
  }

  const manifest = parseAndValidatePinnedManifest(rawManifest, {
    productVersion,
    runtimeVersion,
    conformanceCorpusDigest
  });
  const canonicalRuntimeDirectory = path.dirname(rawManifest.path);
  const assets = await verifyPinnedRuntimeAssets({
    runtimeDirectory: canonicalRuntimeDirectory,
    manifest,
    fileSystem
  });
  return Object.freeze({
    runtimeDirectory: canonicalRuntimeDirectory,
    manifestPath: rawManifest.path,
    launcher: manifest.runtime.launcher,
    executablePath: assets.executable.path,
    manifest,
    runtime: Object.freeze({
      binding_status: "verified",
      family: "node",
      version: runtimeVersion,
      artifact_digest: { ...manifest.artifacts.executable.digest },
      manifest_digest: { ...rawManifest.digest },
      capability_profile: CAPABILITY_PROFILE,
      fixture_suite_version: manifest.compatibility.conformance.fixture_suite_version
    })
  });
}

/**
 * Runs a single operation only after its consumer preflight succeeded.  The
 * caller supplies the launcher so the smoke can prove failed preflight causes
 * zero process launches.
 */
export async function runPinnedNodeRuntimeConsumer({
  args,
  cwd,
  launch,
  ...preflightOptions
} = {}) {
  const preflight = await preflightPinnedNodeRuntimeConsumer(preflightOptions);
  const result = launchPinnedNodeRuntimeConsumer({ preflight, args, cwd, launch });
  return { preflight, result };
}

export function launchPinnedNodeRuntimeConsumer({ preflight, args = [], cwd, launch } = {}) {
  if (!preflight || preflight.launcher !== "node" || !isNonEmptyString(preflight.executablePath)) {
    throw consumerFailure("consumer.launch-descriptor-invalid", "The pinned runtime launch descriptor is invalid.");
  }
  if (!Array.isArray(args) || !args.every((argument) => typeof argument === "string")) {
    throw new TypeError("args must be an array of strings");
  }
  if (typeof launch !== "function") {
    throw new TypeError("launch must be a function");
  }
  return launch(Object.freeze({
    launcher: "node",
    command: process.execPath,
    executablePath: preflight.executablePath,
    args: Object.freeze([...args]),
    cwd
  }));
}

/**
 * Requires the workflow result to repeat the exact runtime binding that was
 * pinned before the process was launched.
 */
export function assertPinnedNodeRuntimeResultBinding(result, preflight) {
  let schemaValid = false;
  try {
    schemaValid = validateCliResult(result);
  } catch {
    schemaValid = false;
  }
  if (!schemaValid) {
    throw consumerFailure(
      "consumer.result-invalid",
      "The CLI result is not a schema-valid v1 result envelope."
    );
  }

  const expected = preflight?.runtime;
  const actual = result.runtime;
  if (!hasExactCanonicalJson(actual, expected)) {
    throw consumerFailure(
      "consumer.result-binding-mismatch",
      "The CLI result runtime binding does not match the pinned runtime manifest."
    );
  }
  return true;
}

async function readPinnedManifest(manifestPath, { fileSystem }) {
  let entry;
  try {
    entry = await fileSystem.lstat(manifestPath);
  } catch {
    throw consumerFailure("consumer.manifest-unavailable", "The distributed runtime manifest is missing or unreadable.");
  }
  if (entry.isSymbolicLink() || !entry.isFile()) {
    throw consumerFailure("consumer.manifest-entry-invalid", "The distributed runtime manifest must be a regular non-symlink file.");
  }
  const raw = await readV1RegularRuntimeFile(manifestPath, { fileSystem });
  if (!raw) {
    throw consumerFailure("consumer.manifest-unavailable", "The distributed runtime manifest is missing or unreadable.");
  }
  return raw;
}

async function verifyPinnedRuntimeAssets({ runtimeDirectory, manifest, fileSystem }) {
  const expected = {
    executable: manifest.artifacts.executable,
    sources: manifest.artifacts.sources
  };
  const assets = {};
  for (const [role, descriptor] of Object.entries(expected)) {
    const candidatePath = path.join(runtimeDirectory, descriptor.path);
    const asset = await readV1RegularRuntimeFile(candidatePath, { fileSystem });
    if (!asset
      || path.dirname(asset.path) !== runtimeDirectory
      || path.basename(asset.path) !== descriptor.path) {
      throw consumerFailure(
        `consumer.runtime-${role}-entry-invalid`,
        `The distributed runtime ${role} must be a direct regular non-symlink member of the manifest directory.`
      );
    }
    if (asset.size_bytes !== descriptor.size_bytes) {
      throw consumerFailure(
        `consumer.runtime-${role}-size-mismatch`,
        `The distributed runtime ${role} size does not match the pinned manifest.`
      );
    }
    if (!matchesV1RuntimeFileDescriptor(asset, descriptor)) {
      throw consumerFailure(
        `consumer.runtime-${role}-digest-mismatch`,
        `The distributed runtime ${role} digest does not match the pinned manifest.`
      );
    }
    assets[role] = asset;
  }
  return Object.freeze(assets);
}

function parseAndValidatePinnedManifest(rawManifest, {
  productVersion,
  runtimeVersion,
  conformanceCorpusDigest
}) {
  let manifest;
  try {
    manifest = parseV1JsonDocument(rawManifest.bytes, {
      option: "--runtime-manifest",
      role: "runtime_manifest"
    });
  } catch {
    throw consumerFailure("consumer.manifest-invalid", "The pinned runtime manifest is not valid strict v1 JSON.");
  }
  let isCanonical;
  try {
    isCanonical = validateRuntimeManifest(manifest)
      && rawManifest.bytes.equals(Buffer.from(`${canonicalJsonText(manifest)}\n`, "utf8"));
  } catch {
    isCanonical = false;
  }
  if (!isCanonical) {
    throw consumerFailure("consumer.manifest-invalid", "The pinned runtime manifest is not canonical schema-valid v1 JSON.");
  }

  const expectedExecutable = `miku-project-node-${runtimeVersion}.mjs`;
  const expectedSources = `miku-project-node-${runtimeVersion}-sources.tgz`;
  if (!hasCanonicalCoreCapabilities(manifest.compatibility.capabilities)
    || manifest.product.name !== "miku-project"
    || manifest.product.release_version !== productVersion
    || manifest.runtime.family !== "node"
    || manifest.runtime.role !== "reference"
    || manifest.runtime.version !== runtimeVersion
    || manifest.runtime.launcher !== "node"
    || manifest.reference_runtime !== null
    || manifest.artifacts.executable.path !== expectedExecutable
    || manifest.artifacts.sources.path !== expectedSources
    || manifest.source.contract.repository !== PRODUCT_REPOSITORY
    || manifest.source.runtime.repository !== PRODUCT_REPOSITORY
    || manifest.source.contract.tag !== `v${productVersion}`
    || manifest.source.runtime.tag !== `v${runtimeVersion}`
    || !sameDigest(manifest.compatibility.conformance.corpus_digest, conformanceCorpusDigest)) {
    throw consumerFailure("consumer.manifest-incompatible", "The pinned runtime manifest is incompatible with the requested Node contract.");
  }
  return manifest;
}

function hasCanonicalCoreCapabilities(capabilities) {
  return Array.isArray(capabilities?.profiles)
    && capabilities.profiles.length === 1
    && capabilities.profiles[0] === CAPABILITY_PROFILE
    && Array.isArray(capabilities.provided)
    && capabilities.provided.length === V1_CORE_CAPABILITIES.length
    && capabilities.provided.every((capability, index) => capability === V1_CORE_CAPABILITIES[index])
    && Array.isArray(capabilities.extensions)
    && capabilities.extensions.length === 0;
}

function assertFileSystem(fileSystem) {
  for (const operation of ["lstat", "realpath", "readFile"]) {
    if (typeof fileSystem?.[operation] !== "function") {
      throw new TypeError(`fileSystem.${operation} must be a function`);
    }
  }
}

function assertDigest(value, name) {
  if (!isDigest(value)) {
    throw consumerFailure("consumer.trust-anchor-invalid", `${name} must be a sha-256 digest with 64 lowercase hexadecimal characters.`);
  }
}

function assertSemver(value, name) {
  if (typeof value !== "string" || !/^(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)(?:-[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*)?(?:\+[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*)?$/.test(value)) {
    throw new TypeError(`${name} must be a semantic version`);
  }
}

function assertNonEmptyString(value, name) {
  if (!isNonEmptyString(value)) {
    throw new TypeError(`${name} must be a non-empty string`);
  }
}

function sameDigest(left, right) {
  return isDigest(left) && isDigest(right)
    && left.algorithm === right.algorithm
    && left.value === right.value;
}

function hasExactCanonicalJson(left, right) {
  try {
    return canonicalJsonText(left) === canonicalJsonText(right);
  } catch {
    return false;
  }
}

function isDigest(value) {
  return value?.algorithm === "sha-256"
    && typeof value.value === "string"
    && /^[0-9a-f]{64}$/.test(value.value);
}

function isNonEmptyString(value) {
  return typeof value === "string" && value.length > 0;
}

function consumerFailure(code, message) {
  return new PinnedRuntimeConsumerError(code, message);
}
