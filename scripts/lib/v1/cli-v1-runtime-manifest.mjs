import fsPromises from "node:fs/promises";
import path from "node:path";

import { canonicalJsonText, sha256RawBytes } from "./cli-v1-canonical-json.mjs";
import { createV1RuntimeError } from "./cli-v1-errors.mjs";
import { parseV1JsonDocument } from "./cli-v1-json-artifact.mjs";
import { validateRuntimeManifest } from "../../generated/cli-v1-schema-validators.mjs";

export const V1_CORE_CAPABILITIES = Object.freeze([
  "miku-project.capability.apply-change.set-task-percent-complete/v1",
  "miku-project.capability.format.ms-project-xml-subset.read/v1",
  "miku-project.capability.format.ms-project-xml-subset.write/v1",
  "miku-project.capability.inspect.project-overview/v1",
  "miku-project.capability.inspect.task-change-context/v1",
  "miku-project.capability.plan-change.set-task-percent-complete/v1",
  "miku-project.capability.publication.exclusive-directory-commit-marker/v1",
  "miku-project.capability.validate.project/v1",
  "miku-project.capability.verify-artifact/v1"
]);

const PRODUCT_REPOSITORY = "https://github.com/igapyon/miku-project";
const CAPABILITY_PROFILE = "miku-project-cli-core/v1";

/**
 * Establishes the self-contained Node runtime binding before a workflow reads
 * any project-domain input.  The caller supplies the executable path rather
 * than allowing a directory scan, PATH lookup, or newest-version selection.
 */
export async function verifyV1VersionedNodeRuntime({
  entryPath,
  runtimeVersion,
  productVersion,
  conformanceCorpusDigest,
  fileSystem = fsPromises
} = {}) {
  if (!isNonEmptyString(entryPath) || !isSemver(runtimeVersion) || !isSemver(productVersion) || !isDigest(conformanceCorpusDigest)) {
    throw new TypeError("versioned Node runtime verification requires entryPath, runtimeVersion, productVersion, and conformanceCorpusDigest");
  }

  const executable = await readRegularRuntimeFile(path.resolve(entryPath), {
    fileSystem,
    code: "runtime.manifest-invalid",
    artifactRole: "runtime_executable",
    message: "The versioned runtime executable is not a regular file."
  });
  const runtimeDirectory = path.dirname(executable.path);
  const expectedExecutableName = `miku-project-node-${runtimeVersion}.mjs`;
  const expectedSourcesName = `miku-project-node-${runtimeVersion}-sources.tgz`;
  if (path.basename(executable.path) !== expectedExecutableName) {
    throw manifestInvalid({
      path: executable.path,
      artifactRole: "runtime_executable",
      message: "The executable filename does not match the embedded runtime version."
    });
  }

  const manifest = await readRuntimeManifest(path.join(runtimeDirectory, "runtime-manifest.json"), { fileSystem });
  validateManifestIdentity({
    manifest: manifest.value,
    runtimeVersion,
    productVersion,
    conformanceCorpusDigest,
    expectedExecutableName,
    expectedSourcesName,
    manifestPath: manifest.path
  });

  const executableDescriptor = manifest.value.artifacts.executable;
  if (!sameFileDescriptor(executable, executableDescriptor)) {
    throw artifactDigestMismatch({
      path: executable.path,
      expected: executableDescriptor.digest.value,
      actual: executable.digest.value
    });
  }

  const sources = await readRegularRuntimeFile(path.join(runtimeDirectory, expectedSourcesName), {
    fileSystem,
    code: "runtime.manifest-invalid",
    artifactRole: "runtime_sources",
    message: "The runtime source archive is missing or is not a regular file."
  });
  if (!sameFileDescriptor(sources, manifest.value.artifacts.sources)) {
    throw manifestInvalid({
      path: sources.path,
      artifactRole: "runtime_sources",
      message: "The runtime source archive does not match its manifest descriptor."
    });
  }

  return Object.freeze({
    binding_status: "verified",
    family: "node",
    version: runtimeVersion,
    artifact_digest: { ...executable.digest },
    manifest_digest: { ...manifest.digest },
    capability_profile: CAPABILITY_PROFILE,
    fixture_suite_version: manifest.value.compatibility.conformance.fixture_suite_version
  });
}

async function readRuntimeManifest(manifestPath, { fileSystem }) {
  const raw = await readRegularRuntimeFile(manifestPath, {
    fileSystem,
    code: "runtime.manifest-invalid",
    artifactRole: "runtime_manifest",
    message: "The runtime manifest is missing or is not a regular file."
  });
  let value;
  try {
    value = parseV1JsonDocument(raw.bytes, { option: "--runtime-manifest", role: "runtime_manifest" });
  } catch {
    throw manifestInvalid({
      path: raw.path,
      artifactRole: "runtime_manifest",
      message: "The runtime manifest is not a valid canonical v1 JSON document."
    });
  }
  if (!validateRuntimeManifest(value)) {
    throw manifestInvalid({
      path: raw.path,
      artifactRole: "runtime_manifest",
      message: "The runtime manifest does not satisfy the v1 runtime manifest schema."
    });
  }
  if (!raw.bytes.equals(Buffer.from(`${canonicalJsonText(value)}\n`, "utf8"))) {
    throw manifestInvalid({
      path: raw.path,
      artifactRole: "runtime_manifest",
      message: "The runtime manifest is not serialized as canonical v1 JSON with one trailing LF."
    });
  }
  return { ...raw, value };
}

/**
 * Reads one direct runtime member as raw bytes after rejecting missing,
 * non-regular, and symlink entries.  The verifier and an outer-pinned
 * consumer share this primitive so both establish the same file boundary;
 * callers own their diagnostic vocabulary and descriptor policy.
 */
export async function readV1RegularRuntimeFile(candidatePath, { fileSystem = fsPromises } = {}) {
  let entry;
  try {
    entry = await fileSystem.lstat(candidatePath);
  } catch {
    return null;
  }
  if (entry.isSymbolicLink() || !entry.isFile()) {
    return null;
  }
  let canonicalPath;
  let bytes;
  try {
    canonicalPath = await fileSystem.realpath(candidatePath);
    bytes = Buffer.from(await fileSystem.readFile(canonicalPath));
  } catch {
    return null;
  }
  return {
    path: canonicalPath,
    bytes,
    size_bytes: bytes.length,
    digest: sha256RawBytes(bytes)
  };
}

export function matchesV1RuntimeFileDescriptor(file, descriptor) {
  return descriptor?.size_bytes === file?.size_bytes
    && descriptor?.digest?.algorithm === file?.digest?.algorithm
    && descriptor?.digest?.value === file?.digest?.value;
}

async function readRegularRuntimeFile(candidatePath, {
  fileSystem,
  code,
  artifactRole,
  message
}) {
  const file = await readV1RegularRuntimeFile(candidatePath, { fileSystem });
  if (!file) {
    throw runtimeFailure({ code, path: candidatePath, artifactRole, message });
  }
  return file;
}

function validateManifestIdentity({
  manifest,
  runtimeVersion,
  productVersion,
  conformanceCorpusDigest,
  expectedExecutableName,
  expectedSourcesName,
  manifestPath
}) {
  if (!isCanonicalCoreCapabilityDeclaration(manifest.compatibility.capabilities)) {
    throw runtimeFailure({
      code: "runtime.capability-missing",
      path: manifestPath,
      artifactRole: "runtime_manifest",
      message: "The runtime manifest does not provide the complete canonical v1 capability profile."
    });
  }
  if (manifest.product.name !== "miku-project"
    || manifest.product.release_version !== productVersion
    || manifest.runtime.family !== "node"
    || manifest.runtime.role !== "reference"
    || manifest.runtime.version !== runtimeVersion
    || manifest.runtime.launcher !== "node"
    || manifest.reference_runtime !== null
    || manifest.artifacts.executable.path !== expectedExecutableName
    || manifest.artifacts.sources.path !== expectedSourcesName
    || manifest.source.contract.repository !== PRODUCT_REPOSITORY
    || manifest.source.runtime.repository !== PRODUCT_REPOSITORY
    || manifest.source.contract.tag !== `v${productVersion}`
    || manifest.source.runtime.tag !== `v${runtimeVersion}`
    || !isDigest(manifest.compatibility.conformance.corpus_digest)
    || manifest.compatibility.conformance.corpus_digest.algorithm !== conformanceCorpusDigest.algorithm
    || manifest.compatibility.conformance.corpus_digest.value !== conformanceCorpusDigest.value) {
    throw manifestInvalid({
      path: manifestPath,
      artifactRole: "runtime_manifest",
      message: "The runtime manifest is incompatible with the embedded Node runtime identity."
    });
  }
}

function isCanonicalCoreCapabilityDeclaration(capabilities) {
  return Array.isArray(capabilities?.profiles)
    && capabilities.profiles.length === 1
    && capabilities.profiles[0] === CAPABILITY_PROFILE
    && Array.isArray(capabilities.provided)
    && capabilities.provided.length === V1_CORE_CAPABILITIES.length
    && capabilities.provided.every((capability, index) => capability === V1_CORE_CAPABILITIES[index])
    && Array.isArray(capabilities.extensions)
    && capabilities.extensions.length === 0;
}

function sameFileDescriptor(file, descriptor) {
  return matchesV1RuntimeFileDescriptor(file, descriptor);
}

function isDigest(value) {
  return value?.algorithm === "sha-256" && typeof value.value === "string" && /^[0-9a-f]{64}$/.test(value.value);
}

function runtimeFailure({ code, path, artifactRole, message, details = {} }) {
  return createV1RuntimeError({ code, message, scope: "runtime", path, artifactRole, details });
}

function manifestInvalid({ path, artifactRole, message }) {
  return runtimeFailure({ code: "runtime.manifest-invalid", path, artifactRole, message });
}

function artifactDigestMismatch({ path, expected, actual }) {
  return runtimeFailure({
    code: "runtime.artifact-digest-mismatch",
    path,
    artifactRole: "runtime_executable",
    message: "The runtime executable does not match its manifest digest.",
    details: { expected_digest: expected, actual_digest: actual }
  });
}

function isNonEmptyString(value) {
  return typeof value === "string" && value.length > 0;
}

function isSemver(value) {
  return typeof value === "string"
    && /^(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)(?:-[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*)?(?:\+[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*)?$/.test(value);
}
