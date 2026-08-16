#!/usr/bin/env node

import {
  cpSync,
  lstatSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  readdirSync,
  rmSync,
  writeFileSync
} from "node:fs";
import os from "node:os";
import path from "node:path";
import { execFileSync, spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

import {
  canonicalJsonText,
  sha256CanonicalJson,
  sha256RawBytes
} from "./lib/v1/cli-v1-canonical-json.mjs";
import { parseV1JsonDocument } from "./lib/v1/cli-v1-json-artifact.mjs";
import { V1_CORE_CAPABILITIES } from "./lib/v1/cli-v1-runtime-manifest.mjs";
import { computeConformanceCorpusDigest } from "./build-cli-v1-runtime.mjs";
import {
  assertPinnedNodeRuntimeResultBinding,
  preflightPinnedNodeRuntimeConsumer,
  runPinnedNodeRuntimeConsumer
} from "../tests/helpers/mikuproject-v1-pinned-runtime-consumer.mjs";

const ROOT = path.resolve(import.meta.dirname, "..");
const PRODUCT_REPOSITORY = "https://github.com/igapyon/miku-project";
const CAPABILITY_PROFILE = "miku-project-cli-core/v1";
const FIXTURE_SUITE_VERSION = "1";
const PROJECT_FIXTURE = "testdata/conformance/v1/fixtures/project/dependency-canonical.xml";
const REQUEST_FIXTURE = "testdata/conformance/v1/fixtures/change/set-task-2-percent-0-to-50.template.json";
const BASE_STATE_DIGEST = "a98f0c8b560382234572a61f360c9e96911bc75fc1d57b79968d6e60b5d751d0";

if (isMainModule()) {
  main().catch((error) => {
    process.stderr.write(`${error?.message ?? String(error)}\n`);
    process.exitCode = 1;
  });
}

export class GateRuntimeVerificationError extends Error {
  constructor(code, message) {
    super(message);
    this.name = "GateRuntimeVerificationError";
    this.code = code;
  }
}

/**
 * Verifies one already-built internal Node reference candidate. The tracked
 * lock is the outer trust anchor; no expected digest is derived from the
 * candidate runtime directory or its manifest.
 */
export async function verifyCliV1ReleaseCandidate({
  root = ROOT,
  runtimeDirectory,
  lockPath,
  launch = defaultLaunch
} = {}) {
  assertNonEmptyString(runtimeDirectory, "runtimeDirectory");
  assertNonEmptyString(lockPath, "lockPath");
  if (typeof launch !== "function") {
    throw new TypeError("launch must be a function");
  }

  const resolvedRoot = path.resolve(root);
  const resolvedRuntimeDirectory = path.resolve(runtimeDirectory);
  const resolvedLockPath = path.resolve(lockPath);
  const rawLock = readRegularFile(resolvedLockPath, "lock");
  const lock = parseCanonicalLock(rawLock.bytes);
  validateGateRuntimeLock(lock);
  validateRepositoryEvidence({ root: resolvedRoot, lock });
  validateCandidateMembers({ runtimeDirectory: resolvedRuntimeDirectory, lock });

  const workRoot = mkdtempSync(path.join(os.tmpdir(), "miku-project-g4-candidate-verifier-"));
  try {
    const consumerRoot = path.join(workRoot, "consumer");
    const distributedRuntime = path.join(consumerRoot, "runtime");
    const consumerWork = path.join(consumerRoot, "work");
    const projectPath = path.join(consumerRoot, "dependency-canonical.xml");
    const requestTemplatePath = path.join(consumerRoot, "set-task-2-percent-0-to-50.template.json");
    mkdirSync(distributedRuntime, { recursive: true });
    mkdirSync(consumerWork, { recursive: true });
    for (const member of lock.candidate.members) {
      cpSync(path.join(resolvedRuntimeDirectory, member.path), path.join(distributedRuntime, member.path));
    }
    cpSync(path.join(resolvedRoot, PROJECT_FIXTURE), projectPath);
    cpSync(path.join(resolvedRoot, REQUEST_FIXTURE), requestTemplatePath);
    if (readdirSync(consumerRoot).includes("node_modules")) {
      throw gateFailure("gate-runtime.consumer-not-isolated", "The candidate consumer must not contain node_modules.");
    }

    const candidatePreflight = await preflightPinnedNodeRuntimeConsumer({
      runtimeDirectory: distributedRuntime,
      trustedManifestDigest: lock.candidate.runtime_manifest_digest,
      productVersion: lock.product.release_version,
      runtimeVersion: lock.product.runtime_version,
      conformanceCorpusDigest: lock.conformance.corpus_digest
    });
    validateLockManifestBinding({ lock, manifest: candidatePreflight.manifest });

    const launches = [];
    const runPinned = async (args) => runPinnedNodeRuntimeConsumer({
      runtimeDirectory: distributedRuntime,
      trustedManifestDigest: lock.candidate.runtime_manifest_digest,
      productVersion: lock.product.release_version,
      runtimeVersion: lock.product.runtime_version,
      conformanceCorpusDigest: lock.conformance.corpus_digest,
      args,
      cwd: consumerWork,
      launch: (descriptor) => {
        launches.push(descriptor);
        return launch(descriptor);
      }
    });

    const validateRun = await runPinned(["validate", "--project", projectPath]);
    const inspectRun = await runPinned([
      "inspect", "--project", projectPath, "--purpose", "task_change_context", "--task-uid", "2"
    ]);
    assertSuccessfulLaunch(validateRun.result, "validate");
    assertSuccessfulLaunch(inspectRun.result, "inspect");

    const request = JSON.parse(readFileSync(requestTemplatePath, "utf8").replace(
      "${BASE_STATE_DIGEST}", BASE_STATE_DIGEST
    ));
    const requestPath = path.join(consumerWork, "change-request.json");
    const planResultPath = path.join(consumerWork, "plan-result.json");
    const approvalPath = path.join(consumerWork, "approval.json");
    const applyResultPath = path.join(consumerWork, "apply-result.json");
    const destination = path.join(consumerWork, "artifact-set");
    writeFileSync(requestPath, `${canonicalJsonText(request)}\n`, "utf8");

    const planRun = await runPinned([
      "plan-change", "--project", projectPath, "--request", requestPath,
      "--destination", destination, "--result", planResultPath
    ]);
    assertSuccessfulLaunch(planRun.result, "plan-change");
    const planResult = readJson(planResultPath);
    writeFileSync(approvalPath, `${canonicalJsonText(approvalForPlan(request, planResult))}\n`, "utf8");

    const applyRun = await runPinned([
      "apply-change", "--project", projectPath, "--request", requestPath,
      "--plan-result", planResultPath, "--approval", approvalPath, "--result", applyResultPath
    ]);
    assertSuccessfulLaunch(applyRun.result, "apply-change");
    const applyResult = readJson(applyResultPath);
    const verifyRun = await runPinned([
      "verify-artifact", "--artifact-set", applyResult.data.artifact_set.path,
      "--expect-plan-result", planResultPath
    ]);
    assertSuccessfulLaunch(verifyRun.result, "verify-artifact");

    const executions = [
      [JSON.parse(validateRun.result.stdout), validateRun.preflight],
      [JSON.parse(inspectRun.result.stdout), inspectRun.preflight],
      [planResult, planRun.preflight],
      [applyResult, applyRun.preflight],
      [JSON.parse(verifyRun.result.stdout), verifyRun.preflight]
    ];
    for (const [result, preflight] of executions) {
      assertPinnedNodeRuntimeResultBinding(result, preflight);
      if (result.status !== "succeeded") {
        throw gateFailure("gate-runtime.workflow-failed", `${result.command} did not return succeeded.`);
      }
    }

    const expectedOutputRuntime = outputRuntimeBindingFromVerified(planRun.preflight.runtime);
    if (!sameCanonicalJson(planResult.data.output_plan.runtime, expectedOutputRuntime)) {
      throw gateFailure("gate-runtime.output-plan-binding-mismatch", "The output plan runtime binding is not exact.");
    }
    const provenance = readJson(path.join(applyResult.data.artifact_set.path, "provenance.json"));
    if (!sameCanonicalJson(provenance.runtime, expectedOutputRuntime)) {
      throw gateFailure("gate-runtime.provenance-binding-mismatch", "The provenance runtime binding is not exact.");
    }
    if (launches.length !== 5) {
      throw gateFailure("gate-runtime.launch-count-mismatch", "The candidate verifier must launch exactly five workflows.");
    }

    return Object.freeze({
      kind: "miku_project_gate_runtime_verification",
      schema_version: "1",
      gate: lock.gate,
      status: "succeeded",
      distribution_status: lock.candidate.distribution_status,
      lock_digest: sha256RawBytes(rawLock.bytes),
      runtime_manifest_digest: { ...lock.candidate.runtime_manifest_digest },
      source: { ...lock.source },
      product: { ...lock.product },
      workflows: ["validate", "inspect", "plan-change", "apply-change", "verify-artifact"]
    });
  } finally {
    rmSync(workRoot, { recursive: true, force: true });
  }
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  const result = await verifyCliV1ReleaseCandidate(options);
  process.stdout.write(`${canonicalJsonText(result)}\n`);
}

function parseArgs(argv) {
  const values = new Map();
  for (let index = 0; index < argv.length; index += 2) {
    const option = argv[index];
    const value = argv[index + 1];
    if (!["--runtime-dir", "--lock"].includes(option) || !value || values.has(option)) {
      throw new Error("usage: node scripts/verify-cli-v1-release-candidate.mjs --runtime-dir <directory> --lock <lock.json>");
    }
    values.set(option, value);
  }
  if (values.size !== 2) {
    throw new Error("usage: node scripts/verify-cli-v1-release-candidate.mjs --runtime-dir <directory> --lock <lock.json>");
  }
  return {
    runtimeDirectory: values.get("--runtime-dir"),
    lockPath: values.get("--lock")
  };
}

function validateRepositoryEvidence({ root, lock }) {
  const packageLock = readRegularFile(path.join(root, "package-lock.json"), "package-lock");
  if (!sameDigest(sha256RawBytes(packageLock.bytes), lock.build.toolchain.package_lock_digest)) {
    throw gateFailure("gate-runtime.package-lock-mismatch", "package-lock.json does not match the candidate build evidence.");
  }
  const corpusDigest = computeConformanceCorpusDigest({ root });
  if (!sameDigest(corpusDigest, lock.conformance.corpus_digest)) {
    throw gateFailure("gate-runtime.corpus-mismatch", "The conformance corpus does not match the candidate lock.");
  }
  let taggedRevision;
  try {
    taggedRevision = execFileSync("git", ["rev-parse", `${lock.source.tag}^{commit}`], {
      cwd: root,
      encoding: "utf8"
    }).trim();
  } catch {
    throw gateFailure("gate-runtime.source-tag-unavailable", "The candidate source tag is unavailable in this repository.");
  }
  if (taggedRevision !== lock.source.revision) {
    throw gateFailure("gate-runtime.source-tag-mismatch", "The candidate source tag does not identify the locked revision.");
  }
}

function validateCandidateMembers({ runtimeDirectory, lock }) {
  let entries;
  try {
    entries = readdirSync(runtimeDirectory, { withFileTypes: true });
  } catch {
    throw gateFailure("gate-runtime.directory-unavailable", "The candidate runtime directory is missing or unreadable.");
  }
  const expectedNames = lock.candidate.members.map((member) => member.path).sort();
  const actualNames = entries.map((entry) => entry.name).sort();
  if (!sameCanonicalJson(actualNames, expectedNames)) {
    throw gateFailure("gate-runtime.member-set-mismatch", "The candidate runtime directory must contain exactly the three locked members.");
  }
  for (const member of lock.candidate.members) {
    const candidatePath = path.join(runtimeDirectory, member.path);
    const file = readRegularFile(candidatePath, member.role);
    if (file.size_bytes !== member.size_bytes || !sameDigest(file.digest, member.digest)) {
      throw gateFailure("gate-runtime.member-digest-mismatch", `${member.role} does not match the candidate lock.`);
    }
  }
}

function validateLockManifestBinding({ lock, manifest }) {
  const executable = lock.candidate.members[1];
  const sources = lock.candidate.members[2];
  const expectedExecutable = {
    path: executable.path,
    media_type: executable.media_type,
    size_bytes: executable.size_bytes,
    digest: executable.digest
  };
  const expectedSources = {
    path: sources.path,
    media_type: sources.media_type,
    size_bytes: sources.size_bytes,
    digest: sources.digest
  };
  const bindings = [
    [manifest.source.contract, lock.source, "contract source"],
    [manifest.source.runtime, lock.source, "runtime source"],
    [manifest.artifacts.executable, expectedExecutable, "executable descriptor"],
    [manifest.artifacts.sources, expectedSources, "sources descriptor"],
    [manifest.compatibility.capabilities.provided, lock.conformance.capabilities, "capability set"],
    [manifest.compatibility.conformance.corpus_digest, lock.conformance.corpus_digest, "corpus digest"]
  ];
  for (const [actual, expected, label] of bindings) {
    if (!sameCanonicalJson(actual, expected)) {
      throw gateFailure("gate-runtime.lock-manifest-mismatch", `The locked ${label} does not match runtime-manifest.json.`);
    }
  }
  if (manifest.product.release_version !== lock.product.release_version
    || manifest.runtime.family !== lock.product.runtime_family
    || manifest.runtime.role !== lock.product.runtime_role
    || manifest.runtime.version !== lock.product.runtime_version
    || manifest.compatibility.conformance.fixture_suite_version !== lock.conformance.fixture_suite_version
    || manifest.compatibility.capabilities.profiles.length !== 1
    || manifest.compatibility.capabilities.profiles[0] !== lock.conformance.profile) {
    throw gateFailure("gate-runtime.lock-manifest-mismatch", "The locked product/runtime/conformance identity does not match runtime-manifest.json.");
  }
}

function parseCanonicalLock(bytes) {
  let value;
  try {
    value = parseV1JsonDocument(bytes, { option: "--lock", role: "gate_runtime_lock" });
  } catch {
    throw gateFailure("gate-runtime.lock-invalid", "The Gate G4 runtime lock is not valid strict JSON.");
  }
  let canonicalBytes;
  try {
    canonicalBytes = Buffer.from(`${canonicalJsonText(value)}\n`, "utf8");
  } catch {
    throw gateFailure("gate-runtime.lock-invalid", "The Gate G4 runtime lock is outside the canonical JSON domain.");
  }
  if (!bytes.equals(canonicalBytes)) {
    throw gateFailure("gate-runtime.lock-invalid", "The Gate G4 runtime lock must be canonical JSON with one trailing LF.");
  }
  return value;
}

function validateGateRuntimeLock(lock) {
  assertExactKeys(lock, ["build", "candidate", "conformance", "gate", "kind", "product", "schema_version", "source"], "$lock");
  assertEqual(lock.kind, "miku_project_gate_runtime_lock", "$lock.kind");
  assertEqual(lock.schema_version, "1", "$lock.schema_version");
  assertEqual(lock.gate, "G4", "$lock.gate");

  assertExactKeys(lock.product, ["release_version", "runtime_family", "runtime_role", "runtime_version"], "$lock.product");
  assertSemver(lock.product.release_version, "$lock.product.release_version");
  assertSemver(lock.product.runtime_version, "$lock.product.runtime_version");
  assertEqual(lock.product.runtime_family, "node", "$lock.product.runtime_family");
  assertEqual(lock.product.runtime_role, "reference", "$lock.product.runtime_role");

  assertExactKeys(lock.source, ["repository", "revision", "tag"], "$lock.source");
  assertEqual(lock.source.repository, PRODUCT_REPOSITORY, "$lock.source.repository");
  if (!/^[0-9a-f]{40}$/.test(lock.source.revision)) {
    throw gateFailure("gate-runtime.lock-invalid", "$lock.source.revision must be a full lowercase Git revision.");
  }
  assertEqual(lock.source.tag, `v${lock.product.release_version}`, "$lock.source.tag");

  assertExactKeys(lock.build, ["command", "toolchain"], "$lock.build");
  assertEqual(lock.build.command, "npm run build:cli-v1-runtime -- --out-dir <fresh-directory>", "$lock.build.command");
  assertExactKeys(lock.build.toolchain, ["esbuild", "node", "npm", "package_lock_digest"], "$lock.build.toolchain");
  for (const field of ["esbuild", "node", "npm"]) {
    assertSemver(lock.build.toolchain[field], `$lock.build.toolchain.${field}`);
  }
  assertDigest(lock.build.toolchain.package_lock_digest, "$lock.build.toolchain.package_lock_digest");

  assertExactKeys(lock.conformance, ["capabilities", "corpus_digest", "fixture_suite_version", "profile"], "$lock.conformance");
  if (!sameCanonicalJson(lock.conformance.capabilities, V1_CORE_CAPABILITIES)) {
    throw gateFailure("gate-runtime.lock-invalid", "$lock.conformance.capabilities must be the canonical v1 core set.");
  }
  assertDigest(lock.conformance.corpus_digest, "$lock.conformance.corpus_digest");
  assertEqual(lock.conformance.fixture_suite_version, FIXTURE_SUITE_VERSION, "$lock.conformance.fixture_suite_version");
  assertEqual(lock.conformance.profile, CAPABILITY_PROFILE, "$lock.conformance.profile");

  assertExactKeys(lock.candidate, ["distribution_status", "members", "runtime_manifest_digest"], "$lock.candidate");
  assertEqual(lock.candidate.distribution_status, "internal-reference-only", "$lock.candidate.distribution_status");
  assertDigest(lock.candidate.runtime_manifest_digest, "$lock.candidate.runtime_manifest_digest");
  if (!Array.isArray(lock.candidate.members) || lock.candidate.members.length !== 3) {
    throw gateFailure("gate-runtime.lock-invalid", "$lock.candidate.members must contain manifest, executable, and sources.");
  }
  const expectedMembers = [
    ["manifest", "runtime-manifest.json", "application/json"],
    ["executable", `miku-project-node-${lock.product.runtime_version}.mjs`, "text/javascript"],
    ["sources", `miku-project-node-${lock.product.runtime_version}-sources.tgz`, "application/gzip"]
  ];
  for (let index = 0; index < expectedMembers.length; index += 1) {
    const member = lock.candidate.members[index];
    const [role, expectedPath, mediaType] = expectedMembers[index];
    assertExactKeys(member, ["digest", "media_type", "path", "role", "size_bytes"], `$lock.candidate.members[${index}]`);
    assertEqual(member.role, role, `$lock.candidate.members[${index}].role`);
    assertEqual(member.path, expectedPath, `$lock.candidate.members[${index}].path`);
    assertEqual(member.media_type, mediaType, `$lock.candidate.members[${index}].media_type`);
    assertDigest(member.digest, `$lock.candidate.members[${index}].digest`);
    if (!Number.isSafeInteger(member.size_bytes) || member.size_bytes <= 0) {
      throw gateFailure("gate-runtime.lock-invalid", `$lock.candidate.members[${index}].size_bytes must be a positive safe integer.`);
    }
  }
  if (!sameDigest(lock.candidate.members[0].digest, lock.candidate.runtime_manifest_digest)) {
    throw gateFailure("gate-runtime.lock-invalid", "The manifest member digest must equal runtime_manifest_digest.");
  }
}

function readRegularFile(filePath, role) {
  let entry;
  try {
    entry = lstatSync(filePath);
  } catch {
    throw gateFailure("gate-runtime.file-unavailable", `${role} is missing or unreadable.`);
  }
  if (entry.isSymbolicLink() || !entry.isFile()) {
    throw gateFailure("gate-runtime.file-entry-invalid", `${role} must be a regular non-symlink file.`);
  }
  const bytes = readFileSync(filePath);
  return Object.freeze({
    bytes,
    size_bytes: bytes.length,
    digest: sha256RawBytes(bytes)
  });
}

function approvalForPlan(request, planResult) {
  return {
    kind: "miku_project_change_approval",
    schema_version: "1",
    semantic_contract_version: "1",
    approved: true,
    base_state_digest: { ...planResult.data.semantic_diff.base_state_digest },
    change_request_digest: sha256CanonicalJson(request),
    semantic_diff_digest: sha256CanonicalJson(planResult.data.semantic_diff),
    output_plan_digest: sha256CanonicalJson(planResult.data.output_plan)
  };
}

function outputRuntimeBindingFromVerified(runtime) {
  return {
    family: runtime.family,
    version: runtime.version,
    artifact_digest: { ...runtime.artifact_digest },
    manifest_digest: { ...runtime.manifest_digest },
    capability_profile: runtime.capability_profile,
    fixture_suite_version: runtime.fixture_suite_version
  };
}

function defaultLaunch({ command, executablePath, args, cwd }) {
  return spawnSync(command, [executablePath, ...args], { cwd, encoding: "utf8" });
}

function assertSuccessfulLaunch(result, command) {
  if (result.status !== 0) {
    throw gateFailure(
      "gate-runtime.workflow-launch-failed",
      `${command} exited with ${String(result.status)}: ${result.stderr || result.stdout || "no output"}`
    );
  }
}

function readJson(filePath) {
  return JSON.parse(readFileSync(filePath, "utf8"));
}

function assertExactKeys(value, expected, location) {
  if (!value || typeof value !== "object" || Array.isArray(value)
    || !sameCanonicalJson(Object.keys(value).sort(), [...expected].sort())) {
    throw gateFailure("gate-runtime.lock-invalid", `${location} has an unexpected object shape.`);
  }
}

function assertEqual(actual, expected, location) {
  if (actual !== expected) {
    throw gateFailure("gate-runtime.lock-invalid", `${location} must equal ${expected}.`);
  }
}

function assertDigest(value, location) {
  if (value?.algorithm !== "sha-256" || typeof value.value !== "string" || !/^[0-9a-f]{64}$/.test(value.value)) {
    throw gateFailure("gate-runtime.lock-invalid", `${location} must be a lowercase SHA-256 digest.`);
  }
}

function assertSemver(value, location) {
  if (typeof value !== "string" || !/^(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)(?:-[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*)?(?:\+[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*)?$/.test(value)) {
    throw gateFailure("gate-runtime.lock-invalid", `${location} must be a semantic version.`);
  }
}

function assertNonEmptyString(value, name) {
  if (typeof value !== "string" || value.length === 0) {
    throw new TypeError(`${name} must be a non-empty string`);
  }
}

function sameDigest(left, right) {
  return left?.algorithm === "sha-256"
    && right?.algorithm === "sha-256"
    && left.value === right.value;
}

function sameCanonicalJson(left, right) {
  try {
    return canonicalJsonText(left) === canonicalJsonText(right);
  } catch {
    return false;
  }
}

function gateFailure(code, message) {
  return new GateRuntimeVerificationError(code, message);
}

function isMainModule() {
  return process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url);
}
