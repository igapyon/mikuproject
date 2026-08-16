import {
  cpSync,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  readdirSync,
  renameSync,
  rmSync,
  symlinkSync,
  writeFileSync
} from "node:fs";
import fsPromises from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { execFileSync, spawnSync } from "node:child_process";

import { afterEach, describe, expect, it } from "vitest";

import {
  buildReleaseNodeRuntime,
  buildNodeRuntimeForTest,
  computeConformanceCorpusDigest
} from "../scripts/build-cli-v1-runtime.mjs";
import { canonicalJsonText, sha256CanonicalJson, sha256RawBytes } from "../scripts/lib/v1/cli-v1-canonical-json.mjs";
import { runV1VersionedRuntime } from "../scripts/lib/v1/cli-v1-router.mjs";
import { verifyCliV1ReleaseCandidate } from "../scripts/verify-cli-v1-release-candidate.mjs";
import { validateCliResult, validateRuntimeManifest } from "../scripts/generated/cli-v1-schema-validators.mjs";
import {
  assertPinnedNodeRuntimeResultBinding,
  preflightPinnedNodeRuntimeConsumer,
  runPinnedNodeRuntimeConsumer
} from "./helpers/mikuproject-v1-pinned-runtime-consumer.mjs";

const repoRoot = path.resolve(import.meta.dirname, "..");
const runtimeVersion = JSON.parse(readFileSync(path.join(repoRoot, "package.json"), "utf8")).version;
const sourceRevision = execFileSync("git", ["rev-parse", "HEAD"], { cwd: repoRoot, encoding: "utf8" }).trim();
const sourceTag = `v${runtimeVersion}`;
const checkedInGateLock = readJson(path.join(repoRoot, "docs/miku-project-node-reference-runtime-lock-v1.0.3.json"));
const canonicalProject = path.join(repoRoot, "testdata/conformance/v1/fixtures/project/hierarchy-canonical.xml");
const flatCanonicalProject = path.join(repoRoot, "testdata/conformance/v1/fixtures/project/dependency-canonical.xml");
const flatChangeRequestTemplate = path.join(repoRoot, "testdata/conformance/v1/fixtures/change/set-task-2-percent-0-to-50.template.json");
const conformanceCases = new Map(readJson(path.join(repoRoot, "testdata/conformance/v1/suite-index.json")).cases.map((testCase) => [testCase.id, testCase]));
const temporaryDirectories = [];

afterEach(() => {
  while (temporaryDirectories.length > 0) {
    rmSync(temporaryDirectories.pop(), { recursive: true, force: true });
  }
});

describe("v1 versioned Node runtime manifest binding", () => {
  it("refuses a dirty or non-exact-tagged release source before making a runtime directory", () => {
    const directory = createTemporaryDirectory("miku-project-v1-runtime-release-preflight-");
    const dirtyRoot = createMiniGitRoot(directory, "dirty", { tag: sourceTag, dirty: true });
    const wrongTagRoot = createMiniGitRoot(directory, "wrong-tag", { tag: "v0.0.0", dirty: false });
    const dirtyOutput = path.join(directory, "dirty-runtime");
    const wrongTagOutput = path.join(directory, "wrong-tag-runtime");

    expect(() => buildReleaseNodeRuntime({ root: dirtyRoot, outDir: dirtyOutput })).toThrow("clean working tree");
    expect(() => buildReleaseNodeRuntime({ root: wrongTagRoot, outDir: wrongTagOutput })).toThrow(`tagged ${sourceTag}`);
    expect(existsSync(dirtyOutput)).toBe(false);
    expect(existsSync(wrongTagOutput)).toBe(false);
  });

  it("builds and invokes a runtime from a clean source at its exact release tag", () => {
    const directory = createTemporaryDirectory("miku-project-v1-runtime-release-success-");
    const releaseSource = createCleanTaggedReleaseSource(directory);
    const built = buildReleaseNodeRuntime({ root: releaseSource });
    const manifest = readJson(built.manifestPath);
    const projectPath = path.join(releaseSource, "testdata/conformance/v1/fixtures/project/dependency-canonical.xml");
    const invoked = runRuntime(built.runtimeDirectory, ["validate", "--project", projectPath], { cwd: directory });

    expect(runGitText(releaseSource, ["status", "--porcelain"])).toBe("");
    expect(manifest.source.contract).toEqual({
      repository: "https://github.com/igapyon/miku-project",
      revision: runGitText(releaseSource, ["rev-parse", "HEAD"]),
      tag: sourceTag
    });
    expect(manifest.source.runtime).toEqual(manifest.source.contract);
    expect(invoked.status).toBe(0);
    expect(readJsonText(invoked.stdout)).toMatchObject({
      command: "validate",
      status: "succeeded",
      runtime: { binding_status: "verified", version: runtimeVersion }
    });
  }, 30000);

  it("runs all v1 workflows from a three-file runtime distributed outside its source checkout", async () => {
    const directory = createTemporaryDirectory("miku-project-v1-runtime-consumer-smoke-");
    const releaseSource = createCleanTaggedReleaseSource(directory);
    const built = buildReleaseNodeRuntime({ root: releaseSource });
    const trustedManifestDigest = { ...built.manifestDigest };
    const consumerRoot = path.join(directory, "consumer");
    const distributedRuntime = path.join(consumerRoot, "runtime");
    const consumerWork = path.join(consumerRoot, "work");
    const consumerProject = path.join(consumerRoot, "dependency-canonical.xml");
    const consumerRequestTemplate = path.join(consumerRoot, "set-task-2-percent-0-to-50.template.json");
    const distributedMembers = ["runtime-manifest.json", executableName(), sourcesName()];

    mkdirSync(distributedRuntime, { recursive: true });
    mkdirSync(consumerWork, { recursive: true });
    for (const member of distributedMembers) {
      cpSync(path.join(built.runtimeDirectory, member), path.join(distributedRuntime, member));
    }
    cpSync(path.join(releaseSource, "testdata/conformance/v1/fixtures/project/dependency-canonical.xml"), consumerProject);
    cpSync(path.join(releaseSource, "testdata/conformance/v1/fixtures/change/set-task-2-percent-0-to-50.template.json"), consumerRequestTemplate);
    rmSync(releaseSource, { recursive: true, force: true });

    expect(existsSync(releaseSource)).toBe(false);
    expect(existsSync(path.join(consumerRoot, "node_modules"))).toBe(false);
    expect(readdirSync(distributedRuntime).sort()).toEqual([...distributedMembers].sort());

    const launches = [];
    const runPinned = async (args) => {
      const invoked = await runPinnedNodeRuntimeConsumer({
        runtimeDirectory: distributedRuntime,
        trustedManifestDigest,
        productVersion: runtimeVersion,
        runtimeVersion,
        conformanceCorpusDigest: computeConformanceCorpusDigest({ root: repoRoot }),
        args,
        cwd: consumerWork,
        launch: ({ launcher, command, executablePath, args: launchArgs, cwd }) => {
          launches.push({ launcher, command, executablePath, args: launchArgs, cwd });
          return spawnSync(command, [executablePath, ...launchArgs], { cwd, encoding: "utf8" });
        }
      });
      return invoked;
    };
    const validateRun = await runPinned(["validate", "--project", consumerProject]);
    const inspectRun = await runPinned([
      "inspect", "--project", consumerProject, "--purpose", "task_change_context", "--task-uid", "2"
    ]);
    const { result: validate, preflight: validatePreflight } = validateRun;
    const { result: inspect, preflight: inspectPreflight } = inspectRun;
    const expectedRuntime = validatePreflight.runtime;
    const expectedOutputRuntime = outputRuntimeBindingFromVerified(expectedRuntime);
    expect(validate.status).toBe(0);
    expect(inspect.status).toBe(0);

    const request = JSON.parse(readFileSync(consumerRequestTemplate, "utf8").replace(
      "${BASE_STATE_DIGEST}", "a98f0c8b560382234572a61f360c9e96911bc75fc1d57b79968d6e60b5d751d0"
    ));
    const requestPath = path.join(consumerWork, "change-request.json");
    const planResultPath = path.join(consumerWork, "plan-result.json");
    const approvalPath = path.join(consumerWork, "approval.json");
    const applyResultPath = path.join(consumerWork, "apply-result.json");
    const destination = path.join(consumerWork, "artifact-set");
    writeFileSync(requestPath, `${canonicalJsonText(request)}\n`, "utf8");

    const planRun = await runPinned([
      "plan-change", "--project", consumerProject, "--request", requestPath,
      "--destination", destination, "--result", planResultPath
    ]);
    const { result: plan, preflight: planPreflight } = planRun;
    expect(plan.status).toBe(0);
    const planResult = readJson(planResultPath);
    expect(planResult.data.output_plan.runtime).toEqual(expectedOutputRuntime);
    writeFileSync(approvalPath, `${canonicalJsonText(approvalForPlan(request, planResult))}\n`, "utf8");

    const applyRun = await runPinned([
      "apply-change", "--project", consumerProject, "--request", requestPath,
      "--plan-result", planResultPath, "--approval", approvalPath, "--result", applyResultPath
    ]);
    const { result: apply, preflight: applyPreflight } = applyRun;
    expect(apply.status).toBe(0);
    const applyResult = readJson(applyResultPath);
    expect(readJson(path.join(applyResult.data.artifact_set.path, "provenance.json")).runtime).toEqual(expectedOutputRuntime);
    const verifyRun = await runPinned([
      "verify-artifact", "--artifact-set", applyResult.data.artifact_set.path,
      "--expect-plan-result", planResultPath
    ]);
    const { result: verify, preflight: verifyPreflight } = verifyRun;
    expect(verify.status).toBe(0);

    for (const [result, preflight] of [
      [readJsonText(validate.stdout), validatePreflight],
      [readJsonText(inspect.stdout), inspectPreflight],
      [planResult, planPreflight],
      [applyResult, applyPreflight],
      [readJsonText(verify.stdout), verifyPreflight]
    ]) {
      assertPinnedNodeRuntimeResultBinding(result, preflight);
      expect(result.status).toBe("succeeded");
      expect(result.runtime).toEqual(preflight.runtime);
    }
    expect(launches).toHaveLength(5);
    expect(launches).toEqual(launches.map(({ args, cwd }) => ({
      launcher: "node",
      command: process.execPath,
      executablePath: path.join(path.dirname(validatePreflight.manifestPath), executableName()),
      args,
      cwd
    })));
  }, 30000);

  it("replays all five workflows from a three-member runtime using a canonical Gate G4 lock", async () => {
    const candidate = createGateRuntimeCandidate();
    const checkedInLockBytes = readFileSync(path.join(repoRoot, "docs/miku-project-node-reference-runtime-lock-v1.0.3.json"));
    expect(checkedInLockBytes.toString("utf8")).toBe(`${canonicalJsonText(checkedInGateLock)}\n`);
    expect(sha256RawBytes(checkedInLockBytes)).toEqual({
      algorithm: "sha-256",
      value: "95cd11cc4460348fa066908994430adba5983384c06c75679855120e5c5ea3d5"
    });
    const result = await verifyCliV1ReleaseCandidate({
      root: repoRoot,
      runtimeDirectory: candidate.runtimeDirectory,
      lockPath: candidate.lockPath
    });

    expect(result).toEqual({
      kind: "miku_project_gate_runtime_verification",
      schema_version: "1",
      gate: "G4",
      status: "succeeded",
      distribution_status: "internal-reference-only",
      lock_digest: sha256RawBytes(readFileSync(candidate.lockPath)),
      runtime_manifest_digest: candidate.lock.candidate.runtime_manifest_digest,
      source: candidate.lock.source,
      product: candidate.lock.product,
      workflows: ["validate", "inspect", "plan-change", "apply-change", "verify-artifact"]
    });
  }, 30000);

  it("rejects a lock whose pinned manifest names a different source revision before launch", async () => {
    const candidate = createGateRuntimeCandidate();
    const manifestPath = path.join(candidate.runtimeDirectory, "runtime-manifest.json");
    const manifest = readJson(manifestPath);
    manifest.source.contract.revision = "f".repeat(40);
    manifest.source.runtime.revision = "f".repeat(40);
    writeFileSync(manifestPath, `${canonicalJsonText(manifest)}\n`, "utf8");
    const manifestBytes = readFileSync(manifestPath);
    candidate.lock.candidate.members[0].size_bytes = manifestBytes.length;
    candidate.lock.candidate.members[0].digest = sha256RawBytes(manifestBytes);
    candidate.lock.candidate.runtime_manifest_digest = sha256RawBytes(manifestBytes);
    writeFileSync(candidate.lockPath, `${canonicalJsonText(candidate.lock)}\n`, "utf8");
    let launches = 0;

    await expect(verifyCliV1ReleaseCandidate({
      root: repoRoot,
      runtimeDirectory: candidate.runtimeDirectory,
      lockPath: candidate.lockPath,
      launch: () => {
        launches += 1;
        throw new Error("a source-inconsistent lock must not launch the runtime");
      }
    })).rejects.toMatchObject({
      name: "GateRuntimeVerificationError",
      code: "gate-runtime.lock-manifest-mismatch"
    });
    expect(launches).toBe(0);
  }, 30000);

  it("rejects an internally inconsistent Gate G4 lock before launching the runtime", async () => {
    const candidate = createGateRuntimeCandidate();
    candidate.lock.candidate.runtime_manifest_digest.value = "0".repeat(64);
    writeFileSync(candidate.lockPath, `${canonicalJsonText(candidate.lock)}\n`, "utf8");
    let launches = 0;

    await expect(verifyCliV1ReleaseCandidate({
      root: repoRoot,
      runtimeDirectory: candidate.runtimeDirectory,
      lockPath: candidate.lockPath,
      launch: () => {
        launches += 1;
        throw new Error("an invalid Gate G4 lock must not launch the runtime");
      }
    })).rejects.toMatchObject({
      name: "GateRuntimeVerificationError",
      code: "gate-runtime.lock-invalid"
    });
    expect(launches).toBe(0);
  }, 30000);

  it("rejects invalid or non-identical CLI result runtime bindings from the pinned consumer", async () => {
    const directory = createTemporaryDirectory("miku-project-v1-runtime-consumer-result-binding-");
    const distributed = createExternalConsumerRuntime(directory);
    const projectPath = path.join(distributed.consumerRoot, "dependency-canonical.xml");
    cpSync(flatCanonicalProject, projectPath);
    const preflight = await preflightPinnedNodeRuntimeConsumer({
      runtimeDirectory: distributed.runtimeDirectory,
      trustedManifestDigest: distributed.trustedManifestDigest,
      productVersion: runtimeVersion,
      runtimeVersion,
      conformanceCorpusDigest: computeConformanceCorpusDigest({ root: repoRoot })
    });
    const invoked = spawnSync(process.execPath, [preflight.executablePath, "validate", "--project", projectPath], {
      cwd: distributed.consumerRoot,
      encoding: "utf8"
    });
    expect(invoked.status).toBe(0);
    const validResult = readJsonText(invoked.stdout);
    expect(validateCliResult(validResult)).toBe(true);
    expect(() => assertPinnedNodeRuntimeResultBinding(validResult, preflight)).not.toThrow();

    for (const [label, mutate, expectedSchemaValid, expectedCode] of [
      ["missing runtime field", (result) => {
        delete result.runtime.family;
      }, false, "consumer.result-invalid"],
      ["additional runtime field", (result) => {
        result.runtime.unexpected = "must-not-be-accepted";
      }, false, "consumer.result-invalid"],
      ["additional nested digest field", (result) => {
        result.runtime.artifact_digest.unexpected = "must-not-be-accepted";
      }, false, "consumer.result-invalid"],
      ["schema-valid runtime value mismatch", (result) => {
        result.runtime.version = `${runtimeVersion}-unexpected`;
      }, true, "consumer.result-binding-mismatch"]
    ]) {
      const mutated = JSON.parse(JSON.stringify(validResult));
      mutate(mutated);
      expect(validateCliResult(mutated), label).toBe(expectedSchemaValid);
      expectConsumerResultBindingFailure(mutated, preflight, expectedCode);
    }
  }, 30000);

  it("rejects a manifest and executable coordinated tamper before consumer launch or domain I/O", async () => {
    const directory = createTemporaryDirectory("miku-project-v1-runtime-consumer-tamper-");
    const distributed = createExternalConsumerRuntime(directory);
    const executablePath = path.join(distributed.runtimeDirectory, executableName());
    const manifestPath = path.join(distributed.runtimeDirectory, "runtime-manifest.json");
    const projectPath = path.join(distributed.consumerRoot, "must-not-read-project.xml");
    const requestPath = path.join(distributed.consumerRoot, "must-not-read-request.json");
    const planPath = path.join(distributed.consumerRoot, "must-not-read-plan.json");
    const approvalPath = path.join(distributed.consumerRoot, "must-not-read-approval.json");
    const resultPath = path.join(distributed.consumerRoot, "must-not-create-result.json");
    const destination = path.join(distributed.consumerRoot, "must-not-create-artifact-set");
    const manifest = readJson(manifestPath);

    writeFileSync(executablePath, "\n// coordinated external consumer tamper\n", { flag: "a" });
    const executableBytes = readFileSync(executablePath);
    manifest.artifacts.executable.size_bytes = executableBytes.length;
    manifest.artifacts.executable.digest = sha256RawBytes(executableBytes);
    writeManifest(distributed.runtimeDirectory, manifest);

    const guarded = createProtectedFileSystemGuard([
      projectPath, requestPath, planPath, approvalPath, resultPath, destination
    ]);
    let launches = 0;
    for (const args of [
      [
        "plan-change", "--project", projectPath, "--request", requestPath,
        "--destination", destination, "--result", resultPath
      ],
      [
        "apply-change", "--project", projectPath, "--request", requestPath,
        "--plan-result", planPath, "--approval", approvalPath, "--result", resultPath
      ],
      ["verify-artifact", "--artifact-set", destination, "--expect-plan-result", planPath]
    ]) {
      await expect(runPinnedNodeRuntimeConsumer({
        runtimeDirectory: distributed.runtimeDirectory,
        trustedManifestDigest: distributed.trustedManifestDigest,
        productVersion: runtimeVersion,
        runtimeVersion,
        conformanceCorpusDigest: computeConformanceCorpusDigest({ root: repoRoot }),
        args,
        cwd: distributed.consumerRoot,
        fileSystem: guarded.fileSystem,
        launch: () => {
          launches += 1;
          throw new Error("the consumer must not launch a tampered runtime");
        }
      })).rejects.toMatchObject({ code: "consumer.manifest-pin-mismatch" });
    }

    expect(launches).toBe(0);
    expect(guarded.accesses).toEqual([]);
    expect(existsSync(resultPath)).toBe(false);
    expect(existsSync(destination)).toBe(false);
  }, 30000);

  it.each([
    ["executable size mismatch", (distributed) => {
      writeFileSync(path.join(distributed.runtimeDirectory, executableName()), "\n// executable size tamper\n", { flag: "a" });
    }, "consumer.runtime-executable-size-mismatch"],
    ["sources size mismatch", (distributed) => {
      writeFileSync(path.join(distributed.runtimeDirectory, sourcesName()), "\nsources size tamper\n", { flag: "a" });
    }, "consumer.runtime-sources-size-mismatch"],
    ["executable digest mismatch", (distributed) => {
      replaceOneByte(path.join(distributed.runtimeDirectory, executableName()));
    }, "consumer.runtime-executable-digest-mismatch"],
    ["sources digest mismatch", (distributed) => {
      replaceOneByte(path.join(distributed.runtimeDirectory, sourcesName()));
    }, "consumer.runtime-sources-digest-mismatch"],
    ["missing executable", (distributed) => {
      renameSync(
        path.join(distributed.runtimeDirectory, executableName()),
        path.join(distributed.runtimeDirectory, `${executableName()}.missing`)
      );
    }, "consumer.runtime-executable-entry-invalid"],
    ["symlinked sources", (distributed) => {
      const sourcesPath = path.join(distributed.runtimeDirectory, sourcesName());
      const savedSourcesPath = path.join(distributed.runtimeDirectory, `${sourcesName()}.saved`);
      renameSync(sourcesPath, savedSourcesPath);
      symlinkSync(path.basename(savedSourcesPath), sourcesPath);
    }, "consumer.runtime-sources-entry-invalid"]
  ])("rejects %s before launching Node or accessing domain paths", async (_label, mutate, expectedCode) => {
    const directory = createTemporaryDirectory("miku-project-v1-runtime-consumer-asset-preflight-");
    const distributed = createExternalConsumerRuntime(directory);
    const projectPath = path.join(distributed.consumerRoot, "must-not-read-project.xml");
    const requestPath = path.join(distributed.consumerRoot, "must-not-read-request.json");
    const resultPath = path.join(distributed.consumerRoot, "must-not-create-result.json");
    const destination = path.join(distributed.consumerRoot, "must-not-create-artifact-set");
    mutate(distributed);
    const guarded = createProtectedFileSystemGuard([projectPath, requestPath, resultPath, destination]);
    let launches = 0;

    await expect(runPinnedNodeRuntimeConsumer({
      runtimeDirectory: distributed.runtimeDirectory,
      trustedManifestDigest: distributed.trustedManifestDigest,
      productVersion: runtimeVersion,
      runtimeVersion,
      conformanceCorpusDigest: computeConformanceCorpusDigest({ root: repoRoot }),
      args: [
        "plan-change", "--project", projectPath, "--request", requestPath,
        "--destination", destination, "--result", resultPath
      ],
      cwd: distributed.consumerRoot,
      fileSystem: guarded.fileSystem,
      launch: () => {
        launches += 1;
        throw new Error("the consumer must not launch an asset-mismatched runtime");
      }
    })).rejects.toMatchObject({ code: expectedCode });

    expect(launches).toBe(0);
    expect(guarded.accesses).toEqual([]);
    expect(existsSync(resultPath)).toBe(false);
    expect(existsSync(destination)).toBe(false);
  }, 30000);

  it.each([
    ["missing trust anchor", (distributed) => ({ trustedManifestDigest: null }), "consumer.trust-anchor-invalid"],
    ["malformed trust anchor", (distributed) => ({
      trustedManifestDigest: { algorithm: "sha-256", value: distributed.trustedManifestDigest.value.toUpperCase() }
    }), "consumer.trust-anchor-invalid"],
    ["missing manifest", (distributed) => {
      renameSync(
        path.join(distributed.runtimeDirectory, "runtime-manifest.json"),
        path.join(distributed.runtimeDirectory, "runtime-manifest.json.missing")
      );
      return {};
    }, "consumer.manifest-unavailable"],
    ["symlinked manifest", (distributed) => {
      const manifestPath = path.join(distributed.runtimeDirectory, "runtime-manifest.json");
      const savedManifestPath = path.join(distributed.runtimeDirectory, "runtime-manifest.json.saved");
      renameSync(manifestPath, savedManifestPath);
      symlinkSync(path.basename(savedManifestPath), manifestPath);
      return {};
    }, "consumer.manifest-entry-invalid"]
  ])("rejects %s before consumer launch", async (_label, prepare, expectedCode) => {
    const directory = createTemporaryDirectory("miku-project-v1-runtime-consumer-preflight-");
    const distributed = createExternalConsumerRuntime(directory);
    const overrides = prepare(distributed);
    let launches = 0;

    await expect(runPinnedNodeRuntimeConsumer({
      runtimeDirectory: distributed.runtimeDirectory,
      trustedManifestDigest: distributed.trustedManifestDigest,
      productVersion: runtimeVersion,
      runtimeVersion,
      conformanceCorpusDigest: computeConformanceCorpusDigest({ root: repoRoot }),
      args: ["validate", "--project", path.join(distributed.consumerRoot, "must-not-read.xml")],
      cwd: distributed.consumerRoot,
      launch: () => {
        launches += 1;
        throw new Error("the consumer must not launch an invalid runtime");
      },
      ...overrides
    })).rejects.toMatchObject({ code: expectedCode });

    expect(launches).toBe(0);
  }, 30000);

  it("builds a deterministic manifest-bound runtime and only then runs a verified workflow", () => {
    const directory = createTemporaryDirectory("miku-project-v1-runtime-valid-");
    const first = buildRuntime(directory, "runtime");
    const firstBytes = new Map(["runtime-manifest.json", executableName(), sourcesName()].map((name) => [
      name,
      readFileSync(path.join(first.runtimeDirectory, name))
    ]));
    rmSync(first.runtimeDirectory, { recursive: true, force: true });
    const second = buildRuntime(directory, "runtime");

    for (const name of ["runtime-manifest.json", executableName(), sourcesName()]) {
      expect(firstBytes.get(name)).toEqual(readFileSync(path.join(second.runtimeDirectory, name)));
    }

    const manifest = readJson(second.manifestPath);
    expect(validateRuntimeManifest(manifest)).toBe(true);
    expect(manifest).toMatchObject({
      runtime: { family: "node", role: "reference", version: runtimeVersion, launcher: "node" },
      compatibility: {
        conformance: { fixture_suite_version: "1", corpus_digest: computeConformanceCorpusDigest({ root: repoRoot }) }
      },
      artifacts: {
        executable: { path: executableName(), media_type: "text/javascript" },
        sources: { path: sourcesName(), media_type: "application/gzip" }
      },
      source: {
        contract: { revision: sourceRevision, tag: sourceTag },
        runtime: { revision: sourceRevision, tag: sourceTag }
      },
      reference_runtime: null
    });
    expect(manifest.artifacts.executable.digest).toEqual(sha256RawBytes(readFileSync(path.join(second.runtimeDirectory, executableName()))));
    expect(manifest.artifacts.sources.digest).toEqual(sha256RawBytes(readFileSync(path.join(second.runtimeDirectory, sourcesName()))));
    expect(readFileSync(second.manifestPath, "utf8")).toBe(`${canonicalJsonText(manifest)}\n`);

    const invoked = runRuntime(second.runtimeDirectory, ["validate", "--project", canonicalProject]);
    expect(invoked.status).toBe(0);
    expect(invoked.stderr).toBe("");
    const result = readJsonText(invoked.stdout);
    expect(result).toMatchObject({
      command: "validate",
      status: "succeeded",
      exit_code: 0,
      runtime: {
        binding_status: "verified",
        family: "node",
        version: runtimeVersion,
        artifact_digest: manifest.artifacts.executable.digest,
        manifest_digest: sha256RawBytes(readFileSync(second.manifestPath)),
        capability_profile: "miku-project-cli-core/v1",
        fixture_suite_version: "1"
      }
    });
  }, 20000);

  it("executes all five v1 workflows with one verified manifest binding", () => {
    const directory = createTemporaryDirectory("miku-project-v1-runtime-workflows-");
    const built = buildRuntime(directory, "node");
    const requestPath = path.join(directory, "change-request.json");
    const planResultPath = path.join(directory, "plan-result.json");
    const approvalPath = path.join(directory, "approval.json");
    const applyResultPath = path.join(directory, "apply-result.json");
    const destination = path.join(directory, "artifact-set");
    const manifest = readJson(built.manifestPath);

    const validate = runRuntime(built.runtimeDirectory, ["validate", "--project", flatCanonicalProject]);
    const inspect = runRuntime(built.runtimeDirectory, [
      "inspect", "--project", flatCanonicalProject, "--purpose", "task_change_context", "--task-uid", "2"
    ]);
    expect(validate.status).toBe(0);
    expect(inspect.status).toBe(0);
    const request = JSON.parse(readFileSync(flatChangeRequestTemplate, "utf8").replace(
      "${BASE_STATE_DIGEST}", "a98f0c8b560382234572a61f360c9e96911bc75fc1d57b79968d6e60b5d751d0"
    ));
    writeFileSync(requestPath, `${canonicalJsonText(request)}\n`, "utf8");

    const plan = runRuntime(built.runtimeDirectory, [
      "plan-change", "--project", flatCanonicalProject, "--request", requestPath,
      "--destination", destination, "--result", planResultPath
    ]);
    expect(plan.status).toBe(0);
    const planResult = readJson(planResultPath);
    const expectedOutputRuntime = outputRuntimeBinding(manifest);
    expect(planResult.data.output_plan.runtime).toEqual(expectedOutputRuntime);
    writeFileSync(approvalPath, `${canonicalJsonText(approvalForPlan(request, planResult))}\n`, "utf8");

    const apply = runRuntime(built.runtimeDirectory, [
      "apply-change", "--project", flatCanonicalProject, "--request", requestPath,
      "--plan-result", planResultPath, "--approval", approvalPath, "--result", applyResultPath
    ]);
    expect(apply.status).toBe(0);
    const applyResult = readJson(applyResultPath);
    expect(readJson(path.join(applyResult.data.artifact_set.path, "provenance.json")).runtime).toEqual(expectedOutputRuntime);
    const verify = runRuntime(built.runtimeDirectory, [
      "verify-artifact", "--artifact-set", applyResult.data.artifact_set.path,
      "--expect-plan-result", planResultPath
    ]);
    expect(verify.status).toBe(0);

    for (const result of [readJsonText(validate.stdout), readJsonText(inspect.stdout), planResult, applyResult, readJsonText(verify.stdout)]) {
      expect(result).toMatchObject({
        status: "succeeded",
        runtime: {
          binding_status: "verified",
          artifact_digest: manifest.artifacts.executable.digest,
          manifest_digest: sha256RawBytes(readFileSync(built.manifestPath)),
          capability_profile: "miku-project-cli-core/v1",
          fixture_suite_version: "1"
        }
      });
    }
  }, 20000);

  it.each([
    ["CR-MANIFEST-INVALID-001", (runtimeDirectory, testCase) => {
      const manifest = readJson(path.join(runtimeDirectory, "runtime-manifest.json"));
      const mutation = testCase.parameters.runtime_setup.manifest_mutations[0];
      expect(mutation).toMatchObject({ operation: "replace", pointer: "/artifacts/executable/path" });
      manifest.artifacts.executable.path = mutation.value;
      writeManifest(runtimeDirectory, manifest);
    }],
    ["CR-ASSET-DIGEST-001", (runtimeDirectory, testCase) => {
      const mutation = testCase.parameters.runtime_setup.filesystem_mutations[0];
      expect(mutation).toMatchObject({ operation: "append-content", artifact_role: "executable" });
      writeFileSync(path.join(runtimeDirectory, executableName()), `\n// ${mutation.content}\n`, { flag: "a" });
    }],
    ["CR-CAPABILITY-MISSING-001", (runtimeDirectory, testCase) => {
      const manifest = readJson(path.join(runtimeDirectory, "runtime-manifest.json"));
      const mutation = testCase.parameters.runtime_setup.manifest_mutations[0];
      expect(mutation).toMatchObject({ operation: "remove", pointer: "/compatibility/capabilities/provided/7" });
      manifest.compatibility.capabilities.provided.splice(7, 1);
      writeManifest(runtimeDirectory, manifest);
    }],
    ["CR-SOURCE-MISSING-001", (runtimeDirectory, testCase) => {
      const mutation = testCase.parameters.runtime_setup.filesystem_mutations[0];
      expect(mutation).toMatchObject({ operation: "remove", artifact_role: "sources" });
      renameSync(path.join(runtimeDirectory, sourcesName()), path.join(runtimeDirectory, `${sourcesName()}.missing`));
    }]
  ])("%s fails closed before reading project input or reserving a result", async (caseId, mutate) => {
    const directory = createTemporaryDirectory("miku-project-v1-runtime-integrity-");
    const built = buildRuntime(directory, "node");
    const unreadProject = path.join(directory, "must-not-be-read.xml");
    const resultPath = path.join(directory, "must-not-be-created.json");
    const testCase = conformanceCases.get(caseId);
    expect(testCase).toMatchObject({
      command: "validate",
      expected_status: "runtime-error",
      expected_exit_code: 3,
      parameters: { expected_project_input_read: false }
    });
    expect(testCase.parameters.runtime_setup.manifest_base).toBe("generated-test-runtime");
    mutate(built.runtimeDirectory, testCase);

    const invoked = runRuntime(built.runtimeDirectory, [
      "validate", "--project", unreadProject, "--result", resultPath
    ]);
    expect(invoked.status).toBe(3);
    expect(invoked.stderr).toBe("");
    const result = readJsonText(invoked.stdout);
    expect(result).toMatchObject({
      command: "validate",
      status: "runtime-error",
      exit_code: 3,
      diagnostics: [{ code: testCase.expected_diagnostic_codes[0] }],
      runtime: { binding_status: "unverified" },
      io: {
        inputs: [{ role: "project", path: unreadProject, digest: null }],
        result: { target: "stdout", path: null }
      }
    });
    expect(existsSync(unreadProject)).toBe(false);
    expect(existsSync(resultPath)).toBe(false);

    const guarded = await runRuntimeWithInputResultReadGuard({
      runtimeDirectory: built.runtimeDirectory,
      args: ["validate", "--project", unreadProject, "--result", resultPath],
      cwd: directory,
      protectedPaths: [unreadProject, resultPath]
    });
    expect(guarded.accesses).toEqual([]);
    expect(guarded.result).toMatchObject({
      command: "validate",
      status: "runtime-error",
      exit_code: 3,
      diagnostics: [{ code: testCase.expected_diagnostic_codes[0] }],
      runtime: { binding_status: "unverified" }
    });
    expect(readJsonText(guarded.stdout)).toEqual(guarded.result);
  });

  it("rejects a source archive digest mismatch with the same no-input/no-result boundary", async () => {
    const directory = createTemporaryDirectory("miku-project-v1-runtime-source-digest-");
    const built = buildRuntime(directory, "node");
    const unreadProject = path.join(directory, "must-not-be-read.xml");
    const resultPath = path.join(directory, "must-not-be-created.json");
    writeFileSync(path.join(built.runtimeDirectory, sourcesName()), "\nsource-digest-mismatch\n", { flag: "a" });

    const invoked = runRuntime(built.runtimeDirectory, [
      "validate", "--project", unreadProject, "--result", resultPath
    ]);
    expect(invoked.status).toBe(3);
    const result = readJsonText(invoked.stdout);
    expect(result).toMatchObject({
      command: "validate",
      status: "runtime-error",
      diagnostics: [{ code: "runtime.manifest-invalid" }],
      runtime: { binding_status: "unverified" },
      io: { result: { target: "stdout", path: null } }
    });
    expect(existsSync(unreadProject)).toBe(false);
    expect(existsSync(resultPath)).toBe(false);

    const guarded = await runRuntimeWithInputResultReadGuard({
      runtimeDirectory: built.runtimeDirectory,
      args: ["validate", "--project", unreadProject, "--result", resultPath],
      cwd: directory,
      protectedPaths: [unreadProject, resultPath]
    });
    expect(guarded.accesses).toEqual([]);
    expect(guarded.result).toMatchObject({
      command: "validate",
      status: "runtime-error",
      exit_code: 3,
      diagnostics: [{ code: "runtime.manifest-invalid" }],
      runtime: { binding_status: "unverified" }
    });
    expect(readJsonText(guarded.stdout)).toEqual(guarded.result);
  });
});

function createExternalConsumerRuntime(parentDirectory) {
  const releaseSource = createCleanTaggedReleaseSource(parentDirectory);
  const built = buildReleaseNodeRuntime({ root: releaseSource });
  const consumerRoot = path.join(parentDirectory, "consumer");
  const runtimeDirectory = path.join(consumerRoot, "runtime");
  const distributedMembers = ["runtime-manifest.json", executableName(), sourcesName()];
  mkdirSync(runtimeDirectory, { recursive: true });
  for (const member of distributedMembers) {
    cpSync(path.join(built.runtimeDirectory, member), path.join(runtimeDirectory, member));
  }
  rmSync(releaseSource, { recursive: true, force: true });
  expect(existsSync(releaseSource)).toBe(false);
  expect(readdirSync(runtimeDirectory).sort()).toEqual([...distributedMembers].sort());
  return {
    consumerRoot,
    runtimeDirectory,
    trustedManifestDigest: { ...built.manifestDigest }
  };
}

function createProtectedFileSystemGuard(protectedPaths) {
  const accesses = [];
  const protectedPathSet = new Set(protectedPaths.map((candidate) => path.resolve(candidate)));
  return {
    accesses,
    fileSystem: {
      lstat: guardFileSystemAccess("lstat"),
      realpath: guardFileSystemAccess("realpath"),
      readFile: guardFileSystemAccess("readFile")
    }
  };

  function guardFileSystemAccess(operation) {
    return async (candidate, ...rest) => {
      const resolvedPath = path.resolve(String(candidate));
      if (protectedPathSet.has(resolvedPath)) {
        accesses.push({ operation, path: resolvedPath });
        throw new Error(`unexpected ${operation} for protected consumer domain path: ${resolvedPath}`);
      }
      return fsPromises[operation](candidate, ...rest);
    };
  }
}

function replaceOneByte(filePath) {
  const bytes = Buffer.from(readFileSync(filePath));
  bytes[0] ^= 0x01;
  writeFileSync(filePath, bytes);
}

function buildRuntime(parentDirectory, name) {
  return buildNodeRuntimeForTest({
    root: repoRoot,
    outDir: path.join(parentDirectory, name),
    sourceRevision,
    sourceTag,
    packageVersion: runtimeVersion,
    runtimeVersion
  });
}

function runRuntime(runtimeDirectory, args, { cwd = repoRoot } = {}) {
  return spawnSync(process.execPath, [path.join(runtimeDirectory, executableName()), ...args], {
    cwd,
    encoding: "utf8"
  });
}

async function runRuntimeWithInputResultReadGuard({ runtimeDirectory, args, cwd, protectedPaths }) {
  const accesses = [];
  const protectedPathSet = new Set(protectedPaths.map((candidate) => path.resolve(candidate)));
  const fileSystem = {
    lstat: guardFileSystemAccess("lstat"),
    realpath: guardFileSystemAccess("realpath"),
    readFile: guardFileSystemAccess("readFile")
  };
  let stdout = "";
  const result = await runV1VersionedRuntime(args, {
    entryPath: path.join(runtimeDirectory, executableName()),
    runtimeVersion,
    productVersion: runtimeVersion,
    conformanceCorpusDigest: computeConformanceCorpusDigest({ root: repoRoot }),
    cwd,
    stdout: {
      write(chunk) {
        stdout += String(chunk);
        return true;
      }
    },
    fileSystem
  });
  return { accesses, result, stdout };

  function guardFileSystemAccess(operation) {
    return async (candidate, ...rest) => {
      const resolvedPath = path.resolve(String(candidate));
      if (protectedPathSet.has(resolvedPath)) {
        accesses.push({ operation, path: resolvedPath });
        throw new Error(`unexpected ${operation} for protected runtime boundary path: ${resolvedPath}`);
      }
      return fsPromises[operation](candidate, ...rest);
    };
  }
}

function writeManifest(runtimeDirectory, manifest) {
  writeFileSync(path.join(runtimeDirectory, "runtime-manifest.json"), `${canonicalJsonText(manifest)}\n`, "utf8");
}

function executableName() {
  return `miku-project-node-${runtimeVersion}.mjs`;
}

function sourcesName() {
  return `miku-project-node-${runtimeVersion}-sources.tgz`;
}

function createTemporaryDirectory(prefix) {
  const directory = mkdtempSync(path.join(os.tmpdir(), prefix));
  temporaryDirectories.push(directory);
  return directory;
}

function createMiniGitRoot(parentDirectory, name, { tag, dirty }) {
  const root = path.join(parentDirectory, name);
  mkdirSync(root, { recursive: true });
  writeFileSync(path.join(root, "package.json"), `${JSON.stringify({ name: "miku-project", version: runtimeVersion })}\n`, "utf8");
  runGit(root, ["init"]);
  runGit(root, ["config", "user.email", "tests@example.invalid"]);
  runGit(root, ["config", "user.name", "miku-project test"]);
  runGit(root, ["add", "package.json"]);
  runGit(root, ["commit", "-m", "test source"]);
  runGit(root, ["tag", tag]);
  if (dirty) {
    writeFileSync(path.join(root, "dirty.txt"), "not a release source\n", "utf8");
  }
  return root;
}

function createCleanTaggedReleaseSource(parentDirectory) {
  const root = path.join(parentDirectory, "release-source");
  cpSync(repoRoot, root, {
    recursive: true,
    filter: (sourcePath) => {
      const firstSegment = path.relative(repoRoot, sourcePath).split(path.sep)[0];
      return ![".git", "bundle", "coverage", "local-data", "node_modules", "runtime", "workplace"].includes(firstSegment);
    }
  });
  cpSync(path.join(repoRoot, "node_modules", "@xmldom"), path.join(root, "node_modules", "@xmldom"), { recursive: true });
  runGit(root, ["init"]);
  runGit(root, ["config", "user.email", "tests@example.invalid"]);
  runGit(root, ["config", "user.name", "miku-project test"]);
  runGit(root, ["add", "-A"]);
  runGit(root, ["commit", "-m", "release source"]);
  runGit(root, ["tag", sourceTag]);
  return root;
}

function createGateRuntimeCandidate() {
  const directory = createTemporaryDirectory("miku-project-v1-gate-runtime-candidate-");
  const runtimeDirectory = path.join(directory, "runtime");
  const built = buildNodeRuntimeForTest({
    root: repoRoot,
    outDir: runtimeDirectory,
    sourceRevision: checkedInGateLock.source.revision,
    sourceTag: checkedInGateLock.source.tag,
    // The checked-in Gate G4 lock is a frozen v1.0.3 release record.  Do not
    // inherit the working tree's current package version when reproducing its
    // candidate shape; both identities must remain pinned to the historical
    // lock while member digests are freshly generated for this test.
    packageVersion: checkedInGateLock.product.release_version,
    runtimeVersion: checkedInGateLock.product.runtime_version
  });
  const manifestBytes = readFileSync(built.manifestPath);
  const manifest = JSON.parse(manifestBytes);
  const lock = structuredClone(checkedInGateLock);
  // This is a synthetic candidate lock used to exercise the verifier against
  // the current checkout.  Its build-evidence digest must describe that
  // checkout; the checked-in v1.0.3 lock remains immutable release history.
  lock.build.toolchain.package_lock_digest = sha256RawBytes(
    readFileSync(path.join(repoRoot, "package-lock.json"))
  );
  lock.candidate.members = [
    {
      role: "manifest",
      path: "runtime-manifest.json",
      media_type: "application/json",
      size_bytes: manifestBytes.length,
      digest: sha256RawBytes(manifestBytes)
    },
    {
      role: "executable",
      ...manifest.artifacts.executable
    },
    {
      role: "sources",
      ...manifest.artifacts.sources
    }
  ];
  lock.candidate.runtime_manifest_digest = sha256RawBytes(manifestBytes);
  lock.conformance.capabilities = [...manifest.compatibility.capabilities.provided];
  lock.conformance.corpus_digest = { ...manifest.compatibility.conformance.corpus_digest };
  lock.conformance.fixture_suite_version = manifest.compatibility.conformance.fixture_suite_version;
  lock.source = { ...manifest.source.runtime };
  const lockPath = path.join(directory, "runtime-lock.json");
  writeFileSync(lockPath, `${canonicalJsonText(lock)}\n`, "utf8");
  return { lock, lockPath, runtimeDirectory };
}

function runGit(root, args) {
  execFileSync("git", args, { cwd: root, encoding: "utf8" });
}

function runGitText(root, args) {
  return execFileSync("git", args, { cwd: root, encoding: "utf8" }).trim();
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

function outputRuntimeBinding(manifest) {
  return {
    family: "node",
    version: runtimeVersion,
    artifact_digest: { ...manifest.artifacts.executable.digest },
    manifest_digest: sha256RawBytes(Buffer.from(`${canonicalJsonText(manifest)}\n`, "utf8")),
    capability_profile: "miku-project-cli-core/v1",
    fixture_suite_version: "1"
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

function expectConsumerResultBindingFailure(result, preflight, expectedCode) {
  try {
    assertPinnedNodeRuntimeResultBinding(result, preflight);
  } catch (error) {
    expect(error).toMatchObject({ code: expectedCode });
    return;
  }
  throw new Error(`Expected pinned consumer result binding to fail with ${expectedCode}.`);
}

function readJson(filePath) {
  return readJsonText(readFileSync(filePath, "utf8"));
}

function readJsonText(text) {
  return JSON.parse(text);
}
