#!/usr/bin/env node

import { runV1R1Harness } from "../../scripts/lib/v1/cli-v1-router.mjs";

const TEST_RUNTIME = Object.freeze({
  binding_status: "verified",
  family: "node",
  version: "1.0.2",
  artifact_digest: digest("0".repeat(64)),
  manifest_digest: digest("1".repeat(64)),
  capability_profile: "miku-project-cli-core/v1",
  fixture_suite_version: "1"
});

try {
  const result = await runV1R1Harness(process.argv.slice(2), { runtime: TEST_RUNTIME });
  process.exitCode = result.exit_code;
} catch (error) {
  // This helper is test-only.  An unstructured failure here means its service
  // boundary is broken, rather than being a user-visible v1 CLI result.
  process.stderr.write(`[miku-project-v1-r1-harness] ${error instanceof Error ? error.message : String(error)}\n`);
  process.exitCode = 99;
}

function digest(value) {
  return { algorithm: "sha-256", value };
}
