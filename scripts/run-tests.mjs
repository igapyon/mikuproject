import { spawnSync } from "node:child_process";
import { SUITES } from "./lib/test-suite-topology.mjs";

const requestedSuite = process.argv[2] || "all";
if (!Object.hasOwn(SUITES, requestedSuite)) {
  console.error(`[run-tests] unknown suite: ${requestedSuite}`);
  console.error("[run-tests] expected one of: fast, full, all");
  process.exit(1);
}

const extraArgs = process.argv.slice(3);
const vitestArgs = [
  "./node_modules/vitest/vitest.mjs",
  "run",
  ...extraArgs
];

vitestArgs.push(...SUITES[requestedSuite]);

const result = spawnSync(process.execPath, vitestArgs, {
  stdio: "inherit",
  cwd: process.cwd(),
  env: {
    ...process.env,
    TZ: "Asia/Tokyo"
  }
});

process.exit(result.status ?? 1);
