import { spawnSync } from "node:child_process";

const FAST_SUITE = [
  "tests/mikuproject-ai-json-util.test.js",
  "tests/mikuproject-main-util.test.js",
  "tests/mikuproject-excel-io.test.js",
  "tests/mikuproject-msproject-xml-roundtrip.test.js",
  "tests/mikuproject-project-workbook-json.test.js",
  "tests/mikuproject-project-xlsx.test.js",
  "tests/miku-project-browser-runtime.test.js",
  "tests/mikuproject-cli.test.js",
  "tests/mikuproject-wbs-markdown.test.js",
  "tests/mikuproject-wbs-xlsx.test.js"
];

const SUITES = {
  fast: FAST_SUITE,
  full: FAST_SUITE,
  all: FAST_SUITE
};

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
