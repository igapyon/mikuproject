import { spawnSync } from "node:child_process";

const SUITES = {
  fast: [
    "tests/mikuproject-ai-json-util.test.js",
    "tests/mikuproject-main-util.test.js",
    "tests/mikuproject-excel-io.test.js",
    "tests/mikuproject-msproject-xml-roundtrip.test.js",
    "tests/mikuproject-project-workbook-json.test.js",
    "tests/mikuproject-project-xlsx.test.js",
    "tests/mikuproject-wbs-markdown.test.js",
    "tests/mikuproject-wbs-xlsx.test.js",
    "tests/mikuproject-single-html.test.js",
    "lht-cmn/components.test.js"
  ],
  ui: [
    "tests/mikuproject-main-file-input-wiring.test.js",
    "tests/mikuproject-main-ai-json.test.js",
    "tests/mikuproject-main-xlsx-import.test.js",
    "tests/mikuproject-main-validation.test.js",
    "tests/mikuproject-main-preview-export.test.js",
    "tests/mikuproject-main.test.js"
  ],
  full: []
};

const requestedSuite = process.argv[2] || "full";
if (!Object.hasOwn(SUITES, requestedSuite)) {
  console.error(`[run-tests] unknown suite: ${requestedSuite}`);
  console.error("[run-tests] expected one of: fast, ui, full");
  process.exit(1);
}

const extraArgs = process.argv.slice(3);
const vitestArgs = [
  "./node_modules/vitest/vitest.mjs",
  "run",
  ...extraArgs
];

if (requestedSuite !== "full") {
  vitestArgs.push(...SUITES[requestedSuite]);
}

const result = spawnSync(process.execPath, vitestArgs, {
  stdio: "inherit",
  cwd: process.cwd()
});

process.exit(result.status ?? 1);
