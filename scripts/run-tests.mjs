import { spawnSync } from "node:child_process";

const FAST_SUITE = [
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
];

const UI_SUITE = [
  "tests/mikuproject-main-file-input-wiring.test.js",
  "tests/mikuproject-main.test.js"
];

const EXTENDED_SUITE = [
  "tests/mikuproject-main-validation.test.js",
  "tests/mikuproject-main-xlsx-import.test.js",
  "tests/mikuproject-main-preview-svg.test.js",
  "tests/mikuproject-main-ai-json.test.js",
  "tests/mikuproject-main-ai-json-patch-task.test.js",
  "tests/mikuproject-main-ai-json-patch-task-structure.test.js",
  "tests/mikuproject-main-ai-json-patch-task-links.test.js",
  "tests/mikuproject-main-ai-json-patch-models.test.js",
  "tests/mikuproject-main-ai-json-patch-assignments.test.js",
  "tests/mikuproject-main-xlsx-import-replace.test.js",
  "tests/mikuproject-main-preview-export.test.js"
];

const SUITES = {
  fast: FAST_SUITE,
  ui: UI_SUITE,
  extended: EXTENDED_SUITE,
  full: [...FAST_SUITE, ...UI_SUITE],
  all: [...FAST_SUITE, ...UI_SUITE, ...EXTENDED_SUITE]
};

const requestedSuite = process.argv[2] || "all";
if (!Object.hasOwn(SUITES, requestedSuite)) {
  console.error(`[run-tests] unknown suite: ${requestedSuite}`);
  console.error("[run-tests] expected one of: fast, ui, extended, full, all");
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
  cwd: process.cwd()
});

process.exit(result.status ?? 1);
