export const FAST_SUITE = [
  "tests/mikuproject-ai-json-util.test.js",
  "tests/mikuproject-main-util.test.js",
  "tests/mikuproject-excel-io.test.js",
  "tests/mikuproject-msproject-xml-roundtrip.test.js",
  "tests/mikuproject-project-workbook-json.test.js",
  "tests/mikuproject-project-xlsx.test.js",
  "tests/mikuproject-core-api.test.js",
  "tests/mikuproject-core-api-loader.test.js",
  "tests/mikuproject-cli-argv.test.js",
  "tests/mikuproject-cli-io.test.js",
  "tests/mikuproject-cli-diagnostics.test.js",
  "tests/mikuproject-cli-compatibility-contract.test.js",
  "tests/mikuproject-wbs-markdown.test.js",
  "tests/mikuproject-wbs-xlsx.test.js",
  "tests/mikuproject-test-suite-topology.test.js"
];

export const FULL_ONLY_SUITE = [
  "tests/miku-project-browser-runtime.test.js",
  "tests/mikuproject-cli.test.js"
];

export const FULL_SUITE = [...FAST_SUITE, ...FULL_ONLY_SUITE];

// `all` is the stable complete-suite alias while every repository test file is
// assigned to either FAST_SUITE or FULL_ONLY_SUITE.
export const ALL_SUITE = [...FULL_SUITE];

export const SUITES = Object.freeze({
  fast: FAST_SUITE,
  full: FULL_SUITE,
  all: ALL_SUITE
});
