import { spawnSync } from "node:child_process";

const testGroups = [
  [
    "tests/mikuproject-excel-io.test.js",
    "tests/mikuproject-project-xlsx.test.js",
    "tests/mikuproject-wbs-xlsx.test.js",
    "lht-cmn/components.test.js"
  ],
  [
    "tests/mikuproject-main.test.js"
  ]
];

for (const files of testGroups) {
  const result = spawnSync(
    process.execPath,
    ["./node_modules/vitest/vitest.mjs", "run", "--testTimeout=15000", "--hookTimeout=15000", ...files],
    {
      stdio: "inherit",
      cwd: process.cwd()
    }
  );
  if (result.status !== 0) {
    process.exit(result.status ?? 1);
  }
}
