import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { afterEach, describe, expect, it } from "vitest";

import { loadMikuprojectCoreApi } from "../scripts/lib/core-api-loader.mjs";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..");
const dependencyXml = readFileSync(
  path.resolve(repoRoot, "testdata/dependency.xml"),
  "utf8"
);

const disposers = [];

afterEach(() => {
  while (disposers.length > 0) {
    disposers.pop()();
  }
});

describe("core api loader", () => {
  it("boots __mikuprojectCoreApi for node-side reuse", () => {
    const loaded = loadMikuprojectCoreApi({ rootDir: repoRoot });
    disposers.push(() => loaded.dispose());

    const result = loaded.api.importExternal({
      source: { format: "ms_project_xml", text: dependencyXml },
      mode: "replace"
    });
    const reportBundle = loaded.api.report.all.export(result.model);

    expect(result.kind).toBe("ms_project_xml");
    expect(result.model.tasks.length).toBeGreaterThan(0);
    expect(reportBundle.zipBytes).toBeInstanceOf(Uint8Array);
    expect(reportBundle.entries.some((entry) => entry.name === "wbs.xlsx")).toBe(true);
  });
});
