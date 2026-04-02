import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { beforeEach, describe, expect, it } from "vitest";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const aiJsonUtilCode = readFileSync(
  path.resolve(__dirname, "../src/js/ai-json-util.js"),
  "utf8"
);

function bootUtil() {
  new Function(aiJsonUtilCode)();
  return globalThis.__mikuprojectAiJsonUtil;
}

describe("mikuproject ai json util", () => {
  beforeEach(() => {
    delete globalThis.__mikuprojectAiJsonUtil;
  });

  it("extracts the last json fenced block", () => {
    const util = bootUtil();

    const result = util.extractLastJsonBlock([
      "before",
      "```json",
      "{\"a\":1}",
      "```",
      "middle",
      "```json",
      "{\"b\":2}",
      "```"
    ].join("\n"));

    expect(result).toBe("{\"b\":2}");
  });

  it("falls back to trimmed raw text when no json fence exists", () => {
    const util = bootUtil();

    expect(util.extractLastJsonBlock("  {\"a\":1}  ")).toBe("{\"a\":1}");
  });

  it("detects workbook json", () => {
    const util = bootUtil();

    expect(util.detectJsonDocumentKind({ format: "mikuproject_workbook_json" })).toBe("workbook_json");
  });

  it("detects project draft view", () => {
    const util = bootUtil();

    expect(util.detectJsonDocumentKind({ view_type: "project_draft_view" })).toBe("project_draft_view");
  });

  it("detects patch json", () => {
    const util = bootUtil();

    expect(util.detectJsonDocumentKind({ operations: [] })).toBe("patch_json");
  });

  it("returns undefined for unknown json shapes", () => {
    const util = bootUtil();

    expect(util.detectJsonDocumentKind({ foo: "bar" })).toBeUndefined();
    expect(util.detectJsonDocumentKind(null)).toBeUndefined();
    expect(util.detectJsonDocumentKind("text")).toBeUndefined();
  });
});
