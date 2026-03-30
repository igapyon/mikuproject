import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { buildSingleHtmlFromSource } from "../scripts/lib/single-html.mjs";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const ROOT = path.resolve(__dirname, "..");

describe("mikuproject single html build", () => {
  it("inlines local app assets and does not reference removed mermaid runtime", () => {
    const srcHtmlPath = path.resolve(ROOT, "mikuproject-src.html");
    const sourceHtml = readFileSync(srcHtmlPath, "utf8");
    const builtHtml = buildSingleHtmlFromSource(sourceHtml, srcHtmlPath);

    expect(builtHtml).not.toContain('src="src/js/main.js"');
    expect(builtHtml).not.toContain("src/vendor/mermaid/");
    expect(builtHtml).not.toContain("mermaid.min.js");
    expect(builtHtml).toContain("function initialize()");
  });
});
