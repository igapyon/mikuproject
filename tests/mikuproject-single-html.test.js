import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { buildSingleHtmlFromSource } from "../scripts/lib/single-html.mjs";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const ROOT = path.resolve(__dirname, "..");

describe("mikuproject single html build", () => {
  it("inlines vendored mermaid runtime into mikuproject html", () => {
    const srcHtmlPath = path.resolve(ROOT, "mikuproject-src.html");
    const sourceHtml = readFileSync(srcHtmlPath, "utf8");
    const builtHtml = buildSingleHtmlFromSource(sourceHtml, srcHtmlPath);
    const vendoredMermaid = readFileSync(
      path.resolve(ROOT, "src/vendor/mermaid/mermaid.min.js"),
      "utf8"
    ).trimEnd();

    expect(builtHtml).toContain(vendoredMermaid.slice(0, 120));
    expect(builtHtml).not.toContain('src="src/vendor/mermaid/mermaid.min.js"');
    expect(builtHtml).not.toContain('src="src/js/main.js"');
  });
});
