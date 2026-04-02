import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const mainUtilCode = readFileSync(
  path.resolve(__dirname, "../src/js/main-util.js"),
  "utf8"
);

function bootMainUtil() {
  new Function(mainUtilCode)();
  return globalThis.__mikuprojectMainUtil;
}

describe("mikuproject main util", () => {
  it("parses optional non-negative integer values", () => {
    const util = bootMainUtil();

    expect(util.parseOptionalNonNegativeInteger("")).toBeUndefined();
    expect(util.parseOptionalNonNegativeInteger("abc")).toBeUndefined();
    expect(util.parseOptionalNonNegativeInteger("-3")).toBe(0);
    expect(util.parseOptionalNonNegativeInteger("7.9")).toBe(7);
    expect(util.parseOptionalNonNegativeInteger(" 12 ")).toBe(12);
  });

  it("formats compact and save timestamps", () => {
    const util = bootMainUtil();
    const date = new Date("2026-04-02T08:59:00+09:00");

    expect(util.formatTimestampCompact(date)).toBe("202604020859");
    expect(util.formatSaveStamp(date)).toBe("2026-04-02 08:59");
  });

  it("packs zip entries with valid signatures", () => {
    const util = bootMainUtil();
    const zipBytes = util.packZipEntries([
      { name: "README.txt", data: util.encodeUtf8("hello\n") },
      { name: "nested/data.json", data: util.encodeUtf8("{\"ok\":true}\n") }
    ]);
    const view = new DataView(zipBytes.buffer, zipBytes.byteOffset, zipBytes.byteLength);

    expect(view.getUint32(0, true)).toBe(0x04034b50);
    expect(view.getUint32(zipBytes.byteLength - 22, true)).toBe(0x06054b50);
    expect(util.computeZipCrc32(util.encodeUtf8("hello\n"))).toBe(0x363a3020);
  });
});
