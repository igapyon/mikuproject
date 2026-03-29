// @vitest-environment jsdom

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { deflateRawSync } from "node:zlib";

import { describe, expect, it } from "vitest";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const typesCode = readFileSync(
  path.resolve(__dirname, "../src/js/types.js"),
  "utf8"
);
const excelIoCode = readFileSync(
  path.resolve(__dirname, "../src/js/excel-io.js"),
  "utf8"
);

function bootExcelIoModule() {
  new Function(`${typesCode}\n${excelIoCode}`)();
  return globalThis.__mikuprojectExcelIo;
}

function decodeUtf8(bytes) {
  return new TextDecoder().decode(bytes);
}

function encodeUtf8(text) {
  return new TextEncoder().encode(text);
}

function buildDeflatedZipWithSingleEntry(name, text) {
  const nameBytes = encodeUtf8(name);
  const rawData = encodeUtf8(text);
  const compressedData = new Uint8Array(deflateRawSync(rawData));

  const localHeader = new Uint8Array(30 + nameBytes.length);
  const localView = new DataView(localHeader.buffer);
  localView.setUint32(0, 0x04034b50, true);
  localView.setUint16(4, 20, true);
  localView.setUint16(6, 0, true);
  localView.setUint16(8, 8, true);
  localView.setUint32(14, 0, true);
  localView.setUint32(18, compressedData.byteLength, true);
  localView.setUint32(22, rawData.byteLength, true);
  localView.setUint16(26, nameBytes.length, true);
  localView.setUint16(28, 0, true);
  localHeader.set(nameBytes, 30);

  const centralHeader = new Uint8Array(46 + nameBytes.length);
  const centralView = new DataView(centralHeader.buffer);
  centralView.setUint32(0, 0x02014b50, true);
  centralView.setUint16(4, 20, true);
  centralView.setUint16(6, 20, true);
  centralView.setUint16(8, 0, true);
  centralView.setUint16(10, 8, true);
  centralView.setUint32(16, 0, true);
  centralView.setUint32(20, compressedData.byteLength, true);
  centralView.setUint32(24, rawData.byteLength, true);
  centralView.setUint16(28, nameBytes.length, true);
  centralView.setUint16(30, 0, true);
  centralView.setUint16(32, 0, true);
  centralView.setUint16(34, 0, true);
  centralView.setUint16(36, 0, true);
  centralView.setUint32(38, 0, true);
  centralView.setUint32(42, 0, true);
  centralHeader.set(nameBytes, 46);

  const end = new Uint8Array(22);
  const endView = new DataView(end.buffer);
  endView.setUint32(0, 0x06054b50, true);
  endView.setUint16(4, 0, true);
  endView.setUint16(6, 0, true);
  endView.setUint16(8, 1, true);
  endView.setUint16(10, 1, true);
  endView.setUint32(12, centralHeader.byteLength, true);
  endView.setUint32(16, localHeader.byteLength + compressedData.byteLength, true);
  endView.setUint16(20, 0, true);

  const bytes = new Uint8Array(localHeader.byteLength + compressedData.byteLength + centralHeader.byteLength + end.byteLength);
  let offset = 0;
  bytes.set(localHeader, offset);
  offset += localHeader.byteLength;
  bytes.set(compressedData, offset);
  offset += compressedData.byteLength;
  bytes.set(centralHeader, offset);
  offset += centralHeader.byteLength;
  bytes.set(end, offset);
  return bytes;
}

describe("mikuproject excel io", () => {
  it("exports a minimal xlsx package with required workbook entries", () => {
    const excelIo = bootExcelIoModule();
    const codec = new excelIo.XlsxWorkbookCodec();
    const workbook = {
      sheets: [
        {
          name: "Project",
          rows: [
            {
              cells: [
                { value: "Name" },
                { value: "Miku Project" }
              ]
            }
          ]
        }
      ]
    };

    const bytes = codec.exportWorkbook(workbook);
    const entryNames = codec.listEntries(bytes);

    expect(bytes).toBeInstanceOf(Uint8Array);
    expect(bytes.byteLength).toBeGreaterThan(0);
    expect(entryNames).toEqual([
      "[Content_Types].xml",
      "_rels/.rels",
      "xl/_rels/workbook.xml.rels",
      "xl/styles.xml",
      "xl/workbook.xml",
      "xl/worksheets/sheet1.xml"
    ]);
  });

  it("round-trips sheet names, sparse cells, formulas, and primitive cell values as text by default", () => {
    const excelIo = bootExcelIoModule();
    const codec = new excelIo.XlsxWorkbookCodec();
    const workbook = {
      sheets: [
        {
          name: "Project",
          rows: [
            {
              cells: [
                { value: "Task" },
                { value: "Days" },
                { value: "Done" },
                { value: "Formula" }
              ]
            },
            {
              cells: [
                { value: "Design" },
                { value: "2" },
                { value: "true" },
                { formula: "B2*2", value: 4 }
              ]
            },
            {
              cells: [
                { value: "Build" },
                {},
                { value: "false" },
                { value: "" }
              ]
            }
          ]
        },
        {
          name: "Resources",
          rows: [
            {
              cells: [
                { value: "Name" },
                { value: "Role" }
              ]
            },
            {
              cells: [
                { value: "Miku" },
                { value: "Engineer" }
              ]
            }
          ]
        }
      ]
    };

    const bytes = codec.exportWorkbook(workbook);
    const imported = codec.importWorkbook(bytes);

    expect(imported).toEqual(workbook);
  });

  it("writes string cells as text-formatted inline strings and preserves leading whitespace", () => {
    const excelIo = bootExcelIoModule();
    const codec = new excelIo.XlsxWorkbookCodec();
    const workbook = {
      sheets: [{
        name: "Sheet1",
        rows: [{
          cells: [
            { value: "  - round-trip" },
            { value: "plain" }
          ]
        }]
      }]
    };

    const bytes = codec.exportWorkbook(workbook);
    const entries = codec.unpackEntries(bytes);
    const sheetXml = new TextDecoder().decode(entries["xl/worksheets/sheet1.xml"]);
    const stylesXml = new TextDecoder().decode(entries["xl/styles.xml"]);

    expect(sheetXml).toContain('t="inlineStr"><is><t xml:space="preserve">  - round-trip</t></is></c>');
    expect(stylesXml).toContain('numFmtId="49"');
    expect(codec.importWorkbook(bytes)).toEqual(workbook);
  });

  it("sanitizes XML-invalid control characters from cell strings", () => {
    const excelIo = bootExcelIoModule();
    const codec = new excelIo.XlsxWorkbookCodec();
    const workbook = {
      sheets: [{
        name: "Sheet1",
        rows: [{
          cells: [
            { value: "ok\u0000bad\u0008text" },
            { value: "line1\nline2\tok" }
          ]
        }]
      }]
    };

    const bytes = codec.exportWorkbook(workbook);
    const entries = codec.unpackEntries(bytes);
    const sheetXml = new TextDecoder().decode(entries["xl/worksheets/sheet1.xml"]);
    const imported = codec.importWorkbook(bytes);

    expect(sheetXml).toContain("okbadtext");
    expect(sheetXml).not.toContain("\u0000");
    expect(sheetXml).not.toContain("\u0008");
    expect(imported.sheets[0].rows[0].cells[0].value).toBe("okbadtext");
    expect(imported.sheets[0].rows[0].cells[1].value).toBe("line1\nline2\tok");
  });

  it("keeps explicitly formatted numeric cells as numeric values", () => {
    const excelIo = bootExcelIoModule();
    const codec = new excelIo.XlsxWorkbookCodec();
    const workbook = {
      sheets: [{
        name: "Numeric",
        rows: [{
          cells: [
            { value: 12, numberFormat: "integer" },
            { value: 0.5, numberFormat: "percent" },
            { formula: "A1*2", value: 24 }
          ]
        }]
      }]
    };

    const bytes = codec.exportWorkbook(workbook);
    const imported = codec.importWorkbook(bytes);

    expect(imported).toEqual(workbook);
  });

  it("round-trips column widths, row heights, and basic cell formatting", () => {
    const excelIo = bootExcelIoModule();
    const codec = new excelIo.XlsxWorkbookCodec();
    const workbook = {
      sheets: [
        {
          name: "Schedule",
          columns: [
            { width: 24 },
            { width: 12 },
            { width: 14 }
          ],
          rows: [
            {
              height: 28,
              cells: [
                { value: "Task", horizontalAlign: "center" },
                { value: "Start", horizontalAlign: "center" },
                { value: "Progress", horizontalAlign: "center" }
              ]
            },
            {
              cells: [
                { value: "Design" },
                { value: 45367, numberFormat: "date" },
                { value: 0.5, numberFormat: "percent" }
              ]
            }
          ]
        }
      ]
    };

    const bytes = codec.exportWorkbook(workbook);
    const imported = codec.importWorkbook(bytes);

    expect(imported).toEqual(workbook);
  });

  it("round-trips hidden columns", () => {
    const excelIo = bootExcelIoModule();
    const codec = new excelIo.XlsxWorkbookCodec();
    const workbook = {
      sheets: [
        {
          name: "HiddenCols",
          columns: [
            { width: 12 },
            { width: 18, hidden: true },
            { hidden: true }
          ],
          rows: [
            {
              cells: [
                { value: "Visible" },
                { value: "Hidden" },
                { value: "HiddenNoWidth" }
              ]
            }
          ]
        }
      ]
    };

    const bytes = codec.exportWorkbook(workbook);
    const imported = codec.importWorkbook(bytes);
    const sheetXml = decodeUtf8(codec.unpackEntries(bytes)["xl/worksheets/sheet1.xml"]);

    expect(imported).toEqual(workbook);
    expect(sheetXml).toContain('min="2" max="2" width="18" customWidth="1" hidden="1"');
    expect(sheetXml).toContain('min="3" max="3" hidden="1"');
  });

  it("round-trips wrapped text alignment", () => {
    const excelIo = bootExcelIoModule();
    const codec = new excelIo.XlsxWorkbookCodec();
    const workbook = {
      sheets: [
        {
          name: "Wrapped",
          rows: [
            {
              cells: [
                { value: "Long task name", horizontalAlign: "left", wrapText: true }
              ]
            }
          ]
        }
      ]
    };

    const bytes = codec.exportWorkbook(workbook);
    const imported = codec.importWorkbook(bytes);
    const stylesXml = decodeUtf8(codec.unpackEntries(bytes)["xl/styles.xml"]);

    expect(imported).toEqual(workbook);
    expect(stylesXml).toContain('wrapText="1"');
  });

  it("round-trips bold, fill color, and border styles", () => {
    const excelIo = bootExcelIoModule();
    const codec = new excelIo.XlsxWorkbookCodec();
    const workbook = {
      sheets: [
        {
          name: "Styled",
          rows: [
            {
              cells: [
                {
                  value: "Header",
                  bold: true,
                  fillColor: "#D9EAF7",
                  border: "thin"
                }
              ]
            }
          ]
        }
      ]
    };

    const bytes = codec.exportWorkbook(workbook);
    const imported = codec.importWorkbook(bytes);

    expect(imported).toEqual(workbook);
  });

  it("round-trips font size styles", () => {
    const excelIo = bootExcelIoModule();
    const codec = new excelIo.XlsxWorkbookCodec();
    const workbook = {
      sheets: [
        {
          name: "Sized",
          rows: [
            {
              cells: [
                {
                  value: "Title",
                  bold: true,
                  fontSize: 16
                }
              ]
            }
          ]
        }
      ]
    };

    const bytes = codec.exportWorkbook(workbook);
    const imported = codec.importWorkbook(bytes);

    expect(imported).toEqual(workbook);
  });

  it("round-trips merged cell ranges", () => {
    const excelIo = bootExcelIoModule();
    const codec = new excelIo.XlsxWorkbookCodec();
    const workbook = {
      sheets: [
        {
          name: "Merged",
          mergedRanges: ["A1:B1", "A3:A4"],
          rows: [
            {
              cells: [
                { value: "Header", bold: true },
                {}
              ]
            },
            {
              cells: [
                { value: "Row1" },
                { value: "Value1" }
              ]
            },
            {
              cells: [
                { value: "Group" },
                { value: "Value2" }
              ]
            },
            {
              cells: [
                {},
                { value: "Value3" }
              ]
            }
          ]
        }
      ]
    };

    const bytes = codec.exportWorkbook(workbook);
    const imported = codec.importWorkbook(bytes);

    expect(imported.sheets[0].name).toBe("Merged");
    expect(imported.sheets[0].mergedRanges).toEqual(["A1:B1", "A3:A4"]);
    expect(imported.sheets[0].rows[0].cells[0].value).toBe("Header");
    expect(imported.sheets[0].rows[2].cells[0].value).toBe("Group");
    expect(imported.sheets[0].rows[3].cells[1].value).toBe("Value3");
  });

  it("round-trips frozen pane settings", () => {
    const excelIo = bootExcelIoModule();
    const codec = new excelIo.XlsxWorkbookCodec();
    const workbook = {
      sheets: [
        {
          name: "Frozen",
          freezePane: {
            rowSplit: 2,
            colSplit: 3
          },
          rows: [
            {
              cells: [
                { value: "A" },
                { value: "B" },
                { value: "C" },
                { value: "D" }
              ]
            },
            {
              cells: [
                { value: "1" },
                { value: "2" },
                { value: "3" },
                { value: "4" }
              ]
            }
          ]
        }
      ]
    };

    const bytes = codec.exportWorkbook(workbook);
    const imported = codec.importWorkbook(bytes);
    const sheetXml = decodeUtf8(codec.unpackEntries(bytes)["xl/worksheets/sheet1.xml"]);

    expect(imported).toEqual(workbook);
    expect(sheetXml).toContain("<sheetViews>");
    expect(sheetXml).toContain('xSplit="3"');
    expect(sheetXml).toContain('ySplit="2"');
    expect(sheetXml).toContain('topLeftCell="D3"');
    expect(sheetXml).toContain('state="frozen"');
  });

  it("rejects duplicate or invalid sheet names", () => {
    const excelIo = bootExcelIoModule();
    const codec = new excelIo.XlsxWorkbookCodec();

    expect(() => codec.exportWorkbook({
      sheets: [
        { name: "Same", rows: [] },
        { name: "Same", rows: [] }
      ]
    })).toThrow(/sheet name/i);

    expect(() => codec.exportWorkbook({
      sheets: [
        { name: "Bad/Name", rows: [] }
      ]
    })).toThrow(/sheet name/i);
  });

  it("exposes workbook xml for inspection after unzip", () => {
    const excelIo = bootExcelIoModule();
    const codec = new excelIo.XlsxWorkbookCodec();
    const bytes = codec.exportWorkbook({
      sheets: [
        {
          name: "One",
          rows: [
            {
              cells: [
                { value: "hello" },
                { value: 1 }
              ]
            }
          ]
        }
      ]
    });

    const entries = codec.unpackEntries(bytes);

    expect(decodeUtf8(entries["xl/workbook.xml"])).toContain("<sheet");
    expect(decodeUtf8(entries["xl/workbook.xml"])).toContain('name="One"');
    expect(decodeUtf8(entries["xl/worksheets/sheet1.xml"])).toContain('r="A1"');
    expect(decodeUtf8(entries["xl/worksheets/sheet1.xml"])).toContain("inlineStr");
    expect(decodeUtf8(entries["xl/worksheets/sheet1.xml"])).toContain("<t>1</t>");
  });

  it("writes styles and sheet layout xml when formatting is present", () => {
    const excelIo = bootExcelIoModule();
    const codec = new excelIo.XlsxWorkbookCodec();
    const bytes = codec.exportWorkbook({
      sheets: [
        {
          name: "Styled",
          columns: [
            { width: 18 }
          ],
          rows: [
            {
              height: 32,
              cells: [
                { value: 45367, numberFormat: "date", horizontalAlign: "center" }
              ]
            }
          ]
        }
      ]
    });

    const entries = codec.unpackEntries(bytes);

    expect(Object.keys(entries).sort()).toContain("xl/styles.xml");
    expect(decodeUtf8(entries["[Content_Types].xml"])).toContain("/xl/styles.xml");
    expect(decodeUtf8(entries["xl/_rels/workbook.xml.rels"])).toContain("styles.xml");
    expect(decodeUtf8(entries["xl/styles.xml"])).toContain("numFmtId=\"14\"");
    expect(decodeUtf8(entries["xl/styles.xml"])).toContain("horizontal=\"center\"");
    expect(decodeUtf8(entries["xl/worksheets/sheet1.xml"])).toContain("<cols>");
    expect(decodeUtf8(entries["xl/worksheets/sheet1.xml"])).toContain("width=\"18\"");
    expect(decodeUtf8(entries["xl/worksheets/sheet1.xml"])).toContain("ht=\"32\"");
    expect(decodeUtf8(entries["xl/worksheets/sheet1.xml"])).toContain("customHeight=\"1\"");
    expect(decodeUtf8(entries["xl/worksheets/sheet1.xml"])).toContain(" s=\"1\"");
  });

  it("writes empty styled cells when a fill-only cell is present", () => {
    const excelIo = bootExcelIoModule();
    const codec = new excelIo.XlsxWorkbookCodec();
    const bytes = codec.exportWorkbook({
      sheets: [
        {
          name: "StyledGap",
          rows: [
            {
              cells: [
                { value: "Header", fillColor: "#BFD7EA", border: "thin" },
                { fillColor: "#BFD7EA", border: "thin" }
              ]
            }
          ]
        }
      ]
    });

    const sheetXml = decodeUtf8(codec.unpackEntries(bytes)["xl/worksheets/sheet1.xml"]);

    expect(sheetXml).toContain('r="A1"');
    expect(sheetXml).toContain('r="B1"');
    expect(sheetXml).toMatch(/<c r="B1" s="\d+"\/>/);
  });

  it("writes font, fill, and border definitions when style options are present", () => {
    const excelIo = bootExcelIoModule();
    const codec = new excelIo.XlsxWorkbookCodec();
    const bytes = codec.exportWorkbook({
      sheets: [
        {
          name: "Styled",
          rows: [
            {
              cells: [
                {
                  value: "Header",
                  bold: true,
                  fillColor: "#D9EAF7",
                  border: "thin"
                }
              ]
            }
          ]
        }
      ]
    });

    const stylesXml = decodeUtf8(codec.unpackEntries(bytes)["xl/styles.xml"]);

    expect(stylesXml).toContain("<b/>");
    expect(stylesXml).toContain('patternType="solid"');
    expect(stylesXml).toContain("FFD9EAF7");
    expect(stylesXml).toContain("<left style=\"thin\"/>");
    expect(stylesXml).toContain("applyFill=\"1\"");
    expect(stylesXml).toContain("applyBorder=\"1\"");
    expect(stylesXml).toContain("applyFont=\"1\"");
  });

  it("writes font size definitions when font size is present", () => {
    const excelIo = bootExcelIoModule();
    const codec = new excelIo.XlsxWorkbookCodec();
    const bytes = codec.exportWorkbook({
      sheets: [
        {
          name: "FontSized",
          rows: [
            {
              cells: [
                {
                  value: "Title",
                  fontSize: 16
                }
              ]
            }
          ]
        }
      ]
    });

    const stylesXml = decodeUtf8(codec.unpackEntries(bytes)["xl/styles.xml"]);

    expect(stylesXml).toContain('<sz val="16"/>');
    expect(stylesXml).toContain("applyFont=\"1\"");
  });

  it("writes mergeCells xml when merged ranges are present", () => {
    const excelIo = bootExcelIoModule();
    const codec = new excelIo.XlsxWorkbookCodec();
    const bytes = codec.exportWorkbook({
      sheets: [
        {
          name: "Merged",
          mergedRanges: ["A1:B1"],
          rows: [
            {
              cells: [
                { value: "Header" },
                {}
              ]
            }
          ]
        }
      ]
    });

    const worksheetXml = decodeUtf8(codec.unpackEntries(bytes)["xl/worksheets/sheet1.xml"]);

    expect(worksheetXml).toContain("<mergeCells");
    expect(worksheetXml).toContain('ref="A1:B1"');
  });

  it("can asynchronously unpack a deflated zip entry", async () => {
    const excelIo = bootExcelIoModule();
    const codec = new excelIo.XlsxWorkbookCodec();
    const bytes = buildDeflatedZipWithSingleEntry("[Content_Types].xml", "<Types></Types>");

    const entries = await codec.unpackEntriesAsync(bytes);

    expect(Object.keys(entries)).toEqual(["[Content_Types].xml"]);
    expect(decodeUtf8(entries["[Content_Types].xml"])).toBe("<Types></Types>");
  });
});
