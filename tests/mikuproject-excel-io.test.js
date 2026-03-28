// @vitest-environment jsdom

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

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
      "xl/workbook.xml",
      "xl/worksheets/sheet1.xml"
    ]);
  });

  it("round-trips sheet names, sparse cells, formulas, and primitive cell values", () => {
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
                { value: 2 },
                { value: true },
                { formula: "B2*2", value: 4 }
              ]
            },
            {
              cells: [
                { value: "Build" },
                {},
                { value: false },
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
                { value: 1 },
                { value: 2 },
                { value: 3 },
                { value: 4 }
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
    expect(decodeUtf8(entries["xl/worksheets/sheet1.xml"])).toContain("<v>1</v>");
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
    expect(sheetXml).toContain('<c r="B1" s="1"/>');
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
});
