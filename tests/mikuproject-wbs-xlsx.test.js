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
const msProjectXmlCode = readFileSync(
  path.resolve(__dirname, "../src/js/msproject-xml.js"),
  "utf8"
);
const wbsXlsxCode = readFileSync(
  path.resolve(__dirname, "../src/js/wbs-xlsx.js"),
  "utf8"
);

function bootModules() {
  new Function(`${typesCode}\n${excelIoCode}\n${msProjectXmlCode}\n${wbsXlsxCode}`)();
  return {
    excelIo: globalThis.__mikuprojectExcelIo,
    xml: globalThis.__mikuprojectXml,
    wbsXlsx: globalThis.__mikuprojectWbsXlsx
  };
}

function findRowIndexByCellValue(sheet, value, columnIndex = 0) {
  return sheet.rows.findIndex((row) => row.cells[columnIndex]?.value === value);
}

function findRowIndexByPredicate(sheet, predicate) {
  return sheet.rows.findIndex((row) => predicate(row.cells));
}

const SAMPLE_HOLIDAY_COUNT = 90;

describe("mikuproject wbs xlsx", () => {
  it("provides Excel-style layout references for WBS worksheet tuning", () => {
    const { wbsXlsx } = bootModules();

    expect(wbsXlsx.layout.columnIndex("A")).toBe(0);
    expect(wbsXlsx.layout.columnIndex("S")).toBe(18);
    expect(wbsXlsx.layout.columnName(18)).toBe("S");
    expect(wbsXlsx.layout.reference(17, 2)).toBe("C17");
    expect(wbsXlsx.layout.range("A1", "C17")).toBe("A1:C17");
    expect(wbsXlsx.layout.parseCellReference("C17")).toEqual({
      reference: "C17",
      rowNumber: 17,
      rowIndex: 16,
      columnName: "C",
      columnIndex: 2
    });
    expect(wbsXlsx.layout.describeCell("C17")).toBe("C17 (row 17, rowIndex 16, column C, columnIndex 2)");
  });

  it("can log WBS layout cell references on demand", () => {
    const { wbsXlsx } = bootModules();
    const messages = [];

    const message = wbsXlsx.layout.logCell("S12", "week header", (line) => {
      messages.push(line);
    });

    expect(message).toBe("week header: S12 (row 12, rowIndex 11, column S, columnIndex 18)");
    expect(messages).toEqual([
      "week header: S12 (row 12, rowIndex 11, column S, columnIndex 18)"
    ]);
  });

  it("collects holiday dates from calendar exceptions", () => {
    const { xml, wbsXlsx } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);
    const holidayDates = wbsXlsx.collectWbsHolidayDates(model);

    expect(holidayDates).toHaveLength(SAMPLE_HOLIDAY_COUNT);
    expect(holidayDates).toContain("2026-03-20");
    expect(holidayDates[0]).toBe("2026-03-20");
    expect(holidayDates.at(-1)).toBe("2031-02-24");
  });

  it("exports a dedicated WBS workbook from ProjectModel", () => {
    const { xml, wbsXlsx } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);

    const workbook = wbsXlsx.exportWbsWorkbook(model);
    const sheet = workbook.sheets[0];

    expect(workbook.sheets.map((item) => item.name)).toEqual(["WBS"]);
    expect(sheet.columns[2].width).toBe(12);
    expect(sheet.columns[3].width).toBe(10);
    expect(sheet.columns[4].width).toBe(10);
    expect(sheet.columns[5].width).toBe(42);
    expect(sheet.columns[6].width).toBe(20);
    expect(sheet.columns[7].width).toBe(18);
    expect(sheet.columns[8].width).toBe(12);
    expect(sheet.columns[9].width).toBe(28);
    expect(sheet.columns[11].width).toBe(18);
    expect(sheet.columns[11].hidden).toBe(true);
    expect(sheet.columns[12].width).toBe(12);
    expect(sheet.columns[12].hidden).toBe(true);
    expect(sheet.columns[13].width).toBe(12);
    expect(sheet.columns[13].hidden).toBe(true);
    expect(sheet.columns[14].width).toBe(12);
    expect(sheet.columns[14].hidden).toBe(true);
    expect(sheet.columns[15].width).toBe(16);
    expect(sheet.columns[16].hidden).toBe(true);
    expect(sheet.columns[16].width).toBe(12);
    expect(sheet.columns[17].hidden).toBe(true);
    expect(sheet.columns[17].width).toBe(20);
    expect(sheet.columns[18].hidden).toBe(true);
    expect(sheet.columns[18].width).toBe(18);
    expect(sheet.mergedRanges).toContain("A3:D3");
    expect(sheet.mergedRanges).toContain("F3:G3");
    expect(sheet.rows[0].cells[0].value).toBe("WBS");
    expect(sheet.rows[0].cells[0].fontSize).toBe(16);
    expect(sheet.rows[0].cells[0].fillColor).toBeUndefined();
    expect(sheet.rows[0].cells[1].fillColor).toBe("#EEF4FA");
    expect(sheet.rows[1].cells[0].value).toBeUndefined();
    const projectInfoHeaderIndex = findRowIndexByCellValue(sheet, "プロジェクト", 0);
    expect(projectInfoHeaderIndex).toBe(2);
    expect(sheet.rows[projectInfoHeaderIndex].cells[0].fontSize).toBe(14);
    expect(sheet.rows[projectInfoHeaderIndex + 1].cells[0].value).toBe("題名");
    expect(sheet.rows[projectInfoHeaderIndex + 1].cells[2].value).toBe("-");
    expect(sheet.rows[projectInfoHeaderIndex + 2].cells[0].value).toBe("カレンダ");
    expect(sheet.rows[projectInfoHeaderIndex + 2].cells[2].value).toBe("1 Standard");
    expect(sheet.rows[projectInfoHeaderIndex + 3].cells[0].value).toBe("基準");
    expect(sheet.rows[projectInfoHeaderIndex + 3].cells[2].value).toBe("開始基準");
    expect(sheet.rows[projectInfoHeaderIndex + 4].cells[0].value).toBe("開始日");
    expect(sheet.rows[projectInfoHeaderIndex + 4].cells[2].value).toBe("2026-03-16");
    expect(sheet.rows[projectInfoHeaderIndex + 5].cells[0].value).toBe("終了日");
    expect(sheet.rows[projectInfoHeaderIndex + 5].cells[2].value).toBe("2026-04-01");
    expect(sheet.rows[projectInfoHeaderIndex + 6].cells[0].value).toBe("現在日");
    expect(sheet.rows[projectInfoHeaderIndex + 6].cells[2].value).toBe("2026-03-23");
    expect(sheet.rows[projectInfoHeaderIndex + 7].cells[0].value).toBe("祝日");
    expect(sheet.rows[projectInfoHeaderIndex + 7].cells[2].value).toBeGreaterThan(0);
    const summaryHeaderIndex = findRowIndexByCellValue(sheet, "サマリ", 5);
    expect(summaryHeaderIndex).toBe(2);
    expect(sheet.rows[summaryHeaderIndex].height).toBe(24);
    expect(sheet.rows[summaryHeaderIndex].cells[5].fontSize).toBe(14);
    expect(sheet.rows[summaryHeaderIndex].cells[5].fillColor).toBe("#E1EDF8");
    expect(sheet.rows[summaryHeaderIndex + 1].cells[5].value).toBe("表示日");
    expect(sheet.rows[summaryHeaderIndex + 1].cells[6].value).toBe(17);
    expect(sheet.rows[summaryHeaderIndex + 1].cells[5].horizontalAlign).toBe("right");
    expect(sheet.rows[summaryHeaderIndex + 1].cells[6].horizontalAlign).toBe("center");
    expect(sheet.rows[summaryHeaderIndex + 1].cells[6].bold).toBe(true);
    expect(sheet.rows[projectInfoHeaderIndex + 1].cells[2].horizontalAlign).toBe("left");
    expect(sheet.rows[projectInfoHeaderIndex + 4].cells[2].horizontalAlign).toBe("left");
    expect(sheet.rows[summaryHeaderIndex + 3].cells[5].value).toBe("営業日");
    expect(sheet.rows[summaryHeaderIndex + 3].cells[6].value).toBe(12);
    expect(sheet.rows[summaryHeaderIndex + 4].cells[5].value).toBe("前日数");
    expect(sheet.rows[summaryHeaderIndex + 4].cells[6].value).toBe("-");
    expect(sheet.rows[summaryHeaderIndex + 5].cells[5].value).toBe("後日数");
    expect(sheet.rows[summaryHeaderIndex + 5].cells[6].value).toBe("-");
    expect(sheet.rows[summaryHeaderIndex + 6].cells[5].value).toBe("表示");
    expect(sheet.rows[summaryHeaderIndex + 6].cells[6].value).toBe("暦日");
    expect(sheet.rows[summaryHeaderIndex + 7].cells[5].value).toBe("進捗");
    expect(sheet.rows[summaryHeaderIndex + 7].cells[6].value).toBe("暦日");
    expect(sheet.rows[summaryHeaderIndex + 1].cells[8].value).toBe("タスク");
    expect(sheet.rows[summaryHeaderIndex + 1].cells[9].value).toBe(13);
    expect(sheet.rows[summaryHeaderIndex + 2].cells[8].value).toBe("リソース");
    expect(sheet.rows[summaryHeaderIndex + 2].cells[9].value).toBe(0);
    expect(sheet.rows[summaryHeaderIndex + 3].cells[8].value).toBe("割当");
    expect(sheet.rows[summaryHeaderIndex + 3].cells[9].value).toBe(0);
    expect(sheet.rows[summaryHeaderIndex + 4].cells[8].value).toBe("カレンダ");
    expect(sheet.rows[summaryHeaderIndex + 4].cells[9].value).toBe(1);
    expect(sheet.rows[summaryHeaderIndex + 8].cells[5].value).toBe("基準日");
    expect(sheet.rows[summaryHeaderIndex + 8].cells[6].value).toBe("2026-03-23");
    const weekRowIndex = findRowIndexByCellValue(sheet, "週", 18);
    const headerRowIndex = findRowIndexByCellValue(sheet, "UID");
    const dateRowIndex = headerRowIndex - 1;
    expect(weekRowIndex).toBe(11);
    expect(headerRowIndex).toBe(13);
    expect(dateRowIndex).toBe(12);
    expect(sheet.rows[weekRowIndex].height).toBe(24);
    expect(sheet.rows[weekRowIndex].cells[5].value).toBeUndefined();
    expect(sheet.rows[weekRowIndex].cells[18].fontSize).toBe(14);
    expect(sheet.rows[weekRowIndex].cells[18].fillColor).toBe("#E3EEF9");
    expect(sheet.rows[weekRowIndex].cells[19].fillColor).toBe("#D9E2EA");
    expect(sheet.rows[weekRowIndex].cells[20].value).toBe("週 03/15");
    expect(sheet.rows[weekRowIndex].cells[20].fontSize).toBe(14);
    expect(sheet.rows[headerRowIndex].cells.slice(0, 19).map((cell) => cell.value)).toEqual([
      "UID",
      "ID",
      "WBS",
      "種別",
      "階層",
      "名称",
      "開始",
      "終了",
      "期間",
      "タスク詳細",
      "進捗",
      "作業進捗",
      "マイル",
      "サマリ",
      "クリティカル",
      "担当",
      "カレンダ",
      "リソース",
      "先行"
    ]);
    expect(sheet.rows[headerRowIndex].cells[19].fillColor).toBe("#D9E2EA");
    expect(sheet.rows[dateRowIndex].cells.slice(20).map((cell) => cell.value)).toEqual([
      "3/16",
      "3/17",
      "3/18",
      "3/19",
      "3/20",
      "3/21",
      "3/22",
      "3/23",
      "3/24",
      "3/25",
      "3/26",
      "3/27",
      "3/28",
      "3/29",
      "3/30",
      "3/31",
      "4/1"
    ]);
    expect(sheet.rows[headerRowIndex].cells.slice(20).map((cell) => cell.value)).toEqual([
      "Mon",
      "Tue",
      "Wed",
      "Thu",
      "Fri",
      "Sat",
      "Sun",
      "Mon",
      "Tue",
      "Wed",
      "Thu",
      "Fri",
      "Sat",
      "Sun",
      "Mon",
      "Tue",
      "Wed"
    ]);
    expect(sheet.rows[headerRowIndex].cells[0].fillColor).toBe("#E1EDF8");
    expect(sheet.rows[headerRowIndex].cells[2].fillColor).toBe("#E6F0DF");
    expect(sheet.rows[headerRowIndex].cells[5].horizontalAlign).toBe("center");
    expect(sheet.rows[headerRowIndex].cells[15].horizontalAlign).toBe("center");
    expect(sheet.rows[headerRowIndex].cells[15].verticalAlign).toBe("center");
    expect(sheet.rows[headerRowIndex].cells[6].fillColor).toBe("#FDE7D3");
    expect(sheet.rows[headerRowIndex].cells[10].fillColor).toBe("#FBE4EC");
    expect(sheet.rows[headerRowIndex].cells[15].fillColor).toBe("#E2F1EF");
    expect(sheet.rows[dateRowIndex].cells[20].fillColor).toBe("#D9EAF7");
    expect(sheet.rows[dateRowIndex].cells[20].verticalAlign).toBe("center");
    const firstTaskRow = sheet.rows[headerRowIndex + 1];
    const secondTaskRow = sheet.rows[headerRowIndex + 2];
    const thirdTaskRow = sheet.rows[headerRowIndex + 3];
    expect(firstTaskRow.cells[3].value).toBe("フェーズ");
    expect(firstTaskRow.cells[3].fillColor).toBe("#EEF7E8");
    expect(firstTaskRow.cells[0].fillColor).toBe("#EEF7E8");
    expect(firstTaskRow.cells[5].bold).toBe(true);
    expect(firstTaskRow.cells[9].value).toBe("-");
    expect(firstTaskRow.cells[9].fillColor).toBe("#F5F7FA");
    expect(firstTaskRow.cells[9].horizontalAlign).toBe("center");
    expect(firstTaskRow.cells[15].value).toBe("-");
    expect(firstTaskRow.cells[15].fillColor).toBe("#F5F7FA");
    expect(firstTaskRow.cells[15].horizontalAlign).toBe("center");
    expect(firstTaskRow.cells[17].value).toBe("-");
    expect(firstTaskRow.cells[18].value).toBe("-");
    expect(firstTaskRow.cells[10].value).toBe("100%\n[##########]");
    expect(firstTaskRow.cells[11].value).toBe("");
    expect(firstTaskRow.cells[13].value).toBe("Sum");
    expect(firstTaskRow.cells[6].value).toBe("2026-03-16");
    expect(firstTaskRow.cells[7].value).toBe("2026-03-17");
    expect(firstTaskRow.cells[8].value).toBe("2日");
    expect(firstTaskRow.cells[20].value).toBe("━");
    expect(firstTaskRow.cells[21].value).toBe("━");
    expect(secondTaskRow.cells[3].value).toBe("マイル");
    expect(secondTaskRow.cells[3].fillColor).toBe("#FFF4E0");
    expect(secondTaskRow.cells[0].fillColor).toBe("#FFF4E0");
    expect(firstTaskRow.cells[5].value).toBe("> 基盤整備");
    expect(secondTaskRow.cells[5].value).toBe("  * 着手");
    expect(secondTaskRow.cells[5].fillColor).toBe("#FFF4E0");
    expect(secondTaskRow.cells[5].horizontalAlign).toBe("center");
    expect(secondTaskRow.cells[5].wrapText).toBe(true);
    expect(secondTaskRow.cells[6].fillColor).toBe("#FFF4E0");
    expect(secondTaskRow.cells[9].value).toBe("-");
    expect(secondTaskRow.cells[9].fillColor).toBe("#F5F7FA");
    expect(secondTaskRow.cells[9].wrapText).toBeUndefined();
    expect(secondTaskRow.cells[15].value).toBe("-");
    expect(secondTaskRow.cells[15].fillColor).toBe("#F5F7FA");
    expect(secondTaskRow.cells[15].horizontalAlign).toBe("center");
    expect(secondTaskRow.cells[15].verticalAlign).toBe("center");
    expect(secondTaskRow.cells[16].value).toBe("1 Standard");
    expect(secondTaskRow.cells[17].value).toBe("-");
    expect(secondTaskRow.cells[17].horizontalAlign).toBe("center");
    expect(secondTaskRow.cells[6].value).toBe("2026-03-16");
    expect(secondTaskRow.cells[7].value).toBe("2026-03-16");
    expect(secondTaskRow.cells[8].value).toBe("1日");
    expect(secondTaskRow.cells[10].value).toBe("100%\n[##########]");
    expect(secondTaskRow.cells[10].fillColor).toBe("#FFF4E0");
    expect(secondTaskRow.cells[11].value).toBe("");
    expect(secondTaskRow.cells[20].value).toBe("◆");
    expect(secondTaskRow.cells[20].fillColor).toBe("#5BAE9C");
    expect(secondTaskRow.cells[21].value).toBe("");
    expect(thirdTaskRow.cells[5].value).toBe("  - 初期実装（MS Project XML 調査・基軸フォーマット選定・内部モデルの概要確定）");
    expect(thirdTaskRow.cells[9].value).toBe("-");
    expect(thirdTaskRow.cells[18].value).toBe("-");
    expect(thirdTaskRow.cells[6].value).toBe("2026-03-16");
    expect(thirdTaskRow.cells[7].value).toBe("2026-03-16");
    expect(thirdTaskRow.cells[8].value).toBe("1日");
    expect(thirdTaskRow.cells[10].value).toBe("100%\n[##########]");
    expect(thirdTaskRow.cells[11].value).toBe("");
    expect(thirdTaskRow.cells[24].value).toBe("");
    const legendHeaderIndex = findRowIndexByCellValue(sheet, "凡例", 0);
    expect(legendHeaderIndex).toBe(headerRowIndex + 15);
    expect(sheet.rows[legendHeaderIndex - 1].height).toBe(28);
    expect(sheet.rows[legendHeaderIndex - 1].cells[0].value).toBeUndefined();
    expect(sheet.rows[legendHeaderIndex].height).toBe(24);
    expect(sheet.rows[legendHeaderIndex].cells[0].fontSize).toBe(14);
    expect(sheet.rows[legendHeaderIndex + 1].height).toBe(24);
    expect(sheet.rows[legendHeaderIndex + 1].cells[0].value).toBe("進捗済み");
    expect(sheet.rows[legendHeaderIndex + 1].cells[0].bold).toBe(true);
    expect(sheet.rows[legendHeaderIndex + 1].cells[0].fillColor).toBe("#5BAE9C");
    expect(sheet.rows[legendHeaderIndex + 7].cells[0].value).toBe("━:フェーズ");
    expect(sheet.rows[legendHeaderIndex + 8].cells[0].value).toBe("■:進捗済みタスク");
    expect(sheet.rows[legendHeaderIndex + 9].cells[0].value).toBe("□:予定タスク");
    expect(sheet.rows[legendHeaderIndex + 10].cells[0].fillColor).toBe("#FFF4E0");
    expect(sheet.rows[legendHeaderIndex + 11].cells[0].fillColor).toBe("#FBE4EC");
    expect(sheet.rows[legendHeaderIndex + 12].cells[0].fillColor).toBe("#F7EAF0");
    expect(sheet.rows[legendHeaderIndex + 13].cells[0].fillColor).toBe("#F3E1E9");
    expect(sheet.rows[legendHeaderIndex + 11].cells[0].value).toBe("Mil:マイルストーン");
    expect(sheet.rows[legendHeaderIndex + 12].cells[0].value).toBe("Sum:サマリ");
    expect(sheet.rows[legendHeaderIndex + 13].cells[0].value).toBe("Crit:クリティカル");
    expect(sheet.rows[legendHeaderIndex + 14].cells[0].value).toBe("-:未設定");
  });

  it("can generate a real xlsx from the dedicated WBS workbook", () => {
    const { excelIo, xml, wbsXlsx } = bootModules();
    const codec = new excelIo.XlsxWorkbookCodec();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);

    const workbook = wbsXlsx.exportWbsWorkbook(model);
    const bytes = codec.exportWorkbook(workbook);
    const entries = codec.listEntries(bytes);
    const unpackedEntries = codec.unpackEntries(bytes);
    const sheetXml = new TextDecoder().decode(unpackedEntries["xl/worksheets/sheet1.xml"]);
    const stylesXml = new TextDecoder().decode(unpackedEntries["xl/styles.xml"]);

    expect(entries).toContain("xl/workbook.xml");
    expect(entries).toContain("xl/worksheets/sheet1.xml");
    expect(sheetXml).toContain('ref="A3:D3"');
    expect(sheetXml).toContain('ref="F3:G3"');
    expect(sheetXml).toContain('s="1"');
    expect(stylesXml).toContain('<sz val="16"/>');
    expect(sheetXml).toContain('ref="U12:Z12"');
    expect(sheetXml).toContain('min="12" max="12" width="18" customWidth="1" hidden="1"');
    expect(sheetXml).toContain('min="13" max="13" width="12" customWidth="1" hidden="1"');
    expect(sheetXml).toContain('min="14" max="14" width="12" customWidth="1" hidden="1"');
    expect(sheetXml).toContain('min="15" max="15" width="12" customWidth="1" hidden="1"');
    expect(sheetXml).toContain('min="17" max="17" width="12" customWidth="1" hidden="1"');
    expect(sheetXml).toContain('min="18" max="18" width="20" customWidth="1" hidden="1"');
    expect(sheetXml).toContain('min="19" max="19" width="18" customWidth="1" hidden="1"');
    expect(sheetXml).not.toContain("<pane");
    expect(sheetXml).toContain("凡例");
    expect(sheetXml).toContain("プロジェクト");
    expect(sheetXml).toContain("週 03/15");
    expect(sheetXml).toContain("1 Standard");
    expect(sheetXml).toContain("階層");
    expect(sheetXml).toContain("3/16");
    expect(sheetXml).toContain("Mon");
  });

  it("marks weekend date-band cells with weekend fill", () => {
    const { xml, wbsXlsx } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);

    model.project.startDate = "2026-03-20T09:00:00";
    model.project.finishDate = "2026-03-23T18:00:00";
    model.project.currentDate = "2026-03-21T09:00:00";

    const workbook = wbsXlsx.exportWbsWorkbook(model);
    const sheet = workbook.sheets[0];

    const headerRowIndex = findRowIndexByCellValue(sheet, "UID");
    const dateRowIndex = headerRowIndex - 1;
    expect(sheet.rows[dateRowIndex].cells.slice(20).map((cell) => cell.value)).toEqual([
      "3/20",
      "3/21",
      "3/22",
      "3/23"
    ]);
    expect(sheet.rows[headerRowIndex].cells.slice(20).map((cell) => cell.value)).toEqual([
      "Fri",
      "Sat",
      "Sun",
      "Mon"
    ]);
    expect(sheet.rows[dateRowIndex].cells[21].fillColor).toBe("#FFE6A7");
    expect(sheet.rows[dateRowIndex].cells[22].fillColor).toBe("#C9D3E1");
  });

  it("uses project calendar weekdays instead of hardcoded weekends for non-working fill", () => {
    const { xml, wbsXlsx } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);

    model.project.startDate = "2026-03-20T09:00:00";
    model.project.finishDate = "2026-03-23T18:00:00";
    model.project.currentDate = "2026-03-20T09:00:00";
    model.calendars[0].weekDays = [
      { dayType: 1, dayWorking: false, workingTimes: [] },
      { dayType: 2, dayWorking: true, workingTimes: [{ fromTime: "09:00:00", toTime: "18:00:00" }] },
      { dayType: 3, dayWorking: true, workingTimes: [{ fromTime: "09:00:00", toTime: "18:00:00" }] },
      { dayType: 4, dayWorking: true, workingTimes: [{ fromTime: "09:00:00", toTime: "18:00:00" }] },
      { dayType: 5, dayWorking: true, workingTimes: [{ fromTime: "09:00:00", toTime: "18:00:00" }] },
      { dayType: 6, dayWorking: false, workingTimes: [] },
      { dayType: 7, dayWorking: true, workingTimes: [{ fromTime: "09:00:00", toTime: "18:00:00" }] }
    ];

    const workbook = wbsXlsx.exportWbsWorkbook(model);
    const sheet = workbook.sheets[0];
    const headerRowIndex = findRowIndexByCellValue(sheet, "UID");
    const dateRowIndex = headerRowIndex - 1;

    expect(sheet.rows[dateRowIndex].cells.slice(20).map((cell) => cell.value)).toEqual([
      "3/20",
      "3/21",
      "3/22",
      "3/23"
    ]);
    expect(sheet.rows[headerRowIndex].cells.slice(20).map((cell) => cell.value)).toEqual([
      "Fri",
      "Sat",
      "Sun",
      "Mon"
    ]);
    expect(sheet.rows[dateRowIndex].cells[20].fillColor).toBe("#FFE6A7");
    expect(sheet.rows[dateRowIndex].cells[21].fillColor).toBe("#D9EAF7");
    expect(sheet.rows[dateRowIndex].cells[22].fillColor).toBe("#C9D3E1");
  });

  it("suppresses task bands on weekly non-working days and configured holidays", () => {
    const { xml, wbsXlsx } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);
    const task = model.tasks.find((item) => !item.summary && !item.milestone);
    if (!task) {
      throw new Error("Expected a non-summary task in sample model");
    }

    model.project.startDate = "2026-03-26T09:00:00";
    model.project.finishDate = "2026-03-29T18:00:00";
    model.project.currentDate = "2026-03-26T09:00:00";
    task.start = "2026-03-26T09:00:00";
    task.finish = "2026-03-29T18:00:00";
    task.percentComplete = 50;
    model.calendars[0].weekDays = [
      { dayType: 1, dayWorking: false, workingTimes: [] },
      { dayType: 2, dayWorking: true, workingTimes: [{ fromTime: "09:00:00", toTime: "18:00:00" }] },
      { dayType: 3, dayWorking: true, workingTimes: [{ fromTime: "09:00:00", toTime: "18:00:00" }] },
      { dayType: 4, dayWorking: true, workingTimes: [{ fromTime: "09:00:00", toTime: "18:00:00" }] },
      { dayType: 5, dayWorking: true, workingTimes: [{ fromTime: "09:00:00", toTime: "18:00:00" }] },
      { dayType: 6, dayWorking: false, workingTimes: [] },
      { dayType: 7, dayWorking: true, workingTimes: [{ fromTime: "09:00:00", toTime: "18:00:00" }] }
    ];
    model.calendars[0].exceptions = [{
      name: "祝日",
      fromDate: "2026-03-28T00:00:00",
      toDate: "2026-03-28T23:59:59",
      dayWorking: false,
      workingTimes: []
    }];

    const workbook = wbsXlsx.exportWbsWorkbook(model);
    const sheet = workbook.sheets[0];
    const headerRowIndex = findRowIndexByCellValue(sheet, "UID");
    const dateRowIndex = headerRowIndex - 1;
    const dateColumns = new Map(
      sheet.rows[dateRowIndex].cells.map((cell, index) => [cell.value, index])
    );
    const taskRowIndex = headerRowIndex + 1 + model.tasks.indexOf(task);

    expect(taskRowIndex).toBeGreaterThan(headerRowIndex);
    expect(sheet.rows[taskRowIndex].cells[dateColumns.get("3/26")].value).toBe("■");
    expect(sheet.rows[taskRowIndex].cells[dateColumns.get("3/27")].value).toBe("");
    expect(sheet.rows[taskRowIndex].cells[dateColumns.get("3/28")].value).toBe("");
    expect(sheet.rows[taskRowIndex].cells[dateColumns.get("3/29")].value).toBe("□");
    expect(sheet.rows[taskRowIndex].cells[dateColumns.get("3/27")].fillColor).toBe("#C9D3E1");
    expect(sheet.rows[taskRowIndex].cells[dateColumns.get("3/28")].fillColor).toBe("#FCE4EC");
    expect(sheet.rows[taskRowIndex].cells[dateColumns.get("3/29")].fillColor).toBe("#9FD5C9");
  });

  it("marks week-start date-band cells with week-start fill", () => {
    const { xml, wbsXlsx } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);

    model.project.startDate = "2026-03-16T09:00:00";
    model.project.finishDate = "2026-03-23T18:00:00";
    model.project.currentDate = "2026-03-18T09:00:00";

    const workbook = wbsXlsx.exportWbsWorkbook(model);
    const sheet = workbook.sheets[0];

    const headerRowIndex = findRowIndexByCellValue(sheet, "UID");
    const dateRowIndex = headerRowIndex - 1;
    expect(sheet.rows[dateRowIndex].cells[27].value).toBe("3/23");
    expect(sheet.rows[headerRowIndex].cells[27].value).toBe("Mon");
    expect(sheet.rows[dateRowIndex].cells[27].fillColor).toBe("#D9EAF7");
    expect(sheet.rows[headerRowIndex + 1].cells[27].fillColor).toBe("#F4F7FB");
  });

  it("emphasizes week labels that contain a month boundary", () => {
    const { xml, wbsXlsx } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);

    model.project.startDate = "2026-03-30T09:00:00";
    model.project.finishDate = "2026-04-03T18:00:00";
    model.project.currentDate = "2026-04-01T09:00:00";

    const workbook = wbsXlsx.exportWbsWorkbook(model);
    const sheet = workbook.sheets[0];

    const weekRowIndex = findRowIndexByCellValue(sheet, "週", 18);
    expect(sheet.mergedRanges).toContain("U12:Y12");
    expect(sheet.rows[weekRowIndex].cells[20].value).toBe("週 03/29 / 04");
    expect(sheet.rows[weekRowIndex].cells[20].fillColor).toBe("#D6E7F8");
  });

  it("emphasizes month-start date headers", () => {
    const { xml, wbsXlsx } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);

    model.project.startDate = "2026-03-30T09:00:00";
    model.project.finishDate = "2026-04-03T18:00:00";
    model.project.currentDate = "2026-03-31T09:00:00";

    const workbook = wbsXlsx.exportWbsWorkbook(model);
    const sheet = workbook.sheets[0];

    const headerRowIndex = findRowIndexByCellValue(sheet, "UID");
    expect(sheet.rows[headerRowIndex - 1].cells[22].value).toBe("4/1");
    expect(sheet.rows[headerRowIndex].cells[22].value).toBe("Wed");
    expect(sheet.rows[headerRowIndex - 1].cells[22].fillColor).toBe("#DCEAF7");
  });

  it("renders milestone bands with a diamond marker", () => {
    const { xml, wbsXlsx } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);

    model.tasks[2].milestone = true;
    model.tasks[2].finish = model.tasks[2].start;

    const workbook = wbsXlsx.exportWbsWorkbook(model);
    const sheet = workbook.sheets[0];

    const headerRowIndex = findRowIndexByCellValue(sheet, "UID");
    const milestoneRow = sheet.rows[headerRowIndex + 3];
    expect(milestoneRow.cells[3].value).toBe("マイル");
    expect(milestoneRow.cells[3].fillColor).toBe("#FFF4E0");
    expect(milestoneRow.cells[12].value).toBe("Mil");
    expect(milestoneRow.cells[20].value).toBe("◆");
  });

  it("renders critical flags with an exclamation marker", () => {
    const { xml, wbsXlsx } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);

    model.tasks[1].critical = true;

    const workbook = wbsXlsx.exportWbsWorkbook(model);
    const sheet = workbook.sheets[0];

    const headerRowIndex = findRowIndexByCellValue(sheet, "UID");
    expect(sheet.rows[headerRowIndex + 2].cells[14].value).toBe("Crit");
  });

  it("marks configured holidays in the date band", () => {
    const { xml, wbsXlsx } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);

    const workbook = wbsXlsx.exportWbsWorkbook(model, {
      holidayDates: ["2026-03-20"]
    });
    const sheet = workbook.sheets[0];

    const projectInfoHeaderIndex = findRowIndexByCellValue(sheet, "プロジェクト", 0);
    expect(sheet.rows[projectInfoHeaderIndex + 7].cells[2].value).toBeGreaterThan(0);
    const headerRowIndex = findRowIndexByCellValue(sheet, "UID");
    expect(sheet.rows[headerRowIndex - 1].cells[24].value).toBe("3/20");
    expect(sheet.rows[headerRowIndex].cells[24].value).toBe("Fri");
    expect(sheet.rows[headerRowIndex - 1].cells[24].fillColor).toBe("#FCE4EC");
    expect(sheet.rows[headerRowIndex + 1].cells[24].value).toBe("");
    expect(sheet.rows[headerRowIndex + 1].cells[24].fillColor).toBe("#FCE4EC");
  });

  it("can limit the displayed date band around base date", () => {
    const { xml, wbsXlsx } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);

    const workbook = wbsXlsx.exportWbsWorkbook(model, {
      displayDaysBeforeBaseDate: 1,
      displayDaysAfterBaseDate: 2
    });
    const sheet = workbook.sheets[0];
    const summaryHeaderIndex = findRowIndexByCellValue(sheet, "サマリ", 5);
    const headerRowIndex = findRowIndexByCellValue(sheet, "UID");
    const dateRowIndex = headerRowIndex - 1;

    expect(sheet.rows[dateRowIndex].cells.slice(20).map((cell) => cell.value)).toEqual([
      "3/22",
      "3/23",
      "3/24",
      "3/25"
    ]);
    expect(sheet.rows[summaryHeaderIndex + 4].cells[6].value).toBe(1);
    expect(sheet.rows[summaryHeaderIndex + 5].cells[6].value).toBe(2);
    expect(sheet.rows[summaryHeaderIndex + 6].cells[6].value).toBe("暦日");
    expect(sheet.rows[summaryHeaderIndex + 7].cells[6].value).toBe("暦日");
  });

  it("can limit the displayed date band around base date using business days", () => {
    const { xml, wbsXlsx } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);

    model.project.startDate = "2026-03-16T09:00:00";
    model.project.finishDate = "2026-03-24T18:00:00";
    model.project.currentDate = "2026-03-18T09:00:00";

    const workbook = wbsXlsx.exportWbsWorkbook(model, {
      holidayDates: ["2026-03-20"],
      displayDaysBeforeBaseDate: 1,
      displayDaysAfterBaseDate: 2,
      useBusinessDaysForDisplayRange: true
    });
    const sheet = workbook.sheets[0];
    const summaryHeaderIndex = findRowIndexByCellValue(sheet, "サマリ", 5);
    const headerRowIndex = findRowIndexByCellValue(sheet, "UID");
    const dateRowIndex = headerRowIndex - 1;

    expect(sheet.rows[dateRowIndex].cells.slice(20).map((cell) => cell.value)).toEqual([
      "3/17",
      "3/18",
      "3/19",
      "3/20",
      "3/21",
      "3/22",
      "3/23"
    ]);
    expect(sheet.rows[summaryHeaderIndex + 3].cells[6].value).toBe(4);
    expect(sheet.rows[summaryHeaderIndex + 6].cells[6].value).toBe("営業日");
    expect(sheet.rows[summaryHeaderIndex + 7].cells[6].value).toBe("暦日");
  });

  it("can calculate progress band using business days", () => {
    const { xml, wbsXlsx } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);

    model.project.currentDate = "2026-03-25T09:00:00";
    model.tasks[2].start = "2026-03-16T09:00:00";
    model.tasks[2].finish = "2026-03-22T18:00:00";
    model.tasks[2].percentComplete = 50;

    const workbook = wbsXlsx.exportWbsWorkbook(model, {
      holidayDates: ["2026-03-20"],
      useBusinessDaysForProgressBand: true
    });
    const sheet = workbook.sheets[0];
    const summaryHeaderIndex = findRowIndexByCellValue(sheet, "サマリ", 5);
    const headerRowIndex = findRowIndexByCellValue(sheet, "UID");
    const designRow = sheet.rows[headerRowIndex + 3];

    expect(sheet.rows[summaryHeaderIndex + 7].cells[6].value).toBe("営業日");
    expect(designRow.cells[8].value).toBe("4営業日");
    expect(designRow.cells[20].fillColor).toBe("#5BAE9C");
    expect(designRow.cells[21].fillColor).toBe("#5BAE9C");
    expect(designRow.cells[22].fillColor).toBe("#9FD5C9");
    expect(designRow.cells[24].fillColor).toBe("#FCE4EC");
    expect(designRow.cells[25].fillColor).toBe("#C9D3E1");
    expect(designRow.cells[26].fillColor).toBe("#9FD5C9");
  });

  it("shows task band on non-working endpoints only", () => {
    const { xml, wbsXlsx } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);

    model.tasks[2].start = "2026-03-20T09:00:00";
    model.tasks[2].finish = "2026-03-23T18:00:00";
    model.tasks[2].percentComplete = 0;

    const workbook = wbsXlsx.exportWbsWorkbook(model, {
      holidayDates: ["2026-03-20"]
    });
    const sheet = workbook.sheets[0];
    const headerRowIndex = findRowIndexByCellValue(sheet, "UID");
    const designRow = sheet.rows[headerRowIndex + 3];

    expect(designRow.cells[24].value).toBe("□");
    expect(designRow.cells[25].value).toBe("");
    expect(designRow.cells[26].value).toBe("");
    expect(designRow.cells[27].value).toBe("□");
  });

  it("truncates long owner, resources, and predecessors labels for wbs display", () => {
    const { xml, wbsXlsx } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);

    model.project.title = "Sample Project Title Very Long";
    model.calendars[0].name = "Standard Calendar Very Long";
    model.tasks[2].predecessors = [{ predecessorUid: "1", type: 1, lag: "0" }];

    const workbook = wbsXlsx.exportWbsWorkbook(model);
    const sheet = workbook.sheets[0];
    const projectInfoHeaderIndex = findRowIndexByCellValue(sheet, "プロジェクト", 0);
    const headerRowIndex = findRowIndexByCellValue(sheet, "UID");
    const secondTaskRow = sheet.rows[headerRowIndex + 3];
    const thirdTaskRow = sheet.rows[headerRowIndex + 4];

    expect(sheet.rows[projectInfoHeaderIndex + 1].cells[2].value).toBe("Sample Project ...");
    expect(secondTaskRow.cells[15].value).toBe("-");
    expect(secondTaskRow.cells[16].value).toBe("1 Standa...");
    expect(secondTaskRow.cells[17].value).toBe("-");
    expect(thirdTaskRow.cells[18].value).toBe("-");
  });

  it("uses taller rows for long task names in wbs display", () => {
    const { xml, wbsXlsx } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);

    model.tasks[2].name = "Design task with a very long title for wrapped display";

    const workbook = wbsXlsx.exportWbsWorkbook(model);
    const sheet = workbook.sheets[0];
    const headerRowIndex = findRowIndexByCellValue(sheet, "UID");
    const designRow = sheet.rows[headerRowIndex + 3];

    expect(designRow.height).toBe(34);
    expect(designRow.cells[5].wrapText).toBe(true);
  });
});
