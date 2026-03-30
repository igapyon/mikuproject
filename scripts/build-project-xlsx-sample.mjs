import fs from "node:fs";
import path from "node:path";
import { JSDOM } from "jsdom";

const ROOT = process.cwd();
const typesCode = fs.readFileSync(path.resolve(ROOT, "src/js/types.js"), "utf8");
const markdownEscapeCode = fs.readFileSync(path.resolve(ROOT, "src/js/markdown-escape.js"), "utf8");
const excelIoCode = fs.readFileSync(path.resolve(ROOT, "src/js/excel-io.js"), "utf8");
const msProjectXmlCode = fs.readFileSync(path.resolve(ROOT, "src/js/msproject-xml.js"), "utf8");
const projectWorkbookSchemaCode = fs.readFileSync(path.resolve(ROOT, "src/js/project-workbook-schema.js"), "utf8");
const projectXlsxCode = fs.readFileSync(path.resolve(ROOT, "src/js/project-xlsx.js"), "utf8");
const wbsXlsxCode = fs.readFileSync(path.resolve(ROOT, "src/js/wbs-xlsx.js"), "utf8");
const wbsMarkdownCode = fs.readFileSync(path.resolve(ROOT, "src/js/wbs-markdown.js"), "utf8");

const dom = new JSDOM("<!DOCTYPE html><html><body></body></html>");
globalThis.window = dom.window;
globalThis.document = dom.window.document;
globalThis.DOMParser = dom.window.DOMParser;
globalThis.XMLSerializer = dom.window.XMLSerializer;
globalThis.Node = dom.window.Node;

globalThis.eval(`${typesCode}\n${markdownEscapeCode}\n${excelIoCode}\n${msProjectXmlCode}\n${projectWorkbookSchemaCode}\n${projectXlsxCode}\n${wbsXlsxCode}\n${wbsMarkdownCode}`);

const excelIo = globalThis.__mikuprojectExcelIo;
const xml = globalThis.__mikuprojectXml;
const projectXlsx = globalThis.__mikuprojectProjectXlsx;
const wbsXlsx = globalThis.__mikuprojectWbsXlsx;
const wbsMarkdown = globalThis.__mikuprojectWbsMarkdown;
if (!excelIo?.XlsxWorkbookCodec) {
  throw new Error("mikuproject excel io module is not loaded");
}
if (!xml?.SAMPLE_XML || typeof xml.importMsProjectXml !== "function") {
  throw new Error("mikuproject xml module is not loaded");
}
if (typeof projectXlsx?.exportProjectWorkbook !== "function") {
  throw new Error("mikuproject project xlsx module is not loaded");
}
if (typeof wbsXlsx?.exportWbsWorkbook !== "function") {
  throw new Error("mikuproject wbs xlsx module is not loaded");
}
if (typeof wbsMarkdown?.exportWbsMarkdown !== "function") {
  throw new Error("mikuproject wbs markdown module is not loaded");
}

const codec = new excelIo.XlsxWorkbookCodec();
const model = xml.importMsProjectXml(xml.SAMPLE_XML);
const workbook = projectXlsx.exportProjectWorkbook(model);
const holidayDates = wbsXlsx.collectWbsHolidayDates(model);
const wbsWorkbook = wbsXlsx.exportWbsWorkbook(model, { holidayDates });
const wbsMarkdownText = wbsMarkdown.exportWbsMarkdown(model, {
  holidayDates,
  useBusinessDaysForDisplayRange: true,
  useBusinessDaysForProgressBand: true
});
const richMarkdownModel = buildRichWbsMarkdownSampleModel(model);
const richHolidayDates = wbsXlsx.collectWbsHolidayDates(richMarkdownModel);
const richWbsMarkdownText = wbsMarkdown.exportWbsMarkdown(richMarkdownModel, {
  holidayDates: richHolidayDates,
  useBusinessDaysForDisplayRange: true,
  useBusinessDaysForProgressBand: true
});

const bytes = codec.exportWorkbook(workbook);
const wbsBytes = codec.exportWorkbook(wbsWorkbook);
const outputPath = path.resolve(ROOT, "local-data/mikuproject-sample.xlsx");
const wbsOutputPath = path.resolve(ROOT, "local-data/mikuproject-wbs-sample.xlsx");
const wbsMarkdownOutputPath = path.resolve(ROOT, "local-data/mikuproject-wbs-sample.md");
const richWbsMarkdownOutputPath = path.resolve(ROOT, "local-data/mikuproject-wbs-sample-rich.md");
fs.mkdirSync(path.dirname(outputPath), { recursive: true });
fs.writeFileSync(outputPath, Buffer.from(bytes));
fs.writeFileSync(wbsOutputPath, Buffer.from(wbsBytes));
fs.writeFileSync(wbsMarkdownOutputPath, wbsMarkdownText, "utf8");
fs.writeFileSync(richWbsMarkdownOutputPath, richWbsMarkdownText, "utf8");
console.log(`[build:project:xlsx-sample] generated ${path.relative(ROOT, outputPath)}`);
console.log(`[build:project:xlsx-sample] generated ${path.relative(ROOT, wbsOutputPath)}`);
console.log(`[build:project:xlsx-sample] generated ${path.relative(ROOT, wbsMarkdownOutputPath)}`);
console.log(`[build:project:xlsx-sample] generated ${path.relative(ROOT, richWbsMarkdownOutputPath)}`);

function buildRichWbsMarkdownSampleModel(model) {
  const cloned = JSON.parse(JSON.stringify(model));
  const longNameTarget = cloned.tasks.find((task) => task.outlineNumber === "1.2");
  if (longNameTarget) {
    longNameTarget.name = "初期実装（MS Project XML 調査・基軸フォーマット選定・内部モデルの概要確定）";
    longNameTarget.notes = "前提と制約を整理する\n関係者レビューを実施する";
  }
  const insertIndex = cloned.tasks.findIndex((task) => task.outlineNumber === "1.2");
  if (insertIndex >= 0) {
    const templateTask = cloned.tasks[insertIndex];
    cloned.tasks.splice(
      insertIndex + 1,
      0,
      createDerivedTask(templateTask, {
        uid: "901",
        id: "901",
        name: "内部 JSON 形式への写像方針を確認する",
        outlineLevel: 3,
        outlineNumber: "1.2.1",
        wbs: "1.2.1",
        start: "2026-03-17T09:00:00",
        finish: "2026-03-18T18:00:00",
        duration: "PT16H0M0S",
        percentComplete: 40,
        notes: "説明責務と round-trip 観点を切り分ける"
      }),
      createDerivedTask(templateTask, {
        uid: "902",
        id: "902",
        name: "フィールド差分の洗い出し結果を整理する",
        outlineLevel: 4,
        outlineNumber: "1.2.1.1",
        wbs: "1.2.1.1",
        start: "2026-03-18T09:00:00",
        finish: "2026-03-18T18:00:00",
        duration: "PT8H0M0S",
        percentComplete: 10,
        notes: "XML と内部モデルの差分を箇条書きで残す"
      }),
      createDerivedTask(templateTask, {
        uid: "903",
        id: "903",
        name: "長い説明文の折り返しを確認するための task",
        outlineLevel: 5,
        outlineNumber: "1.2.1.1.1",
        wbs: "1.2.1.1.1",
        start: "2026-03-18T09:00:00",
        finish: "2026-03-19T18:00:00",
        duration: "PT16H0M0S",
        percentComplete: 0,
        notes: "かなり長い補足説明をここへ入れて、Markdown tree の見え方と table の見え方を同時に確認する"
      })
    );
  }
  return cloned;
}

function createDerivedTask(baseTask, overrides) {
  return {
    ...baseTask,
    ...overrides,
    milestone: false,
    summary: false,
    critical: false,
    percentWorkComplete: overrides.percentComplete,
    predecessors: [],
    extendedAttributes: [],
    baselines: [],
    timephasedData: []
  };
}
