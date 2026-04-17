import fs from "node:fs";
import { createRequire } from "node:module";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { JSDOM } from "jsdom";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const DEFAULT_ROOT_DIR = path.resolve(__dirname, "../..");
const require = createRequire(import.meta.url);

export const CORE_API_MODULE_RELATIVE_PATHS = [
  "src/js/types.js",
  "src/js/ai-json-util.js",
  "src/js/ai-json-spec.js",
  "src/js/main-util.js",
  "src/js/msproject-ai-views.js",
  "src/js/msproject-calendar.js",
  "src/js/msproject-samples.js",
  "src/js/msproject-csv.js",
  "src/js/msproject-validate-helpers.js",
  "src/js/msproject-validate.js",
  "src/js/msproject-xml-dom.js",
  "src/js/msproject-codec.js",
  "src/js/msproject-mermaid.js",
  "src/js/msproject-xml.js",
  "src/js/markdown-escape.js",
  "src/js/project-workbook-schema.js",
  "src/js/excel-io-util.js",
  "src/js/excel-io-zip.js",
  "src/js/excel-io-normalize.js",
  "src/js/excel-io-package-xml.js",
  "src/js/excel-io-worksheet-build.js",
  "src/js/excel-io-worksheet-parse.js",
  "src/js/excel-io-workbook-parse.js",
  "src/js/excel-io-workbook-build.js",
  "src/js/excel-io-styles-build.js",
  "src/js/excel-io-styles-parse.js",
  "src/js/excel-io.js",
  "src/js/project-xlsx-import-util.js",
  "src/js/project-xlsx-import-project.js",
  "src/js/project-xlsx-import-calendars.js",
  "src/js/project-xlsx-import-entities.js",
  "src/js/project-xlsx-import.js",
  "src/js/project-xlsx-export-util.js",
  "src/js/project-xlsx-export-project.js",
  "src/js/project-xlsx-export-entities.js",
  "src/js/project-xlsx-export-calendars.js",
  "src/js/project-xlsx-export.js",
  "src/js/project-xlsx.js",
  "src/js/project-workbook-json-validate.js",
  "src/js/project-workbook-json-import.js",
  "src/js/project-workbook-json-export.js",
  "src/js/project-workbook-json.js",
  "src/js/project-patch-json-util.js",
  "src/js/project-patch-json-links.js",
  "src/js/project-patch-json-tasks.js",
  "src/js/project-patch-json-entities.js",
  "src/js/project-patch-json-updates.js",
  "src/js/project-patch-json-core.js",
  "src/js/project-patch-json.js",
  "src/js/core-api-msproject-ai.js",
  "src/js/core-api-msproject.js",
  "src/js/core-api-workbook-xlsx.js",
  "src/js/core-api-workbook.js",
  "src/js/core-api-ai-json-import.js",
  "src/js/core-api-ai-json.js",
  "src/js/core-api-external-binary.js",
  "src/js/core-api-external-document.js",
  "src/js/core-api-external-import.js",
  "src/js/core-api-import.js",
  "src/js/wbs-dateband.js",
  "src/js/wbs-xlsx-base.js",
  "src/js/wbs-xlsx-taskmeta.js",
  "src/js/wbs-xlsx-layout.js",
  "src/js/wbs-xlsx-sections.js",
  "src/js/wbs-xlsx-cells.js",
  "src/js/wbs-xlsx-export.js",
  "src/js/wbs-xlsx-public.js",
  "src/js/wbs-xlsx.js",
  "src/js/wbs-svg-zip.js",
  "src/js/wbs-svg-bars.js",
  "src/js/wbs-svg-viewport.js",
  "src/js/wbs-svg-axis.js",
  "src/js/wbs-svg-scaffold.js",
  "src/js/wbs-svg-labels.js",
  "src/js/wbs-svg-calendar.js",
  "src/js/wbs-svg-timeline.js",
  "src/js/wbs-svg-render.js",
  "src/js/wbs-svg-public.js",
  "src/js/wbs-svg.js",
  "src/js/wbs-markdown.js",
  "src/js/core-api-report.js",
  "src/js/core-api-report-adapters.js",
  "src/js/core-api-report-public.js",
  "src/js/core-api-registry.js",
  "src/js/core-api-public.js",
  "src/js/core-api.js"
];

function installWindowGlobals(window) {
  const previous = new Map();
  const keys = [
    "window",
    "document",
    "navigator",
    "Node",
    "Element",
    "HTMLElement",
    "DOMParser",
    "XMLSerializer",
    "XMLDocument",
    "Blob",
    "File"
  ];

  for (const key of keys) {
    previous.set(key, globalThis[key]);
    Object.defineProperty(globalThis, key, {
      configurable: true,
      writable: true,
      value: window[key]
    });
  }

  return () => {
    for (const key of keys) {
      if (previous.get(key) === undefined) {
        delete globalThis[key];
        continue;
      }
      Object.defineProperty(globalThis, key, {
        configurable: true,
        writable: true,
        value: previous.get(key)
      });
    }
  };
}

function clearMikuprojectGlobals() {
  delete globalThis.__mikuprojectAiJsonUtil;
  delete globalThis.__mikuprojectAiJsonSpec;
  delete globalThis.__mikuprojectMainUtil;
  delete globalThis.__mikuprojectMsprojectAiViews;
  delete globalThis.__mikuprojectMsprojectCalendar;
  delete globalThis.__mikuprojectMsprojectSamples;
  delete globalThis.__mikuprojectMsprojectCsv;
  delete globalThis.__mikuprojectMsprojectValidateHelpers;
  delete globalThis.__mikuprojectMsprojectValidate;
  delete globalThis.__mikuprojectMsprojectXmlDom;
  delete globalThis.__mikuprojectMsprojectCodec;
  delete globalThis.__mikuprojectMsprojectMermaid;
  delete globalThis.__mikuprojectXmlDom;
  delete globalThis.__mikuprojectXml;
  delete globalThis.__mikuprojectMarkdownEscape;
  delete globalThis.__mikuprojectProjectWorkbookSchema;
  delete globalThis.__mikuprojectExcelIoUtil;
  delete globalThis.__mikuprojectExcelIoZip;
  delete globalThis.__mikuprojectExcelIoNormalize;
  delete globalThis.__mikuprojectExcelIoPackageXml;
  delete globalThis.__mikuprojectExcelIoWorksheetBuild;
  delete globalThis.__mikuprojectExcelIoWorksheetParse;
  delete globalThis.__mikuprojectExcelIoWorkbookParse;
  delete globalThis.__mikuprojectExcelIoWorkbookBuild;
  delete globalThis.__mikuprojectExcelIoStylesBuild;
  delete globalThis.__mikuprojectExcelIoStylesParse;
  delete globalThis.__mikuprojectExcelIo;
  delete globalThis.__mikuprojectProjectXlsxImportUtil;
  delete globalThis.__mikuprojectProjectXlsxImportProject;
  delete globalThis.__mikuprojectProjectXlsxImportCalendars;
  delete globalThis.__mikuprojectProjectXlsxImportEntities;
  delete globalThis.__mikuprojectProjectXlsxImport;
  delete globalThis.__mikuprojectProjectXlsxExportUtil;
  delete globalThis.__mikuprojectProjectXlsxExportProject;
  delete globalThis.__mikuprojectProjectXlsxExportEntities;
  delete globalThis.__mikuprojectProjectXlsxExportCalendars;
  delete globalThis.__mikuprojectProjectXlsxExport;
  delete globalThis.__mikuprojectProjectXlsx;
  delete globalThis.__mikuprojectProjectWorkbookJsonValidate;
  delete globalThis.__mikuprojectProjectWorkbookJsonImport;
  delete globalThis.__mikuprojectProjectWorkbookJsonExport;
  delete globalThis.__mikuprojectProjectWorkbookJson;
  delete globalThis.__mikuprojectProjectPatchJson;
  delete globalThis.__mikuprojectCoreApiMsprojectAi;
  delete globalThis.__mikuprojectCoreApiMsproject;
  delete globalThis.__mikuprojectCoreApiWorkbookXlsx;
  delete globalThis.__mikuprojectCoreApiWorkbook;
  delete globalThis.__mikuprojectCoreApiAiJsonImport;
  delete globalThis.__mikuprojectCoreApiAiJson;
  delete globalThis.__mikuprojectCoreApiExternalBinary;
  delete globalThis.__mikuprojectCoreApiExternalDocument;
  delete globalThis.__mikuprojectCoreApiExternalImport;
  delete globalThis.__mikuprojectCoreApiImport;
  delete globalThis.__mikuprojectWbsXlsxLayout;
  delete globalThis.__mikuprojectCoreApiReport;
  delete globalThis.__mikuprojectCoreApiReportAdapters;
  delete globalThis.__mikuprojectCoreApiReportPublic;
  delete globalThis.__mikuprojectCoreApiRegistry;
  delete globalThis.__mikuprojectCoreApiPublic;
  delete globalThis.__mikuprojectWbsXlsx;
  delete globalThis.__mikuprojectNativeSvg;
  delete globalThis.__mikuprojectWbsMarkdown;
  delete globalThis.__mikuprojectCoreApi;
}

function loadXmlDomGlobals(window) {
  try {
    const xmldom = require("@xmldom/xmldom");
    const parser = new xmldom.DOMParser();
    const xmlDocument = parser.parseFromString("<root/>", "application/xml");
    return {
      DOMParser: xmldom.DOMParser,
      XMLSerializer: xmldom.XMLSerializer,
      createDocument(rootName) {
        return new xmldom.DOMImplementation().createDocument("", rootName, null);
      },
      XMLDocument: xmlDocument.constructor
    };
  } catch (_error) {
    const xmlDocument = new window.DOMParser().parseFromString("<root/>", "application/xml");
    return {
      DOMParser: window.DOMParser,
      XMLSerializer: window.XMLSerializer,
      createDocument(rootName) {
        return window.document.implementation.createDocument("", rootName, null);
      },
      XMLDocument: xmlDocument.constructor
    };
  }
}

export function loadMikuprojectCoreApi(options = {}) {
  const rootDir = options.rootDir ? path.resolve(options.rootDir) : DEFAULT_ROOT_DIR;
  const dom = new JSDOM("<!doctype html><html><body></body></html>", {
    url: "http://localhost/"
  });
  const restoreWindowGlobals = installWindowGlobals(dom.window);
  clearMikuprojectGlobals();
  const xmlDomGlobals = loadXmlDomGlobals(dom.window);
  Object.assign(globalThis, {
    DOMParser: xmlDomGlobals.DOMParser,
    XMLSerializer: xmlDomGlobals.XMLSerializer,
    XMLDocument: xmlDomGlobals.XMLDocument,
    __mikuprojectXmlDom: {
      DOMParser: xmlDomGlobals.DOMParser,
      XMLSerializer: xmlDomGlobals.XMLSerializer,
      createDocument: xmlDomGlobals.createDocument
    }
  });
  const combinedCode = CORE_API_MODULE_RELATIVE_PATHS
    .map((relativePath) => fs.readFileSync(path.resolve(rootDir, relativePath), "utf8"))
    .join("\n");
  new Function(combinedCode)();

  const api = globalThis.__mikuprojectCoreApi;
  if (!api) {
    restoreWindowGlobals();
    dom.window.close();
    throw new Error("Failed to boot __mikuprojectCoreApi");
  }

  return {
    api,
    dispose() {
      clearMikuprojectGlobals();
      restoreWindowGlobals();
      dom.window.close();
    }
  };
}
