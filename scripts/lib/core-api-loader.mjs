import fs from "node:fs";
import { createRequire } from "node:module";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { JSDOM } from "jsdom";

import { CORE_API_MODULE_RELATIVE_PATHS } from "./runtime-module-paths.mjs";

export { CORE_API_MODULE_RELATIVE_PATHS } from "./runtime-module-paths.mjs";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const DEFAULT_ROOT_DIR = path.resolve(__dirname, "../..");
const require = createRequire(import.meta.url);

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
  delete globalThis.__mikuprojectMsOfficeCore;
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
  delete globalThis.__mikuProjectCoreApi;
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

  const api = globalThis.__mikuProjectCoreApi || globalThis.__mikuprojectCoreApi;
  if (!api) {
    restoreWindowGlobals();
    dom.window.close();
    throw new Error("Failed to boot __mikuProjectCoreApi");
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
