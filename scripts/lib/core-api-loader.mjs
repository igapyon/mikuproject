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
  "src/js/msproject-xml.js",
  "src/js/markdown-escape.js",
  "src/js/project-workbook-schema.js",
  "src/js/excel-io.js",
  "src/js/project-xlsx.js",
  "src/js/project-workbook-json.js",
  "src/js/project-patch-json.js",
  "src/js/wbs-xlsx.js",
  "src/js/wbs-svg.js",
  "src/js/wbs-markdown.js",
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
  delete globalThis.__mikuprojectXmlDom;
  delete globalThis.__mikuprojectXml;
  delete globalThis.__mikuprojectMarkdownEscape;
  delete globalThis.__mikuprojectProjectWorkbookSchema;
  delete globalThis.__mikuprojectExcelIo;
  delete globalThis.__mikuprojectProjectXlsx;
  delete globalThis.__mikuprojectProjectWorkbookJson;
  delete globalThis.__mikuprojectProjectPatchJson;
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
