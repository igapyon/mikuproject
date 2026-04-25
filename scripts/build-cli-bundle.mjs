import fs from "node:fs";
import path from "node:path";

const ROOT = process.cwd();
const args = parseArgs(process.argv.slice(2));
const outFile = path.resolve(args.out || path.join(ROOT, "bundle", "mikuproject.mjs"));
const CORE_API_MODULE_RELATIVE_PATHS = readCoreApiModuleRelativePaths();

const XMLDOM_MODULE_RELATIVE_PATHS = [
  "node_modules/@xmldom/xmldom/lib/conventions.js",
  "node_modules/@xmldom/xmldom/lib/dom.js",
  "node_modules/@xmldom/xmldom/lib/entities.js",
  "node_modules/@xmldom/xmldom/lib/sax.js",
  "node_modules/@xmldom/xmldom/lib/dom-parser.js",
  "node_modules/@xmldom/xmldom/lib/index.js"
];

buildCliBundle();

function buildCliBundle() {
  assertRequiredFilesExist();
  fs.mkdirSync(path.dirname(outFile), { recursive: true });
  fs.writeFileSync(outFile, buildSingleMjsRuntime(), "utf8");
  fs.chmodSync(outFile, 0o755);
  console.log(`[build:cli-bundle] generated ${path.relative(ROOT, outFile)}`);
}

function parseArgs(argv) {
  const options = {};
  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (!token.startsWith("--")) {
      throw new Error(`未対応の引数です: ${token}`);
    }

    const key = token.slice(2);
    const value = argv[index + 1];
    if (value === undefined || value.startsWith("--")) {
      throw new Error(`オプション ${token} には値が必要です`);
    }
    options[key] = value;
    index += 1;
  }
  return options;
}

function assertRequiredFilesExist() {
  const requiredFiles = [
    "scripts/mikuproject-cli.mjs",
    ...CORE_API_MODULE_RELATIVE_PATHS,
    ...XMLDOM_MODULE_RELATIVE_PATHS
  ];
  for (const relativePath of requiredFiles) {
    if (!fs.existsSync(path.resolve(ROOT, relativePath))) {
      throw new Error(`CLI bundle に必要なファイルが見つかりません: ${relativePath}`);
    }
  }
}

function buildSingleMjsRuntime() {
  const cliSource = stripCliImports(readRepoFile("scripts/mikuproject-cli.mjs"));
  const coreModuleSources = Object.fromEntries(
    CORE_API_MODULE_RELATIVE_PATHS.map((relativePath) => [relativePath, readRepoFile(relativePath)])
  );
  const xmldomModuleSources = Object.fromEntries(
    XMLDOM_MODULE_RELATIVE_PATHS.map((relativePath) => [
      toXmldomModuleId(relativePath),
      readRepoFile(relativePath)
    ])
  );

  return [
    "#!/usr/bin/env node",
    "",
    "import fs from \"node:fs\";",
    "import path from \"node:path\";",
    "import { fileURLToPath } from \"node:url\";",
    "import { Blob as NodeBlob, File as NodeFile } from \"node:buffer\";",
    "",
    "const BUNDLED_CORE_API_MODULE_RELATIVE_PATHS = Object.freeze(",
    `${JSON.stringify(CORE_API_MODULE_RELATIVE_PATHS, null, 2)});`,
    "",
    "const BUNDLED_CORE_API_MODULE_SOURCES = Object.freeze(",
    `${JSON.stringify(coreModuleSources, null, 2)});`,
    "",
    "const BUNDLED_XMLDOM_MODULE_SOURCES = Object.freeze(",
    `${JSON.stringify(xmldomModuleSources, null, 2)});`,
    "",
    String.raw`function normalizeBundledXmldomModuleId(parentId, request) {
  if (request === "@xmldom/xmldom") {
    return "/node_modules/@xmldom/xmldom/lib/index.js";
  }
  if (request.startsWith("./") || request.startsWith("../")) {
    const resolved = path.posix.normalize(path.posix.join(path.posix.dirname(parentId), request));
    return resolved.endsWith(".js") ? resolved : resolved + ".js";
  }
  throw new Error("Unsupported bundled require: " + request);
}`,
    "",
    String.raw`const BUNDLED_XMLDOM_MODULE_CACHE = new Map();

function requireBundledXmldom(request, parentId = "/node_modules/@xmldom/xmldom/lib/index.js") {
  const moduleId = normalizeBundledXmldomModuleId(parentId, request);
  if (BUNDLED_XMLDOM_MODULE_CACHE.has(moduleId)) {
    return BUNDLED_XMLDOM_MODULE_CACHE.get(moduleId).exports;
  }
  const source = BUNDLED_XMLDOM_MODULE_SOURCES[moduleId];
  if (source === undefined) {
    throw new Error("Bundled xmldom module not found: " + moduleId);
  }
  const module = { exports: {} };
  BUNDLED_XMLDOM_MODULE_CACHE.set(moduleId, module);
  const localRequire = (nextRequest) => requireBundledXmldom(nextRequest, moduleId);
  new Function("require", "module", "exports", source)(localRequire, module, module.exports);
  return module.exports;
}`,
    "",
    String.raw`function createBundledXmlDomGlobals() {
  const xmldom = requireBundledXmldom("@xmldom/xmldom");
  const parser = new xmldom.DOMParser();
  const xmlDocument = parser.parseFromString("<root/>", "application/xml");
  return {
    DOMParser: xmldom.DOMParser,
    XMLSerializer: xmldom.XMLSerializer,
    XMLDocument: xmlDocument.constructor,
    createDocument(rootName) {
      return new xmldom.DOMImplementation().createDocument("", rootName, null);
    }
  };
}`,
    "",
    String.raw`function installBundledWindowGlobals(xmlDomGlobals) {
  const previous = new Map();
  const fakeDocument = {
    implementation: {
      createDocument(_namespace, rootName) {
        return xmlDomGlobals.createDocument(rootName);
      }
    }
  };
  const fakeWindow = {
    document: fakeDocument,
    DOMParser: xmlDomGlobals.DOMParser,
    XMLSerializer: xmlDomGlobals.XMLSerializer,
    XMLDocument: xmlDomGlobals.XMLDocument,
    Blob: globalThis.Blob || NodeBlob,
    File: globalThis.File || NodeFile,
    Node: globalThis.Node || { ELEMENT_NODE: 1, TEXT_NODE: 3 },
    Element: globalThis.Element || class Element {},
    HTMLElement: globalThis.HTMLElement || class HTMLElement {},
    navigator: globalThis.navigator || {},
    setTimeout: globalThis.setTimeout,
    clearTimeout: globalThis.clearTimeout
  };
  const values = {
    window: fakeWindow,
    document: fakeDocument,
    navigator: fakeWindow.navigator,
    Node: fakeWindow.Node,
    Element: fakeWindow.Element,
    HTMLElement: fakeWindow.HTMLElement,
    DOMParser: xmlDomGlobals.DOMParser,
    XMLSerializer: xmlDomGlobals.XMLSerializer,
    XMLDocument: xmlDomGlobals.XMLDocument,
    Blob: fakeWindow.Blob,
    File: fakeWindow.File
  };

  for (const [key, value] of Object.entries(values)) {
    previous.set(key, globalThis[key]);
    Object.defineProperty(globalThis, key, {
      configurable: true,
      writable: true,
      value
    });
  }

  return () => {
    for (const [key, value] of previous.entries()) {
      if (value === undefined) {
        delete globalThis[key];
        continue;
      }
      Object.defineProperty(globalThis, key, {
        configurable: true,
        writable: true,
        value
      });
    }
  };
}`,
    "",
    buildClearGlobalsFunctionSource(),
    "",
    String.raw`function loadMikuprojectCoreApi() {
  const xmlDomGlobals = createBundledXmlDomGlobals();
  const restoreWindowGlobals = installBundledWindowGlobals(xmlDomGlobals);
  clearMikuprojectGlobals();
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

  const combinedCode = BUNDLED_CORE_API_MODULE_RELATIVE_PATHS
    .map((relativePath) => BUNDLED_CORE_API_MODULE_SOURCES[relativePath])
    .join("\n");
  new Function(combinedCode)();

  const api = globalThis.__mikuprojectCoreApi;
  if (!api) {
    restoreWindowGlobals();
    throw new Error("Failed to boot __mikuprojectCoreApi");
  }

  return {
    api,
    dispose() {
      clearMikuprojectGlobals();
      restoreWindowGlobals();
    }
  };
}`,
    "",
    cliSource
  ].join("\n");
}

function buildClearGlobalsFunctionSource() {
  const loaderSource = readRepoFile("scripts/lib/core-api-loader.mjs");
  const match = loaderSource.match(/function clearMikuprojectGlobals\(\) \{[\s\S]*?\n\}/);
  if (!match) {
    throw new Error("core-api-loader.mjs から clearMikuprojectGlobals を抽出できませんでした");
  }
  return match[0];
}

function readCoreApiModuleRelativePaths() {
  const loaderSource = readRepoFile("scripts/lib/core-api-loader.mjs");
  const match = loaderSource.match(/export const CORE_API_MODULE_RELATIVE_PATHS = (\[[\s\S]*?\n\]);/);
  if (!match) {
    throw new Error("core-api-loader.mjs から CORE_API_MODULE_RELATIVE_PATHS を抽出できませんでした");
  }
  return Function(`"use strict"; return ${match[1]};`)();
}

function stripCliImports(source) {
  return source
    .replace(/^#!.*\n/, "")
    .split("\n")
    .filter((line) => !line.startsWith("import "))
    .join("\n")
    .trimStart();
}

function toXmldomModuleId(relativePath) {
  const prefix = "node_modules/@xmldom/xmldom/";
  if (!relativePath.startsWith(prefix)) {
    throw new Error(`Unexpected xmldom path: ${relativePath}`);
  }
  return `/node_modules/@xmldom/xmldom/${relativePath.slice(prefix.length)}`;
}

function readRepoFile(relativePath) {
  return fs.readFileSync(path.resolve(ROOT, relativePath), "utf8");
}
