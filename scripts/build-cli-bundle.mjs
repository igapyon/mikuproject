import fs from "node:fs";
import path from "node:path";
import zlib from "node:zlib";

import { CORE_API_MODULE_RELATIVE_PATHS } from "./lib/runtime-module-paths.mjs";

const ROOT = process.cwd();
const args = parseArgs(process.argv.slice(2));
const outFile = path.resolve(args.out || path.join(ROOT, "bundle", "miku-project.mjs"));
const sourcesOutFile = path.resolve(args["sources-out"] || path.join(path.dirname(outFile), "miku-project-sources.tgz"));

const XMLDOM_MODULE_RELATIVE_PATHS = [
  "node_modules/@xmldom/xmldom/lib/conventions.js",
  "node_modules/@xmldom/xmldom/lib/dom.js",
  "node_modules/@xmldom/xmldom/lib/entities.js",
  "node_modules/@xmldom/xmldom/lib/sax.js",
  "node_modules/@xmldom/xmldom/lib/dom-parser.js",
  "node_modules/@xmldom/xmldom/lib/index.js"
];

const SOURCE_ARCHIVE_ROOT = "miku-project-sources";
const SOURCE_ARCHIVE_PATHS = [
  "package.json",
  "package-lock.json",
  "README.md",
  "LICENSE",
  "THIRD-PARTY-NOTICES.md",
  "CONTRIBUTING.md",
  "CONTRIBUTORS.md",
  "CODE_OF_CONDUCT.md",
  "docs",
  "scripts",
  "src",
  "testdata",
  "tests"
];

buildCliBundle();

function buildCliBundle() {
  assertRequiredFilesExist();
  fs.mkdirSync(path.dirname(outFile), { recursive: true });
  fs.writeFileSync(outFile, buildSingleMjsRuntime(), "utf8");
  fs.chmodSync(outFile, 0o755);
  console.log(`[build:cli-bundle] generated ${path.relative(ROOT, outFile)}`);
  fs.mkdirSync(path.dirname(sourcesOutFile), { recursive: true });
  fs.writeFileSync(sourcesOutFile, buildSourcesTgz());
  console.log(`[build:cli-bundle] generated ${path.relative(ROOT, sourcesOutFile)}`);
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
    "scripts/miku-project-cli.mjs",
    ...CORE_API_MODULE_RELATIVE_PATHS,
    ...XMLDOM_MODULE_RELATIVE_PATHS
  ];
  for (const relativePath of requiredFiles) {
    if (!fs.existsSync(path.resolve(ROOT, relativePath))) {
      throw new Error(`CLI bundle に必要なファイルが見つかりません: ${relativePath}`);
    }
  }
  for (const relativePath of SOURCE_ARCHIVE_PATHS) {
    if (!fs.existsSync(path.resolve(ROOT, relativePath))) {
      throw new Error(`source archive に必要なパスが見つかりません: ${relativePath}`);
    }
  }
}

function buildSingleMjsRuntime() {
  const packageJson = JSON.parse(readRepoFile("package.json"));
  const cliSource = stripCliImports(readRepoFile("scripts/miku-project-cli.mjs"));
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
    `const BUNDLED_PACKAGE_VERSION = ${JSON.stringify(packageJson.version || "unknown")};`,
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

  const api = globalThis.__mikuProjectCoreApi || globalThis.__mikuprojectCoreApi;
  if (!api) {
    restoreWindowGlobals();
    throw new Error("Failed to boot __mikuProjectCoreApi");
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

function buildSourcesTgz() {
  const tarParts = [];
  for (const sourcePath of collectSourceArchiveFiles()) {
    const absolutePath = path.resolve(ROOT, sourcePath);
    const archivePath = path.posix.join(SOURCE_ARCHIVE_ROOT, toPosixPath(sourcePath));
    const data = fs.readFileSync(absolutePath);
    tarParts.push(buildTarFileEntry(archivePath, data, getArchiveMode(absolutePath)));
  }
  tarParts.push(Buffer.alloc(1024));
  return zlib.gzipSync(Buffer.concat(tarParts), { level: 9, mtime: 0 });
}

function collectSourceArchiveFiles() {
  const files = [];
  for (const relativePath of SOURCE_ARCHIVE_PATHS) {
    const absolutePath = path.resolve(ROOT, relativePath);
    const stat = fs.statSync(absolutePath);
    if (stat.isDirectory()) {
      collectFilesRecursive(relativePath, files);
      continue;
    }
    if (stat.isFile()) {
      files.push(relativePath);
    }
  }
  return files.sort(compareNormalizedRelativePaths);
}

function collectFilesRecursive(relativeDir, files) {
  const absoluteDir = path.resolve(ROOT, relativeDir);
  const entries = fs.readdirSync(absoluteDir, { withFileTypes: true })
    .sort((a, b) => compareUtf16CodeUnits(toPosixPath(a.name), toPosixPath(b.name)));
  for (const entry of entries) {
    const relativePath = path.join(relativeDir, entry.name);
    if (entry.name === ".DS_Store") {
      continue;
    }
    if (entry.isDirectory()) {
      collectFilesRecursive(relativePath, files);
      continue;
    }
    if (entry.isFile()) {
      files.push(relativePath);
    }
  }
}

function buildTarFileEntry(name, data, mode) {
  const header = Buffer.alloc(512);
  const { namePart, prefixPart } = splitTarPath(name);
  writeTarString(header, namePart, 0, 100);
  writeTarOctal(header, mode, 100, 8);
  writeTarOctal(header, 0, 108, 8);
  writeTarOctal(header, 0, 116, 8);
  writeTarOctal(header, data.length, 124, 12);
  writeTarOctal(header, 0, 136, 12);
  header.fill(0x20, 148, 156);
  header[156] = "0".charCodeAt(0);
  writeTarString(header, "ustar", 257, 6);
  writeTarString(header, "00", 263, 2);
  writeTarString(header, "miku-project", 265, 32);
  writeTarString(header, "miku-project", 297, 32);
  writeTarString(header, prefixPart, 345, 155);

  let checksum = 0;
  for (const byte of header) {
    checksum += byte;
  }
  writeTarChecksum(header, checksum);

  const paddingLength = (512 - (data.length % 512)) % 512;
  return Buffer.concat([header, data, Buffer.alloc(paddingLength)]);
}

function splitTarPath(name) {
  const nameBytes = Buffer.byteLength(name);
  if (nameBytes <= 100) {
    return { namePart: name, prefixPart: "" };
  }

  const parts = name.split("/");
  for (let index = 1; index < parts.length; index += 1) {
    const prefixPart = parts.slice(0, index).join("/");
    const namePart = parts.slice(index).join("/");
    if (Buffer.byteLength(prefixPart) <= 155 && Buffer.byteLength(namePart) <= 100) {
      return { namePart, prefixPart };
    }
  }
  throw new Error(`tar path が長すぎます: ${name}`);
}

function writeTarString(buffer, value, offset, length) {
  const written = buffer.write(value, offset, length, "utf8");
  if (written < length) {
    buffer[offset + written] = 0;
  }
}

function writeTarOctal(buffer, value, offset, length) {
  const text = value.toString(8).padStart(length - 1, "0");
  buffer.write(text.slice(-length + 1), offset, length - 1, "ascii");
  buffer[offset + length - 1] = 0;
}

function writeTarChecksum(buffer, checksum) {
  const text = checksum.toString(8).padStart(6, "0");
  buffer.write(text.slice(-6), 148, 6, "ascii");
  buffer[154] = 0;
  buffer[155] = 0x20;
}

function getArchiveMode(absolutePath) {
  const stat = fs.statSync(absolutePath);
  return stat.mode & 0o111 ? 0o755 : 0o644;
}

function toPosixPath(relativePath) {
  return relativePath.split(path.sep).join("/");
}

function compareNormalizedRelativePaths(left, right) {
  return compareUtf16CodeUnits(toPosixPath(left), toPosixPath(right));
}

function compareUtf16CodeUnits(left, right) {
  if (left < right) {
    return -1;
  }
  if (left > right) {
    return 1;
  }
  return 0;
}

function readRepoFile(relativePath) {
  return fs.readFileSync(path.resolve(ROOT, relativePath), "utf8");
}
