import fs from "node:fs";
import path from "node:path";

import {
  CORE_API_MODULE_RELATIVE_PATHS,
  WEB_SURFACE_MODULE_RELATIVE_PATHS
} from "./lib/runtime-module-paths.mjs";

const ROOT = process.cwd();
const args = parseArgs(process.argv.slice(2));
const outFile = path.resolve(args.out || path.join(ROOT, "bundle", "miku-project-runtime.mjs"));

buildBrowserRuntime();

function buildBrowserRuntime() {
  assertRequiredFilesExist();
  const packageJson = JSON.parse(readRepoFile("package.json"));
  const runtimeSource = buildRuntimeSource(packageJson.version || "unknown");
  assertBrowserRuntimeBoundary(runtimeSource);
  fs.mkdirSync(path.dirname(outFile), { recursive: true });
  fs.writeFileSync(outFile, runtimeSource, "utf8");
  console.log(`[build:browser-runtime] generated ${path.relative(ROOT, outFile)}`);
}

function parseArgs(argv) {
  const options = {};
  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (token !== "--out") {
      throw new Error(`未対応の引数です: ${token}`);
    }
    const value = argv[index + 1];
    if (value === undefined || value.startsWith("--")) {
      throw new Error(`オプション ${token} には値が必要です`);
    }
    options.out = value;
    index += 1;
  }
  return options;
}

function assertRequiredFilesExist() {
  for (const relativePath of ["package.json", ...CORE_API_MODULE_RELATIVE_PATHS]) {
    if (!fs.existsSync(path.resolve(ROOT, relativePath))) {
      throw new Error(`browser runtime に必要なファイルが見つかりません: ${relativePath}`);
    }
  }
}

function buildRuntimeSource(packageVersion) {
  const moduleBlocks = CORE_API_MODULE_RELATIVE_PATHS.map((relativePath) => {
    const source = prepareBrowserModuleSource(relativePath, readRepoFile(relativePath));
    return `  // embedded core module: ${relativePath}\n${indentSource(source, "  ")}`;
  });

  return [
    `export const version = ${JSON.stringify(packageVersion)};`,
    "",
    "export const embeddedCorePaths = Object.freeze(",
    `${JSON.stringify(CORE_API_MODULE_RELATIVE_PATHS, null, 2)});`,
    "",
    "let loadedCoreApi;",
    "",
    "export function loadMikuProjectRuntime(options = {}) {",
    "  if (options === null || typeof options !== \"object\" || Array.isArray(options)) {",
    "    throw new TypeError(\"loadMikuProjectRuntime options must be an object\");",
    "  }",
    "  if (options.expectedVersion !== undefined && options.expectedVersion !== version) {",
    "    throw new Error(`miku-project runtime version mismatch: expected ${options.expectedVersion}, actual ${version}`);",
    "  }",
    "  if (loadedCoreApi) {",
    "    return loadedCoreApi;",
    "  }",
    "  const existingCoreApi = globalThis.__mikuProjectCoreApi || globalThis.__mikuprojectCoreApi;",
    "  if (existingCoreApi) {",
    "    if (options.reuseExisting !== true) {",
    "      throw new Error(\"miku-project core API is already installed\");",
    "    }",
    "    loadedCoreApi = existingCoreApi;",
    "    return loadedCoreApi;",
    "  }",
    "",
    ...moduleBlocks,
    "",
    "  const coreApi = globalThis.__mikuProjectCoreApi || globalThis.__mikuprojectCoreApi;",
    "  if (!coreApi) {",
    "    throw new Error(\"Failed to boot __mikuProjectCoreApi\");",
    "  }",
    "  loadedCoreApi = coreApi;",
    "  return loadedCoreApi;",
    "}",
    "",
    "export default loadMikuProjectRuntime;",
    ""
  ].join("\n");
}

function prepareBrowserModuleSource(relativePath, source) {
  if (relativePath !== "src/js/ms-office-core.js") {
    return source;
  }

  const nodeFallback = /  function getNodeZlib\(\) \{[\s\S]*?\n  \}\n  function readOfficePackage\(/;
  if (!nodeFallback.test(source)) {
    throw new Error("ms-office-core.js の Node ZIP fallback を特定できませんでした");
  }

  return source
    .replace(
      nodeFallback,
      [
        "  function getSynchronousZipCodec() {",
        "    throw new Error(\"Synchronous ZIP deflate operations are unavailable in the browser runtime. Use stored entries, readZipPackageAsync with DecompressionStream, or inject an async inflater.\");",
        "  }",
        "  function readOfficePackage("
      ].join("\n")
    )
    .replaceAll("getNodeZlib", "getSynchronousZipCodec");
}

function assertBrowserRuntimeBoundary(source) {
  const forbiddenPatterns = [
    { label: "Node.js module reference", pattern: /node:/ },
    { label: "process reference", pattern: /(^|[^A-Za-z0-9_$])process([^A-Za-z0-9_$]|$)/m },
    { label: "CLI entrypoint", pattern: /miku-project-cli\.mjs/ }
  ];
  for (const forbidden of forbiddenPatterns) {
    if (forbidden.pattern.test(source)) {
      throw new Error(`browser runtime に ${forbidden.label} が残っています`);
    }
  }
  for (const relativePath of WEB_SURFACE_MODULE_RELATIVE_PATHS) {
    if (source.includes(relativePath)) {
      throw new Error(`browser runtime に Web surface module が混入しています: ${relativePath}`);
    }
  }
}

function indentSource(source, indent) {
  return source
    .trimEnd()
    .split("\n")
    .map((line) => `${indent}${line}`)
    .join("\n");
}

function readRepoFile(relativePath) {
  return fs.readFileSync(path.resolve(ROOT, relativePath), "utf8");
}
