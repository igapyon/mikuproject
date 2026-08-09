import fs from "node:fs";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { JSDOM } from "jsdom";

import { WEB_SURFACE_MODULE_RELATIVE_PATHS } from "./lib/runtime-module-paths.mjs";

const args = parseArgs(process.argv.slice(2));
const runtimePath = path.resolve(args.runtimePath);
const runtimeSource = fs.readFileSync(runtimePath, "utf8");

assertRuntimeBoundary(runtimeSource);

const dom = new JSDOM("<!doctype html><html><body></body></html>", {
  url: "http://localhost/"
});
const restoreGlobals = installWindowGlobals(dom.window);

try {
  const runtime = await import(`${pathToFileURL(runtimePath).href}?smoke=${Date.now()}`);
  if (args.expectedVersion && runtime.version !== args.expectedVersion) {
    throw new Error(`runtime version mismatch: expected ${args.expectedVersion}, actual ${runtime.version}`);
  }
  if (!Array.isArray(runtime.embeddedCorePaths) || runtime.embeddedCorePaths.length === 0) {
    throw new Error("runtime embeddedCorePaths is missing or empty");
  }
  const api = runtime.loadMikuProjectRuntime({ expectedVersion: runtime.version });
  if (!api || api !== globalThis.__mikuProjectCoreApi) {
    throw new Error("runtime did not install __mikuProjectCoreApi");
  }
  if (runtime.default !== runtime.loadMikuProjectRuntime) {
    throw new Error("runtime default export is not loadMikuProjectRuntime");
  }
  console.log(`[smoke:browser-runtime] ok ${runtime.version} ${path.basename(runtimePath)}`);
} finally {
  restoreGlobals();
  dom.window.close();
}

function parseArgs(argv) {
  if (argv.length === 0 || argv[0].startsWith("--")) {
    throw new Error("usage: node scripts/smoke-browser-runtime.mjs <runtime.mjs> [--expected-version <version>]");
  }
  const options = { runtimePath: argv[0] };
  for (let index = 1; index < argv.length; index += 1) {
    const token = argv[index];
    if (token !== "--expected-version") {
      throw new Error(`unknown argument: ${token}`);
    }
    const value = argv[index + 1];
    if (value === undefined || value.startsWith("--")) {
      throw new Error(`${token} requires a value`);
    }
    options.expectedVersion = value;
    index += 1;
  }
  return options;
}

function assertRuntimeBoundary(source) {
  const checks = [
    ["node: reference", /node:/],
    ["process reference", /(^|[^A-Za-z0-9_$])process([^A-Za-z0-9_$]|$)/m],
    ["CLI entrypoint", /miku-project-cli\.mjs/]
  ];
  for (const [label, pattern] of checks) {
    if (pattern.test(source)) {
      throw new Error(`browser runtime contains forbidden ${label}`);
    }
  }
  for (const relativePath of WEB_SURFACE_MODULE_RELATIVE_PATHS) {
    if (source.includes(relativePath)) {
      throw new Error(`browser runtime contains Web surface module: ${relativePath}`);
    }
  }
}

function installWindowGlobals(window) {
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
  const previous = new Map();
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
      const value = previous.get(key);
      if (value === undefined) {
        delete globalThis[key];
      } else {
        Object.defineProperty(globalThis, key, {
          configurable: true,
          writable: true,
          value
        });
      }
    }
  };
}
