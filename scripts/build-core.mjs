import fs from "node:fs";
import path from "node:path";
import { transform } from "esbuild";

import { CORE_API_MODULE_RELATIVE_PATHS } from "./lib/runtime-module-paths.mjs";

const ROOT = process.cwd();
const MS_OFFICE_CORE_VENDOR_PATH = "src/vendor/miku-ms-office-core-0.5.0.1.mjs";
const CORE_TYPESCRIPT_PATHS = CORE_API_MODULE_RELATIVE_PATHS
  .filter((relativePath) => relativePath !== "src/js/ms-office-core.js")
  .map((relativePath) => relativePath.replace("/js/", "/ts/").replace(/\.js$/, ".ts"));

const tsModule = await loadTypeScriptModule();
await generateMsOfficeCoreAdapter();
for (const relativePath of CORE_TYPESCRIPT_PATHS) {
  transpileTypeScript(relativePath, tsModule);
}

async function loadTypeScriptModule() {
  const module = await import("typescript");
  return module.default || module;
}

async function generateMsOfficeCoreAdapter() {
  const vendorPath = path.resolve(ROOT, MS_OFFICE_CORE_VENDOR_PATH);
  const outputPath = path.resolve(ROOT, "src/js/ms-office-core.js");
  const source = fs.readFileSync(vendorPath, "utf8");
  const result = await transform(source, {
    format: "iife",
    globalName: "__mikuMsOfficeCoreRelease",
    target: "es2019"
  });
  const adapter = `${result.code}
(() => {
  (globalThis).__mikuprojectMsOfficeCore = __mikuMsOfficeCoreRelease;
})();
`;
  fs.writeFileSync(outputPath, adapter, "utf8");
}

function transpileTypeScript(relativePath, typescript) {
  const tsPath = path.resolve(ROOT, relativePath);
  const jsPath = path.resolve(ROOT, relativePath.replace("/ts/", "/js/").replace(/\.ts$/, ".js"));
  const source = applyTemplateValues(fs.readFileSync(tsPath, "utf8"));
  const result = typescript.transpileModule(source, {
    compilerOptions: {
      target: typescript.ScriptTarget.ES2019,
      module: typescript.ModuleKind.None,
      lib: ["ES2020", "DOM"],
      strict: false,
      skipLibCheck: true
    },
    reportDiagnostics: true,
    fileName: tsPath
  });
  const errors = (result.diagnostics || [])
    .filter((diagnostic) => diagnostic.category === typescript.DiagnosticCategory.Error)
    .map((diagnostic) => typescript.flattenDiagnosticMessageText(diagnostic.messageText, "\n"));
  if (errors.length > 0) {
    throw new Error(`TypeScript transpile error in ${relativePath}:\n${errors.join("\n")}`);
  }
  fs.writeFileSync(jsPath, result.outputText, "utf8");
}

function applyTemplateValues(source) {
  return source.replaceAll("{{AI_PROMPT_TEXT_JSON}}", JSON.stringify(loadAiPromptText()));
}

function loadAiPromptText() {
  return fs.readFileSync(path.resolve(ROOT, "docs/miku-project-ai-json-spec.md"), "utf8");
}
