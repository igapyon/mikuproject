import fs from "node:fs";
import path from "node:path";
import { buildSingleHtmlFromSource } from "./lib/single-html.mjs";

const ROOT = process.cwd();
const args = new Set(process.argv.slice(2));
const buildJs = !args.has("--html-only");
const buildHtml = !args.has("--js-only");

const TARGETS = [
  {
    id: "index",
    srcHtml: "index-src.html",
    outHtml: "index.html"
  },
  {
    id: "mikuproject",
    srcHtml: "mikuproject-src.html",
    outHtml: "mikuproject.html",
    tsOrder: [
      "src/ts/types.ts",
      "src/ts/markdown-escape.ts",
      "src/ts/ai-json-util.ts",
      "src/ts/main-util.ts",
      "src/ts/excel-io.ts",
      "src/ts/msproject-xml.ts",
      "src/ts/project-workbook-schema.ts",
      "src/ts/project-xlsx.ts",
      "src/ts/project-workbook-json.ts",
      "src/ts/project-patch-json.ts",
      "src/ts/wbs-xlsx.ts",
      "src/ts/wbs-markdown.ts",
      "src/ts/wbs-svg.ts",
      "src/ts/main-render.ts",
      "src/ts/main.ts"
    ]
  }
];

const tsModule = await loadTypeScriptModule();

if (buildJs) {
  for (const target of TARGETS) {
    transpileTypeScript(target, tsModule);
  }
}

if (buildHtml) {
  for (const target of TARGETS) {
    const srcPath = path.resolve(ROOT, target.srcHtml);
    const outPath = path.resolve(ROOT, target.outHtml);
    const source = fs.readFileSync(srcPath, "utf8");
    const output = buildSingleHtmlFromSource(applyTemplateValues(source), srcPath, ROOT);
    fs.mkdirSync(path.dirname(outPath), { recursive: true });
    fs.writeFileSync(outPath, output, "utf8");
    console.log(`[build:project] generated ${target.outHtml}`);
  }
}

async function loadTypeScriptModule() {
  try {
    const module = await import("typescript");
    return module.default || module;
  } catch (_error) {
    return null;
  }
}

function transpileTypeScript(target, tsModule) {
  for (const relTsPath of target.tsOrder || []) {
    const tsPath = path.resolve(ROOT, relTsPath);
    const jsPath = path.resolve(
      ROOT,
      relTsPath.replace("/ts/", "/js/").replace(/\.ts$/, ".js")
    );

    const source = fs.readFileSync(tsPath, "utf8");
    let outputText = source;
    if (tsModule) {
      const result = tsModule.transpileModule(source, {
        compilerOptions: {
          target: tsModule.ScriptTarget.ES2019,
          module: tsModule.ModuleKind.None,
          lib: ["ES2020", "DOM"],
          strict: false,
          skipLibCheck: true
        },
        reportDiagnostics: true,
        fileName: tsPath
      });

      if (result.diagnostics && result.diagnostics.length > 0) {
        const errors = result.diagnostics
          .filter((diagnostic) => diagnostic.category === tsModule.DiagnosticCategory.Error)
          .map((diagnostic) => tsModule.flattenDiagnosticMessageText(diagnostic.messageText, "\n"));
        if (errors.length > 0) {
          throw new Error(`TypeScript transpile error in ${relTsPath}:\n${errors.join("\n")}`);
        }
      }
      outputText = result.outputText;
    } else {
      console.warn(
        `[build:project] typescript not found. copied ${relTsPath} -> ${relTsPath.replace("/ts/", "/js/").replace(/\.ts$/, ".js")}`
      );
    }

    fs.mkdirSync(path.dirname(jsPath), { recursive: true });
    fs.writeFileSync(jsPath, outputText, "utf8");
  }
}

function applyTemplateValues(source) {
  return source
    .replaceAll("{{BUILD_DATE}}", formatBuildDate(new Date()))
    .replaceAll("{{AI_PROMPT_TEXT}}", escapeHtml(loadAiPromptText()));
}

function formatBuildDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function loadAiPromptText() {
  const promptPath = path.resolve(ROOT, "docs/mikuproject-ai-json-spec.md");
  return fs.readFileSync(promptPath, "utf8");
}

function escapeHtml(text) {
  return text
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}
