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
      "src/ts/ai-json-spec.ts",
      "src/ts/msproject-ai-views.ts",
      "src/ts/msproject-calendar.ts",
      "src/ts/msproject-samples.ts",
      "src/ts/msproject-csv.ts",
      "src/ts/msproject-validate-helpers.ts",
      "src/ts/msproject-validate.ts",
      "src/ts/msproject-xml-dom.ts",
      "src/ts/msproject-codec.ts",
      "src/ts/msproject-mermaid.ts",
      "src/ts/main-util.ts",
      "src/ts/excel-io-util.ts",
      "src/ts/excel-io-zip.ts",
      "src/ts/excel-io-normalize.ts",
      "src/ts/excel-io-package-xml.ts",
      "src/ts/excel-io-worksheet-build.ts",
      "src/ts/excel-io-worksheet-parse.ts",
      "src/ts/excel-io-workbook-parse.ts",
      "src/ts/excel-io-workbook-build.ts",
      "src/ts/excel-io-styles-build.ts",
      "src/ts/excel-io-styles-parse.ts",
      "src/ts/excel-io.ts",
      "src/ts/msproject-xml.ts",
      "src/ts/project-workbook-schema.ts",
      "src/ts/project-xlsx-import-util.ts",
      "src/ts/project-xlsx-import-project.ts",
      "src/ts/project-xlsx-import-calendars.ts",
      "src/ts/project-xlsx-import-entities.ts",
      "src/ts/project-xlsx-import.ts",
      "src/ts/project-xlsx-export-util.ts",
      "src/ts/project-xlsx-export-project.ts",
      "src/ts/project-xlsx-export-entities.ts",
      "src/ts/project-xlsx-export-calendars.ts",
      "src/ts/project-xlsx-export.ts",
      "src/ts/project-xlsx.ts",
      "src/ts/project-workbook-json-validate.ts",
      "src/ts/project-workbook-json-import.ts",
      "src/ts/project-workbook-json-export.ts",
      "src/ts/project-workbook-json.ts",
      "src/ts/project-patch-json-util.ts",
      "src/ts/project-patch-json-links.ts",
      "src/ts/project-patch-json-tasks.ts",
      "src/ts/project-patch-json-entities.ts",
      "src/ts/project-patch-json-updates.ts",
      "src/ts/project-patch-json-core.ts",
      "src/ts/project-patch-json.ts",
      "src/ts/core-api-msproject-ai.ts",
      "src/ts/core-api-msproject.ts",
      "src/ts/core-api-workbook-xlsx.ts",
  "src/ts/core-api-workbook.ts",
  "src/ts/core-api-ai-json-import.ts",
  "src/ts/core-api-ai-json.ts",
      "src/ts/core-api-external-binary.ts",
      "src/ts/core-api-external-document.ts",
      "src/ts/core-api-external-import.ts",
      "src/ts/core-api-import.ts",
      "src/ts/wbs-dateband.ts",
      "src/ts/wbs-xlsx-base.ts",
      "src/ts/wbs-xlsx-taskmeta.ts",
      "src/ts/wbs-xlsx-layout.ts",
      "src/ts/wbs-xlsx-sections.ts",
      "src/ts/wbs-xlsx-cells.ts",
      "src/ts/wbs-xlsx-export.ts",
      "src/ts/wbs-xlsx-public.ts",
      "src/ts/wbs-xlsx.ts",
      "src/ts/wbs-markdown.ts",
      "src/ts/wbs-svg-zip.ts",
      "src/ts/wbs-svg-bars.ts",
      "src/ts/wbs-svg-viewport.ts",
      "src/ts/wbs-svg-axis.ts",
      "src/ts/wbs-svg-scaffold.ts",
      "src/ts/wbs-svg-labels.ts",
      "src/ts/wbs-svg-calendar.ts",
      "src/ts/wbs-svg-timeline.ts",
      "src/ts/wbs-svg-render.ts",
      "src/ts/wbs-svg-public.ts",
      "src/ts/wbs-svg.ts",
      "src/ts/core-api-report.ts",
      "src/ts/core-api-report-adapters.ts",
      "src/ts/core-api-report-public.ts",
      "src/ts/core-api-registry.ts",
      "src/ts/core-api-public.ts",
      "src/ts/core-api.ts",
      "src/ts/main-io.ts",
      "src/ts/main-import.ts",
      "src/ts/main-export.ts",
      "src/ts/main-input-events.ts",
      "src/ts/main-button-events.ts",
      "src/ts/main-events.ts",
      "src/ts/main-preview.ts",
      "src/ts/main-ui.ts",
      "src/ts/main-support.ts",
      "src/ts/main-samples.ts",
      "src/ts/main-downloads.ts",
      "src/ts/main-output-actions.ts",
      "src/ts/main-import-actions.ts",
      "src/ts/main-xml-actions.ts",
      "src/ts/main-archive-actions.ts",
      "src/ts/main-save-state.ts",
      "src/ts/main-tab-actions.ts",
      "src/ts/main-preview-actions.ts",
      "src/ts/main-transform.ts",
      "src/ts/main-flow.ts",
      "src/ts/main-model.ts",
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

    const source = applyTemplateValues(fs.readFileSync(tsPath, "utf8"));
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
    .replaceAll("{{AI_PROMPT_TEXT_JSON}}", JSON.stringify(loadAiPromptText()))
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
