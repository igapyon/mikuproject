#!/usr/bin/env node

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { loadMikuprojectCoreApi } from "./lib/core-api-loader.mjs";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..");

async function main() {
  const { command, options } = parseArgs(process.argv.slice(2));
  if (options.help || command.length === 0) {
    writeHelp(process.stdout);
    return;
  }

  const loaded = loadMikuprojectCoreApi({ rootDir: repoRoot });
  try {
    const result = await runCommand(command, options, loaded.api);
    if (Array.isArray(result.diagnostics) && result.diagnostics.length > 0) {
      writeDiagnostics(process.stderr, result.diagnostics);
    }
    writeOutput(result.output, options.out);
  } finally {
    loaded.dispose();
  }
}

function parseArgs(argv) {
  const command = [];
  const options = {};

  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (!token.startsWith("--")) {
      command.push(token);
      continue;
    }

    const key = token.slice(2);
    if (key === "help") {
      options.help = true;
      continue;
    }

    const value = argv[index + 1];
    if (value === undefined || value.startsWith("--")) {
      throw new Error(`オプション ${token} には値が必要です`);
    }
    options[key] = value;
    index += 1;
  }

  return { command, options };
}

async function runCommand(command, options, api) {
  const [scope, action] = command;
  const subaction = command[2];

  if (scope === "ai" && action === "spec" && subaction === undefined) {
    return {
      output: `${api.getAiJsonSpecText().trim()}\n`,
      diagnostics: []
    };
  }

  if (scope === "state" && action === "from-draft" && subaction === undefined) {
    const draftDocument = parseJsonLike(await readTextInput(options.in), api);
    ensureKind(api, draftDocument, "project_draft_view", "state from-draft");
    const imported = api.importExternal({
      source: { format: "project_draft_view", document: draftDocument },
      mode: "replace"
    });
    const workbookDocument = api.workbookJson.exportDocument(imported.model);
    return {
      output: `${JSON.stringify(workbookDocument, null, 2)}\n`,
      diagnostics: []
    };
  }

  if (scope === "state" && action === "apply-patch" && subaction === undefined) {
    if (!options.state) {
      throw new Error("state apply-patch には --state workbook.json が必要です");
    }
    const stateDocument = parsePlainJson(await readTextInput(options.state));
    const patchDocument = parseJsonLike(await readTextInput(options.in), api);
    ensureWorkbookJson(api, stateDocument, "state apply-patch");
    ensureKind(api, patchDocument, "patch_json", "state apply-patch");

    const baseModel = api.workbookJson.importAsProjectModel(stateDocument).model;
    const patched = api.importExternal({
      source: { format: "patch_json", document: patchDocument },
      mode: "patch",
      baseModel
    });
    const workbookDocument = api.workbookJson.exportDocument(patched.model);
    return {
      output: `${JSON.stringify(workbookDocument, null, 2)}\n`,
      diagnostics: buildPatchDiagnostics(patched)
    };
  }

  if (scope === "export" && subaction === undefined) {
    if (action === "workbook-json") {
      const stateDocument = parsePlainJson(await readTextInput(options.in));
      ensureWorkbookJson(api, stateDocument, "export workbook-json");
      const model = api.workbookJson.importAsProjectModel(stateDocument).model;
      const exported = api.workbookJson.exportDocument(model);
      return {
        output: `${JSON.stringify(exported, null, 2)}\n`,
        diagnostics: []
      };
    }

    if (action === "xml") {
      const stateDocument = parsePlainJson(await readTextInput(options.in));
      ensureWorkbookJson(api, stateDocument, "export xml");
      const model = api.workbookJson.importAsProjectModel(stateDocument).model;
      return {
        output: `${api.msProject.exportToXml(model)}\n`,
        diagnostics: []
      };
    }

    if (action === "xlsx") {
      const stateDocument = parsePlainJson(await readTextInput(options.in));
      ensureWorkbookJson(api, stateDocument, "export xlsx");
      const model = api.workbookJson.importAsProjectModel(stateDocument).model;
      const workbook = api.xlsx.exportWorkbook(model);
      return {
        output: api.xlsx.encodeWorkbook(workbook),
        diagnostics: []
      };
    }
  }

  throw new Error(`未対応のコマンドです: ${command.join(" ")}`);
}

async function readTextInput(inputPath) {
  if (inputPath) {
    return fs.readFileSync(path.resolve(inputPath), "utf8");
  }

  const chunks = [];
  for await (const chunk of process.stdin) {
    chunks.push(typeof chunk === "string" ? Buffer.from(chunk) : chunk);
  }
  if (chunks.length === 0) {
    throw new Error("入力が必要です。--in を指定するか標準入力を渡してください");
  }
  return Buffer.concat(chunks).toString("utf8");
}

function parsePlainJson(sourceText) {
  return JSON.parse(sourceText);
}

function parseJsonLike(sourceText, api) {
  try {
    return JSON.parse(sourceText);
  } catch (_error) {
    return api.parseAiJsonText(sourceText).document;
  }
}

function ensureWorkbookJson(api, documentLike, context) {
  const kind = api.detectAiJsonDocumentKind(documentLike);
  if (kind !== "workbook_json") {
    throw new Error(`${context} は mikuproject_workbook_json を入力してください`);
  }
}

function ensureKind(api, documentLike, expectedKind, context) {
  const kind = api.detectAiJsonDocumentKind(documentLike);
  if (kind !== expectedKind) {
    throw new Error(`${context} は ${expectedKind} を入力してください`);
  }
}

function buildPatchDiagnostics(result) {
  const lines = [];
  const warningCount = Array.isArray(result.warnings) ? result.warnings.length : 0;
  const changeCount = Array.isArray(result.changes) ? result.changes.length : 0;

  lines.push(`[mikuproject-cli] applied patch_json changes=${changeCount} warnings=${warningCount}`);
  for (const warning of result.warnings || []) {
    lines.push(`[warning] ${formatWarning(warning)}`);
  }
  return lines;
}

function formatWarning(warning) {
  const fragments = [warning.message];
  if (warning.scope) {
    fragments.push(`scope=${warning.scope}`);
  }
  if (warning.uid) {
    fragments.push(`uid=${warning.uid}`);
  }
  if (warning.label) {
    fragments.push(`label=${warning.label}`);
  }
  return fragments.join(" ");
}

function writeDiagnostics(stream, diagnostics) {
  for (const line of diagnostics) {
    stream.write(`${line}\n`);
  }
}

function writeOutput(output, outPath) {
  if (outPath) {
    if (output instanceof Uint8Array) {
      fs.writeFileSync(path.resolve(outPath), output);
      return;
    }
    fs.writeFileSync(path.resolve(outPath), output, "utf8");
    return;
  }

  if (output instanceof Uint8Array) {
    process.stdout.write(Buffer.from(output));
    return;
  }
  process.stdout.write(output);
}

function writeHelp(stream) {
  stream.write([
    "Usage:",
    "  mikuproject ai spec",
    "  mikuproject state from-draft [--in draft.json] [--out workbook.json]",
    "  mikuproject state apply-patch --state workbook.json [--in patch.json] [--out workbook.next.json]",
    "  mikuproject export workbook-json [--in workbook.json] [--out workbook.json]",
    "  mikuproject export xml [--in workbook.json] [--out project.xml]",
    "  mikuproject export xlsx [--in workbook.json] [--out project.xlsx]"
  ].join("\n"));
  stream.write("\n");
}

main().catch((error) => {
  process.stderr.write(`[mikuproject-cli] ${error.message}\n`);
  process.exit(1);
});
