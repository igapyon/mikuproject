#!/usr/bin/env node

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { loadMikuprojectCoreApi } from "./lib/core-api-loader.mjs";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..");
const DIAGNOSTICS_VERSION = 1;

class CliUsageError extends Error {
  constructor(message, code = "usage_error", details = undefined) {
    super(message);
    this.name = "CliUsageError";
    this.code = code;
    this.details = details;
  }
}

class CliProcessingError extends Error {
  constructor(message, code = "processing_error", details = undefined) {
    super(message);
    this.name = "CliProcessingError";
    this.code = code;
    this.details = details;
  }
}

async function main() {
  const rawArgv = process.argv.slice(2);
  const { command, options } = parseArgs(rawArgv);
  if (options.version) {
    process.stdout.write(`miku-project ${getCliVersion()}\n`);
    return;
  }
  if (options.help || command.length === 0) {
    writeHelp(process.stdout);
    return;
  }

  const loaded = loadMikuprojectCoreApi({ rootDir: repoRoot });
  try {
    const result = await runCommand(command, options, loaded.api);
    if (Array.isArray(result.diagnostics) && result.diagnostics.length > 0) {
      writeDiagnostics(process.stderr, result.diagnostics);
    } else if (result.diagnostics && !Array.isArray(result.diagnostics)) {
      writeDiagnostics(process.stderr, result.diagnostics);
    }
    writeOutput(result.output, options);
    if (typeof result.exitCode === "number") {
      process.exitCode = result.exitCode;
    }
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
    if (key === "version") {
      options.version = true;
      continue;
    }

    const value = argv[index + 1];
    if (value === undefined || value.startsWith("--")) {
      throw new CliUsageError(`オプション ${token} には値が必要です`, "missing_option_value", {
        option: token
      });
    }
    options[key] = value;
    index += 1;
  }

  return { command, options };
}

function getCliVersion() {
  if (typeof BUNDLED_PACKAGE_VERSION === "string" && BUNDLED_PACKAGE_VERSION) {
    return BUNDLED_PACKAGE_VERSION;
  }
  try {
    const packageJson = JSON.parse(fs.readFileSync(path.join(repoRoot, "package.json"), "utf8"));
    if (typeof packageJson.version === "string" && packageJson.version) {
      return packageJson.version;
    }
  } catch {
    // Fall through to an explicit unknown marker when package metadata is unavailable.
  }
  return "unknown";
}

async function runCommand(command, options, api) {
  const [scope, action, subject, detail] = command;

  if (scope === "ai" && action === "spec" && subject === undefined) {
    return {
      output: `${api.getAiJsonSpecText().trim()}\n`,
      diagnostics: []
    };
  }

  if (scope === "ai" && action === "detect-kind" && subject === undefined) {
    ensureSingleStdinSource([
      { optionName: "--in", value: options.in, allowImplicitStdin: true }
    ]);
    const diagnosticsFormat = parseDiagnosticsFormat(options.diagnostics);
    const documentLike = parseJsonLike(await readTextInput(options.in), api, "ai detect-kind");
    const kind = api.detectAiJsonDocumentKind(documentLike);
    if (!kind) {
      throw new CliProcessingError("入力 JSON の kind を判定できませんでした", "document_kind_not_detected");
    }
    return {
      output: `${kind}\n`,
      diagnostics: diagnosticsFormat === "json"
        ? buildCommandDiagnostics("detect-kind", {
          io: buildIoDiagnostics({
            inputs: [{ optionName: "--in", value: options.in, allowImplicitStdin: true }],
            output: null
          }),
          detected_kind: kind
        })
        : []
    };
  }

  if (scope === "ai" && action === "export" && detail === undefined) {
    ensureSingleStdinSource([
      { optionName: "--in", value: options.in, allowImplicitStdin: true }
    ]);
    const diagnosticsFormat = parseDiagnosticsFormat(options.diagnostics);
    if (subject === "project-overview") {
      const model = await loadWorkbookModel(api, options.in, "ai export project-overview");
      const exported = api.aiViews.exportProjectOverviewView(model);
      return {
        output: `${JSON.stringify(exported, null, 2)}\n`,
        diagnostics: diagnosticsFormat === "json"
          ? buildCommandDiagnostics("ai export project-overview", {
            io: buildIoDiagnostics({
              inputs: [{ optionName: "--in", value: options.in, allowImplicitStdin: true }],
              output: options.out
            }),
            output_kind: "project_overview_view",
            phase_count: Array.isArray(exported.phases) ? exported.phases.length : 0,
            milestone_count: exported.summary?.milestone_count
          })
          : []
      };
    }

    if (subject === "task-edit") {
      const model = await loadWorkbookModel(api, options.in, "ai export task-edit");
      const requestedTaskUid = resolveTaskEditUid(model, options);
      const exported = api.aiViews.exportTaskEditView(model, requestedTaskUid);
      return {
        output: `${JSON.stringify(exported, null, 2)}\n`,
        diagnostics: diagnosticsFormat === "json"
          ? buildCommandDiagnostics("ai export task-edit", {
            io: buildIoDiagnostics({
              inputs: [{ optionName: "--in", value: options.in, allowImplicitStdin: true }],
              output: options.out
            }),
            output_kind: "task_edit_view",
            target_task_uid: exported.target_task?.uid,
            phase_uid: exported.phase?.uid || null
          })
          : []
      };
    }

    if (subject === "phase-detail") {
      const model = await loadWorkbookModel(api, options.in, "ai export phase-detail");
      const mode = parsePhaseDetailMode(options.mode);
      const requestedPhaseUid = resolvePhaseDetailUid(model, options);
      const exported = api.aiViews.exportPhaseDetailView(model, requestedPhaseUid, {
        mode,
        rootUid: options["root-task-uid"],
        maxDepth: parseOptionalNonNegativeInteger(options["max-depth"], "--max-depth")
      });
      return {
        output: `${JSON.stringify(exported, null, 2)}\n`,
        diagnostics: diagnosticsFormat === "json"
          ? buildCommandDiagnostics("ai export phase-detail", {
            io: buildIoDiagnostics({
              inputs: [{ optionName: "--in", value: options.in, allowImplicitStdin: true }],
              output: options.out
            }),
            output_kind: "phase_detail_view",
            phase_uid: exported.phase?.uid,
            mode: exported.scope?.mode,
            root_task_uid: exported.scope?.root_uid ?? null,
            max_depth: exported.scope?.max_depth ?? null,
            task_count: Array.isArray(exported.tasks) ? exported.tasks.length : 0
          })
          : []
      };
    }

    if (subject === "bundle") {
      const model = await loadWorkbookModel(api, options.in, "ai export bundle");
      const exported = buildAiProjectionBundle(api, model);
      return {
        output: `${JSON.stringify(exported, null, 2)}\n`,
        diagnostics: diagnosticsFormat === "json"
          ? buildCommandDiagnostics("ai export bundle", {
            io: buildIoDiagnostics({
              inputs: [{ optionName: "--in", value: options.in, allowImplicitStdin: true }],
              output: options.out
            }),
            output_kind: "ai_projection_bundle",
            phase_count: Array.isArray(exported.phase_detail_views_full) ? exported.phase_detail_views_full.length : 0,
            task_count: Array.isArray(exported.task_edit_views_full) ? exported.task_edit_views_full.length : 0
          })
          : []
      };
    }

    throw new CliUsageError(`未対応の ai export コマンドです: ${subject || "(missing)"}`, "unsupported_ai_export_command");
  }

  if (scope === "ai" && action === "validate-patch" && subject === undefined) {
    ensureSingleStdinSource([
      { optionName: "--state", value: options.state, allowImplicitStdin: false },
      { optionName: "--in", value: options.in, allowImplicitStdin: true }
    ]);
    const diagnosticsFormat = parseDiagnosticsFormat(options.diagnostics);
    const report = await validatePatchCommand(api, options);
    return {
      output: formatValidationOutput(report, diagnosticsFormat),
      diagnostics: [],
      exitCode: report.ok ? 0 : 1
    };
  }

  if (scope === "state" && action === "from-draft" && subject === undefined) {
    ensureSingleStdinSource([
      { optionName: "--in", value: options.in, allowImplicitStdin: true }
    ]);
    const draftDocument = parseJsonLike(await readTextInput(options.in), api, "state from-draft");
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

  if (scope === "state" && action === "summarize" && subject === undefined) {
    ensureSingleStdinSource([
      { optionName: "--in", value: options.in, allowImplicitStdin: true }
    ]);
    const diagnosticsFormat = parseDiagnosticsFormat(options.diagnostics);
    const model = await loadWorkbookModel(api, options.in, "state summarize");
    const summary = buildStateSummary(api, model);
    return {
      output: `${JSON.stringify(summary, null, 2)}\n`,
      diagnostics: diagnosticsFormat === "json"
        ? buildCommandDiagnostics("state summarize", {
          io: buildIoDiagnostics({
            inputs: [{ optionName: "--in", value: options.in, allowImplicitStdin: true }],
            output: options.out
          }),
          project_name: summary.project?.name,
          phase_count: summary.phase_count,
          task_count: summary.summary?.task_count,
          milestone_count: summary.summary?.milestone_count
        })
        : []
    };
  }

  if (scope === "state" && action === "diff" && subject === undefined) {
    ensureSingleStdinSource([
      { optionName: "--before", value: options.before, allowImplicitStdin: false },
      { optionName: "--after", value: options.after, allowImplicitStdin: false }
    ]);
    const diagnosticsFormat = parseDiagnosticsFormat(options.diagnostics);
    if (!options.before || !options.after) {
      throw new CliUsageError("state diff には --before と --after が必要です", "missing_diff_inputs", {
        required_options: ["--before", "--after"]
      });
    }
    const beforeDocument = parsePlainJson(await readTextInput(options.before), "state diff --before");
    const afterDocument = parsePlainJson(await readTextInput(options.after), "state diff --after");
    ensureWorkbookJson(api, beforeDocument, "state diff --before");
    ensureWorkbookJson(api, afterDocument, "state diff --after");
    const beforeModel = api.workbookJson.importAsProjectModel(beforeDocument).model;
    const diffed = api.workbookJson.importIntoProjectModel(afterDocument, beforeModel);
    const summary = buildStateDiffSummary(diffed);
    return {
      output: `${JSON.stringify(summary, null, 2)}\n`,
      diagnostics: diagnosticsFormat === "json"
        ? buildCommandDiagnostics("state diff", {
          io: buildIoDiagnostics({
            inputs: [
              { optionName: "--before", value: options.before, allowImplicitStdin: false },
              { optionName: "--after", value: options.after, allowImplicitStdin: false }
            ],
            output: options.out
          }),
          warnings: summary.warnings,
          changes_summary: summary.changes_summary
        })
        : []
    };
  }

  if (scope === "state" && action === "apply-patch" && subject === undefined) {
    ensureSingleStdinSource([
      { optionName: "--state", value: options.state, allowImplicitStdin: false },
      { optionName: "--in", value: options.in, allowImplicitStdin: true }
    ]);
    if (!options.state) {
      throw new CliUsageError("state apply-patch には --state workbook.json が必要です", "missing_state_option", {
        option: "--state"
      });
    }
    const stateDocument = parsePlainJson(await readTextInput(options.state), "state apply-patch --state");
    const patchDocument = parseJsonLike(await readTextInput(options.in), api, "state apply-patch --in");
    ensureWorkbookJson(api, stateDocument, "state apply-patch");
    ensureKind(api, patchDocument, "patch_json", "state apply-patch");

    const baseModel = api.workbookJson.importAsProjectModel(stateDocument).model;
    const patched = api.importExternal({
      source: { format: "patch_json", document: patchDocument },
      mode: "patch",
      baseModel
    });
    const workbookDocument = api.workbookJson.exportDocument(patched.model);
    const diagnosticsFormat = parseDiagnosticsFormat(options.diagnostics);
    return {
      output: `${JSON.stringify(workbookDocument, null, 2)}\n`,
      diagnostics: diagnosticsFormat === "json"
        ? buildPatchDiagnosticsJson(patched, "apply-patch", buildIoDiagnostics({
          inputs: [
            { optionName: "--state", value: options.state, allowImplicitStdin: false },
            { optionName: "--in", value: options.in, allowImplicitStdin: true }
          ],
          output: options.out
        }))
        : buildPatchDiagnosticsText(patched, "apply-patch")
    };
  }

  if (scope === "export" && subject === undefined) {
    ensureSingleStdinSource([
      { optionName: "--in", value: options.in, allowImplicitStdin: true }
    ]);
    const diagnosticsFormat = parseDiagnosticsFormat(options.diagnostics);
    if (action === "workbook-json") {
      const stateDocument = parsePlainJson(await readTextInput(options.in), "export workbook-json");
      ensureWorkbookJson(api, stateDocument, "export workbook-json");
      const model = api.workbookJson.importAsProjectModel(stateDocument).model;
      const exported = api.workbookJson.exportDocument(model);
      return {
        output: `${JSON.stringify(exported, null, 2)}\n`,
        diagnostics: diagnosticsFormat === "json"
          ? buildCommandDiagnostics("export workbook-json", {
            io: buildIoDiagnostics({
              inputs: [{ optionName: "--in", value: options.in, allowImplicitStdin: true }],
              output: options.out
            }),
            output_kind: "workbook_json",
            sheet_count: Object.keys(exported.sheets || {}).length
          })
          : []
      };
    }

    if (action === "xml") {
      const stateDocument = parsePlainJson(await readTextInput(options.in), "export xml");
      ensureWorkbookJson(api, stateDocument, "export xml");
      const model = api.workbookJson.importAsProjectModel(stateDocument).model;
      const output = `${api.msProject.exportToXml(model)}\n`;
      return {
        output,
        diagnostics: diagnosticsFormat === "json"
          ? buildCommandDiagnostics("export xml", {
            io: buildIoDiagnostics({
              inputs: [{ optionName: "--in", value: options.in, allowImplicitStdin: true }],
              output: options.out
            }),
            output_kind: "ms_project_xml",
            output_length: output.length
          })
          : []
      };
    }

    if (action === "xlsx") {
      ensureBinaryOutputTarget(options, "export xlsx");
      const stateDocument = parsePlainJson(await readTextInput(options.in), "export xlsx");
      ensureWorkbookJson(api, stateDocument, "export xlsx");
      const model = api.workbookJson.importAsProjectModel(stateDocument).model;
      const workbook = api.xlsx.exportWorkbook(model);
      const output = api.xlsx.encodeWorkbook(workbook);
      return {
        output,
        diagnostics: diagnosticsFormat === "json"
          ? buildCommandDiagnostics("export xlsx", {
            io: buildIoDiagnostics({
              inputs: [{ optionName: "--in", value: options.in, allowImplicitStdin: true }],
              output: options.out,
              outputBase64: options["out-base64"]
            }),
            output_kind: "xlsx",
            byte_length: output.length
          })
          : []
      };
    }
  }

  if (scope === "import" && subject === undefined) {
    const diagnosticsFormat = parseDiagnosticsFormat(options.diagnostics);
    if (action === "xlsx") {
      ensureBinaryInputSource(options, "import xlsx");
      const sourceBytes = await readBinaryInput(options, "import xlsx");
      const imported = api.importExternal({
        source: { format: "xlsx", bytes: sourceBytes },
        mode: "replace"
      });
      const workbookDocument = api.workbookJson.exportDocument(imported.model);
      return {
        output: `${JSON.stringify(workbookDocument, null, 2)}\n`,
        diagnostics: diagnosticsFormat === "json"
          ? buildCommandDiagnostics("import xlsx", {
            io: buildIoDiagnostics({
              inputs: [describeBinaryInputForDiagnostics(options, "import xlsx")],
              output: options.out
            }),
            input_kind: "xlsx",
            output_kind: "workbook_json",
            byte_length: sourceBytes.length,
            sheet_count: Object.keys(workbookDocument.sheets || {}).length
          })
          : []
      };
    }
  }

  if (scope === "report" && subject === undefined) {
    ensureSingleStdinSource([
      { optionName: "--in", value: options.in, allowImplicitStdin: true }
    ]);
    const diagnosticsFormat = parseDiagnosticsFormat(options.diagnostics);
    const stateDocument = parsePlainJson(await readTextInput(options.in), `report ${action}`);
    ensureWorkbookJson(api, stateDocument, `report ${action}`);
    const model = api.workbookJson.importAsProjectModel(stateDocument).model;

    if (action === "wbs-xlsx") {
      ensureBinaryOutputTarget(options, "report wbs-xlsx");
      const output = api.report.wbsXlsx.exportBytes(model);
      return {
        output,
        diagnostics: diagnosticsFormat === "json"
          ? buildCommandDiagnostics("report wbs-xlsx", {
            io: buildIoDiagnostics({
              inputs: [{ optionName: "--in", value: options.in, allowImplicitStdin: true }],
              output: options.out,
              outputBase64: options["out-base64"]
            }),
            output_kind: "wbs_xlsx",
            byte_length: output.length
          })
          : []
      };
    }

    if (action === "daily-svg") {
      const output = `${api.report.svg.exportDaily(model)}\n`;
      return {
        output,
        diagnostics: diagnosticsFormat === "json"
          ? buildCommandDiagnostics("report daily-svg", {
            io: buildIoDiagnostics({
              inputs: [{ optionName: "--in", value: options.in, allowImplicitStdin: true }],
              output: options.out
            }),
            output_kind: "daily_svg",
            output_length: output.length
          })
          : []
      };
    }

    if (action === "weekly-svg") {
      const output = `${api.report.svg.exportWeekly(model)}\n`;
      return {
        output,
        diagnostics: diagnosticsFormat === "json"
          ? buildCommandDiagnostics("report weekly-svg", {
            io: buildIoDiagnostics({
              inputs: [{ optionName: "--in", value: options.in, allowImplicitStdin: true }],
              output: options.out
            }),
            output_kind: "weekly_svg",
            output_length: output.length
          })
          : []
      };
    }

    if (action === "monthly-calendar-svg") {
      ensureBinaryOutputTarget(options, "report monthly-calendar-svg");
      const archive = api.report.svg.exportMonthlyCalendar(model);
      return {
        output: archive.zipBytes,
        diagnostics: diagnosticsFormat === "json"
          ? buildCommandDiagnostics("report monthly-calendar-svg", {
            io: buildIoDiagnostics({
              inputs: [{ optionName: "--in", value: options.in, allowImplicitStdin: true }],
              output: options.out,
              outputBase64: options["out-base64"]
            }),
            output_kind: "monthly_calendar_svg_zip",
            byte_length: archive.zipBytes.length
          })
          : []
      };
    }

    if (action === "all") {
      ensureBinaryOutputTarget(options, "report all");
      const archive = api.report.all.export(model);
      return {
        output: archive.zipBytes,
        diagnostics: diagnosticsFormat === "json"
          ? buildCommandDiagnostics("report all", {
            io: buildIoDiagnostics({
              inputs: [{ optionName: "--in", value: options.in, allowImplicitStdin: true }],
              output: options.out,
              outputBase64: options["out-base64"]
            }),
            output_kind: "report_bundle_zip",
            byte_length: archive.zipBytes.length
          })
          : []
      };
    }

    if (action === "wbs-markdown") {
      const output = `${api.report.wbsMarkdown.export(model)}\n`;
      return {
        output,
        diagnostics: diagnosticsFormat === "json"
          ? buildCommandDiagnostics("report wbs-markdown", {
            io: buildIoDiagnostics({
              inputs: [{ optionName: "--in", value: options.in, allowImplicitStdin: true }],
              output: options.out
            }),
            output_kind: "wbs_markdown",
            output_length: output.length
          })
          : []
      };
    }

    if (action === "mermaid") {
      const output = `${api.report.mermaid.exportGantt(model)}\n`;
      return {
        output,
        diagnostics: diagnosticsFormat === "json"
          ? buildCommandDiagnostics("report mermaid", {
            io: buildIoDiagnostics({
              inputs: [{ optionName: "--in", value: options.in, allowImplicitStdin: true }],
              output: options.out
            }),
            output_kind: "mermaid",
            output_length: output.length
          })
          : []
      };
    }
  }

  throw new CliUsageError(`未対応のコマンドです: ${command.join(" ")}`, "unsupported_command");
}

async function readTextInput(inputPath) {
  if (inputPath && inputPath !== "-") {
    return fs.readFileSync(path.resolve(inputPath), "utf8");
  }

  const chunks = [];
  for await (const chunk of process.stdin) {
    chunks.push(typeof chunk === "string" ? Buffer.from(chunk) : chunk);
  }
  if (chunks.length === 0) {
    throw new CliUsageError("入力が必要です。--in を指定するか標準入力を渡してください", "missing_input", {
      option: "--in"
    });
  }
  return Buffer.concat(chunks).toString("utf8");
}

async function readStdinBytes() {
  const chunks = [];
  for await (const chunk of process.stdin) {
    chunks.push(typeof chunk === "string" ? Buffer.from(chunk) : chunk);
  }
  if (chunks.length === 0) {
    throw new CliUsageError("入力が必要です。--in を指定するか標準入力を渡してください", "missing_input", {
      option: "--in"
    });
  }
  return Buffer.concat(chunks);
}

function decodeBase64Input(sourceText, context) {
  const normalized = sourceText.replace(/\s+/g, "");
  if (!normalized) {
    throw new CliUsageError(`${context} の Base64 入力が空です`, "invalid_base64_input", {
      context
    });
  }
  if (!/^[A-Za-z0-9+/]*={0,2}$/.test(normalized) || normalized.length % 4 !== 0) {
    throw new CliUsageError(`${context} の Base64 入力を解析できませんでした`, "invalid_base64_input", {
      context
    });
  }
  return Buffer.from(normalized, "base64");
}

async function readBinaryInput(options, context) {
  if (options["in-base64"]) {
    if (options["in-base64"] === "-") {
      return decodeBase64Input((await readStdinBytes()).toString("utf8"), context);
    }
    return decodeBase64Input(fs.readFileSync(path.resolve(options["in-base64"]), "utf8"), context);
  }
  if (options.in && options.in !== "-") {
    return fs.readFileSync(path.resolve(options.in));
  }
  throw new CliUsageError(`${context} には --in <path> または --in-base64 - が必要です`, "missing_binary_input", {
    required_options: ["--in", "--in-base64"]
  });
}

function parsePlainJson(sourceText, context = "input") {
  try {
    return JSON.parse(sourceText);
  } catch (_error) {
    throw new CliProcessingError(`${context} の JSON を解析できませんでした`, "invalid_json_input", {
      context
    });
  }
}

function parseJsonLike(sourceText, api, context = "input") {
  try {
    return JSON.parse(sourceText);
  } catch (_error) {
    try {
      return api.parseAiJsonText(sourceText).document;
    } catch (_parseError) {
      throw new CliProcessingError(`${context} の JSON を解析できませんでした`, "invalid_json_input", {
        context
      });
    }
  }
}

function ensureSingleStdinSource(inputs) {
  const stdinOptions = inputs
    .filter((input) => input.value === "-" || (input.allowImplicitStdin && input.value === undefined))
    .map((input) => input.optionName);

  if (stdinOptions.length > 1) {
    throw new CliUsageError(`標準入力を使える入力オプションは 1 つだけです: ${stdinOptions.join(", ")}`, "multiple_stdin_sources", {
      options: stdinOptions
    });
  }
}

function ensureBinaryInputSource(options, commandLabel) {
  if (options.in && options["in-base64"]) {
    throw new CliUsageError(`${commandLabel} では --in と --in-base64 を同時に指定できません`, "multiple_input_sources", {
      options: ["--in", "--in-base64"]
    });
  }
  if (options.in === "-") {
    throw new CliUsageError(`${commandLabel} の binary stdin には --in-base64 - を指定してください`, "binary_stdin_requires_base64", {
      option: "--in-base64"
    });
  }
  if (options["in-base64"] && options["in-base64"] !== "-") {
    return;
  }
  if (options.in && options.in !== "-") {
    return;
  }
  if (options["in-base64"] === "-") {
    return;
  }
  throw new CliUsageError(`${commandLabel} には --in <path> または --in-base64 - が必要です`, "missing_binary_input", {
    required_options: ["--in", "--in-base64"]
  });
}

function ensureBinaryOutputTarget(options, commandLabel) {
  if (options.out && options["out-base64"]) {
    throw new CliUsageError(`${commandLabel} では --out と --out-base64 を同時に指定できません`, "multiple_output_targets", {
      options: ["--out", "--out-base64"]
    });
  }
  if (options["out-base64"] && options["out-base64"] !== "-") {
    throw new CliUsageError(`${commandLabel} の --out-base64 は - のみ対応です`, "invalid_base64_output_target", {
      option: "--out-base64",
      expected: "-"
    });
  }
  if (options["out-base64"] === "-") {
    return;
  }
  if (options.out && options.out !== "-") {
    return;
  }
  throw new CliUsageError(`${commandLabel} は binary artifact のため --out <path> が必要です`, "binary_stdout_not_supported", {
    option: "--out",
    command: commandLabel,
    expected: "file_or_base64_stdout"
  });
}

async function loadWorkbookModel(api, inputPath, context) {
  const stateDocument = parsePlainJson(await readTextInput(inputPath), context);
  ensureWorkbookJson(api, stateDocument, context);
  return api.workbookJson.importAsProjectModel(stateDocument).model;
}

function parseOptionalNonNegativeInteger(value, optionName) {
  if (value === undefined) {
    return undefined;
  }
  if (!/^\d+$/.test(value)) {
    throw new CliUsageError(`${optionName} には 0 以上の整数を指定してください`, "invalid_integer_option", {
      option: optionName,
      expected: "non_negative_integer"
    });
  }
  return Number(value);
}

function parseDiagnosticsFormat(value) {
  if (value === undefined) {
    return "text";
  }
  if (value === "text" || value === "json") {
    return value;
  }
  throw new CliUsageError("--diagnostics には text または json を指定してください", "invalid_diagnostics_option", {
    option: "--diagnostics",
    expected_values: ["text", "json"]
  });
}

function parsePhaseDetailMode(value) {
  if (value === undefined) {
    return "scoped";
  }
  if (value === "scoped" || value === "full") {
    return value;
  }
  throw new CliUsageError("--mode には scoped または full を指定してください", "invalid_mode_option", {
    option: "--mode",
    expected_values: ["scoped", "full"]
  });
}

function parseSelectMode(value) {
  if (value === undefined) {
    return "auto";
  }
  if (value === "auto" || value === "first-task" || value === "first-phase" || value === "uid") {
    return value;
  }
  throw new CliUsageError("--select には auto / first-task / first-phase / uid を指定してください", "invalid_select_option", {
    option: "--select",
    expected_values: ["auto", "first-task", "first-phase", "uid"]
  });
}

function resolveTaskEditUid(model, options) {
  const select = parseSelectMode(options.select);
  if (options["task-uid"]) {
    return options["task-uid"];
  }
  if (select === "auto" || select === "first-task") {
    return findFirstTaskUid(model);
  }
  if (select === "uid") {
    throw new CliUsageError("ai export task-edit --select uid には --task-uid が必要です", "missing_task_uid", {
      option: "--task-uid"
    });
  }
  throw new CliUsageError(`ai export task-edit では --select ${select} を使えません`, "invalid_select_option");
}

function resolvePhaseDetailUid(model, options) {
  const select = parseSelectMode(options.select);
  if (options["phase-uid"]) {
    return options["phase-uid"];
  }
  if (select === "auto" || select === "first-phase") {
    return findFirstPhaseUid(model);
  }
  if (select === "uid") {
    throw new CliUsageError("ai export phase-detail --select uid には --phase-uid が必要です", "missing_phase_uid", {
      option: "--phase-uid"
    });
  }
  throw new CliUsageError(`ai export phase-detail では --select ${select} を使えません`, "invalid_select_option");
}

function findFirstTaskUid(model) {
  const tasks = (model.tasks || []).filter((task) => String(task.uid || "").trim() && String(task.uid || "").trim() !== "0");
  const firstTask = tasks.find((task) => !task.summary) || tasks[0];
  return firstTask?.uid;
}

function findFirstPhaseUid(model) {
  const firstPhase = (model.tasks || []).find((task) =>
    String(task.uid || "").trim() !== "0" && task.summary && task.outlineLevel === 1
  );
  return firstPhase?.uid;
}

function ensureWorkbookJson(api, documentLike, context) {
  const kind = api.detectAiJsonDocumentKind(documentLike);
  if (kind !== "workbook_json") {
    throw new CliProcessingError(`${context} は mikuproject_workbook_json を入力してください`, "invalid_workbook_kind", {
      context,
      expected_kind: "workbook_json",
      actual_kind: kind || null
    });
  }
}

function ensureKind(api, documentLike, expectedKind, context) {
  const kind = api.detectAiJsonDocumentKind(documentLike);
  if (kind !== expectedKind) {
    const code = expectedKind === "patch_json" ? "invalid_patch_kind" : "invalid_document_kind";
    throw new CliProcessingError(`${context} は ${expectedKind} を入力してください`, code, {
      context,
      expected_kind: expectedKind,
      actual_kind: kind || null
    });
  }
}

function buildPatchDiagnosticsText(result, contextLabel) {
  const lines = [];
  const warningCount = Array.isArray(result.warnings) ? result.warnings.length : 0;
  const changeCount = Array.isArray(result.changes) ? result.changes.length : 0;
  const status = determineStatus({
    ok: true,
    warnings: result.warnings || [],
    errors: [],
    changes_summary: summarizeChanges(result.changes || [])
  });

  lines.push(`[miku-project-cli] ${contextLabel} patch_json status=${status} changes=${changeCount} warnings=${warningCount}`);
  for (const warning of result.warnings || []) {
    lines.push(`[warning] ${formatWarning(warning)}`);
  }
  return lines;
}

function buildPatchDiagnosticsJson(result, contextLabel, io) {
  return buildCommandDiagnostics(contextLabel, {
    io,
    warnings: result.warnings || [],
    errors: [],
    changes_summary: summarizeChanges(result.changes || [])
  });
}

function buildIoDiagnostics({ inputs, output, outputBase64 }) {
  return {
    inputs: (inputs || []).map((input) => describeInputSource(input)),
    output: describeOutputTarget(output, outputBase64)
  };
}

function describeInputSource(input) {
  if (input.base64 && input.value === "-") {
    return {
      option: input.optionName,
      mode: "stdin_base64"
    };
  }
  if (input.base64) {
    return {
      option: input.optionName,
      mode: "file_base64",
      path: input.value
    };
  }
  if (input.value === "-") {
    return {
      option: input.optionName,
      mode: "stdin"
    };
  }
  if (input.value === undefined && input.allowImplicitStdin) {
    return {
      option: input.optionName,
      mode: "stdin_implicit"
    };
  }
  return {
    option: input.optionName,
    mode: "file",
    path: input.value
  };
}

function describeBinaryInputForDiagnostics(options, _context) {
  if (options["in-base64"]) {
    return {
      optionName: "--in-base64",
      value: options["in-base64"],
      base64: true
    };
  }
  return {
    optionName: "--in",
    value: options.in,
    allowImplicitStdin: false
  };
}

function describeOutputTarget(outPath, outBase64Path) {
  if (outBase64Path === "-") {
    return {
      mode: "stdout_base64"
    };
  }
  if (!outPath || outPath === "-") {
    return {
      mode: "stdout"
    };
  }
  return {
    mode: "file",
    path: outPath
  };
}

function buildIoDiagnosticsFromArgv(argv) {
  const command = summarizeCommandFromArgv(argv);
  const inputs = [];
  let output;

  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (!token.startsWith("--")) {
      continue;
    }
    const value = argv[index + 1];
    if (token === "--in" || token === "--state" || token === "--before" || token === "--after") {
      inputs.push({
        option: token,
        mode: value === "-" ? "stdin" : "file",
        ...(value && value !== "-" ? { path: value } : {})
      });
      index += 1;
      continue;
    }
    if (token === "--in-base64") {
      inputs.push({
        option: token,
        mode: value === "-" ? "stdin_base64" : "file_base64",
        ...(value && value !== "-" ? { path: value } : {})
      });
      index += 1;
      continue;
    }
    if (token === "--out") {
      output = value === "-" ? { mode: "stdout" } : { mode: "file", path: value };
      index += 1;
      continue;
    }
    if (token === "--out-base64") {
      output = value === "-" ? { mode: "stdout_base64" } : { mode: "file_base64", path: value };
      index += 1;
      continue;
    }
    if (token !== "--help") {
      index += 1;
    }
  }

  for (const implicitOption of inferImplicitInputOptions(command, inputs)) {
    inputs.push({
      option: implicitOption,
      mode: "stdin_implicit"
    });
  }

  return {
    inputs,
    output: output || { mode: "stdout" }
  };
}

function inferImplicitInputOptions(command, existingInputs) {
  const existingOptions = new Set(existingInputs.map((input) => input.option));

  const implicitCandidatesByCommand = {
    "ai detect-kind": ["--in"],
    "ai export project-overview": ["--in"],
    "ai export task-edit": ["--in"],
    "ai export phase-detail": ["--in"],
    "ai export bundle": ["--in"],
    "ai validate-patch": ["--in"],
    "state from-draft": ["--in"],
    "state summarize": ["--in"],
    "state apply-patch": ["--in"],
    "import xlsx": [],
    "export workbook-json": ["--in"],
    "export xml": ["--in"],
    "export xlsx": ["--in"],
    "report wbs-xlsx": ["--in"],
    "report daily-svg": ["--in"],
    "report weekly-svg": ["--in"],
    "report monthly-calendar-svg": ["--in"],
    "report all": ["--in"],
    "report wbs-markdown": ["--in"],
    "report mermaid": ["--in"]
  };

  return (implicitCandidatesByCommand[command] || []).filter((option) => !existingOptions.has(option));
}

function summarizeChanges(changes) {
  const byScope = {
    project: 0,
    tasks: 0,
    resources: 0,
    assignments: 0,
    calendars: 0
  };
  const affectedItems = {
    project: new Set(),
    tasks: new Set(),
    resources: new Set(),
    assignments: new Set(),
    calendars: new Set()
  };

  for (const change of changes) {
    if (!Object.hasOwn(byScope, change.scope)) {
      continue;
    }
    byScope[change.scope] += 1;
    affectedItems[change.scope].add(change.uid);
  }

  return {
    total_changes: changes.length,
    by_scope: byScope,
    affected_items: {
      project: affectedItems.project.size,
      tasks: affectedItems.tasks.size,
      resources: affectedItems.resources.size,
      assignments: affectedItems.assignments.size,
      calendars: affectedItems.calendars.size
    }
  };
}

function buildStateSummary(api, model) {
  const overview = api.aiViews.exportProjectOverviewView(model);
  return {
    kind: "state_summary",
    project: overview.project,
    summary: overview.summary,
    phase_count: Array.isArray(overview.phases) ? overview.phases.length : 0,
    top_level_dependency_count: Array.isArray(overview.top_level_dependencies) ? overview.top_level_dependencies.length : 0,
    phases: (overview.phases || []).map((phase) => ({
      uid: phase.uid,
      name: phase.name,
      task_count: phase.task_count,
      milestone_count: phase.milestone_count,
      planned_start: phase.planned_start,
      planned_finish: phase.planned_finish
    })),
    major_milestones: (overview.milestones || []).slice(0, 10)
  };
}

function buildStateDiffSummary(result) {
  return {
    kind: "state_diff_summary",
    warnings: result.warnings || [],
    changes_summary: summarizeChanges(result.changes || []),
    changed_items: buildChangedItems(result.changes || [])
  };
}

function buildAiProjectionBundle(api, model) {
  const projectOverview = api.aiViews.exportProjectOverviewView(model);
  const phaseDetailViewsFull = (projectOverview.phases || [])
    .map((phase) => phase?.uid)
    .filter(Boolean)
    .map((phaseUid) => api.aiViews.exportPhaseDetailView(model, phaseUid, { mode: "full" }));
  const taskEditViewsFull = (model.tasks || [])
    .filter((task) => !(task.uid === "0" || task.summary))
    .map((task) => api.aiViews.exportTaskEditView(model, task.uid));

  return {
    view_type: "ai_projection_bundle",
    project_overview_view: projectOverview,
    phase_detail_views_full: phaseDetailViewsFull,
    task_edit_views_full: taskEditViewsFull
  };
}

function buildCommandDiagnostics(context, extra = {}) {
  const warnings = Array.isArray(extra.warnings) ? extra.warnings : [];
  const errors = Array.isArray(extra.errors) ? extra.errors : [];
  const status = typeof extra.status === "string"
    ? extra.status
    : determineStatus({
      ok: extra.ok !== false,
      warnings,
      errors,
      changes_summary: extra.changes_summary
    });
  const exitCode = typeof extra.exit_code === "number"
    ? extra.exit_code
    : determineExitCodeFromStatus(status);
  return {
    ok: extra.ok !== false,
    diagnostics_version: DIAGNOSTICS_VERSION,
    command: context,
    context,
    status,
    exit_code: exitCode,
    warning_count: warnings.length,
    error_count: errors.length,
    warnings,
    errors,
    ...extra
  };
}

function determineExitCodeFromStatus(status) {
  if (status === "error") {
    return 1;
  }
  return 0;
}

function determineStatus(input) {
  if (!input.ok || (input.errors || []).length > 0) {
    return "error";
  }
  if ((input.warnings || []).length > 0) {
    return "warning";
  }
  if (input.changes_summary && input.changes_summary.total_changes === 0) {
    return "noop";
  }
  return "success";
}

function buildChangedItems(changes) {
  const grouped = {
    project: [],
    tasks: [],
    resources: [],
    assignments: [],
    calendars: []
  };

  for (const change of changes) {
    if (!Object.hasOwn(grouped, change.scope)) {
      continue;
    }
    grouped[change.scope].push({
      uid: change.uid,
      label: change.label,
      field: change.field,
      before: change.before,
      after: change.after
    });
  }

  return grouped;
}

async function validatePatchCommand(api, options) {
  const diagnosticsFormat = parseDiagnosticsFormat(options.diagnostics);
  try {
    if (!options.state) {
      throw new CliUsageError("ai validate-patch には --state workbook.json が必要です", "missing_state_option", {
        option: "--state"
      });
    }
    const stateDocument = parsePlainJson(await readTextInput(options.state), "ai validate-patch --state");
    const patchDocument = parseJsonLike(await readTextInput(options.in), api, "ai validate-patch --in");
    ensureWorkbookJson(api, stateDocument, "ai validate-patch");
    ensureKind(api, patchDocument, "patch_json", "ai validate-patch");

    const baseModel = api.workbookJson.importAsProjectModel(stateDocument).model;
    const patched = api.patchJson.applyToProjectModel(patchDocument, baseModel);
    return {
      ok: true,
      diagnostics_version: DIAGNOSTICS_VERSION,
      command: "ai validate-patch",
      status: determineStatus({
        ok: true,
        warnings: patched.warnings || [],
        errors: [],
        changes_summary: summarizeChanges(patched.changes || [])
      }),
      exit_code: 0,
      warning_count: (patched.warnings || []).length,
      error_count: 0,
      io: buildIoDiagnostics({
        inputs: [
          { optionName: "--state", value: options.state, allowImplicitStdin: false },
          { optionName: "--in", value: options.in, allowImplicitStdin: true }
        ],
        output: null
      }),
      warnings: patched.warnings || [],
      errors: [],
      changes_summary: summarizeChanges(patched.changes || []),
      diagnostics_format: diagnosticsFormat
    };
  } catch (error) {
    if (error instanceof CliUsageError) {
      throw error;
    }
    return {
      ok: false,
      diagnostics_version: DIAGNOSTICS_VERSION,
      command: "ai validate-patch",
      status: "error",
      exit_code: 1,
      warning_count: 0,
      error_count: 1,
      io: buildIoDiagnostics({
        inputs: [
          { optionName: "--state", value: options.state, allowImplicitStdin: false },
          { optionName: "--in", value: options.in, allowImplicitStdin: true }
        ],
        output: null
      }),
      error_type: "processing_error",
      error_code: inferErrorCode(error, "ai validate-patch"),
      error_details: extractErrorDetails(error),
      warnings: [],
      errors: [buildErrorItem(error, "ai validate-patch")],
      changes_summary: summarizeChanges([]),
      diagnostics_format: diagnosticsFormat
    };
  }
}

function formatValidationOutput(report, diagnosticsFormat) {
  if (diagnosticsFormat === "json") {
    return `${JSON.stringify(report, null, 2)}\n`;
  }

  const lines = [
    `[miku-project-cli] validate-patch ok=${report.ok ? "true" : "false"} status=${report.status} warnings=${report.warnings.length} errors=${report.errors.length} changes=${report.changes_summary.total_changes}`
  ];
  for (const error of report.errors) {
    lines.push(`[error] ${error.message}`);
  }
  for (const warning of report.warnings) {
    lines.push(`[warning] ${formatWarning(warning)}`);
  }
  lines.push(`[changes] project=${report.changes_summary.by_scope.project} tasks=${report.changes_summary.by_scope.tasks} resources=${report.changes_summary.by_scope.resources} assignments=${report.changes_summary.by_scope.assignments} calendars=${report.changes_summary.by_scope.calendars}`);
  return `${lines.join("\n")}\n`;
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
  if (!Array.isArray(diagnostics)) {
    stream.write(`${JSON.stringify(diagnostics, null, 2)}\n`);
    return;
  }
  for (const line of diagnostics) {
    stream.write(`${line}\n`);
  }
}

function writeOutput(output, options) {
  const outPath = options.out;
  const outBase64Path = options["out-base64"];
  if (outBase64Path) {
    if (outBase64Path !== "-") {
      throw new CliUsageError("--out-base64 は - のみ対応です", "invalid_base64_output_target", {
        option: "--out-base64",
        expected: "-"
      });
    }
    if (!(output instanceof Uint8Array)) {
      throw new CliUsageError("--out-base64 は binary output 専用です", "base64_output_requires_binary", {
        option: "--out-base64"
      });
    }
    process.stdout.write(`${Buffer.from(output).toString("base64")}\n`);
    return;
  }

  if (outPath && outPath !== "-") {
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
    "miku-project - local MS Project XML bridge for WBS, AI workflows, and reports",
    "",
    "Usage:",
    "  miku-project --version",
    "  miku-project ai spec",
    "  miku-project ai export project-overview [--in workbook.json|-] [--diagnostics text|json] [--out overview.editjson|-]",
    "  miku-project ai export task-edit [--in workbook.json|-] [--task-uid 123] [--select auto|first-task|uid] [--diagnostics text|json] [--out task.editjson|-]",
    "  miku-project ai export phase-detail [--in workbook.json|-] [--phase-uid 100] [--select auto|first-phase|uid] [--mode scoped|full] [--root-task-uid 123] [--max-depth 2] [--diagnostics text|json] [--out phase.editjson|-]",
    "  miku-project ai export bundle [--in workbook.json|-] [--diagnostics text|json] [--out bundle.editjson|-]",
    "  miku-project ai detect-kind [--in document.json|-] [--diagnostics text|json]",
    "  miku-project ai validate-patch --state workbook.json [--in patch.json] [--diagnostics text|json]",
    "  miku-project state from-draft [--in draft.json|-] [--out workbook.json|-]",
    "  miku-project state summarize [--in workbook.json|-] [--diagnostics text|json]",
    "  miku-project state diff --before workbook.before.json --after workbook.after.json [--diagnostics text|json]",
    "  miku-project state apply-patch --state workbook.json|- [--in patch.json|-] [--diagnostics text|json] [--out workbook.next.json|-]",
    "  miku-project import xlsx (--in workbook.xlsx|--in-base64 -) [--diagnostics text|json] [--out workbook.json|-]",
    "  miku-project export workbook-json [--in workbook.json|-] [--diagnostics text|json] [--out workbook.json|-]",
    "  miku-project export xml [--in workbook.json|-] [--diagnostics text|json] [--out project.xml|-]",
    "  miku-project export xlsx [--in workbook.json|-] [--diagnostics text|json] (--out project.xlsx|--out-base64 -)",
    "  miku-project report wbs-xlsx [--in workbook.json|-] [--diagnostics text|json] (--out report.xlsx|--out-base64 -)",
    "  miku-project report daily-svg [--in workbook.json|-] [--diagnostics text|json] [--out report.svg|-]",
    "  miku-project report weekly-svg [--in workbook.json|-] [--diagnostics text|json] [--out report.svg|-]",
    "  miku-project report monthly-calendar-svg [--in workbook.json|-] [--diagnostics text|json] (--out report.zip|--out-base64 -)",
    "  miku-project report all [--in workbook.json|-] [--diagnostics text|json] (--out report-bundle.zip|--out-base64 -)",
    "  miku-project report wbs-markdown [--in workbook.json|-] [--diagnostics text|json] [--out report.md|-]",
    "  miku-project report mermaid [--in workbook.json|-] [--diagnostics text|json] [--out report.mmd|-]",
    "",
    "Description:",
    "  Read local MS Project-related data, then inspect, validate, transform, or export it.",
    "  The CLI is the artifact-based path for AI projection, patch validation, application, and diff workflows.",
    "",
    "Default behavior:",
    "  Text and JSON commands write their primary result to stdout unless --out <path> is supplied.",
    "  Where the Usage form permits it, --in - or an omitted text input reads UTF-8 text from stdin.",
    "  --diagnostics defaults to text. Use --diagnostics json for structured diagnostics on stderr.",
    "",
    "Inputs:",
    "  File paths resolve from the current working directory. Use - only where the Usage form allows stdin.",
    "  XLSX binary stdin requires --in-base64 -. JSON, XML, and text inputs use --in <path> or stdin as shown.",
    "",
    "Outputs:",
    "  JSON, XML, SVG, Markdown, and Mermaid text are written to stdout or --out <path>.",
    "  XLSX and ZIP artifacts require --out <path>, or --out-base64 - for Base64 on stdout.",
    "",
    "Generated artifacts:",
    "  This CLI writes only the output explicitly named by --out. Existing output files are generated artifacts, not source files.",
    "  npm run build:cli-bundle regenerates bundle/miku-project.mjs and bundle/miku-project-sources.tgz; do not edit them by hand.",
    "",
    "Overwrite behavior:",
    "  --out <path> overwrites an existing file without an interactive prompt.",
    "",
    "Machine-readable output contract:",
    "  JSON commands emit their primary document on stdout or --out. With --diagnostics json, diagnostics are JSON on stderr.",
    "  JSON field definitions are documented in docs/miku-project-ai-json-spec.md and docs/import-export-workflows.md.",
    "",
    "Diagnostics / warnings:",
    "  Text diagnostics and usage errors go to stderr. Structured diagnostics include command, I/O, status, and error details.",
    "",
    "Exit codes:",
    "  0 success; 1 processing or validation failure; 2 invalid command usage.",
    "",
    "Examples:",
    "  miku-project --version",
    "  miku-project ai spec",
    "",
    "References:",
    "  README.md, docs/miku-project-ai-json-spec.md, docs/import-export-workflows.md"
  ].join("\n"));
  stream.write("\n");
}

main().catch((error) => {
  const rawArgv = process.argv.slice(2);
  const diagnosticsFormat = detectRequestedDiagnosticsFormat(rawArgv);
  const exitCode = error instanceof CliUsageError ? 2 : 1;
  if (diagnosticsFormat === "json") {
    process.stderr.write(`${JSON.stringify(buildErrorDiagnostics(rawArgv, error, exitCode), null, 2)}\n`);
  } else {
    process.stderr.write(`[miku-project-cli] ${error.message}\n`);
  }
  process.exit(exitCode);
});

function detectRequestedDiagnosticsFormat(argv) {
  for (let index = 0; index < argv.length; index += 1) {
    if (argv[index] === "--diagnostics" && argv[index + 1] === "json") {
      return "json";
    }
  }
  return "text";
}

function buildErrorDiagnostics(argv, error, exitCode) {
  const context = summarizeCommandFromArgv(argv);
  const errorCode = inferErrorCode(error, context);
  return {
    ok: false,
    diagnostics_version: DIAGNOSTICS_VERSION,
    command: context,
    context,
    status: "error",
    exit_code: exitCode,
    warning_count: 0,
    error_count: 1,
    io: buildIoDiagnosticsFromArgv(argv),
    error_type: error instanceof CliUsageError ? "usage_error" : "processing_error",
    error_code: errorCode,
    error_details: extractErrorDetails(error),
    warnings: [],
    errors: [buildErrorItem(error, context)]
  };
}

function summarizeCommandFromArgv(argv) {
  const command = [];
  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (!token.startsWith("--")) {
      command.push(token);
      continue;
    }
    const key = token.slice(2);
    if (key !== "help") {
      index += 1;
    }
  }
  return command.join(" ") || "cli";
}

function inferErrorCode(error, context) {
  if (error instanceof CliUsageError && typeof error.code === "string") {
    return error.code;
  }
  if (error instanceof CliProcessingError && typeof error.code === "string") {
    return error.code;
  }
  const message = error instanceof Error ? error.message : String(error);

  if (message.includes("オプション --") && message.includes("には値が必要")) {
    return "missing_option_value";
  }
  if (message.includes("入力が必要です。--in を指定するか標準入力を渡してください")) {
    return "missing_input";
  }
  if (message.includes("標準入力を使える入力オプションは 1 つだけです")) {
    return "multiple_stdin_sources";
  }
  if (message.includes("--diagnostics には text または json")) {
    return "invalid_diagnostics_option";
  }
  if (message.includes("--mode には scoped または full")) {
    return "invalid_mode_option";
  }
  if (message.includes("--select には auto / first-task / first-phase / uid")) {
    return "invalid_select_option";
  }
  if (message.includes("0 以上の整数を指定してください")) {
    return "invalid_integer_option";
  }
  if (message.includes("には --state workbook.json が必要です")) {
    return "missing_state_option";
  }
  if (message.includes("state diff には --before と --after が必要です")) {
    return "missing_diff_inputs";
  }
  if (message.includes("--in と --in-base64 を同時に指定できません")) {
    return "multiple_input_sources";
  }
  if (message.includes("--out と --out-base64 を同時に指定できません")) {
    return "multiple_output_targets";
  }
  if (message.includes("binary stdin には --in-base64 -")) {
    return "binary_stdin_requires_base64";
  }
  if (message.includes("には --in <path> または --in-base64 - が必要です")) {
    return "missing_binary_input";
  }
  if (message.includes("Base64 入力")) {
    return "invalid_base64_input";
  }
  if (message.includes("--out-base64 は - のみ対応")) {
    return "invalid_base64_output_target";
  }
  if (message.includes("--out-base64 は binary output 専用")) {
    return "base64_output_requires_binary";
  }
  if (message.includes("binary artifact のため --out <path> が必要です")) {
    return "binary_stdout_not_supported";
  }
  if (message.includes("--select uid には --task-uid が必要です")) {
    return "missing_task_uid";
  }
  if (message.includes("--select uid には --phase-uid が必要です")) {
    return "missing_phase_uid";
  }
  if (message.includes("未対応の ai export コマンドです")) {
    return "unsupported_ai_export_command";
  }
  if (message.includes("未対応のコマンドです")) {
    return "unsupported_command";
  }
  if (context === "ai validate-patch" && !(error instanceof CliUsageError)) {
    return "patch_validation_failed";
  }
  return error instanceof CliUsageError ? "usage_error" : "processing_error";
}

function buildErrorItem(error, context) {
  const code = inferErrorCode(error, context);
  const item = {
    code,
    message: error instanceof Error ? error.message : String(error)
  };
  const details = extractErrorDetails(error);
  if (details) {
    item.details = details;
  }
  return item;
}

function extractErrorDetails(error) {
  if ((error instanceof CliUsageError || error instanceof CliProcessingError) && error.details && typeof error.details === "object") {
    return error.details;
  }
  return undefined;
}
