import {
  CliProcessingError,
  CliUsageError,
  extractErrorDetails,
  inferErrorCode
} from "./cli-errors.mjs";
import {
  buildIoDiagnostics,
  ensureSingleStdinSource,
  readTextInput
} from "./cli-io.mjs";
import {
  buildCommandDiagnostics,
  buildErrorItem,
  determineStatus,
  DIAGNOSTICS_VERSION,
  formatValidationOutput,
  parseDiagnosticsFormat,
  summarizeChanges
} from "./cli-diagnostics.mjs";
import {
  ensureKind,
  ensureWorkbookJson,
  loadWorkbookModel,
  parseJsonLike,
  parseOptionalNonNegativeInteger,
  parsePhaseDetailMode,
  parsePlainJson,
  resolvePhaseDetailUid,
  resolveTaskEditUid
} from "./cli-command-utils.mjs";

export async function runAiCommand(command, options, api) {
  const [scope, action, subject, detail] = command;
  if (scope !== "ai") {
    return undefined;
  }

  if (action === "spec" && subject === undefined) {
    return {
      output: `${api.getAiJsonSpecText().trim()}\n`,
      diagnostics: []
    };
  }

  if (action === "detect-kind" && subject === undefined) {
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

  if (action === "export" && detail === undefined) {
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

  if (action === "validate-patch" && subject === undefined) {
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

  return undefined;
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
