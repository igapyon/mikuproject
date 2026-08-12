import { CliUsageError } from "./cli-errors.mjs";
import {
  buildIoDiagnostics,
  ensureSingleStdinSource,
  readTextInput
} from "./cli-io.mjs";
import {
  buildChangedItems,
  buildCommandDiagnostics,
  buildPatchDiagnosticsJson,
  buildPatchDiagnosticsText,
  parseDiagnosticsFormat,
  summarizeChanges
} from "./cli-diagnostics.mjs";
import {
  ensureKind,
  ensureWorkbookJson,
  loadWorkbookModel,
  parseJsonLike,
  parsePlainJson
} from "./cli-command-utils.mjs";

export async function runStateCommand(command, options, api) {
  const [scope, action, subject] = command;
  if (scope !== "state") {
    return undefined;
  }

  if (action === "from-draft" && subject === undefined) {
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

  if (action === "summarize" && subject === undefined) {
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

  if (action === "diff" && subject === undefined) {
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

  if (action === "apply-patch" && subject === undefined) {
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

  return undefined;
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
