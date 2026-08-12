import { CliProcessingError, CliUsageError } from "./cli-errors.mjs";
import { readTextInput } from "./cli-io.mjs";

export function parsePlainJson(sourceText, context = "input") {
  try {
    return JSON.parse(sourceText);
  } catch (_error) {
    throw new CliProcessingError(`${context} の JSON を解析できませんでした`, "invalid_json_input", {
      context
    });
  }
}

export function parseJsonLike(sourceText, api, context = "input") {
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

export async function loadWorkbookModel(api, inputPath, context) {
  const stateDocument = parsePlainJson(await readTextInput(inputPath), context);
  ensureWorkbookJson(api, stateDocument, context);
  return api.workbookJson.importAsProjectModel(stateDocument).model;
}

export function parseOptionalNonNegativeInteger(value, optionName) {
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

export function parsePhaseDetailMode(value) {
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

export function parseSelectMode(value) {
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

export function resolveTaskEditUid(model, options) {
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

export function resolvePhaseDetailUid(model, options) {
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

export function ensureWorkbookJson(api, documentLike, context) {
  const kind = api.detectAiJsonDocumentKind(documentLike);
  if (kind !== "workbook_json") {
    throw new CliProcessingError(`${context} は mikuproject_workbook_json を入力してください`, "invalid_workbook_kind", {
      context,
      expected_kind: "workbook_json",
      actual_kind: kind || null
    });
  }
}

export function ensureKind(api, documentLike, expectedKind, context) {
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
