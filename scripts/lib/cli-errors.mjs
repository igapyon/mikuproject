export class CliUsageError extends Error {
  constructor(message, code = "usage_error", details = undefined) {
    super(message);
    this.name = "CliUsageError";
    this.code = code;
    this.details = details;
  }
}

export class CliProcessingError extends Error {
  constructor(message, code = "processing_error", details = undefined) {
    super(message);
    this.name = "CliProcessingError";
    this.code = code;
    this.details = details;
  }
}

export function inferErrorCode(error, context) {
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

export function extractErrorDetails(error) {
  if ((error instanceof CliUsageError || error instanceof CliProcessingError) && error.details && typeof error.details === "object") {
    return error.details;
  }
  return undefined;
}
