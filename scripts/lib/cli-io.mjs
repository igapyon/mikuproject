import fs from "node:fs";
import path from "node:path";

import { summarizeCommandFromArgv } from "./cli-argv.mjs";
import { CliUsageError } from "./cli-errors.mjs";

export async function readTextInput(inputPath) {
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

export async function readStdinBytes() {
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

export function decodeBase64Input(sourceText, context) {
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

export async function readBinaryInput(options, context) {
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

export function ensureSingleStdinSource(inputs) {
  const stdinOptions = inputs
    .filter((input) => input.value === "-" || (input.allowImplicitStdin && input.value === undefined))
    .map((input) => input.optionName);

  if (stdinOptions.length > 1) {
    throw new CliUsageError(`標準入力を使える入力オプションは 1 つだけです: ${stdinOptions.join(", ")}`, "multiple_stdin_sources", {
      options: stdinOptions
    });
  }
}

export function ensureBinaryInputSource(options, commandLabel) {
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

export function ensureBinaryOutputTarget(options, commandLabel) {
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

export function writeOutput(output, options) {
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

export function buildIoDiagnostics({ inputs, output, outputBase64 }) {
  return {
    inputs: (inputs || []).map((input) => describeInputSource(input)),
    output: describeOutputTarget(output, outputBase64)
  };
}

export function describeInputSource(input) {
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

export function describeBinaryInputForDiagnostics(options) {
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

export function describeOutputTarget(outPath, outBase64Path) {
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

export function buildIoDiagnosticsFromArgv(argv) {
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

export function inferImplicitInputOptions(command, existingInputs) {
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
