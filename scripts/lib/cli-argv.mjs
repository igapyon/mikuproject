import { CliUsageError } from "./cli-errors.mjs";

export function parseArgs(argv) {
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

export function detectRequestedDiagnosticsFormat(argv) {
  for (let index = 0; index < argv.length; index += 1) {
    if (argv[index] === "--diagnostics" && argv[index + 1] === "json") {
      return "json";
    }
  }
  return "text";
}

export function summarizeCommandFromArgv(argv) {
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
