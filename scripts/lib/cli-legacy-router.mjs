import { CliUsageError } from "./cli-errors.mjs";
import { runAiCommand } from "./cli-ai-commands.mjs";
import { runExchangeCommand } from "./cli-exchange-commands.mjs";
import { runReportCommand } from "./cli-report-commands.mjs";
import { runStateCommand } from "./cli-state-commands.mjs";

const LEGACY_COMMAND_SERVICES = [
  runAiCommand,
  runStateCommand,
  runExchangeCommand,
  runReportCommand
];

export async function runLegacyCommand(command, options, api) {
  for (const runService of LEGACY_COMMAND_SERVICES) {
    const result = await runService(command, options, api);
    if (result !== undefined) {
      return result;
    }
  }
  throw new CliUsageError(`未対応のコマンドです: ${command.join(" ")}`, "unsupported_command");
}
