import { createV1UsageError } from "./cli-v1-errors.mjs";

export const V1_WORKFLOW_COMMANDS = Object.freeze([
  "inspect",
  "validate",
  "plan-change",
  "apply-change",
  "verify-artifact"
]);

const COMMAND_SPECS = Object.freeze({
  inspect: Object.freeze({
    sideEffectClass: "read-only",
    required: Object.freeze(["project", "purpose"]),
    allowed: Object.freeze(["project", "purpose", "task-uid", "result"]),
    stdinOptions: Object.freeze(["project"])
  }),
  validate: Object.freeze({
    sideEffectClass: "read-only",
    required: Object.freeze(["project"]),
    allowed: Object.freeze(["project", "result"]),
    stdinOptions: Object.freeze(["project"])
  }),
  "plan-change": Object.freeze({
    sideEffectClass: "exchange-artifact-generation",
    required: Object.freeze(["project", "request", "destination"]),
    allowed: Object.freeze(["project", "request", "destination", "result"]),
    stdinOptions: Object.freeze(["project", "request"])
  }),
  "apply-change": Object.freeze({
    sideEffectClass: "meaning-change-and-project-artifact-generation",
    required: Object.freeze(["project", "request", "plan-result", "approval"]),
    allowed: Object.freeze(["project", "request", "plan-result", "approval", "result"]),
    stdinOptions: Object.freeze(["project", "request", "plan-result", "approval"])
  }),
  "verify-artifact": Object.freeze({
    sideEffectClass: "read-only",
    required: Object.freeze(["artifact-set"]),
    allowed: Object.freeze(["artifact-set", "expect-plan-result", "result"]),
    stdinOptions: Object.freeze(["expect-plan-result"])
  })
});

export function isV1WorkflowCommand(value) {
  return typeof value === "string" && V1_WORKFLOW_COMMANDS.includes(value);
}

export function isV1ControlInvocation(argv) {
  return Array.isArray(argv)
    && ((argv.length === 1 && (argv[0] === "--help" || argv[0] === "--version"))
      || (argv.length === 2 && isV1WorkflowCommand(argv[0]) && argv[1] === "--help"));
}

export function parseV1Invocation(argv) {
  if (!Array.isArray(argv)) {
    throw new TypeError("v1 argv must be an array");
  }
  if (argv.length === 1 && argv[0] === "--help") {
    return Object.freeze({ kind: "control", control: "help", command: null });
  }
  if (argv.length === 1 && argv[0] === "--version") {
    return Object.freeze({ kind: "control", control: "version", command: null });
  }

  const command = argv[0];
  if (typeof command === "string" && command.startsWith("--") && command !== "--") {
    throw createV1UsageError({
      code: "cli.unknown-option",
      message: `Unsupported v1 option: ${command}`,
      option: command,
      details: { received: command }
    });
  }
  if (!isV1WorkflowCommand(command)) {
    throw createV1UsageError({
      code: "cli.unknown-command",
      message: command === undefined ? "A v1 command is required." : `Unsupported v1 command: ${String(command)}`,
      details: { received: command ?? null },
      command: command ?? null
    });
  }
  if (argv.length === 2 && argv[1] === "--help") {
    return Object.freeze({ kind: "control", control: "command-help", command });
  }
  if (argv.slice(1).includes("--help") || argv.slice(1).includes("--version")) {
    throw createV1UsageError({
      code: "cli.unexpected-argument",
      message: "--help and --version cannot be combined with workflow options.",
      details: { received: argv.slice(1) },
      command
    });
  }

  const spec = COMMAND_SPECS[command];
  const options = parseV1Options(argv.slice(1), command, spec);
  validateCommandOptions(command, options, spec);
  return Object.freeze({
    kind: "workflow",
    command,
    sideEffectClass: spec.sideEffectClass,
    options: Object.freeze({ ...options, result: options.result ?? "-" })
  });
}

export function getV1CommandSpec(command) {
  return COMMAND_SPECS[command] || null;
}

function parseV1Options(tokens, command, spec) {
  const options = {};
  for (let index = 0; index < tokens.length; index += 1) {
    const token = tokens[index];
    if (typeof token !== "string" || !token.startsWith("--") || token === "--") {
      throw createV1UsageError({
        code: "cli.unexpected-argument",
        message: `Unexpected positional argument: ${String(token)}`,
        details: { received: token ?? null },
        command
      });
    }
    if (token.includes("=")) {
      throw createV1UsageError({
        code: "cli.unknown-option",
        message: `The --key=value form is not supported: ${token}`,
        option: token,
        details: { received: token },
        command
      });
    }
    const optionName = token.slice(2);
    if (!spec.allowed.includes(optionName)) {
      throw createV1UsageError({
        code: "cli.unknown-option",
        message: `Unsupported option for ${command}: ${token}`,
        option: token,
        details: { command, received: token },
        command
      });
    }
    if (Object.hasOwn(options, optionName)) {
      throw createV1UsageError({
        code: "cli.duplicate-option",
        message: `Option ${token} must not be repeated.`,
        option: token,
        details: { command, option: token },
        command
      });
    }
    const value = tokens[index + 1];
    if (value === undefined || (typeof value === "string" && value.startsWith("--"))) {
      throw createV1UsageError({
        code: "cli.missing-option",
        message: `Option ${token} requires a value.`,
        option: token,
        details: { command, option: token },
        command
      });
    }
    assertOptionValue(optionName, value, command);
    options[optionName] = value;
    index += 1;
  }
  return options;
}

function validateCommandOptions(command, options, spec) {
  for (const optionName of spec.required) {
    if (!Object.hasOwn(options, optionName)) {
      throw createV1UsageError({
        code: "cli.missing-option",
        message: `Command ${command} requires --${optionName}.`,
        option: `--${optionName}`,
        details: { command, option: `--${optionName}` },
        command
      });
    }
  }

  const stdinOptions = spec.stdinOptions.filter((optionName) => options[optionName] === "-");
  if (stdinOptions.length > 1) {
    throw createV1UsageError({
      code: "cli.multiple-stdin-sources",
      message: "Only one explicit v1 input option may use stdin.",
      details: { command, options: stdinOptions.map((optionName) => `--${optionName}`) },
      command
    });
  }
  if (options["artifact-set"] === "-") {
    throw createV1UsageError({
      code: "cli.invalid-option-value",
      message: "--artifact-set does not accept stdin.",
      option: "--artifact-set",
      details: { command, option: "--artifact-set", received: "-" },
      command
    });
  }

  if (command === "inspect") {
    if (options.purpose === "project_overview" && Object.hasOwn(options, "task-uid")) {
      throw createV1UsageError({
        code: "cli.invalid-option-value",
        message: "--task-uid is not allowed with --purpose project_overview.",
        option: "--task-uid",
        details: { command, purpose: options.purpose },
        command
      });
    }
    if (options.purpose === "task_change_context" && !Object.hasOwn(options, "task-uid")) {
      throw createV1UsageError({
        code: "cli.missing-option",
        message: "--purpose task_change_context requires --task-uid.",
        option: "--task-uid",
        details: { command, purpose: options.purpose },
        command
      });
    }
  }
}

function assertOptionValue(optionName, value, command) {
  if (typeof value !== "string" || value.length === 0 || value.includes("\0")) {
    throw createV1UsageError({
      code: "cli.invalid-option-value",
      message: `Option --${optionName} has an invalid value.`,
      option: `--${optionName}`,
      details: { command, option: `--${optionName}`, received: typeof value === "string" ? value : null },
      command
    });
  }
  if (optionName === "purpose" && value !== "project_overview" && value !== "task_change_context") {
    throw createV1UsageError({
      code: "cli.invalid-option-value",
      message: "--purpose must be project_overview or task_change_context.",
      option: "--purpose",
      details: { command, option: "--purpose", received: value },
      command
    });
  }
}
