import path from "node:path";

import { isV1WorkflowCommand, parseV1Invocation } from "./cli-v1-argv.mjs";
import { runV1ApplyChange } from "./cli-v1-apply.mjs";
import { createV1RuntimeError, isCliV1Error } from "./cli-v1-errors.mjs";
import { reserveV1ResultTransport } from "./cli-v1-io.mjs";
import { runV1Inspect, runV1PlanChange, runV1Validate } from "./cli-v1-r1-commands.mjs";
import { createUnverifiedRuntimeBinding, createV1ErrorResult, serializeV1Result } from "./cli-v1-result.mjs";

/**
 * Detects only the explicit v1 command words. Legacy routing remains the
 * public entrypoint until ZB-P4.4.6 wires this boundary into the CLI.
 */
export function recognizesV1Workflow(argv) {
  return Array.isArray(argv) && isV1WorkflowCommand(argv[0]);
}

export async function prepareV1WorkflowInvocation(argv, options = {}) {
  const invocation = parseV1Invocation(argv);
  if (invocation.kind !== "workflow") {
    return { invocation, resultTransport: null };
  }
  const resultTransport = await reserveV1ResultTransport(invocation.options.result, options);
  return { invocation, resultTransport };
}

/**
 * Executes exactly the implemented R1/C1 surface with an explicit, caller-owned
 * runtime binding.  The public product entrypoint must not call this until a
 * complete, manifest-verified core runtime exists; this boundary is for the
 * fixed-binding conformance harness only.
 */
export async function runV1R1Harness(argv, {
  runtime,
  cwd = process.cwd(),
  stdin = process.stdin,
  stdout = process.stdout,
  fileSystem
} = {}) {
  let invocation;
  try {
    invocation = parseV1Invocation(argv);
  } catch (error) {
    if (!isCliV1Error(error)) {
      throw error;
    }
    const result = createV1ErrorResult({ error, command: "cli", runtime });
    stdout.write(serializeV1Result(result));
    return result;
  }
  if (invocation.kind !== "workflow") {
    throw new TypeError("The fixed v1 service harness accepts only a v1 workflow invocation");
  }
  assertImplementedServiceInvocation(invocation);

  let resultTransport;
  try {
    resultTransport = await reserveV1ResultTransport(invocation.options.result, { cwd, stdout, fileSystem });
  } catch (error) {
    if (!isCliV1Error(error)) {
      throw error;
    }
    const result = createV1ErrorResult({
      error,
      command: invocation.command,
      runtime,
      io: invocation.command === "plan-change"
        ? unreadPlanChangeIo(invocation, cwd)
        : invocation.command === "apply-change"
          ? unreadApplyChangeIo(invocation, cwd)
          : unreadR1ProjectIo(invocation, cwd),
      data: invocation.command === "validate" && error.status === "rejected"
        ? { validation: { valid: false, format_profile: null, state_digest: null } }
        : null
    });
    stdout.write(serializeV1Result(result));
    return result;
  }
  try {
    if (invocation.command === "validate") {
      return await runV1Validate({ invocation, resultTransport, runtime, cwd, stdin, fileSystem });
    }
    if (invocation.command === "inspect") {
      return await runV1Inspect({ invocation, resultTransport, runtime, cwd, stdin, fileSystem });
    }
    if (invocation.command === "plan-change") {
      return await runV1PlanChange({ invocation, resultTransport, runtime, cwd, stdin, fileSystem });
    }
    if (invocation.command === "apply-change") {
      return await runV1ApplyChange({ invocation, resultTransport, runtime, cwd, stdin, fileSystem });
    }
    throw new TypeError(`The fixed v1 service harness does not implement ${invocation.command} ${invocation.options.purpose ?? ""}`.trim());
  } catch (error) {
    await resultTransport.abort();
    throw error;
  }
}

/**
 * The repository source CLI and its development bundle deliberately expose a
 * fail-closed v1 boundary.  It must identify v1 before legacy routing but may
 * not read a project or reserve a result file while the nine-capability
 * runtime manifest does not exist.  P4.9 replaces this gate with verified
 * manifest/asset binding before any R1/C1 service is enabled publicly.
 */
export function rejectUnreleasedV1Workflow(argv, { version, stdout = process.stdout } = {}) {
  if (!recognizesV1Workflow(argv)) {
    throw new TypeError("rejectUnreleasedV1Workflow requires a recognized v1 workflow command");
  }
  const error = createV1RuntimeError({
    code: "runtime.capability-missing",
    message: "This development CLI does not provide a complete manifest-verified miku-project-cli-core/v1 runtime.",
    scope: "runtime",
    details: {
      required_capability_profile: "miku-project-cli-core/v1",
      implementation_state: "R1-service-only",
      requested_command: argv[0]
    }
  });
  const result = createV1ErrorResult({
    error,
    // No v1 invocation grammar or input role is trusted before a complete
    // runtime binding exists.  `cli` therefore records this pre-dispatch
    // runtime gate, while details.requested_command preserves the recognized
    // command word without inventing an unread project input.
    command: "cli",
    runtime: createUnverifiedRuntimeBinding({ version })
  });
  stdout.write(serializeV1Result(result));
  return result;
}

function assertImplementedServiceInvocation(invocation) {
  if (invocation.command === "validate") {
    return;
  }
  if (invocation.command === "inspect") {
    return;
  }
  if (invocation.command === "plan-change") {
    return;
  }
  if (invocation.command === "apply-change") {
    return;
  }
  throw new TypeError(`The fixed v1 service harness does not implement ${invocation.command} ${invocation.options.purpose ?? ""}`.trim());
}

function unreadR1ProjectIo(invocation, cwd) {
  const source = invocation.options.project === "-" ? "stdin" : "file";
  return {
    stdin_option: source === "stdin" ? "--project" : null,
    inputs: [{
      role: "project",
      option: "--project",
      source,
      path: source === "stdin" ? null : resolveUnopenedPath(cwd, invocation.options.project),
      digest: null
    }],
    result: { target: "stdout", path: null },
    destination: null
  };
}

function unreadPlanChangeIo(invocation, cwd) {
  const projectSource = invocation.options.project === "-" ? "stdin" : "file";
  const requestSource = invocation.options.request === "-" ? "stdin" : "file";
  return {
    stdin_option: projectSource === "stdin" ? "--project" : requestSource === "stdin" ? "--request" : null,
    inputs: [{
      role: "project",
      option: "--project",
      source: projectSource,
      path: projectSource === "stdin" ? null : resolveUnopenedPath(cwd, invocation.options.project),
      digest: null
    }, {
      role: "change_request",
      option: "--request",
      source: requestSource,
      path: requestSource === "stdin" ? null : resolveUnopenedPath(cwd, invocation.options.request),
      digest: null
    }],
    result: { target: "stdout", path: null },
    destination: {
      requested_path: invocation.options.destination,
      path: resolveUnopenedPath(cwd, invocation.options.destination)
    }
  };
}

function unreadApplyChangeIo(invocation, cwd) {
  const projectSource = invocation.options.project === "-" ? "stdin" : "file";
  const requestSource = invocation.options.request === "-" ? "stdin" : "file";
  const planResultSource = invocation.options["plan-result"] === "-" ? "stdin" : "file";
  const approvalSource = invocation.options.approval === "-" ? "stdin" : "file";
  const inputSources = [
    ["project", "--project", projectSource, invocation.options.project],
    ["change_request", "--request", requestSource, invocation.options.request],
    ["plan_result", "--plan-result", planResultSource, invocation.options["plan-result"]],
    ["approval", "--approval", approvalSource, invocation.options.approval]
  ];
  const stdinInput = inputSources.find(([, , source]) => source === "stdin");
  return {
    stdin_option: stdinInput ? stdinInput[1] : null,
    inputs: inputSources.map(([role, option, source, requestedPath]) => ({
      role,
      option,
      source,
      path: source === "stdin" ? null : resolveUnopenedPath(cwd, requestedPath),
      digest: null
    })),
    result: { target: "stdout", path: null },
    // apply-change takes its canonical destination only from a successful
    // plan result. Before result-channel reservation that artifact remains
    // unread, so no destination path can honestly be reported yet.
    destination: null
  };
}

function resolveUnopenedPath(cwd, requestedPath) {
  // This preflight metadata records the unresolved absolute candidate without
  // lstat/opening the project.  The R1 service replaces it with a realpath
  // and raw digest only after result-channel reservation succeeds.
  return path.resolve(cwd, requestedPath);
}
