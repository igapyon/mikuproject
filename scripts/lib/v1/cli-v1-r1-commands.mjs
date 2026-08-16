import fsPromises from "node:fs/promises";
import path from "node:path";

import { sha256RawBytes, sha256SemanticState } from "./cli-v1-canonical-json.mjs";
import { readV1CommittedArtifactSetProject } from "./cli-v1-artifact-verifier.mjs";
import { planV1SetTaskPercentComplete } from "./cli-v1-change.mjs";
import { preflightV1NewDestination } from "./cli-v1-destination.mjs";
import { createV1IoError, createV1RejectedError, createV1RuntimeError, isCliV1Error } from "./cli-v1-errors.mjs";
import { readV1JsonArtifact } from "./cli-v1-json-artifact.mjs";
import {
  createV1ProjectOverviewProjection,
  createV1TaskChangeContextProjection
} from "./cli-v1-projection.mjs";
import { createV1DiagnosticFromError, createV1ErrorResult, createV1Result } from "./cli-v1-result.mjs";
import { semanticIssuesToV1Errors, validateV1SemanticState } from "./cli-v1-semantic-validator.mjs";
import { decodeMsProjectXmlSubset } from "./cli-v1-xml-adapter.mjs";

/**
 * Executes the R1 validate service after strict argv parsing and result
 * transport reservation.  This module deliberately has no process.argv or
 * legacy-router dependency; ZB-P4.4.6 owns public entrypoint wiring.
 */
export async function runV1Validate({ invocation, resultTransport, runtime, cwd = process.cwd(), stdin = process.stdin, fileSystem = fsPromises } = {}) {
  assertValidateInvocation(invocation);
  assertResultTransport(resultTransport);

  const prepared = await prepareV1ExternalProjectInput(invocation.options.project, { cwd, stdin, fileSystem });
  if (prepared.error) {
    const result = createV1ErrorResult({
      error: prepared.error,
      command: "validate",
      runtime,
      io: readOnlyProjectIo(prepared.input, resultTransport.target),
      data: prepared.error.status === "rejected" ? rejectedValidationData(null) : null
    });
    await resultTransport.writeResult(result);
    return result;
  }

  const { input, decoded, validation } = prepared;
  let result;
  if (validation.valid) {
    result = createV1Result({
      command: "validate",
      runtime,
      status: "succeeded",
      io: readOnlyProjectIo(input, resultTransport.target),
      observations: {
        normalizations: decoded.normalizations,
        losses: [],
        unsupported: []
      },
      data: {
        validation: {
          valid: true,
          format_profile: decoded.format_profile,
          state_digest: sha256SemanticState(decoded.state)
        }
      }
    });
  } else {
    const errors = semanticIssuesToV1Errors(validation);
    result = createV1Result({
      command: "validate",
      runtime,
      status: "rejected",
      io: readOnlyProjectIo(input, resultTransport.target),
      diagnostics: errors.map(createV1DiagnosticFromError),
      observations: {
        normalizations: decoded.normalizations,
        losses: [],
        unsupported: validation.issues
          .filter((issue) => issue.code === "semantic.unsupported")
          .map((issue) => ({ code: issue.code, path: issue.path, description: issue.message }))
      },
      data: rejectedValidationData(decoded.format_profile)
    });
  }
  await resultTransport.writeResult(result);
  return result;
}

/**
 * Executes the R1 `inspect --purpose project_overview` service.  It calls the
 * same prepared-input helper as validate: no alternate inspect-only XML or
 * semantic path is permitted.
 */
export async function runV1Inspect({ invocation, resultTransport, runtime, cwd = process.cwd(), stdin = process.stdin, fileSystem = fsPromises } = {}) {
  assertInspectInvocation(invocation);
  assertResultTransport(resultTransport);

  const prepared = await prepareV1ExternalProjectInput(invocation.options.project, { cwd, stdin, fileSystem });
  if (prepared.error) {
    const result = createV1ErrorResult({
      error: prepared.error,
      command: "inspect",
      runtime,
      io: readOnlyProjectIo(prepared.input, resultTransport.target),
      data: null
    });
    await resultTransport.writeResult(result);
    return result;
  }

  const { input, decoded, validation } = prepared;
  let result;
  if (validation.valid) {
    try {
      const projection = createV1InspectProjection(decoded.state, invocation);
      result = createV1Result({
        command: "inspect",
        runtime,
        status: "succeeded",
        io: readOnlyProjectIo(input, resultTransport.target),
        observations: {
          normalizations: decoded.normalizations,
          losses: [],
          unsupported: []
        },
        data: { projection }
      });
    } catch (error) {
      result = createV1ErrorResult({
        error: toUnexpectedV1RuntimeError(error),
        command: "inspect",
        runtime,
        io: readOnlyProjectIo(input, resultTransport.target),
        data: null
      });
    }
  } else {
    const errors = semanticIssuesToV1Errors(validation);
    result = createV1Result({
      command: "inspect",
      runtime,
      status: "rejected",
      io: readOnlyProjectIo(input, resultTransport.target),
      diagnostics: errors.map(createV1DiagnosticFromError),
      observations: {
        normalizations: decoded.normalizations,
        losses: [],
        unsupported: validation.issues
          .filter((issue) => issue.code === "semantic.unsupported")
          .map((issue) => ({ code: issue.code, path: issue.path, description: issue.message }))
      },
      data: null
    });
  }
  await resultTransport.writeResult(result);
  return result;
}

/**
 * Executes the C1 human-gate boundary.  It performs no publication: the
 * returned semantic diff/output plan is the only success payload, while the
 * planned state and preflight XML remain internal to this invocation.
 */
export async function runV1PlanChange({ invocation, resultTransport, runtime, cwd = process.cwd(), stdin = process.stdin, fileSystem = fsPromises } = {}) {
  assertPlanChangeInvocation(invocation);
  assertResultTransport(resultTransport);
  let projectInput = unreadProjectInput(invocation.options.project, cwd);
  let requestInput = unreadJsonInput(invocation.options.request, "change_request", "--request", cwd);
  let destination = unresolvedDestination(invocation.options.destination, cwd);

  const prepared = await prepareV1ExternalProjectInput(invocation.options.project, { cwd, stdin, fileSystem });
  projectInput = prepared.input ?? projectInput;
  if (prepared.error) {
    const result = createV1ErrorResult({
      error: prepared.error,
      command: "plan-change",
      runtime,
      io: planChangeIo({ invocation, projectInput, requestInput, destination, resultTarget: resultTransport.target }),
      data: prepared.error.status === "rejected" ? rejectedValidationData(null) : null
    });
    await resultTransport.writeResult(result);
    return result;
  }

  const { decoded, validation } = prepared;
  if (!validation.valid) {
    const errors = semanticIssuesToV1Errors(validation);
    const result = createV1Result({
      command: "plan-change",
      runtime,
      status: "rejected",
      io: planChangeIo({ invocation, projectInput, requestInput, destination, resultTarget: resultTransport.target }),
      diagnostics: errors.map(createV1DiagnosticFromError),
      observations: semanticObservations(decoded, validation),
      data: rejectedValidationData(decoded.format_profile)
    });
    await resultTransport.writeResult(result);
    return result;
  }

  const requestRead = await readV1JsonArtifact(invocation.options.request, {
    role: "change_request",
    option: "--request",
    cwd,
    stdin,
    fileSystem
  });
  requestInput = requestRead.input ?? requestInput;
  if (requestRead.error) {
    const result = createV1ErrorResult({
      error: requestRead.error,
      command: "plan-change",
      runtime,
      io: planChangeIo({ invocation, projectInput, requestInput, destination, resultTarget: resultTransport.target }),
      observations: { normalizations: decoded.normalizations, losses: [], unsupported: [] },
      data: null
    });
    await resultTransport.writeResult(result);
    return result;
  }

  try {
    // Validate request semantics and complete an encode/redecode dry run before
    // reporting a destination problem. This preserves the C1 ordering: an
    // invalid request never becomes conditionally valid because a path happens
    // to be unsafe. The provisional path is internal only; no result is made
    // until the canonical destination preflight below succeeds.
    planV1SetTaskPercentComplete({
      state: decoded.state,
      changeRequest: requestRead.value,
      runtime,
      destination
    });
    destination = await preflightV1NewDestination(invocation.options.destination, {
      cwd,
      fileSystem,
      projectInput
    });
    const plan = planV1SetTaskPercentComplete({
      state: decoded.state,
      changeRequest: requestRead.value,
      runtime,
      destination
    });
    const result = createV1Result({
      command: "plan-change",
      runtime,
      status: "succeeded",
      io: planChangeIo({ invocation, projectInput, requestInput, destination, resultTarget: resultTransport.target }),
      observations: {
        // Input transport normalizations (for example a permitted XML BOM)
        // remain observations of this invocation. The output plan itself
        // records only the encode/preflight normalizations the human must
        // approve for the prospective artifact.
        normalizations: [...decoded.normalizations, ...plan.output_plan.preflight.normalizations],
        losses: [],
        unsupported: []
      },
      data: {
        semantic_diff: plan.semantic_diff,
        output_plan: plan.output_plan
      }
    });
    await resultTransport.writeResult(result);
    return result;
  } catch (error) {
    const result = createV1ErrorResult({
      error: toUnexpectedV1RuntimeError(error),
      command: "plan-change",
      runtime,
      io: planChangeIo({ invocation, projectInput, requestInput, destination, resultTarget: resultTransport.target }),
      observations: { normalizations: decoded.normalizations, losses: [], unsupported: [] },
      data: null
    });
    await resultTransport.writeResult(result);
    return result;
  }
}

/**
 * The sole R1 preparation pipeline for external XML direct input.  Both
 * validate and project_overview inspection call it so a result cannot be
 * projected from a state that validate would interpret differently.
 */
export async function prepareV1ExternalProjectInput(projectOption, { cwd, stdin, fileSystem }) {
  const inputRead = await readV1ExternalProjectInput(projectOption, { cwd, stdin, fileSystem });
  if (inputRead.error) {
    return inputRead;
  }
  let decoded;
  try {
    decoded = decodeMsProjectXmlSubset(inputRead.bytes);
  } catch (error) {
    return { input: inputRead.input, error: toUnexpectedV1RuntimeError(error) };
  }
  try {
    return {
      input: inputRead.input,
      decoded,
      validation: validateV1SemanticState(decoded.state, { adapterIssues: decoded.adapter_issues })
    };
  } catch (error) {
    return { input: inputRead.input, error: toUnexpectedV1RuntimeError(error) };
  }
}

/**
 * Reads an external XML direct entry or, for a directory, only a project.xml
 * member that the artifact verifier has confirmed belongs to a committed set.
 * Both sources return the same raw-byte input to the sole semantic pipeline.
 */
export async function readV1ExternalProjectInput(projectOption, { cwd = process.cwd(), stdin = process.stdin, fileSystem = fsPromises } = {}) {
  if (projectOption === "-") {
    const input = projectInputMetadata({ source: "stdin", path: null, digest: null });
    try {
      const bytes = await readV1ProjectStdinBytes(stdin);
      input.digest = sha256RawBytes(bytes);
      return { input, bytes };
    } catch (error) {
      return {
        input,
        error: createV1IoError({
          code: "io.input-read-failed",
          status: "runtime-error",
          message: "The v1 project input could not be read from stdin.",
          path: null,
          option: "--project",
          details: { error_code: error?.code ?? null }
        })
      };
    }
  }

  if (typeof projectOption !== "string" || projectOption.length === 0 || projectOption.includes("\0")) {
    throw new TypeError("v1 R1 project input requires a parsed --project path or stdin marker");
  }
  const candidatePath = path.resolve(cwd, projectOption);
  let entry;
  try {
    entry = await fileSystem.lstat(candidatePath);
  } catch (error) {
    const input = projectInputMetadata({ source: "file", path: candidatePath, digest: null });
    if (error?.code === "ENOENT") {
      return {
        input,
        error: createV1IoError({
          code: "io.input-not-found",
          status: "rejected",
          message: "The --project input file does not exist.",
          path: candidatePath,
          option: "--project",
          details: { requested_path: projectOption }
        })
      };
    }
    return {
      input,
      error: createV1IoError({
        code: "io.input-read-failed",
        status: "runtime-error",
        message: "The --project input could not be inspected.",
        path: candidatePath,
        option: "--project",
        details: { requested_path: projectOption, error_code: error?.code ?? null }
      })
    };
  }
  if (entry.isSymbolicLink()) {
    const input = projectInputMetadata({ source: "file", path: candidatePath, digest: null });
    return {
      input,
      error: createV1IoError({
        code: "io.input-symlink-rejected",
        status: "rejected",
        message: "A direct --project input must not be a symbolic link.",
        path: candidatePath,
        option: "--project",
        details: { requested_path: projectOption }
      })
    };
  }
  if (entry.isDirectory()) {
    return readV1CommittedArtifactSetProject(projectOption, { cwd, fileSystem });
  }
  if (!entry.isFile()) {
    const source = "file";
    let canonicalPath = candidatePath;
    try {
      canonicalPath = await fileSystem.realpath(candidatePath);
    } catch {
      // The original absolute path is still a safe diagnostic location; its
      // type has already been observed without following a direct symlink.
    }
    const input = projectInputMetadata({ source, path: canonicalPath, digest: null });
    return {
      input,
      error: createV1IoError({
        code: "io.input-type-invalid",
        status: "rejected",
        message: "A v1 project input must be an external XML regular file, a committed artifact-set directory, or stdin.",
        path: canonicalPath,
        option: "--project",
        details: { requested_path: projectOption, observed_type: "other" }
      })
    };
  }

  let canonicalPath;
  try {
    canonicalPath = await fileSystem.realpath(candidatePath);
  } catch (error) {
    const input = projectInputMetadata({ source: "file", path: candidatePath, digest: null });
    return {
      input,
      error: createV1IoError({
        code: "io.input-read-failed",
        status: "runtime-error",
        message: "The --project input file could not be canonicalized.",
        path: candidatePath,
        option: "--project",
        details: { requested_path: projectOption, error_code: error?.code ?? null }
      })
    };
  }
  const input = projectInputMetadata({ source: "file", path: canonicalPath, digest: null });
  try {
    const bytes = asV1ProjectRawBytes(await fileSystem.readFile(canonicalPath));
    input.digest = sha256RawBytes(bytes);
    return { input, bytes };
  } catch (error) {
    return {
      input,
      error: createV1IoError({
        code: "io.input-read-failed",
        status: "runtime-error",
        message: "The --project input file could not be read.",
        path: canonicalPath,
        option: "--project",
        details: { requested_path: projectOption, error_code: error?.code ?? null }
      })
    };
  }
}

function toUnexpectedV1RuntimeError(error) {
  if (isCliV1Error(error)) {
    return error;
  }
  return createV1RuntimeError({
    message: "The v1 validation service encountered an unexpected internal error.",
    option: "--project",
    artifactRole: "external_project",
    details: {
      error_name: error instanceof Error ? error.name : typeof error
    }
  });
}

function planChangeIo({ invocation, projectInput, requestInput, destination, resultTarget }) {
  return {
    stdin_option: invocation.options.project === "-" ? "--project" : invocation.options.request === "-" ? "--request" : null,
    inputs: [copyInput(projectInput), copyInput(requestInput)],
    result: { target: resultTarget.target, path: resultTarget.path },
    destination: { requested_path: invocation.options.destination, path: destination.path }
  };
}

function semanticObservations(decoded, validation) {
  return {
    normalizations: decoded.normalizations,
    losses: [],
    unsupported: validation.issues
      .filter((issue) => issue.code === "semantic.unsupported")
      .map((issue) => ({ code: issue.code, path: issue.path, description: issue.message }))
  };
}

function unreadProjectInput(projectOption, cwd) {
  return projectInputMetadata({
    source: projectOption === "-" ? "stdin" : "file",
    path: projectOption === "-" ? null : path.resolve(cwd, projectOption),
    digest: null
  });
}

function unreadJsonInput(optionValue, role, option, cwd) {
  return {
    role,
    option,
    source: optionValue === "-" ? "stdin" : "file",
    path: optionValue === "-" ? null : path.resolve(cwd, optionValue),
    digest: null
  };
}

function unresolvedDestination(requestedPath, cwd) {
  return { requested_path: requestedPath, path: path.resolve(cwd, requestedPath) };
}

function copyInput(input) {
  return {
    role: input.role,
    option: input.option,
    source: input.source,
    path: input.path,
    digest: input.digest ? { ...input.digest } : null
  };
}

function readOnlyProjectIo(input, resultTarget) {
  return {
    stdin_option: input.source === "stdin" ? "--project" : null,
    inputs: [{
      role: input.role,
      option: input.option,
      source: input.source,
      path: input.path,
      digest: input.digest ? { ...input.digest } : null
    }],
    result: { target: resultTarget.target, path: resultTarget.path },
    destination: null
  };
}

function rejectedValidationData(formatProfile) {
  return {
    validation: {
      valid: false,
      format_profile: formatProfile,
      state_digest: null
    }
  };
}

function projectInputMetadata({ source, path: inputPath, digest }) {
  return {
    role: "project",
    option: "--project",
    source,
    path: inputPath,
    digest
  };
}

async function readV1ProjectStdinBytes(stdin) {
  if (Buffer.isBuffer(stdin) || stdin instanceof Uint8Array) {
    return Buffer.from(stdin);
  }
  if (!stdin || typeof stdin[Symbol.asyncIterator] !== "function") {
    throw new TypeError("v1 stdin must be a Buffer, Uint8Array, or async byte iterable");
  }
  const chunks = [];
  for await (const chunk of stdin) {
    chunks.push(asV1ProjectRawBytes(chunk));
  }
  return Buffer.concat(chunks);
}

function asV1ProjectRawBytes(value) {
  if (Buffer.isBuffer(value) || value instanceof Uint8Array) {
    return Buffer.from(value);
  }
  throw new TypeError("v1 text input must provide raw bytes");
}

function assertValidateInvocation(invocation) {
  if (!invocation || invocation.kind !== "workflow" || invocation.command !== "validate") {
    throw new TypeError("runV1Validate requires a parsed v1 validate workflow invocation");
  }
}

function createV1InspectProjection(state, invocation) {
  if (invocation.options.purpose === "project_overview") {
    return createV1ProjectOverviewProjection(state);
  }
  const target = state.tasks.find((task) => task.uid === invocation.options["task-uid"]);
  if (!target) {
    throw createV1RejectedError({
      code: "change.request-invalid",
      message: "The task_change_context target task does not exist in the current semantic state.",
      scope: "semantic",
      path: `tasks[uid=${invocation.options["task-uid"]}]`,
      option: "--task-uid",
      artifactRole: "external_project",
      details: { target_task_uid: invocation.options["task-uid"] }
    });
  }
  if (target.summary !== false) {
    throw createV1RejectedError({
      code: "change.operation-unsupported",
      message: "task_change_context is available only for a leaf task.",
      scope: "semantic",
      path: `tasks[uid=${target.uid}].summary`,
      option: "--task-uid",
      artifactRole: "external_project",
      details: { target_task_uid: target.uid, summary: target.summary }
    });
  }
  return createV1TaskChangeContextProjection(state, target.uid);
}

function assertInspectInvocation(invocation) {
  if (!invocation
    || invocation.kind !== "workflow"
    || invocation.command !== "inspect"
    || !["project_overview", "task_change_context"].includes(invocation.options?.purpose)) {
    throw new TypeError("runV1Inspect requires a parsed v1 inspect workflow invocation");
  }
}

function assertPlanChangeInvocation(invocation) {
  if (!invocation || invocation.kind !== "workflow" || invocation.command !== "plan-change") {
    throw new TypeError("runV1PlanChange requires a parsed v1 plan-change workflow invocation");
  }
}

function assertResultTransport(resultTransport) {
  if (!resultTransport || !resultTransport.target || typeof resultTransport.writeResult !== "function") {
    throw new TypeError("runV1Validate requires a reserved v1 result transport");
  }
}
