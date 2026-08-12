import fsPromises from "node:fs/promises";
import path from "node:path";

import { createV1IoError } from "./cli-v1-errors.mjs";
import { serializeV1Result } from "./cli-v1-result.mjs";

export async function reserveV1ResultTransport(resultOption = "-", options = {}) {
  const normalizedResultOption = resultOption ?? "-";
  if (normalizedResultOption === "-") {
    return createStdoutResultTransport(options.stdout ?? process.stdout);
  }
  if (typeof normalizedResultOption !== "string" || normalizedResultOption.length === 0 || normalizedResultOption.includes("\0")) {
    throw createV1IoError({
      code: "io.result-path-unsafe",
      status: "rejected",
      message: "--result must be - or a non-empty file path.",
      option: "--result",
      details: { requested_path: typeof normalizedResultOption === "string" ? normalizedResultOption : null }
    });
  }

  const fileSystem = options.fileSystem ?? fsPromises;
  const cwd = options.cwd ?? process.cwd();
  const resultPath = await resolveNewResultPath(normalizedResultOption, { cwd, fileSystem });
  let handle;
  try {
    handle = await fileSystem.open(resultPath, "wx");
  } catch (error) {
    if (error && error.code === "EEXIST") {
      throw createV1IoError({
        code: "io.result-path-exists",
        status: "rejected",
        message: "The --result path already exists and will not be overwritten.",
        path: resultPath,
        option: "--result",
        details: { requested_path: normalizedResultOption }
      });
    }
    throw createV1IoError({
      code: "io.result-reservation-failed",
      status: "runtime-error",
      message: "The v1 result file could not be reserved exclusively.",
      path: resultPath,
      option: "--result",
      details: { requested_path: normalizedResultOption, error_code: error?.code ?? null }
    });
  }
  return createFileResultTransport({ fileSystem, handle, path: resultPath });
}

export async function resolveNewResultPath(requestedPath, { cwd = process.cwd(), fileSystem = fsPromises } = {}) {
  if (typeof requestedPath !== "string" || requestedPath.length === 0 || requestedPath.includes("\0")) {
    throw createV1IoError({
      code: "io.result-path-unsafe",
      status: "rejected",
      message: "--result must be a non-empty file path.",
      option: "--result",
      details: { requested_path: typeof requestedPath === "string" ? requestedPath : null }
    });
  }
  if (requestedPath.endsWith(path.sep) || (path.sep !== "/" && requestedPath.endsWith("/"))) {
    throw createV1IoError({
      code: "io.result-path-unsafe",
      status: "rejected",
      message: "--result must name a file, not a directory path.",
      option: "--result",
      details: { requested_path: requestedPath }
    });
  }

  const requestedBasename = path.basename(requestedPath);
  if (!requestedBasename || requestedBasename === "." || requestedBasename === "..") {
    throw createV1IoError({
      code: "io.result-path-unsafe",
      status: "rejected",
      message: "--result must have a normal filename.",
      option: "--result",
      details: { requested_path: requestedPath }
    });
  }

  const candidatePath = path.resolve(cwd, requestedPath);
  const basename = path.basename(candidatePath);
  if (!basename || basename === "." || basename === ".." || candidatePath === path.parse(candidatePath).root) {
    throw createV1IoError({
      code: "io.result-path-unsafe",
      status: "rejected",
      message: "--result must have a normal filename.",
      option: "--result",
      details: { requested_path: requestedPath }
    });
  }

  const parentPath = path.dirname(candidatePath);
  let parentStat;
  try {
    parentStat = await fileSystem.lstat(parentPath);
  } catch (error) {
    if (error && error.code === "ENOENT") {
      throw createV1IoError({
        code: "io.result-path-unsafe",
        status: "rejected",
        message: "The parent directory for --result must already exist.",
        option: "--result",
        details: { requested_path: requestedPath, parent_path: parentPath }
      });
    }
    throw createV1IoError({
      code: "io.result-reservation-failed",
      status: "runtime-error",
      message: "The parent directory for --result could not be inspected.",
      option: "--result",
      details: { requested_path: requestedPath, parent_path: parentPath, error_code: error?.code ?? null }
    });
  }
  if (!parentStat.isDirectory()) {
    throw createV1IoError({
      code: "io.result-path-unsafe",
      status: "rejected",
      message: "The parent path for --result must be a directory.",
      option: "--result",
      details: { requested_path: requestedPath, parent_path: parentPath }
    });
  }

  let canonicalParentPath;
  try {
    canonicalParentPath = await fileSystem.realpath(parentPath);
  } catch (error) {
    throw createV1IoError({
      code: "io.result-reservation-failed",
      status: "runtime-error",
      message: "The parent directory for --result could not be canonicalized.",
      option: "--result",
      details: { requested_path: requestedPath, parent_path: parentPath, error_code: error?.code ?? null }
    });
  }
  const canonicalResultPath = path.join(canonicalParentPath, basename);
  try {
    await fileSystem.lstat(canonicalResultPath);
    throw createV1IoError({
      code: "io.result-path-exists",
      status: "rejected",
      message: "The --result path already exists and will not be overwritten.",
      path: canonicalResultPath,
      option: "--result",
      details: { requested_path: requestedPath }
    });
  } catch (error) {
    if (error && error.code === "ENOENT") {
      return canonicalResultPath;
    }
    if (error?.code === "io.result-path-exists") {
      throw error;
    }
    throw createV1IoError({
      code: "io.result-reservation-failed",
      status: "runtime-error",
      message: "The requested --result path could not be inspected.",
      path: canonicalResultPath,
      option: "--result",
      details: { requested_path: requestedPath, error_code: error?.code ?? null }
    });
  }
}

function createStdoutResultTransport(stdout) {
  let completed = false;
  return Object.freeze({
    target: Object.freeze({ target: "stdout", path: null }),
    async writeResult(result) {
      if (completed) {
        throw new Error("v1 result transport has already completed");
      }
      stdout.write(serializeV1Result(result));
      completed = true;
    },
    async abort() {
      completed = true;
    }
  });
}

function createFileResultTransport({ fileSystem, handle, path: resultPath }) {
  let completed = false;
  let closed = false;
  return Object.freeze({
    target: Object.freeze({ target: "file", path: resultPath }),
    async writeResult(result) {
      if (completed) {
        throw new Error("v1 result transport has already completed");
      }
      try {
        await handle.writeFile(serializeV1Result(result), "utf8");
        await closeHandle();
        completed = true;
      } catch (error) {
        await closeAndRemoveReservedFile();
        throw error;
      }
    },
    async abort() {
      if (completed) {
        return;
      }
      await closeAndRemoveReservedFile();
      completed = true;
    }
  });

  async function closeHandle() {
    if (!closed) {
      await handle.close();
      closed = true;
    }
  }

  async function closeAndRemoveReservedFile() {
    try {
      await closeHandle();
    } catch {
      // Best effort: a failed close must not prevent removal of our own file.
    }
    try {
      await fileSystem.unlink(resultPath);
    } catch (error) {
      if (error?.code !== "ENOENT") {
        // The caller receives the original write failure. Cleanup diagnostics
        // belong to the later publication workflow, not this channel helper.
      }
    }
  }
}
