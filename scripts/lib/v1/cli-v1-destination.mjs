import fsPromises from "node:fs/promises";
import path from "node:path";

import { createV1RejectedError, createV1RuntimeError } from "./cli-v1-errors.mjs";

/**
 * Read-only C1 destination preflight.  It never creates or reserves the
 * directory; apply-change repeats the check and owns exclusive creation.
 */
export async function preflightV1NewDestination(requestedPath, {
  cwd = process.cwd(),
  fileSystem = fsPromises,
  projectInput = null
} = {}) {
  if (typeof requestedPath !== "string" || requestedPath.length === 0 || requestedPath.includes("\0")) {
    throw destinationUnsafe(requestedPath, "Destination must be a non-empty directory path.");
  }
  if (requestedPath.endsWith(path.sep) || (path.sep !== "/" && requestedPath.endsWith("/"))) {
    throw destinationUnsafe(requestedPath, "Destination must name a new directory basename, not a trailing separator path.");
  }
  const candidatePath = path.resolve(cwd, requestedPath);
  const basename = path.basename(candidatePath);
  if (!basename || basename === "." || basename === ".." || candidatePath === path.parse(candidatePath).root) {
    throw destinationUnsafe(requestedPath, "Destination must have a normal non-root basename.");
  }
  const parentPath = path.dirname(candidatePath);
  let parentEntry;
  try {
    parentEntry = await fileSystem.lstat(parentPath);
  } catch (error) {
    if (error?.code === "ENOENT") {
      throw destinationUnsafe(requestedPath, "Destination parent must already exist.", { parent_path: parentPath });
    }
    throw createV1RuntimeError({
      code: "publication.capability-unsupported",
      message: "Destination parent could not be inspected for exclusive publication support.",
      scope: "filesystem",
      path: parentPath,
      option: "--destination",
      details: { requested_path: requestedPath, error_code: error?.code ?? null }
    });
  }
  if (parentEntry.isSymbolicLink() || !parentEntry.isDirectory()) {
    throw destinationUnsafe(requestedPath, "Destination parent must be a non-symlink directory.", { parent_path: parentPath });
  }
  let canonicalParent;
  try {
    canonicalParent = await fileSystem.realpath(parentPath);
  } catch (error) {
    throw createV1RuntimeError({
      code: "publication.capability-unsupported",
      message: "Destination parent could not be canonicalized for exclusive publication support.",
      scope: "filesystem",
      path: parentPath,
      option: "--destination",
      details: { requested_path: requestedPath, error_code: error?.code ?? null }
    });
  }
  const destinationPath = path.join(canonicalParent, basename);
  try {
    await fileSystem.lstat(destinationPath);
    throw createV1RejectedError({
      code: "publication.destination-exists",
      message: "Destination already exists and will not be replaced.",
      scope: "filesystem",
      path: destinationPath,
      option: "--destination",
      artifactRole: null,
      details: { requested_path: requestedPath }
    });
  } catch (error) {
    if (error?.code === "ENOENT") {
      assertOutsideProjectInput(destinationPath, projectInput, requestedPath);
      return Object.freeze({ requested_path: requestedPath, path: destinationPath });
    }
    if (error?.code === "publication.destination-exists") throw error;
    throw createV1RuntimeError({
      code: "publication.capability-unsupported",
      message: "Destination availability could not be determined safely.",
      scope: "filesystem",
      path: destinationPath,
      option: "--destination",
      details: { requested_path: requestedPath, error_code: error?.code ?? null }
    });
  }
}

function assertOutsideProjectInput(destinationPath, projectInput, requestedPath) {
  if (!projectInput || projectInput.source !== "directory" || typeof projectInput.path !== "string") return;
  const relative = path.relative(projectInput.path, destinationPath);
  if (relative === "" || (!relative.startsWith(`..${path.sep}`) && relative !== ".." && !path.isAbsolute(relative))) {
    throw destinationUnsafe(requestedPath, "Destination must not be an input project artifact set or its descendant.", {
      input_project_path: projectInput.path
    });
  }
}

function destinationUnsafe(requestedPath, message, details = {}) {
  return createV1RejectedError({
    code: "publication.destination-unsafe",
    message,
    scope: "filesystem",
    option: "--destination",
    artifactRole: null,
    details: { requested_path: typeof requestedPath === "string" ? requestedPath : null, ...details }
  });
}
