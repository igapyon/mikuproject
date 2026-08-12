import fsPromises from "node:fs/promises";
import path from "node:path";

import { sha256RawBytes } from "./cli-v1-canonical-json.mjs";
import { createV1IoError, createV1RejectedError } from "./cli-v1-errors.mjs";

/**
 * Reads a versioned JSON exchange artifact without inheriting JSON.parse's
 * duplicate-key blind spot.  The caller owns schema/kind validation, while
 * this transport layer fixes byte decoding, direct-entry safety and I/O
 * metadata identically for request, plan-result and approval inputs.
 */
export async function readV1JsonArtifact(optionValue, {
  role,
  option,
  cwd = process.cwd(),
  stdin = process.stdin,
  fileSystem = fsPromises
} = {}) {
  assertDescriptor(role, option);
  const raw = await readV1JsonArtifactBytes(optionValue, { role, option, cwd, stdin, fileSystem });
  if (raw.error) {
    return raw;
  }
  try {
    return { input: raw.input, value: parseV1JsonDocument(raw.bytes, { option, role }) };
  } catch (error) {
    return { input: raw.input, error };
  }
}

export function parseV1JsonDocument(rawInput, { option = "--request", role = "change_request" } = {}) {
  const bytes = asV1JsonArtifactRawBytes(rawInput);
  if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    throw createJsonError({
      code: "json.bom-not-allowed",
      message: "v1 JSON artifacts must not begin with a UTF-8 BOM.",
      option,
      role,
      details: {}
    });
  }
  let text;
  try {
    text = new TextDecoder("utf-8", { fatal: true, ignoreBOM: true }).decode(bytes);
  } catch (cause) {
    throw createJsonError({
      code: "text.invalid-utf8",
      message: "v1 JSON artifacts must be valid UTF-8.",
      option,
      role,
      details: { cause: cause instanceof Error ? cause.name : "decode-failed" }
    });
  }
  if (text.trim().length === 0) {
    throw createJsonError({
      code: "json.invalid",
      message: "A v1 JSON artifact must contain exactly one JSON document.",
      option,
      role,
      details: { reason: "empty-document" }
    });
  }
  try {
    assertNoDuplicateJsonObjectKeys(text, { option, role });
    return JSON.parse(text);
  } catch (error) {
    if (error?.code === "json.duplicate-key") {
      throw error;
    }
    throw createJsonError({
      code: "json.invalid",
      message: "A v1 JSON artifact must contain exactly one valid JSON document.",
      option,
      role,
      details: { reason: error instanceof Error ? error.name : "parse-failed" }
    });
  }
}

async function readV1JsonArtifactBytes(optionValue, { role, option, cwd, stdin, fileSystem }) {
  if (optionValue === "-") {
    const input = jsonArtifactInputMetadata({ role, option, source: "stdin", path: null, digest: null });
    try {
      const bytes = await readV1JsonArtifactStdinBytes(stdin);
      input.digest = sha256RawBytes(bytes);
      return { input, bytes };
    } catch (error) {
      return {
        input,
        error: createV1IoError({
          code: "io.input-read-failed",
          status: "runtime-error",
          message: `The ${option} JSON artifact could not be read from stdin.`,
          option,
          details: { error_code: error?.code ?? null }
        })
      };
    }
  }
  if (typeof optionValue !== "string" || optionValue.length === 0 || optionValue.includes("\0")) {
    throw new TypeError(`${option} requires a parsed file path or stdin marker`);
  }
  const candidatePath = path.resolve(cwd, optionValue);
  let entry;
  try {
    entry = await fileSystem.lstat(candidatePath);
  } catch (error) {
    const input = jsonArtifactInputMetadata({ role, option, source: "file", path: candidatePath, digest: null });
    if (error?.code === "ENOENT") {
      return {
        input,
        error: createV1IoError({
          code: "io.input-not-found",
          status: "rejected",
          message: `The ${option} JSON artifact does not exist.`,
          path: candidatePath,
          option,
          details: { requested_path: optionValue }
        })
      };
    }
    return {
      input,
      error: createV1IoError({
        code: "io.input-read-failed",
        status: "runtime-error",
        message: `The ${option} JSON artifact could not be inspected.`,
        path: candidatePath,
        option,
        details: { requested_path: optionValue, error_code: error?.code ?? null }
      })
    };
  }
  if (entry.isSymbolicLink()) {
    const input = jsonArtifactInputMetadata({ role, option, source: "file", path: candidatePath, digest: null });
    return {
      input,
      error: createV1IoError({
        code: "io.input-symlink-rejected",
        status: "rejected",
        message: `A direct ${option} JSON artifact must not be a symbolic link.`,
        path: candidatePath,
        option,
        details: { requested_path: optionValue }
      })
    };
  }
  if (!entry.isFile()) {
    let canonicalPath = candidatePath;
    try { canonicalPath = await fileSystem.realpath(candidatePath); } catch { /* report inspected entry */ }
    const input = jsonArtifactInputMetadata({ role, option, source: entry.isDirectory() ? "directory" : "file", path: canonicalPath, digest: null });
    return {
      input,
      error: createV1IoError({
        code: "io.input-type-invalid",
        status: "rejected",
        message: `${option} must name a regular JSON file or stdin.`,
        path: canonicalPath,
        option,
        details: { requested_path: optionValue, observed_type: entry.isDirectory() ? "directory" : "other" }
      })
    };
  }
  let canonicalPath;
  try {
    canonicalPath = await fileSystem.realpath(candidatePath);
  } catch (error) {
    const input = jsonArtifactInputMetadata({ role, option, source: "file", path: candidatePath, digest: null });
    return {
      input,
      error: createV1IoError({
        code: "io.input-read-failed",
        status: "runtime-error",
        message: `The ${option} JSON artifact could not be canonicalized.`,
        path: candidatePath,
        option,
        details: { requested_path: optionValue, error_code: error?.code ?? null }
      })
    };
  }
  const input = jsonArtifactInputMetadata({ role, option, source: "file", path: canonicalPath, digest: null });
  try {
    const bytes = asV1JsonArtifactRawBytes(await fileSystem.readFile(canonicalPath));
    input.digest = sha256RawBytes(bytes);
    return { input, bytes };
  } catch (error) {
    return {
      input,
      error: createV1IoError({
        code: "io.input-read-failed",
        status: "runtime-error",
        message: `The ${option} JSON artifact could not be read.`,
        path: canonicalPath,
        option,
        details: { requested_path: optionValue, error_code: error?.code ?? null }
      })
    };
  }
}

function assertNoDuplicateJsonObjectKeys(text, { option, role }) {
  let offset = 0;
  const skipWhitespace = () => {
    while (offset < text.length && /[\u0009\u000a\u000d\u0020]/u.test(text[offset])) offset += 1;
  };
  const readString = () => {
    const start = offset;
    if (text[offset] !== '"') throw new SyntaxError("JSON string expected");
    offset += 1;
    let escaped = false;
    while (offset < text.length) {
      const character = text[offset];
      offset += 1;
      if (escaped) {
        escaped = false;
      } else if (character === "\\") {
        escaped = true;
      } else if (character === '"') {
        return JSON.parse(text.slice(start, offset));
      } else if (character < " ") {
        throw new SyntaxError("control character in JSON string");
      }
    }
    throw new SyntaxError("unterminated JSON string");
  };
  const readLiteral = () => {
    const start = offset;
    while (offset < text.length && !/[\u0009\u000a\u000d\u0020,\]\}]/u.test(text[offset])) offset += 1;
    if (start === offset) throw new SyntaxError("JSON value expected");
    // JSON.parse on the complete document supplies the authoritative lexical
    // validation. This scanner only needs to traverse its already-valid shape.
  };
  const readValue = () => {
    skipWhitespace();
    if (text[offset] === "{") {
      offset += 1;
      skipWhitespace();
      const keys = new Set();
      if (text[offset] === "}") { offset += 1; return; }
      while (true) {
        skipWhitespace();
        const key = readString();
        if (keys.has(key)) {
          throw createJsonError({
            code: "json.duplicate-key",
            message: "A v1 JSON artifact must not contain duplicate object keys.",
            option,
            role,
            details: { key }
          });
        }
        keys.add(key);
        skipWhitespace();
        if (text[offset] !== ":") throw new SyntaxError("JSON object colon expected");
        offset += 1;
        readValue();
        skipWhitespace();
        if (text[offset] === "}") { offset += 1; return; }
        if (text[offset] !== ",") throw new SyntaxError("JSON object separator expected");
        offset += 1;
      }
    }
    if (text[offset] === "[") {
      offset += 1;
      skipWhitespace();
      if (text[offset] === "]") { offset += 1; return; }
      while (true) {
        readValue();
        skipWhitespace();
        if (text[offset] === "]") { offset += 1; return; }
        if (text[offset] !== ",") throw new SyntaxError("JSON array separator expected");
        offset += 1;
      }
    }
    if (text[offset] === '"') {
      readString();
      return;
    }
    readLiteral();
  };
  readValue();
  skipWhitespace();
  if (offset !== text.length) throw new SyntaxError("trailing JSON content");
}

function createJsonError({ code, message, option, role, details }) {
  return createV1RejectedError({
    code,
    message,
    scope: "artifact",
    option,
    artifactRole: role,
    details
  });
}

function jsonArtifactInputMetadata({ role, option, source, path: inputPath, digest }) {
  return { role, option, source, path: inputPath, digest };
}

async function readV1JsonArtifactStdinBytes(stdin) {
  if (Buffer.isBuffer(stdin) || stdin instanceof Uint8Array) return Buffer.from(stdin);
  if (!stdin || typeof stdin[Symbol.asyncIterator] !== "function") {
    throw new TypeError("v1 stdin must be a Buffer, Uint8Array, or async byte iterable");
  }
  const chunks = [];
  for await (const chunk of stdin) chunks.push(asV1JsonArtifactRawBytes(chunk));
  return Buffer.concat(chunks);
}

function asV1JsonArtifactRawBytes(value) {
  if (Buffer.isBuffer(value) || value instanceof Uint8Array) return Buffer.from(value);
  throw new TypeError("v1 input must provide raw bytes");
}

function assertDescriptor(role, option) {
  if (typeof role !== "string" || role.length === 0 || typeof option !== "string" || !option.startsWith("--")) {
    throw new TypeError("v1 JSON artifact reader requires a logical role and long option");
  }
}
