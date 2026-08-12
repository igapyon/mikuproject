import { createHash } from "node:crypto";

export class CanonicalJsonError extends TypeError {
  constructor(message) {
    super(message);
    this.name = "CanonicalJsonError";
  }
}

/**
 * Serializes the v1 JSON domain without relying on a host JSON serializer.
 * It intentionally accepts only values that can be represented identically by
 * the Node and Java v1 runtimes.
 */
export function canonicalJsonText(value) {
  return serializeValue(value, new Set(), "$");
}

export function canonicalJsonBytes(value) {
  return Buffer.from(canonicalJsonText(value), "utf8");
}

export function sha256CanonicalJson(value) {
  return sha256Bytes(canonicalJsonBytes(value));
}

export function sha256RawBytes(bytes) {
  if (!Buffer.isBuffer(bytes) && !(bytes instanceof Uint8Array)) {
    throw new CanonicalJsonError("raw digest input must be a Buffer or Uint8Array");
  }
  return sha256Bytes(bytes);
}

export function canonicalizeSemanticState(state) {
  assertPlainObject(state, "$", "semantic state");
  const canonicalState = cloneJsonValue(state, "$", new Set());

  canonicalState.tasks = copyArray(canonicalState.tasks, "$.tasks");
  canonicalState.dependencies = copyArray(canonicalState.dependencies, "$.dependencies")
    .sort((left, right) => compareTuple(
      left,
      right,
      ["predecessor_uid", "successor_uid", "type", "lag"],
      "$.dependencies"
    ));
  canonicalState.resources = copyArray(canonicalState.resources, "$.resources")
    .sort((left, right) => compareUid(left, right, "$.resources"));
  canonicalState.assignments = copyArray(canonicalState.assignments, "$.assignments")
    .sort((left, right) => compareUid(left, right, "$.assignments"));
  canonicalState.calendars = copyArray(canonicalState.calendars, "$.calendars")
    .sort((left, right) => compareUid(left, right, "$.calendars"));

  return canonicalState;
}

export function sha256SemanticState(state) {
  return sha256CanonicalJson(canonicalizeSemanticState(state));
}

export function compareUnicodeScalars(left, right) {
  assertUnicodeScalarString(left, "left comparator value");
  assertUnicodeScalarString(right, "right comparator value");
  let leftIndex = 0;
  let rightIndex = 0;
  while (leftIndex < left.length && rightIndex < right.length) {
    const leftCodePoint = left.codePointAt(leftIndex);
    const rightCodePoint = right.codePointAt(rightIndex);
    if (leftCodePoint !== rightCodePoint) {
      return leftCodePoint < rightCodePoint ? -1 : 1;
    }
    leftIndex += leftCodePoint > 0xffff ? 2 : 1;
    rightIndex += rightCodePoint > 0xffff ? 2 : 1;
  }
  if (leftIndex === left.length && rightIndex === right.length) {
    return 0;
  }
  return leftIndex === left.length ? -1 : 1;
}

function serializeValue(value, ancestors, location) {
  if (value === null) {
    return "null";
  }
  switch (typeof value) {
    case "boolean":
      return value ? "true" : "false";
    case "string":
      return quoteString(value, location);
    case "number":
      if (!Number.isSafeInteger(value) || Object.is(value, -0)) {
        throw new CanonicalJsonError(`${location} must be a safe integer other than -0`);
      }
      return String(value);
    case "object":
      if (Array.isArray(value)) {
        return serializeArray(value, ancestors, location);
      }
      assertPlainObject(value, location, "canonical JSON object");
      return serializeObject(value, ancestors, location);
    default:
      throw new CanonicalJsonError(`${location} has unsupported canonical JSON type: ${typeof value}`);
  }
}

function serializeArray(value, ancestors, location) {
  assertArrayShape(value, location);
  assertAcyclic(value, ancestors, location);
  const result = [];
  for (let index = 0; index < value.length; index += 1) {
    result.push(serializeValue(value[index], ancestors, `${location}[${index}]`));
  }
  ancestors.delete(value);
  return `[${result.join(",")}]`;
}

function serializeObject(value, ancestors, location) {
  assertObjectShape(value, location);
  assertAcyclic(value, ancestors, location);
  const keys = Object.keys(value).sort(compareUnicodeScalars);
  const result = keys.map((key) => `${quoteString(key, `${location} key`)}:${serializeValue(value[key], ancestors, `${location}.${key}`)}`);
  ancestors.delete(value);
  return `{${result.join(",")}}`;
}

function cloneJsonValue(value, location, ancestors) {
  if (value === null || typeof value === "boolean" || typeof value === "string" || typeof value === "number") {
    // Use the serializer's checks so an invalid JS-domain value cannot be
    // silently retained in a supposedly canonical semantic state.
    serializeValue(value, ancestors, location);
    return value;
  }
  if (Array.isArray(value)) {
    assertArrayShape(value, location);
    assertAcyclic(value, ancestors, location);
    const copy = value.map((item, index) => cloneJsonValue(item, `${location}[${index}]`, ancestors));
    ancestors.delete(value);
    return copy;
  }
  assertPlainObject(value, location, "semantic state object");
  assertObjectShape(value, location);
  assertAcyclic(value, ancestors, location);
  const copy = {};
  for (const key of Object.keys(value)) {
    copy[key] = cloneJsonValue(value[key], `${location}.${key}`, ancestors);
  }
  ancestors.delete(value);
  return copy;
}

function quoteString(value, location) {
  assertUnicodeScalarString(value, location);
  let result = '"';
  for (let index = 0; index < value.length; index += 1) {
    const codeUnit = value.charCodeAt(index);
    if (codeUnit === 0x22) {
      result += '\\"';
    } else if (codeUnit === 0x5c) {
      result += "\\\\";
    } else if (codeUnit === 0x08) {
      result += "\\b";
    } else if (codeUnit === 0x09) {
      result += "\\t";
    } else if (codeUnit === 0x0a) {
      result += "\\n";
    } else if (codeUnit === 0x0c) {
      result += "\\f";
    } else if (codeUnit === 0x0d) {
      result += "\\r";
    } else if (codeUnit <= 0x1f) {
      result += `\\u${codeUnit.toString(16).padStart(4, "0")}`;
    } else {
      result += value[index];
    }
  }
  return `${result}"`;
}

function assertUnicodeScalarString(value, location) {
  if (typeof value !== "string") {
    throw new CanonicalJsonError(`${location} must be a string`);
  }
  for (let index = 0; index < value.length; index += 1) {
    const codeUnit = value.charCodeAt(index);
    if (codeUnit >= 0xd800 && codeUnit <= 0xdbff) {
      const next = value.charCodeAt(index + 1);
      if (!(next >= 0xdc00 && next <= 0xdfff)) {
        throw new CanonicalJsonError(`${location} contains an unpaired high surrogate`);
      }
      index += 1;
    } else if (codeUnit >= 0xdc00 && codeUnit <= 0xdfff) {
      throw new CanonicalJsonError(`${location} contains an unpaired low surrogate`);
    }
  }
}

function assertPlainObject(value, location, description) {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new CanonicalJsonError(`${location} must be a ${description}`);
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new CanonicalJsonError(`${location} must not use a custom object prototype`);
  }
}

function assertObjectShape(value, location) {
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== "string") {
      throw new CanonicalJsonError(`${location} must not contain symbol keys`);
    }
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor.enumerable || !Object.hasOwn(descriptor, "value")) {
      throw new CanonicalJsonError(`${location}.${key} must be an enumerable data property`);
    }
  }
}

function assertArrayShape(value, location) {
  if (Object.getPrototypeOf(value) !== Array.prototype) {
    throw new CanonicalJsonError(`${location} must not use a custom array prototype`);
  }
  for (let index = 0; index < value.length; index += 1) {
    if (!Object.hasOwn(value, index)) {
      throw new CanonicalJsonError(`${location} must not be a sparse array`);
    }
  }
  for (const key of Reflect.ownKeys(value)) {
    if (key === "length") {
      continue;
    }
    if (typeof key !== "string" || !isArrayIndex(key) || Number(key) >= value.length) {
      throw new CanonicalJsonError(`${location} must not contain non-index array properties`);
    }
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor.enumerable || !Object.hasOwn(descriptor, "value")) {
      throw new CanonicalJsonError(`${location}[${key}] must be an enumerable data property`);
    }
  }
}

function assertAcyclic(value, ancestors, location) {
  if (ancestors.has(value)) {
    throw new CanonicalJsonError(`${location} contains a cycle`);
  }
  ancestors.add(value);
}

function isArrayIndex(key) {
  if (!/^(0|[1-9][0-9]*)$/.test(key)) {
    return false;
  }
  const index = Number(key);
  return Number.isSafeInteger(index) && index >= 0 && index < 2 ** 32 - 1 && String(index) === key;
}

function copyArray(value, location) {
  if (!Array.isArray(value)) {
    throw new CanonicalJsonError(`${location} must be an array`);
  }
  return [...value];
}

function compareTuple(left, right, keys, location) {
  for (const key of keys) {
    const comparison = compareUnicodeScalars(readStringField(left, key, location), readStringField(right, key, location));
    if (comparison !== 0) {
      return comparison;
    }
  }
  return 0;
}

function compareUid(left, right, location) {
  return compareUnicodeScalars(readStringField(left, "uid", location), readStringField(right, "uid", location));
}

function readStringField(value, key, location) {
  assertPlainObject(value, location, "semantic collection member");
  if (!Object.hasOwn(value, key) || typeof value[key] !== "string") {
    throw new CanonicalJsonError(`${location} member must have string ${key}`);
  }
  return value[key];
}

function sha256Bytes(bytes) {
  return {
    algorithm: "sha-256",
    value: createHash("sha256").update(bytes).digest("hex")
  };
}
