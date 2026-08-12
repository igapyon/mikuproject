import { DOMParser } from "@xmldom/xmldom";

import { sha256RawBytes } from "./cli-v1-canonical-json.mjs";
import { createV1RejectedError } from "./cli-v1-errors.mjs";

export const MS_PROJECT_XML_SUBSET_PROFILE = "miku-project-ms-project-xml-subset/v1";
export const MS_PROJECT_XML_ADAPTER = "ms-project-xml-adapter/v1";

const PROJECT_NAMESPACE = "http://schemas.microsoft.com/project";
const XMLNS_NAMESPACE = "http://www.w3.org/2000/xmlns/";

const PROJECT_CHILDREN = new Set([
  "Name", "StartDate", "FinishDate", "CurrentDate", "ScheduleFromStart", "CalendarUID",
  "Tasks", "Resources", "Assignments", "Calendars"
]);
const TASK_CHILDREN = new Set([
  "UID", "ID", "Name", "OutlineLevel", "OutlineNumber", "Start", "Finish", "Duration",
  "Milestone", "Summary", "PercentComplete", "CalendarUID", "PredecessorLink"
]);
const PREDECESSOR_CHILDREN = new Set(["PredecessorUID", "Type", "LinkLag", "LagFormat"]);
const RESOURCE_CHILDREN = new Set(["UID", "ID", "Name", "Type", "CalendarUID"]);
const ASSIGNMENT_CHILDREN = new Set(["UID", "TaskUID", "ResourceUID", "Start", "Finish", "Units", "Work"]);
const CALENDAR_CHILDREN = new Set(["UID", "Name", "IsBaseCalendar"]);

/**
 * Decodes only the explicitly documented MS Project XML subset.  It records
 * raw-input provenance and adapter normalizations, while leaving semantic
 * invariants to cli-v1-semantic-validator.mjs.
 */
export function decodeMsProjectXmlSubset(rawInput) {
  const rawBytes = asMsProjectXmlRawBytes(rawInput);
  const decoded = decodeUtf8Xml(rawBytes);
  const document = parseXml(decoded.text);
  const root = requireProjectRoot(document);
  const adapterIssues = [];

  const projectChildren = inspectDirectChildren(root, PROJECT_CHILDREN, new Set(), "Project");
  appendUnknownIssues(adapterIssues, projectChildren.unknown, "project");

  const calendars = decodeCalendars(projectChildren.values.get("Calendars"), adapterIssues);
  const tasks = decodeTasks(projectChildren.values.get("Tasks"), adapterIssues);
  const resources = decodeResources(projectChildren.values.get("Resources"), adapterIssues);
  const assignments = decodeAssignments(projectChildren.values.get("Assignments"), adapterIssues);
  const project = decodeProject(projectChildren.values, adapterIssues);
  const state = {
    kind: "miku_project_semantic_state",
    schema_version: "1",
    semantic_contract_version: "1",
    project,
    tasks: tasks.tasks,
    dependencies: tasks.dependencies,
    resources,
    assignments,
    calendars
  };

  return Object.freeze({
    format_profile: MS_PROJECT_XML_SUBSET_PROFILE,
    adapter: MS_PROJECT_XML_ADAPTER,
    raw_digest: sha256RawBytes(rawBytes),
    normalizations: Object.freeze(decoded.normalizations),
    adapter_issues: Object.freeze(adapterIssues.map((issue) => Object.freeze({ ...issue }))),
    state
  });
}

function asMsProjectXmlRawBytes(rawInput) {
  if (Buffer.isBuffer(rawInput)) {
    return Buffer.from(rawInput);
  }
  if (rawInput instanceof Uint8Array) {
    return Buffer.from(rawInput);
  }
  throw new TypeError("MS Project XML input must be a Buffer or Uint8Array");
}

function decodeUtf8Xml(rawBytes) {
  let offset = 0;
  const normalizations = [];
  if (rawBytes.length >= 3 && rawBytes[0] === 0xef && rawBytes[1] === 0xbb && rawBytes[2] === 0xbf) {
    offset = 3;
    normalizations.push({
      code: "text.utf8-bom-removed",
      path: "Project",
      before: "utf8-bom",
      after: "no-bom"
    });
  }
  let text;
  try {
    text = new TextDecoder("utf-8", { fatal: true, ignoreBOM: true }).decode(rawBytes.subarray(offset));
  } catch (cause) {
    throw createV1RejectedError({
      code: "text.invalid-utf8",
      message: "Project XML must be valid UTF-8.",
      scope: "input",
      path: null,
      details: { cause: cause instanceof Error ? cause.name : "decode-failed" }
    });
  }
  const declaration = /^\s*<\?xml\s+([^?]*)\?>/i.exec(text);
  if (declaration) {
    const encoding = /\bencoding\s*=\s*(['"])(.*?)\1/i.exec(declaration[1]);
    if (encoding && encoding[2].toUpperCase() !== "UTF-8") {
      throw createV1RejectedError({
        code: "xml.encoding-unsupported",
        message: "Project XML declarations may specify only UTF-8 encoding.",
        scope: "input",
        path: "Project",
        details: { encoding: encoding[2] }
      });
    }
  }
  return { text, normalizations };
}

function parseXml(text) {
  const messages = [];
  let document;
  try {
    document = new DOMParser({
      errorHandler: {
        warning() {},
        error(message) { messages.push(String(message)); },
        fatalError(message) { messages.push(String(message)); }
      }
    }).parseFromString(text, "application/xml");
  } catch (cause) {
    throw createInvalidXmlError("Project XML could not be parsed.", { cause: cause instanceof Error ? cause.name : "parse-failed" });
  }
  if (messages.length > 0 || !document || !document.documentElement || document.getElementsByTagName("parsererror").length > 0) {
    throw createInvalidXmlError("Project XML is not well-formed.", { parser_messages: messages });
  }
  for (const child of Array.from(document.childNodes)) {
    if (child.nodeType === 1) {
      continue;
    }
    if (child.nodeType === 3 && /^\s*$/.test(child.data)) {
      continue;
    }
    if (child.nodeType === 7 || child.nodeType === 8) {
      continue;
    }
    throw createInvalidXmlError("Project XML contains unsupported document-level content.", { node_type: child.nodeType });
  }
  return document;
}

function requireProjectRoot(document) {
  const root = document.documentElement;
  if (localName(root) !== "Project" || root.namespaceURI !== PROJECT_NAMESPACE) {
    throw createV1RejectedError({
      code: "xml.profile-unsupported",
      message: "Project XML root must be Project in the Microsoft Project namespace.",
      scope: "input",
      path: "Project",
      details: { local_name: localName(root), namespace: root.namespaceURI ?? null }
    });
  }
  assertAllowedAttributes(root, "Project");
  return root;
}

function decodeProject(values, adapterIssues) {
  const project = {};
  copyText(project, "name", values.get("Name"));
  copyText(project, "start", values.get("StartDate"));
  copyText(project, "finish", values.get("FinishDate"));
  copyText(project, "current_date", values.get("CurrentDate"));
  copyBoolean(project, "schedule_from_start", values.get("ScheduleFromStart"));
  copyIdentity(project, "calendar_uid", values.get("CalendarUID"));
  return project;
}

function decodeTasks(containers, adapterIssues) {
  const taskElements = requireCollectionMembers(containers, "Tasks", "Task");
  const decoded = taskElements.map((element, index) => decodeTask(element, index, adapterIssues));
  const tasks = [];
  const dependencies = [];
  const levelStack = [];
  const siblingPositions = [];
  const taskIds = new Set();
  let sawPseudoTask = false;

  for (let index = 0; index < decoded.length; index += 1) {
    const source = decoded[index];
    if (source.id !== undefined && Number.isSafeInteger(source.id) && source.id > 0) {
      if (taskIds.has(source.id)) {
        adapterIssues.push(invalidIssue("S-I003", `${xmlTaskPath(source, index)}.id`, "Task ID must be unique when present."));
      }
      taskIds.add(source.id);
    }
    if (source.uid === "0") {
      const validPseudo = index === 0 && !sawPseudoTask && source.outline_level === 0 && source.summary === true;
      if (!validPseudo) {
        adapterIssues.push(invalidIssue("S-I003", `tasks[uid=0]`, "The project summary pseudo task is not in its required form."));
      }
      sawPseudoTask = true;
      continue;
    }
    const level = source.outline_level;
    let parentUid = null;
    if (!Number.isSafeInteger(level) || level < 1) {
      adapterIssues.push(invalidIssue("S-I003", xmlTaskPath(source, index), "Task outline level must begin at 1 for a semantic task."));
    } else if (tasks.length === 0 && level !== 1) {
      adapterIssues.push(invalidIssue("S-I003", xmlTaskPath(source, index), "The first semantic task must be a root task."));
    } else if (tasks.length > 0 && level > levelStack.length + 1) {
      adapterIssues.push(invalidIssue("S-I003", xmlTaskPath(source, index), "Task outline levels may increase by at most one."));
    } else if (level > 1) {
      parentUid = levelStack[level - 2] ?? null;
      if (parentUid === null) {
        adapterIssues.push(invalidIssue("S-I003", xmlTaskPath(source, index), "Task outline parent is missing."));
      }
    }
    if (Number.isSafeInteger(level) && level >= 1) {
      siblingPositions.length = level;
      siblingPositions[level - 1] = (siblingPositions[level - 1] ?? 0) + 1;
      const expectedOutlineNumber = siblingPositions.join(".");
      if (source.outline_number !== undefined && source.outline_number !== expectedOutlineNumber) {
        adapterIssues.push(invalidIssue("S-I003", `${xmlTaskPath(source, index)}.outline_number`, "Task OutlineNumber must match the derived sibling position."));
      }
    }
    levelStack.length = Math.max(0, Number.isSafeInteger(level) ? level - 1 : 0);
    levelStack.push(source.uid);

    const task = {};
    copyValue(task, "uid", source.uid);
    copyValue(task, "name", source.name);
    task.parent_uid = parentUid;
    copyValue(task, "start", source.start);
    copyValue(task, "finish", source.finish);
    copyValue(task, "duration", source.duration);
    copyValue(task, "milestone", source.milestone);
    copyValue(task, "summary", source.summary);
    copyValue(task, "percent_complete", source.percent_complete);
    copyValue(task, "calendar_uid", source.calendar_uid);
    tasks.push(task);
    dependencies.push(...source.dependencies.map((dependency) => ({ ...dependency, successor_uid: source.uid })));
  }

  return { tasks, dependencies };
}

function decodeTask(element, index, adapterIssues) {
  const children = inspectDirectChildren(element, TASK_CHILDREN, new Set(["PredecessorLink"]), `Tasks/Task[${index + 1}]`);
  const values = children.values;
  const uid = readIdentity(values.get("UID"));
  appendUnknownIssues(adapterIssues, children.unknown, `${xmlTaskPath({ uid }, index)}`);
  const dependencies = decodePredecessors(values.get("PredecessorLink"), uid, adapterIssues);
  const task = {
    uid,
    name: readText(values.get("Name")),
    outline_level: readInteger(values.get("OutlineLevel")),
    start: readText(values.get("Start")),
    finish: readText(values.get("Finish")),
    duration: readDuration(values.get("Duration")),
    milestone: readBoolean(values.get("Milestone")),
    summary: readBoolean(values.get("Summary")),
    percent_complete: readInteger(values.get("PercentComplete")),
    calendar_uid: readIdentity(values.get("CalendarUID")),
    dependencies
  };
  validateTaskAdapterFields(values, task, index, adapterIssues);
  return task;
}

function validateTaskAdapterFields(values, task, index, adapterIssues) {
  const taskLocation = xmlTaskPath(task, index);
  const id = readInteger(values.get("ID"));
  task.id = id;
  if (id !== undefined && (!Number.isSafeInteger(id) || id <= 0)) {
    adapterIssues.push(invalidIssue("S-I003", `${taskLocation}.id`, "Task ID must be a positive integer when present."));
  }
  const outlineNumber = readText(values.get("OutlineNumber"));
  if (outlineNumber !== undefined && Number.isSafeInteger(task.outline_level) && task.outline_level > 0) {
    // Exact outline-number comparison depends on sibling positions.  The
    // semantic validator owns that forest check once parent relations exist.
    task.outline_number = outlineNumber;
  }
}

function decodePredecessors(elements, successorUid, adapterIssues) {
  const dependencies = [];
  for (const [index, element] of (elements ?? []).entries()) {
    const children = inspectDirectChildren(element, PREDECESSOR_CHILDREN, new Set(), `Task/PredecessorLink[${index + 1}]`);
    const predecessorUid = readIdentity(children.values.get("PredecessorUID"));
    const path = `dependencies[predecessor_uid=${predecessorUid ?? "?"},successor_uid=${successorUid ?? "?"}]`;
    appendUnknownIssues(adapterIssues, children.unknown, path);
    const type = readInteger(children.values.get("Type"));
    const rawLag = readText(children.values.get("LinkLag"));
    const lagFormat = readInteger(children.values.get("LagFormat"));
    let lag = rawLag;
    if (rawLag === "0" && lagFormat === 3) {
      lag = "PT0H0M0S";
    } else if (rawLag === "PT0H0M0S" && lagFormat === undefined) {
      lag = "PT0H0M0S";
    } else if (rawLag !== undefined) {
      adapterIssues.push(unsupportedIssue("S-I019", `${path}.lag`, "Only FS dependencies with zero lag are supported."));
    }
    if (type !== undefined && type !== 1) {
      adapterIssues.push(unsupportedIssue("S-I019", `${path}.type`, "Only FS dependencies are supported."));
    }
    const dependency = {};
    copyValue(dependency, "predecessor_uid", predecessorUid);
    copyValue(dependency, "type", type === 1 ? "FS" : type);
    copyValue(dependency, "lag", lag);
    dependencies.push(dependency);
  }
  return dependencies;
}

function decodeResources(containers, adapterIssues) {
  const elements = requireCollectionMembers(containers, "Resources", "Resource");
  return elements.map((element, index) => {
    const children = inspectDirectChildren(element, RESOURCE_CHILDREN, new Set(), `Resources/Resource[${index + 1}]`);
    const uid = readIdentity(children.values.get("UID"));
    const path = `resources[uid=${uid ?? "?"}]`;
    appendUnknownIssues(adapterIssues, children.unknown, path);
    const resource = {};
    copyValue(resource, "uid", uid);
    copyValue(resource, "name", readText(children.values.get("Name")));
    const type = readInteger(children.values.get("Type"));
    if (type !== undefined) {
      const resourceType = { 0: "material", 1: "work", 2: "cost" }[type];
      if (resourceType) {
        resource.type = resourceType;
      } else {
        adapterIssues.push(unsupportedIssue("S-I020", `${path}.type`, "This external resource type is outside the v1 XML subset."));
      }
    }
    copyValue(resource, "calendar_uid", readIdentity(children.values.get("CalendarUID")));
    return resource;
  });
}

function decodeAssignments(containers, adapterIssues) {
  const elements = requireCollectionMembers(containers, "Assignments", "Assignment");
  return elements.map((element, index) => {
    const children = inspectDirectChildren(element, ASSIGNMENT_CHILDREN, new Set(), `Assignments/Assignment[${index + 1}]`);
    const uid = readIdentity(children.values.get("UID"));
    const path = `assignments[uid=${uid ?? "?"}]`;
    appendUnknownIssues(adapterIssues, children.unknown, path);
    const assignment = {};
    copyValue(assignment, "uid", uid);
    copyValue(assignment, "task_uid", readIdentity(children.values.get("TaskUID")));
    const resourceUid = readText(children.values.get("ResourceUID"));
    if (resourceUid !== "-65535") {
      copyValue(assignment, "resource_uid", normalizeIdentity(resourceUid));
    }
    copyValue(assignment, "start", readText(children.values.get("Start")));
    copyValue(assignment, "finish", readText(children.values.get("Finish")));
    copyValue(assignment, "units", readUnits(children.values.get("Units")));
    copyValue(assignment, "work", readDuration(children.values.get("Work")));
    return assignment;
  });
}

function decodeCalendars(containers, adapterIssues) {
  const elements = requireCollectionMembers(containers, "Calendars", "Calendar");
  return elements.map((element, index) => {
    const children = inspectDirectChildren(element, CALENDAR_CHILDREN, new Set(), `Calendars/Calendar[${index + 1}]`);
    const uid = readIdentity(children.values.get("UID"));
    const path = `calendars[uid=${uid ?? "?"}]`;
    appendUnknownIssues(adapterIssues, children.unknown, path);
    const calendar = {};
    copyValue(calendar, "uid", uid);
    copyValue(calendar, "name", readText(children.values.get("Name")));
    copyValue(calendar, "is_base_calendar", readBoolean(children.values.get("IsBaseCalendar")));
    return calendar;
  });
}

function requireCollectionMembers(containers, containerName, memberName) {
  if (!containers || containers.length === 0) {
    return [];
  }
  const members = [];
  for (const container of containers) {
    const children = inspectDirectChildren(container, new Set([memberName]), new Set([memberName]), containerName);
    if (children.unknown.length > 0) {
      throw createV1RejectedError({
        code: "xml.profile-unsupported",
        message: `${containerName} contains an unsupported member.`,
        scope: "input",
        path: containerName,
        details: { elements: children.unknown.map((element) => localName(element)) }
      });
    }
    members.push(...(children.values.get(memberName) ?? []));
  }
  if (members.length === 0) {
    throw createInvalidXmlError(`${containerName} must not be empty when present.`, { path: containerName });
  }
  return members;
}

function inspectDirectChildren(parent, allowedNames, collectionNames, path) {
  assertElementNamespace(parent, path);
  assertAllowedAttributes(parent, path);
  const values = new Map();
  const unknown = [];
  for (const child of Array.from(parent.childNodes)) {
    if (child.nodeType === 3 || child.nodeType === 4) {
      if (!/^\s*$/.test(child.data)) {
        throw createInvalidXmlError(`${path} contains text where child elements are required.`, { path });
      }
      continue;
    }
    if (child.nodeType === 7 || child.nodeType === 8) {
      continue;
    }
    if (child.nodeType !== 1) {
      throw createInvalidXmlError(`${path} contains an invalid XML node.`, { path, node_type: child.nodeType });
    }
    assertElementNamespace(child, path);
    assertAllowedAttributes(child, `${path}/${localName(child)}`);
    const name = localName(child);
    if (!allowedNames.has(name)) {
      unknown.push(child);
      continue;
    }
    const current = values.get(name) ?? [];
    current.push(child);
    if (!collectionNames.has(name) && current.length > 1) {
      throw createInvalidXmlError(`${path}/${name} must not be repeated.`, { path: `${path}/${name}` });
    }
    values.set(name, current);
  }
  return { values, unknown };
}

function readText(elements) {
  if (!elements || elements.length === 0) {
    return undefined;
  }
  const element = elements[0];
  for (const child of Array.from(element.childNodes)) {
    if (child.nodeType === 1) {
      throw createInvalidXmlError(`${localName(element)} must contain text only.`, { path: localName(element) });
    }
    if (child.nodeType !== 3 && child.nodeType !== 4 && child.nodeType !== 7 && child.nodeType !== 8) {
      throw createInvalidXmlError(`${localName(element)} contains an invalid XML node.`, { path: localName(element) });
    }
  }
  return element.textContent ?? "";
}

function readIdentity(elements) {
  return normalizeIdentity(readText(elements));
}

function normalizeIdentity(value) {
  if (value === undefined) {
    return undefined;
  }
  return /^(?:0|[1-9][0-9]*)$/.test(value) ? value : value;
}

function readInteger(elements) {
  const value = readText(elements);
  if (value === undefined) {
    return undefined;
  }
  return /^(?:0|[1-9][0-9]*)$/.test(value) && Number.isSafeInteger(Number(value)) ? Number(value) : value;
}

function readBoolean(elements) {
  const value = readText(elements);
  if (value === undefined) {
    return undefined;
  }
  if (value === "0") {
    return false;
  }
  if (value === "1") {
    return true;
  }
  return value;
}

function readDuration(elements) {
  const value = readText(elements);
  if (value === undefined) {
    return undefined;
  }
  const match = /^PT([0-9]+)H([0-9]+)M([0-9]+)S$/.exec(value);
  if (!match || Number(match[2]) > 59 || Number(match[3]) > 59) {
    return value;
  }
  return `PT${Number(match[1])}H${Number(match[2])}M${Number(match[3])}S`;
}

function readUnits(elements) {
  const value = readText(elements);
  if (value === undefined) {
    return undefined;
  }
  if (!/^(?:0|[1-9][0-9]*)(?:\.[0-9]+)?$/.test(value)) {
    return value;
  }
  const [integer, fraction] = value.split(".");
  if (!fraction) {
    return integer;
  }
  const trimmed = fraction.replace(/0+$/, "");
  return trimmed ? `${integer}.${trimmed}` : integer;
}

function copyText(target, property, elements) {
  copyValue(target, property, readText(elements));
}

function copyIdentity(target, property, elements) {
  copyValue(target, property, readIdentity(elements));
}

function copyBoolean(target, property, elements) {
  copyValue(target, property, readBoolean(elements));
}

function copyValue(target, property, value) {
  if (value !== undefined) {
    target[property] = value;
  }
}

function appendUnknownIssues(adapterIssues, elements, semanticPath) {
  for (const element of elements) {
    adapterIssues.push(unsupportedIssue(
      "S-I020",
      `${semanticPath}.${toSnakeCase(localName(element))}`,
      `${localName(element)} is outside the v1 XML subset.`
    ));
  }
}

function xmlTaskPath(task, index) {
  return `tasks[uid=${task.uid ?? `?${index + 1}`}]`;
}

function invalidIssue(rule_id, path, message) {
  return { code: "semantic.invalid", rule_id, path, message };
}

function unsupportedIssue(rule_id, path, message) {
  return { code: "semantic.unsupported", rule_id, path, message };
}

function assertElementNamespace(element, path) {
  if (element.namespaceURI !== PROJECT_NAMESPACE) {
    throw createV1RejectedError({
      code: "xml.profile-unsupported",
      message: `${path} uses a namespace outside the v1 XML subset.`,
      scope: "input",
      path,
      details: { namespace: element.namespaceURI ?? null }
    });
  }
}

function assertAllowedAttributes(element, path) {
  for (const attribute of Array.from(element.attributes ?? [])) {
    if (attribute.namespaceURI !== XMLNS_NAMESPACE && attribute.name !== "xmlns") {
      throw createV1RejectedError({
        code: "xml.profile-unsupported",
        message: `${path} contains an unsupported attribute.`,
        scope: "input",
        path,
        details: { attribute: attribute.name }
      });
    }
  }
}

function localName(element) {
  return element.localName || element.nodeName;
}

function toSnakeCase(value) {
  return value.replace(/([a-z0-9])([A-Z])/g, "$1_$2").toLowerCase();
}

function createInvalidXmlError(message, details) {
  return createV1RejectedError({
    code: "xml.invalid",
    message,
    scope: "input",
    path: "Project",
    details
  });
}
