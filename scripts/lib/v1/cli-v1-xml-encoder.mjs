import { canonicalJsonText, canonicalizeSemanticState, compareUnicodeScalars, sha256RawBytes } from "./cli-v1-canonical-json.mjs";
import { decodeMsProjectXmlSubset, MS_PROJECT_XML_ADAPTER, MS_PROJECT_XML_SUBSET_PROFILE } from "./cli-v1-xml-adapter.mjs";
import { validateV1SemanticState } from "./cli-v1-semantic-validator.mjs";

/**
 * Canonically encodes the approved MS Project XML subset.  This does not use
 * the legacy ProjectModel codec: the v1 semantic state is the sole source,
 * and the result is immediately re-decoded by preflight callers.
 */
export function encodeMsProjectXmlSubset(state) {
  const canonicalState = canonicalizeSemanticState(state);
  const validation = validateV1SemanticState(canonicalState);
  if (!validation.valid) {
    throw new TypeError("MS Project XML v1 encoding requires a valid semantic state");
  }
  const lines = [
    '<?xml version="1.0" encoding="UTF-8"?>',
    '<Project xmlns="http://schemas.microsoft.com/project">'
  ];
  appendProject(lines, canonicalState);
  lines.push("</Project>", "");
  const bytes = Buffer.from(lines.join("\n"), "utf8");
  return Object.freeze({
    format_profile: MS_PROJECT_XML_SUBSET_PROFILE,
    adapter: MS_PROJECT_XML_ADAPTER,
    bytes,
    raw_digest: sha256RawBytes(bytes),
    normalizations: Object.freeze([])
  });
}

/**
 * Converts semantic encode/redecode equivalence into a boolean suitable for
 * the C1 preflight boundary.  The decoded adapter may choose a different
 * collection order, so compare canonical semantic JSON rather than bytes.
 */
export function isV1XmlSemanticRoundTripEquivalent(state, encodedBytes) {
  try {
    const decoded = decodeMsProjectXmlSubset(encodedBytes);
    const validation = validateV1SemanticState(decoded.state, { adapterIssues: decoded.adapter_issues });
    return validation.valid
      && canonicalJsonText(canonicalizeSemanticState(state)) === canonicalJsonText(canonicalizeSemanticState(decoded.state));
  } catch {
    return false;
  }
}

function appendProject(lines, state) {
  const project = state.project;
  appendElement(lines, 1, "Name", project.name);
  appendOptionalBoolean(lines, 1, "ScheduleFromStart", project.schedule_from_start);
  appendElement(lines, 1, "StartDate", project.start);
  appendElement(lines, 1, "FinishDate", project.finish);
  appendOptional(lines, 1, "CalendarUID", project.calendar_uid);
  appendOptional(lines, 1, "CurrentDate", project.current_date);
  appendCalendars(lines, state.calendars);
  appendTasks(lines, state.tasks, state.dependencies);
  appendResources(lines, state.resources);
  appendAssignments(lines, state.assignments);
}

function appendCalendars(lines, calendars) {
  if (calendars.length === 0) return;
  lines.push(`${indent(1)}<Calendars>`);
  for (const calendar of sortByNumericUid(calendars)) {
    lines.push(`${indent(2)}<Calendar>`);
    appendElement(lines, 3, "UID", calendar.uid);
    appendOptional(lines, 3, "Name", calendar.name);
    appendOptionalBoolean(lines, 3, "IsBaseCalendar", calendar.is_base_calendar);
    lines.push(`${indent(2)}</Calendar>`);
  }
  lines.push(`${indent(1)}</Calendars>`);
}

function appendTasks(lines, tasks, dependencies) {
  if (tasks.length === 0) return;
  const dependenciesBySuccessor = new Map();
  for (const dependency of dependencies) {
    const current = dependenciesBySuccessor.get(dependency.successor_uid) ?? [];
    current.push(dependency);
    dependenciesBySuccessor.set(dependency.successor_uid, current);
  }
  lines.push(`${indent(1)}<Tasks>`);
  const outline = deriveTaskOutline(tasks);
  for (const [index, task] of tasks.entries()) {
    lines.push(`${indent(2)}<Task>`);
    appendElement(lines, 3, "UID", task.uid);
    appendElement(lines, 3, "ID", String(index + 1));
    appendElement(lines, 3, "Name", task.name);
    appendElement(lines, 3, "OutlineLevel", String(outline[index].level));
    appendElement(lines, 3, "OutlineNumber", outline[index].number);
    appendElement(lines, 3, "Start", task.start);
    appendElement(lines, 3, "Finish", task.finish);
    appendElement(lines, 3, "Duration", task.duration);
    appendElement(lines, 3, "Milestone", boolText(task.milestone));
    appendElement(lines, 3, "Summary", boolText(task.summary));
    appendElement(lines, 3, "PercentComplete", String(task.percent_complete));
    appendOptional(lines, 3, "CalendarUID", task.calendar_uid);
    for (const dependency of sortDependencies(dependenciesBySuccessor.get(task.uid) ?? [])) {
      lines.push(`${indent(3)}<PredecessorLink>`);
      appendElement(lines, 4, "PredecessorUID", dependency.predecessor_uid);
      appendElement(lines, 4, "Type", "1");
      appendElement(lines, 4, "LinkLag", "0");
      appendElement(lines, 4, "LagFormat", "3");
      lines.push(`${indent(3)}</PredecessorLink>`);
    }
    lines.push(`${indent(2)}</Task>`);
  }
  lines.push(`${indent(1)}</Tasks>`);
}

function appendResources(lines, resources) {
  if (resources.length === 0) return;
  lines.push(`${indent(1)}<Resources>`);
  for (const [index, resource] of sortByNumericUid(resources).entries()) {
    lines.push(`${indent(2)}<Resource>`);
    appendElement(lines, 3, "UID", resource.uid);
    appendElement(lines, 3, "ID", String(index + 1));
    appendOptional(lines, 3, "Name", resource.name);
    if (Object.hasOwn(resource, "type")) appendElement(lines, 3, "Type", resourceTypeCode(resource.type));
    appendOptional(lines, 3, "CalendarUID", resource.calendar_uid);
    lines.push(`${indent(2)}</Resource>`);
  }
  lines.push(`${indent(1)}</Resources>`);
}

function appendAssignments(lines, assignments) {
  if (assignments.length === 0) return;
  lines.push(`${indent(1)}<Assignments>`);
  for (const assignment of sortByNumericUid(assignments)) {
    lines.push(`${indent(2)}<Assignment>`);
    appendElement(lines, 3, "UID", assignment.uid);
    appendElement(lines, 3, "TaskUID", assignment.task_uid);
    appendOptional(lines, 3, "ResourceUID", assignment.resource_uid);
    appendOptional(lines, 3, "Start", assignment.start);
    appendOptional(lines, 3, "Finish", assignment.finish);
    appendOptional(lines, 3, "Units", assignment.units);
    appendOptional(lines, 3, "Work", assignment.work);
    lines.push(`${indent(2)}</Assignment>`);
  }
  lines.push(`${indent(1)}</Assignments>`);
}

function deriveTaskOutline(tasks) {
  const activeAncestors = [];
  const siblingPositions = [];
  return tasks.map((task) => {
    let level;
    if (task.parent_uid === null) {
      level = 1;
      activeAncestors.length = 0;
    } else {
      const parentIndex = activeAncestors.lastIndexOf(task.parent_uid);
      if (parentIndex === -1) throw new TypeError("cannot derive outline from a non-preorder semantic task state");
      level = parentIndex + 2;
      activeAncestors.length = level - 1;
    }
    siblingPositions.length = level;
    siblingPositions[level - 1] = (siblingPositions[level - 1] ?? 0) + 1;
    activeAncestors.push(task.uid);
    return { level, number: siblingPositions.join(".") };
  });
}

function sortByNumericUid(items) {
  return [...items].sort((left, right) => compareNumericIdentity(left.uid, right.uid));
}

function sortDependencies(dependencies) {
  return [...dependencies].sort((left, right) => compareNumericIdentity(left.predecessor_uid, right.predecessor_uid)
    || compareNumericIdentity(left.successor_uid, right.successor_uid)
    || compareUnicodeScalars(left.type, right.type)
    || compareUnicodeScalars(left.lag, right.lag));
}

function compareNumericIdentity(left, right) {
  const leftNumber = BigInt(left);
  const rightNumber = BigInt(right);
  if (leftNumber === rightNumber) return 0;
  return leftNumber < rightNumber ? -1 : 1;
}

function resourceTypeCode(type) {
  return ({ material: "0", work: "1", cost: "2" })[type];
}

function appendOptional(lines, level, name, value) {
  if (value !== undefined) appendElement(lines, level, name, value);
}

function appendOptionalBoolean(lines, level, name, value) {
  if (value !== undefined) appendElement(lines, level, name, boolText(value));
}

function appendElement(lines, level, name, value) {
  lines.push(`${indent(level)}<${name}>${escapeXmlText(String(value))}</${name}>`);
}

function boolText(value) {
  return value ? "1" : "0";
}

function escapeXmlText(value) {
  return value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

function indent(level) {
  return "  ".repeat(level);
}
