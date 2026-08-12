import { createV1RejectedError } from "./cli-v1-errors.mjs";
import { compareUnicodeScalars } from "./cli-v1-canonical-json.mjs";

/**
 * Applies the G1 semantic contract to a decoded v1 state.  This layer does
 * not inspect XML syntax or vocabulary; adapter findings are supplied as
 * structured input so unsupported external data remains fail-closed.
 */
export function validateV1SemanticState(state, { adapterIssues = [] } = {}) {
  const issues = [];
  appendAdapterIssues(issues, adapterIssues);

  if (!isPlainObject(state)) {
    issues.push(invalid("S-I008", "$", "Semantic state must be an object."));
    return freezeValidation(issues);
  }
  validateEnvelope(state, issues);
  validateProject(state.project, issues);

  const tasks = Array.isArray(state.tasks) ? state.tasks : [];
  if (!Array.isArray(state.tasks)) {
    issues.push(invalid("S-I008", "tasks", "Semantic state requires a task array."));
  }
  const taskContext = validateTasks(tasks, issues);

  const dependencies = Array.isArray(state.dependencies) ? state.dependencies : [];
  if (!Array.isArray(state.dependencies)) {
    issues.push(invalid("S-I008", "dependencies", "Semantic state requires a dependency array."));
  }
  validateDependencies(dependencies, taskContext.uids, issues);

  const resources = Array.isArray(state.resources) ? state.resources : [];
  if (!Array.isArray(state.resources)) {
    issues.push(invalid("S-I008", "resources", "Semantic state requires a resource array."));
  }
  const resourceUids = validateUidCollection(resources, "resources", "S-I005", issues);

  const assignments = Array.isArray(state.assignments) ? state.assignments : [];
  if (!Array.isArray(state.assignments)) {
    issues.push(invalid("S-I008", "assignments", "Semantic state requires an assignment array."));
  }
  validateAssignments(assignments, taskContext.uids, resourceUids, issues);

  const calendars = Array.isArray(state.calendars) ? state.calendars : [];
  if (!Array.isArray(state.calendars)) {
    issues.push(invalid("S-I008", "calendars", "Semantic state requires a calendar array."));
  }
  const calendarUids = validateCalendars(calendars, issues);
  validateCalendarReferences(state.project, tasks, resources, calendarUids, issues);

  return freezeValidation(issues);
}

/** Converts stable semantic issues to the diagnostic-error form used by P4.4.4. */
export function semanticIssuesToV1Errors(validation) {
  if (!validation || !Array.isArray(validation.issues)) {
    throw new TypeError("semantic validation result with issues is required");
  }
  return validation.issues.map((issue) => createV1RejectedError({
    code: issue.code,
    message: issue.message,
    scope: "semantic",
    path: issue.path,
    ruleId: issue.rule_id,
    details: { rule_id: issue.rule_id }
  }));
}

function validateEnvelope(state, issues) {
  if (state.kind !== "miku_project_semantic_state") {
    issues.push(invalid("S-I008", "kind", "Semantic state kind must be miku_project_semantic_state."));
  }
  if (state.schema_version !== "1" || state.semantic_contract_version !== "1") {
    issues.push(invalid("S-I008", "schema_version", "Semantic state must use contract version 1."));
  }
}

function validateProject(project, issues) {
  if (!isPlainObject(project)) {
    issues.push(invalid("S-I008", "project", "Project data is required."));
    return;
  }
  validateRequiredText(project, "name", "project", issues);
  validateRequiredDateTime(project, "start", "project", issues);
  validateRequiredDateTime(project, "finish", "project", issues);
  if (isDateTime(project.start) && isDateTime(project.finish) && project.start > project.finish) {
    issues.push(invalid("S-I006", "project.start", "Project start must not be after project finish."));
  }
  validateOptionalDateTime(project, "current_date", "project", issues);
  validateOptionalBoolean(project, "schedule_from_start", "project", issues);
  validateOptionalIdentity(project, "calendar_uid", "project", issues);
}

function validateTasks(tasks, issues) {
  const uids = new Set();
  const taskByUid = new Map();
  const activeAncestors = [];
  const childCount = new Map();
  const parents = new Map();

  for (const [index, task] of tasks.entries()) {
    const path = semanticTaskPath(task, index);
    if (!isPlainObject(task)) {
      issues.push(invalid("S-I008", `tasks[${index}]`, "Every task must be an object."));
      continue;
    }
    const uid = task.uid;
    if (!isIdentity(uid)) {
      issues.push(invalid("S-I001", `${path}.uid`, "Task UID must be a non-empty v1 identity token."));
    } else if (uids.has(uid)) {
      issues.push(invalid("S-I001", `${path}.uid`, "Task UID must be unique."));
    } else {
      uids.add(uid);
      taskByUid.set(uid, task);
      childCount.set(uid, 0);
    }
    validateRequiredText(task, "name", path, issues);
    validateRequiredDateTime(task, "start", path, issues);
    validateRequiredDateTime(task, "finish", path, issues);
    if (isDateTime(task.start) && isDateTime(task.finish) && task.start > task.finish) {
      issues.push(invalid("S-I007", `${path}.start`, "Task start must not be after task finish."));
    }
    validateDuration(task, "duration", path, issues, true);
    validateRequiredBoolean(task, "milestone", path, issues);
    validateRequiredBoolean(task, "summary", path, issues);
    validatePercent(task, path, issues);
    validateOptionalIdentity(task, "calendar_uid", path, issues);

    if (!Object.hasOwn(task, "parent_uid")) {
      issues.push(invalid("S-I003", `${path}.parent_uid`, "Every task must declare a parent relation."));
      continue;
    }
    if (task.parent_uid === null) {
      activeAncestors.length = 0;
    } else if (!isIdentity(task.parent_uid)) {
      issues.push(invalid("S-I003", `${path}.parent_uid`, "Task parent must be null or a task UID."));
      activeAncestors.length = 0;
    } else if (task.parent_uid === uid) {
      issues.push(invalid("S-I003", `${path}.parent_uid`, "A task cannot be its own parent."));
      activeAncestors.length = 0;
    } else {
      const parentIndex = activeAncestors.lastIndexOf(task.parent_uid);
      if (parentIndex === -1 || !taskByUid.has(task.parent_uid)) {
        issues.push(invalid("S-I003", `${path}.parent_uid`, "Task parent must be a preceding active ancestor."));
        activeAncestors.length = 0;
      } else {
        activeAncestors.length = parentIndex + 1;
        childCount.set(task.parent_uid, (childCount.get(task.parent_uid) ?? 0) + 1);
      }
    }
    if (isIdentity(uid)) {
      activeAncestors.push(uid);
      parents.set(uid, task.parent_uid);
    }
  }

  for (const [uid, task] of taskByUid.entries()) {
    const path = semanticTaskPath(task, tasks.indexOf(task));
    const hasChildren = (childCount.get(uid) ?? 0) > 0;
    if (typeof task.summary === "boolean" && task.summary !== hasChildren) {
      issues.push(invalid("S-I004", `${path}.summary`, "Task summary must exactly match whether it has children."));
    }
    if (task.milestone === true && (task.start !== task.finish || task.duration !== "PT0H0M0S")) {
      issues.push(invalid("S-I010", path, "A milestone must have equal start/finish and zero duration."));
    }
  }
  return { uids, taskByUid, parents };
}

function validateDependencies(dependencies, taskUids, issues) {
  const tuples = new Set();
  const adjacency = new Map();
  for (const [index, dependency] of dependencies.entries()) {
    const path = dependencyPath(dependency, index);
    if (!isPlainObject(dependency)) {
      issues.push(invalid("S-I008", `dependencies[${index}]`, "Every dependency must be an object."));
      continue;
    }
    const predecessor = dependency.predecessor_uid;
    const successor = dependency.successor_uid;
    if (!isIdentity(predecessor) || !isIdentity(successor)) {
      issues.push(invalid("S-I008", path, "Dependency endpoints are required."));
      continue;
    }
    if (dependency.type !== "FS" || dependency.lag !== "PT0H0M0S") {
      issues.push(unsupported("S-I019", path, "Only FS dependencies with zero lag are supported."));
    }
    if (!taskUids.has(predecessor) || !taskUids.has(successor)) {
      issues.push(invalid("S-I014", path, "Dependency endpoints must reference existing tasks."));
    }
    if (predecessor === successor) {
      issues.push(invalid("S-I015", path, "A dependency cannot reference the same task twice."));
    }
    const tuple = `${predecessor}\u0000${successor}\u0000${String(dependency.type)}\u0000${String(dependency.lag)}`;
    if (tuples.has(tuple)) {
      issues.push(invalid("S-I025", path, "Duplicate dependency tuples are not allowed."));
    }
    tuples.add(tuple);
    if (taskUids.has(predecessor) && taskUids.has(successor) && predecessor !== successor) {
      const successors = adjacency.get(predecessor) ?? new Set();
      successors.add(successor);
      adjacency.set(predecessor, successors);
    }
  }
  if (hasCycle(adjacency)) {
    issues.push(invalid("S-I016", "dependencies", "Dependency graph must be acyclic."));
  }
}

function validateUidCollection(collection, domain, ruleId, issues) {
  const uids = new Set();
  for (const [index, member] of collection.entries()) {
    const path = `${domain}[${index}]`;
    if (!isPlainObject(member) || !isIdentity(member.uid)) {
      issues.push(invalid(ruleId, `${path}.uid`, `${domain} UID is required.`));
      continue;
    }
    if (uids.has(member.uid)) {
      issues.push(invalid(ruleId, `${domain}[uid=${member.uid}].uid`, `${domain} UID must be unique.`));
      continue;
    }
    uids.add(member.uid);
  }
  return uids;
}

function validateAssignments(assignments, taskUids, resourceUids, issues) {
  const assignmentUids = validateUidCollection(assignments, "assignments", "S-I005", issues);
  void assignmentUids;
  for (const [index, assignment] of assignments.entries()) {
    const path = assignmentPath(assignment, index);
    if (!isPlainObject(assignment)) {
      continue;
    }
    if (!isIdentity(assignment.task_uid)) {
      issues.push(invalid("S-I008", `${path}.task_uid`, "Assignment task UID is required."));
    } else if (!taskUids.has(assignment.task_uid)) {
      issues.push(invalid("S-I017", `${path}.task_uid`, "Assignment must reference an existing task."));
    }
    if (Object.hasOwn(assignment, "resource_uid")) {
      if (!isIdentity(assignment.resource_uid) || !resourceUids.has(assignment.resource_uid)) {
        issues.push(invalid("S-I017", `${path}.resource_uid`, "Assignment resource must reference an existing resource."));
      }
    }
    validateOptionalDateTime(assignment, "start", path, issues);
    validateOptionalDateTime(assignment, "finish", path, issues);
    if (isDateTime(assignment.start) && isDateTime(assignment.finish) && assignment.start > assignment.finish) {
      issues.push(invalid("S-I017", `${path}.start`, "Assignment start must not be after assignment finish."));
    }
    if (Object.hasOwn(assignment, "units") && !isUnits(assignment.units)) {
      issues.push(invalid("S-I024", `${path}.units`, "Assignment units must be a non-negative decimal string."));
    }
    validateDuration(assignment, "work", path, issues, false);
  }
}

function validateCalendars(calendars, issues) {
  const uids = validateUidCollection(calendars, "calendars", "S-I005", issues);
  for (const [index, calendar] of calendars.entries()) {
    const path = calendarPath(calendar, index);
    if (!isPlainObject(calendar)) {
      continue;
    }
    validateOptionalText(calendar, "name", path, issues);
    validateOptionalBoolean(calendar, "is_base_calendar", path, issues);
  }
  return uids;
}

function validateCalendarReferences(project, tasks, resources, calendarUids, issues) {
  validateCalendarReference(project, "calendar_uid", "project", calendarUids, issues);
  for (const [index, task] of tasks.entries()) {
    validateCalendarReference(task, "calendar_uid", semanticTaskPath(task, index), calendarUids, issues);
  }
  for (const [index, resource] of resources.entries()) {
    validateCalendarReference(resource, "calendar_uid", resourcePath(resource, index), calendarUids, issues);
    if (isPlainObject(resource)) {
      validateOptionalText(resource, "name", resourcePath(resource, index), issues);
      if (Object.hasOwn(resource, "type") && !["work", "material", "cost"].includes(resource.type)) {
        issues.push(invalid("S-I024", `${resourcePath(resource, index)}.type`, "Resource type must be work, material, or cost."));
      }
    }
  }
}

function validateCalendarReference(entity, property, path, calendarUids, issues) {
  if (!isPlainObject(entity) || !Object.hasOwn(entity, property)) {
    return;
  }
  if (!isIdentity(entity[property]) || !calendarUids.has(entity[property])) {
    issues.push(invalid("S-I018", `${path}.${property}`, "Calendar reference must identify an existing calendar."));
  }
}

function validateRequiredText(entity, property, path, issues) {
  if (!Object.hasOwn(entity, property)) {
    issues.push(invalid("S-I008", `${path}.${property}`, `${property} is required.`));
  } else if (!isText(entity[property])) {
    issues.push(invalid("S-I024", `${path}.${property}`, `${property} must be non-empty text.`));
  }
}

function validateOptionalText(entity, property, path, issues) {
  if (Object.hasOwn(entity, property) && !isText(entity[property])) {
    issues.push(invalid("S-I024", `${path}.${property}`, `${property} must be non-empty text when present.`));
  }
}

function validateRequiredDateTime(entity, property, path, issues) {
  if (!Object.hasOwn(entity, property)) {
    issues.push(invalid("S-I008", `${path}.${property}`, `${property} is required.`));
  } else if (!isDateTime(entity[property])) {
    issues.push(invalid("S-I024", `${path}.${property}`, `${property} must be a valid local civil datetime.`));
  }
}

function validateOptionalDateTime(entity, property, path, issues) {
  if (Object.hasOwn(entity, property) && !isDateTime(entity[property])) {
    issues.push(invalid("S-I024", `${path}.${property}`, `${property} must be a valid local civil datetime when present.`));
  }
}

function validateRequiredBoolean(entity, property, path, issues) {
  if (!Object.hasOwn(entity, property)) {
    issues.push(invalid("S-I008", `${path}.${property}`, `${property} is required.`));
  } else if (typeof entity[property] !== "boolean") {
    issues.push(invalid("S-I024", `${path}.${property}`, `${property} must be boolean.`));
  }
}

function validateOptionalBoolean(entity, property, path, issues) {
  if (Object.hasOwn(entity, property) && typeof entity[property] !== "boolean") {
    issues.push(invalid("S-I024", `${path}.${property}`, `${property} must be boolean when present.`));
  }
}

function validateOptionalIdentity(entity, property, path, issues) {
  if (Object.hasOwn(entity, property) && !isIdentity(entity[property])) {
    issues.push(invalid("S-I024", `${path}.${property}`, `${property} must be a non-empty identity token when present.`));
  }
}

function validateDuration(entity, property, path, issues, required) {
  if (!Object.hasOwn(entity, property)) {
    if (required) {
      issues.push(invalid("S-I008", `${path}.${property}`, `${property} is required.`));
    }
    return;
  }
  if (!isDuration(entity[property])) {
    issues.push(invalid("S-I009", `${path}.${property}`, `${property} must be a non-negative working duration.`));
  }
}

function validatePercent(task, path, issues) {
  if (!Object.hasOwn(task, "percent_complete")) {
    issues.push(invalid("S-I008", `${path}.percent_complete`, "percent_complete is required."));
    return;
  }
  const value = task.percent_complete;
  if (!Number.isSafeInteger(value)) {
    issues.push(invalid("S-I013", `${path}.percent_complete`, "percent_complete must be an integer."));
  } else if (value < 0) {
    issues.push(invalid("S-I011", `${path}.percent_complete`, "percent_complete must not be negative."));
  } else if (value > 100) {
    issues.push(invalid("S-I012", `${path}.percent_complete`, "percent_complete must not exceed 100."));
  }
}

function appendAdapterIssues(issues, adapterIssues) {
  if (!Array.isArray(adapterIssues)) {
    throw new TypeError("adapterIssues must be an array");
  }
  for (const issue of adapterIssues) {
    if (!issue || (issue.code !== "semantic.invalid" && issue.code !== "semantic.unsupported") || typeof issue.rule_id !== "string" || typeof issue.path !== "string" || typeof issue.message !== "string") {
      throw new TypeError("adapter issues must use stable semantic codes, rule IDs, paths, and messages");
    }
    issues.push({ ...issue });
  }
}

function freezeValidation(issues) {
  const sorted = issues.sort(compareIssue).map((issue) => Object.freeze({ ...issue }));
  const status = sorted.some((issue) => issue.code === "semantic.unsupported")
    ? "unsupported"
    : sorted.length > 0 ? "invalid" : "valid";
  return Object.freeze({ status, valid: status === "valid", issues: Object.freeze(sorted) });
}

function compareIssue(left, right) {
  return compareUnicodeScalars(left.code, right.code)
    || compareUnicodeScalars(left.path, right.path)
    || compareUnicodeScalars(left.rule_id, right.rule_id);
}

function invalid(rule_id, path, message) {
  return { code: "semantic.invalid", rule_id, path, message };
}

function unsupported(rule_id, path, message) {
  return { code: "semantic.unsupported", rule_id, path, message };
}

function hasCycle(adjacency) {
  const visited = new Set();
  const active = new Set();
  const visit = (uid) => {
    if (active.has(uid)) {
      return true;
    }
    if (visited.has(uid)) {
      return false;
    }
    visited.add(uid);
    active.add(uid);
    for (const successor of adjacency.get(uid) ?? []) {
      if (visit(successor)) {
        return true;
      }
    }
    active.delete(uid);
    return false;
  };
  return [...adjacency.keys()].some(visit);
}

function isPlainObject(value) {
  return value !== null && typeof value === "object" && !Array.isArray(value)
    && (Object.getPrototypeOf(value) === Object.prototype || Object.getPrototypeOf(value) === null);
}

function isIdentity(value) {
  return typeof value === "string" && /^(?:0|[1-9][0-9]*)$/.test(value) && hasOnlyUnicodeScalars(value);
}

function isText(value) {
  return typeof value === "string" && value.length > 0 && hasOnlyUnicodeScalars(value);
}

function isDateTime(value) {
  if (typeof value !== "string") {
    return false;
  }
  const match = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})$/.exec(value);
  if (!match) {
    return false;
  }
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const hour = Number(match[4]);
  const minute = Number(match[5]);
  const second = Number(match[6]);
  return month >= 1 && month <= 12 && day >= 1 && day <= daysInMonth(year, month)
    && hour <= 23 && minute <= 59 && second <= 59;
}

function daysInMonth(year, month) {
  if (month === 2) {
    return year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0) ? 29 : 28;
  }
  return [4, 6, 9, 11].includes(month) ? 30 : 31;
}

function isDuration(value) {
  return typeof value === "string" && /^PT[0-9]+H(?:[0-5]?[0-9])M(?:[0-5]?[0-9])S$/.test(value);
}

function isUnits(value) {
  return typeof value === "string" && /^(?:0|[1-9][0-9]*)(?:\.[0-9]+)?$/.test(value);
}

function hasOnlyUnicodeScalars(value) {
  for (let index = 0; index < value.length; index += 1) {
    const codeUnit = value.charCodeAt(index);
    if (codeUnit >= 0xd800 && codeUnit <= 0xdbff) {
      const next = value.charCodeAt(index + 1);
      if (!(next >= 0xdc00 && next <= 0xdfff)) {
        return false;
      }
      index += 1;
    } else if (codeUnit >= 0xdc00 && codeUnit <= 0xdfff) {
      return false;
    }
  }
  return true;
}

function semanticTaskPath(task, index) {
  return isPlainObject(task) && isIdentity(task.uid) ? `tasks[uid=${task.uid}]` : `tasks[${index}]`;
}

function dependencyPath(dependency, index) {
  return isPlainObject(dependency) && isIdentity(dependency.predecessor_uid) && isIdentity(dependency.successor_uid)
    ? `dependencies[predecessor_uid=${dependency.predecessor_uid},successor_uid=${dependency.successor_uid}]`
    : `dependencies[${index}]`;
}

function resourcePath(resource, index) {
  return isPlainObject(resource) && isIdentity(resource.uid) ? `resources[uid=${resource.uid}]` : `resources[${index}]`;
}

function assignmentPath(assignment, index) {
  return isPlainObject(assignment) && isIdentity(assignment.uid) ? `assignments[uid=${assignment.uid}]` : `assignments[${index}]`;
}

function calendarPath(calendar, index) {
  return isPlainObject(calendar) && isIdentity(calendar.uid) ? `calendars[uid=${calendar.uid}]` : `calendars[${index}]`;
}
