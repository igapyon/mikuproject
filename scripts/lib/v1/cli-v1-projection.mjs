import { canonicalizeSemanticState, canonicalJsonText, sha256SemanticState } from "./cli-v1-canonical-json.mjs";
import { validateArtifact } from "../../generated/cli-v1-schema-validators.mjs";

const PROJECT_OVERVIEW_SCOPE = Object.freeze({
  target_task_uid: null,
  included_domains: Object.freeze(["project", "tasks", "dependencies"]),
  omitted_domains: Object.freeze(["task_edit_details", "raw_external_artifact", "unsupported_data"])
});

const TASK_CHANGE_CONTEXT_INCLUDED_DOMAINS = Object.freeze([
  "project", "target_task", "ancestors", "dependencies", "assignments", "resources"
]);
const TASK_CHANGE_CONTEXT_OMITTED_DOMAINS = Object.freeze([
  "other_task_details", "raw_external_artifact", "unsupported_data"
]);
const SUPPORTED_CHANGE_REQUESTS = Object.freeze([Object.freeze({
  kind: "set_task_percent_complete",
  required_preconditions: Object.freeze(["source_state_digest", "expected_percent_complete"])
})]);

/**
 * Creates the deliberately small, external `project_overview` Projection.
 *
 * The semantic state is first canonicalized so collections whose state order
 * is explicitly non-semantic (notably dependencies) cannot leak host/input
 * ordering into the exchange artifact.  Task order is intentionally retained:
 * it is semantic preorder and is represented again by the zero-based order.
 */
export function createV1ProjectOverviewProjection(state) {
  const canonicalState = canonicalizeSemanticState(state);
  const projection = {
    kind: "miku_project_projection",
    schema_version: "1",
    semantic_contract_version: "1",
    purpose: "project_overview",
    source_state_digest: sha256SemanticState(canonicalState),
    scope: {
      target_task_uid: PROJECT_OVERVIEW_SCOPE.target_task_uid,
      included_domains: [...PROJECT_OVERVIEW_SCOPE.included_domains],
      omitted_domains: [...PROJECT_OVERVIEW_SCOPE.omitted_domains]
    },
    project: copyProjectOverview(canonicalState.project),
    tasks: canonicalState.tasks.map((task, order) => ({
      uid: task.uid,
      name: task.name,
      parent_uid: task.parent_uid,
      order,
      summary: task.summary,
      percent_complete: task.percent_complete
    })),
    dependencies: canonicalState.dependencies.map((dependency) => ({
      predecessor_uid: dependency.predecessor_uid,
      successor_uid: dependency.successor_uid,
      type: dependency.type,
      lag: dependency.lag
    })),
    capability: { unsupported_data: [] }
  };
  if (!validateArtifact(projection)) {
    throw new TypeError("project_overview builder produced an artifact outside miku_project_artifacts/v1");
  }
  return projection;
}

/**
 * Checks the complete R1 portion of RB-012 without accepting a looser view
 * shape.  This is intentionally an exact canonical comparison against the
 * deterministic builder: source digest, fixed scope, all overview fields,
 * task preorder/order, and dependency content are one binding.
 */
export function validateV1ProjectOverviewBinding({ state, projection } = {}) {
  try {
    return canonicalJsonText(projection) === canonicalJsonText(createV1ProjectOverviewProjection(state));
  } catch {
    return false;
  }
}

/**
 * Creates the C1 editing Projection for exactly one valid leaf task.  It is
 * intentionally not a generic task view: all collections are filtered to the
 * task's decision context, and the only advertised request is the approved
 * set_task_percent_complete operation.  The semantic validator owns the
 * broader state validity check; this builder enforces the additional C1
 * target constraint so a summary task never receives an edit-capable view.
 */
export function createV1TaskChangeContextProjection(state, taskUid) {
  const canonicalState = canonicalizeSemanticState(state);
  const targetTask = canonicalState.tasks.find((task) => task.uid === taskUid);
  if (!targetTask) {
    throw new TypeError(`task_change_context target task does not exist: ${String(taskUid)}`);
  }
  if (targetTask.summary !== false) {
    throw new TypeError(`task_change_context target task must be a leaf: ${String(taskUid)}`);
  }

  const tasksByUid = new Map(canonicalState.tasks.map((task) => [task.uid, task]));
  const ancestors = [];
  let parentUid = targetTask.parent_uid;
  while (parentUid !== null) {
    const parent = tasksByUid.get(parentUid);
    if (!parent) {
      throw new TypeError(`task_change_context target has an unresolved ancestor: ${String(parentUid)}`);
    }
    ancestors.unshift(copyArtifactValue(parent));
    parentUid = parent.parent_uid;
  }

  const assignments = canonicalState.assignments
    .filter((assignment) => assignment.task_uid === targetTask.uid)
    .map(copyArtifactValue);
  const referencedResourceUids = new Set(assignments
    .filter((assignment) => Object.hasOwn(assignment, "resource_uid"))
    .map((assignment) => assignment.resource_uid));
  const projection = {
    kind: "miku_project_projection",
    schema_version: "1",
    semantic_contract_version: "1",
    purpose: "task_change_context",
    source_state_digest: sha256SemanticState(canonicalState),
    scope: {
      target_task_uid: targetTask.uid,
      included_domains: [...TASK_CHANGE_CONTEXT_INCLUDED_DOMAINS],
      omitted_domains: [...TASK_CHANGE_CONTEXT_OMITTED_DOMAINS]
    },
    project: copyArtifactValue(canonicalState.project),
    target_task: copyArtifactValue(targetTask),
    ancestors,
    dependencies: canonicalState.dependencies
      .filter((dependency) => dependency.predecessor_uid === targetTask.uid || dependency.successor_uid === targetTask.uid)
      .map(copyArtifactValue),
    resources: canonicalState.resources
      .filter((resource) => referencedResourceUids.has(resource.uid))
      .map(copyArtifactValue),
    assignments,
    capability: { unsupported_data: [] },
    supported_change_requests: SUPPORTED_CHANGE_REQUESTS.map(copyArtifactValue)
  };
  if (!validateArtifact(projection)) {
    throw new TypeError("task_change_context builder produced an artifact outside miku_project_artifacts/v1");
  }
  return projection;
}

/**
 * The task_change_context half of RB-012.  Compare against the deterministic
 * builder rather than accepting a subset comparison, which detects both data
 * leaks and missing decision-context members.
 */
export function validateV1TaskChangeContextBinding({ state, projection } = {}) {
  try {
    return projection?.purpose === "task_change_context"
      && canonicalJsonText(projection) === canonicalJsonText(
        createV1TaskChangeContextProjection(state, projection.scope?.target_task_uid)
      );
  } catch {
    return false;
  }
}

function copyProjectOverview(project) {
  const overview = {
    name: project.name,
    start: project.start,
    finish: project.finish
  };
  for (const key of ["current_date", "schedule_from_start", "calendar_uid"]) {
    if (Object.hasOwn(project, key)) {
      overview[key] = project[key];
    }
  }
  return overview;
}

function copyArtifactValue(value) {
  return JSON.parse(JSON.stringify(value));
}
