(() => {
  function normalizeMermaidText(value: string | undefined, fallback: string): string {
    const text = String(value || fallback).replace(/[:：#,，]/g, " ").replace(/\s+/g, " ").trim();
    return text || fallback;
  }

  function normalizeMermaidGanttLabel(
    value: string | undefined,
    fallback: string,
    leadingPrefix: string
  ): string {
    const text = normalizeMermaidText(value, fallback);
    return /^\d/.test(text) ? `${leadingPrefix} ${text}` : text;
  }

  function normalizeMermaidTaskId(value: string | undefined, fallback: string): string {
    return String(value || fallback).replace(/[^A-Za-z0-9_]/g, "_");
  }

  function toMermaidDuration(duration: string | undefined): string | null {
    const text = String(duration || "").trim();
    if (!text) {
      return null;
    }
    const match = /^P(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)$/.exec(text);
    if (!match) {
      return null;
    }
    const hours = Number(match[1] || 0);
    const minutes = Number(match[2] || 0);
    const seconds = Number(match[3] || 0);
    const parts: string[] = [];
    if (hours > 0) {
      parts.push(`${hours}h`);
    }
    if (minutes > 0) {
      parts.push(`${minutes}m`);
    }
    if (seconds > 0) {
      parts.push(`${seconds}s`);
    }
    return parts.length > 0 ? parts.join(" ") : null;
  }

  function formatMermaidLag(duration: string | undefined): string {
    const short = toMermaidDuration(duration);
    if (short) {
      return short;
    }
    return String(duration || "").trim();
  }

  function isZeroDuration(duration: string | undefined): boolean {
    const text = String(duration || "").trim();
    return text === "" || text === "PT0H0M0S" || text === "PT0M0S" || text === "PT0S";
  }

  function buildTaskSectionMap(
    tasks: TaskModel[],
    projectName: string
  ): Map<string, string> {
    const sectionMap = new Map<string, string>();
    const summaryStack: TaskModel[] = [];
    for (const task of tasks) {
      while (summaryStack.length > 0 && task.outlineLevel <= summaryStack[summaryStack.length - 1].outlineLevel) {
        summaryStack.pop();
      }
      if (task.summary) {
        summaryStack.push(task);
        continue;
      }
      const sectionName = summaryStack.length > 0
        ? normalizeMermaidGanttLabel(summaryStack[summaryStack.length - 1].name, "Summary", "Section")
        : normalizeMermaidGanttLabel(projectName, "Tasks", "Section");
      sectionMap.set(task.uid, sectionName);
    }
    return sectionMap;
  }

  globalThis.__mikuprojectMsprojectMermaid = {
    exportMermaidGantt(input: {
      model: ProjectModel;
      describePredecessorType: (type: number | undefined) => string;
    }): string {
      const { model, describePredecessorType } = input;
      const lines: string[] = [
        "gantt",
        `  title ${normalizeMermaidGanttLabel(model.project.name, "Project", "Project")}`,
        "  dateFormat YYYY-MM-DDTHH:mm:ss",
        "  axisFormat %m/%d"
      ];
      const sectionMap = buildTaskSectionMap(model.tasks, model.project.name);
      const taskNameMap = new Map(
        model.tasks.map((task) => [
          task.uid,
          normalizeMermaidGanttLabel(task.name, `Task ${task.uid}`, "Task")
        ])
      );
      const exportedTasks = model.tasks.filter((task) => !task.summary && task.start && task.finish);
      let currentSection = "";
      for (const task of exportedTasks) {
        const section = sectionMap.get(task.uid) || "Tasks";
        if (section !== currentSection) {
          currentSection = section;
          lines.push(`  section ${section}`);
        }
        const tags: string[] = [];
        if (task.critical) {
          tags.push("crit");
        }
        if (task.milestone) {
          tags.push("milestone");
        } else if (task.percentComplete >= 100) {
          tags.push("done");
        } else if (task.percentComplete > 0) {
          tags.push("active");
        }
        const taskId = `task_${normalizeMermaidTaskId(task.uid || task.id || "x", "x")}`;
        const singlePredecessor = task.predecessors.length === 1 ? task.predecessors[0] : null;
        const nativeDependencyTarget = singlePredecessor
          ? `task_${normalizeMermaidTaskId(singlePredecessor.predecessorUid, "x")}`
          : null;
        const nativeDuration = !task.milestone ? toMermaidDuration(task.duration) : null;
        const useNativeDependency = Boolean(
          singlePredecessor
          && nativeDependencyTarget
          && nativeDuration
          && isZeroDuration(singlePredecessor.linkLag)
          && (singlePredecessor.type === undefined || singlePredecessor.type === 1)
        );
        const fields = useNativeDependency
          ? [...tags, taskId, `after ${nativeDependencyTarget}`, nativeDuration]
          : [...tags, taskId, task.start, task.finish].filter(Boolean);
        lines.push(`  ${normalizeMermaidGanttLabel(task.name, `Task ${task.uid}`, "Task")} :${fields.join(", ")}`);
        for (const predecessor of task.predecessors) {
          const predecessorTaskId = `task_${normalizeMermaidTaskId(predecessor.predecessorUid, "x")}`;
          const predecessorName = taskNameMap.get(predecessor.predecessorUid) || `Task ${predecessor.predecessorUid}`;
          if (useNativeDependency && predecessorTaskId === nativeDependencyTarget) {
            lines.push(`  %% dependency(native): ${task.name || taskId} after ${predecessorName} (${taskId} after ${predecessorTaskId})`);
            continue;
          }
          const details = [
            `type=${describePredecessorType(predecessor.type)}`,
            !isZeroDuration(predecessor.linkLag) ? `lag=${formatMermaidLag(predecessor.linkLag)}` : ""
          ].filter(Boolean).join(", ");
          lines.push(`  %% dependency: ${task.name || taskId} after ${predecessorName}${details ? ` (${details})` : ""} [${taskId} after ${predecessorTaskId}]`);
          if (!isZeroDuration(predecessor.linkLag)) {
            lines.push(`  %% dependency(pseudo): ${task.name || taskId} ~= after ${predecessorName} + ${formatMermaidLag(predecessor.linkLag)}`);
          }
        }
        if (task.predecessors.length > 1) {
          lines.push(`  %% dependency(note): ${task.name || taskId} has multiple predecessors`);
        } else if (singlePredecessor && !useNativeDependency) {
          const reasons = [
            !isZeroDuration(singlePredecessor.linkLag) ? `lag=${formatMermaidLag(singlePredecessor.linkLag)}` : "",
            singlePredecessor.type !== undefined && singlePredecessor.type !== 1 ? `type=${describePredecessorType(singlePredecessor.type)}` : "",
            !nativeDuration && !task.milestone ? `duration=${task.duration || "(empty)"}` : ""
          ].filter(Boolean).join(", ");
          if (reasons) {
            lines.push(`  %% dependency(note): ${task.name || taskId} kept as comment because ${reasons}`);
          }
        }
      }
      if (exportedTasks.length === 0) {
        lines.push("  section Tasks");
        lines.push("  No tasks :milestone, empty_0, 1970-01-01T00:00:00, 1970-01-01T00:00:00");
      }
      return `${lines.join("\n")}\n`;
    }
  };
})();
