/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  const mainExport = {
    buildTimestamp(date: Date): string {
      return [
        date.getFullYear(),
        String(date.getMonth() + 1).padStart(2, "0"),
        String(date.getDate()).padStart(2, "0"),
        String(date.getHours()).padStart(2, "0"),
        String(date.getMinutes()).padStart(2, "0")
      ].join("");
    },

    buildDateStamp(date: Date): string {
      return [
        date.getFullYear(),
        String(date.getMonth() + 1).padStart(2, "0"),
        String(date.getDate()).padStart(2, "0")
      ].join("");
    },

    buildCsvExport(input: {
      model: ProjectModel;
      exportCsvParentId: (model: ProjectModel) => string;
    }): { fileName: string; text: string } {
      const stamp = mainExport.buildTimestamp(new Date());
      return {
        fileName: `mikuproject-export-${stamp}.csv`,
        text: `${input.exportCsvParentId(input.model)}\n`
      };
    },

    buildProjectOverviewExport(input: {
      model: ProjectModel;
      exportProjectOverviewView: (model: ProjectModel) => unknown;
    }): { fileName: string; text: string } {
      return {
        fileName: "mikuproject-project-overview-view.editjson",
        text: `${JSON.stringify(input.exportProjectOverviewView(input.model), null, 2)}\n`
      };
    },

    buildTaskEditExport(input: {
      model: ProjectModel;
      requestedTaskUid?: string;
      exportTaskEditView: (model: ProjectModel, taskUid?: string) => unknown;
    }): {
      fileName: string;
      text: string;
      resolvedTaskUid?: string;
    } {
      const view = input.exportTaskEditView(input.model, input.requestedTaskUid) as {
        target_task?: { uid?: string };
      };
      const resolvedTaskUid = view.target_task?.uid;
      const taskSuffix = resolvedTaskUid ? `-${resolvedTaskUid}` : "";
      return {
        fileName: `mikuproject-task-edit-view${taskSuffix}.editjson`,
        text: `${JSON.stringify(view, null, 2)}\n`,
        resolvedTaskUid
      };
    },

    buildAiProjectionBundleExport(input: {
      model: ProjectModel;
      exportProjectOverviewView: (model: ProjectModel) => unknown;
      exportPhaseDetailView: (
        model: ProjectModel,
        phaseUid?: string,
        options?: { mode?: "full" | "scoped"; rootUid?: string; maxDepth?: number }
      ) => unknown;
      exportTaskEditView: (model: ProjectModel, taskUid?: string) => unknown;
    }): { fileName: string; text: string; phaseCount: number; taskCount: number } {
      const projectOverview = input.exportProjectOverviewView(input.model) as {
        phases?: Array<{ uid?: string }>;
      };
      const phaseDetailViewsFull = (projectOverview.phases || [])
        .map((phase) => phase?.uid)
        .filter((uid): uid is string => Boolean(uid))
        .map((phaseUid) => input.exportPhaseDetailView(input.model, phaseUid, { mode: "full" }));
      const taskEditViewsFull = input.model.tasks
        .filter((task) => !(task.uid === "0" || task.summary))
        .map((task) => input.exportTaskEditView(input.model, task.uid));
      const bundle = {
        view_type: "ai_projection_bundle",
        project_overview_view: projectOverview,
        phase_detail_views_full: phaseDetailViewsFull,
        task_edit_views_full: taskEditViewsFull
      };
      return {
        fileName: "mikuproject-full-bundle.editjson",
        text: `${JSON.stringify(bundle, null, 2)}\n`,
        phaseCount: phaseDetailViewsFull.length,
        taskCount: taskEditViewsFull.length
      };
    },

    buildPhaseDetailExport(input: {
      model: ProjectModel;
      mode: "full" | "scoped";
      requestedPhaseUid?: string;
      requestedRootUid?: string;
      requestedMaxDepth?: number;
      exportPhaseDetailView: (
        model: ProjectModel,
        phaseUid?: string,
        options?: { mode?: "full" | "scoped"; rootUid?: string; maxDepth?: number }
      ) => unknown;
    }): {
      fileName: string;
      text: string;
      resolvedPhaseUid?: string;
      resolvedRootUid?: string;
      resolvedMaxDepth?: number;
      resolvedMode: "full" | "scoped";
    } {
      const view = input.exportPhaseDetailView(input.model, input.requestedPhaseUid, {
        mode: input.mode,
        rootUid: input.requestedRootUid,
        maxDepth: input.requestedMaxDepth
      }) as {
        phase?: { uid?: string };
        scope?: { mode?: "full" | "scoped"; root_uid?: string | null; max_depth?: number | null };
      };
      const resolvedPhaseUid = view.phase?.uid;
      const resolvedMode = view.scope?.mode === "scoped" ? "scoped" : "full";
      const resolvedRootUid = view.scope?.root_uid || undefined;
      const resolvedMaxDepth = typeof view.scope?.max_depth === "number" ? view.scope.max_depth : undefined;
      const phaseSuffix = resolvedPhaseUid ? `-${resolvedPhaseUid}` : "";
      const modeSuffix = resolvedMode === "scoped" ? "-scoped" : "-full";
      const rootSuffix = resolvedRootUid ? `-root-${resolvedRootUid}` : "";
      const depthSuffix = typeof resolvedMaxDepth === "number" ? `-depth-${resolvedMaxDepth}` : "";
      return {
        fileName: `mikuproject-phase-detail-view${phaseSuffix}${modeSuffix}${rootSuffix}${depthSuffix}.editjson`,
        text: `${JSON.stringify(view, null, 2)}\n`,
        resolvedPhaseUid,
        resolvedRootUid,
        resolvedMaxDepth,
        resolvedMode
      };
    },

    buildXlsxExport(input: {
      model: ProjectModel;
      createWorkbookCodec: () => { exportWorkbook: (workbook: unknown) => Uint8Array };
      exportProjectWorkbook: (model: ProjectModel) => unknown;
    }): { fileName: string; bytes: Uint8Array } {
      const stamp = mainExport.buildTimestamp(new Date());
      const codec = input.createWorkbookCodec();
      return {
        fileName: `mikuproject-export-${stamp}.xlsx`,
        bytes: codec.exportWorkbook(input.exportProjectWorkbook(input.model))
      };
    },

    buildWorkbookJsonExport(input: {
      model: ProjectModel;
      exportProjectWorkbookJson: (model: ProjectModel) => unknown;
    }): { fileName: string; text: string } {
      const stamp = mainExport.buildTimestamp(new Date());
      return {
        fileName: `mikuproject-workbook-${stamp}.json`,
        text: JSON.stringify(input.exportProjectWorkbookJson(input.model), null, 2)
      };
    },

    buildWbsXlsxExport(input: {
      model: ProjectModel;
      options: {
        holidayDates: string[];
        displayDaysBeforeBaseDate?: number;
        displayDaysAfterBaseDate?: number;
        useBusinessDaysForDisplayRange?: boolean;
        useBusinessDaysForProgressBand?: boolean;
      };
      createWorkbookCodec: () => { exportWorkbook: (workbook: unknown) => Uint8Array };
      exportWbsWorkbook: (
        model: ProjectModel,
        options: {
          holidayDates: string[];
          displayDaysBeforeBaseDate?: number;
          displayDaysAfterBaseDate?: number;
          useBusinessDaysForDisplayRange?: boolean;
          useBusinessDaysForProgressBand?: boolean;
        }
      ) => unknown;
    }): { fileName: string; bytes: Uint8Array } {
      const stamp = mainExport.buildTimestamp(new Date());
      const codec = input.createWorkbookCodec();
      return {
        fileName: `mikuproject-wbs-${stamp}.xlsx`,
        bytes: codec.exportWorkbook(input.exportWbsWorkbook(input.model, input.options))
      };
    },

    buildXmlExport(input: {
      xmlText: string;
    }): { fileName: string; text: string } {
      const stamp = mainExport.buildTimestamp(new Date());
      return {
        fileName: `mikuproject-export-${stamp}.xml`,
        text: `${input.xmlText}\n`
      };
    },

    buildDailySvgExport(input: {
      svg: string;
    }): { fileName: string; text: string } {
      const stamp = mainExport.buildTimestamp(new Date());
      return {
        fileName: `mikuproject-wbs-daily-${stamp}.svg`,
        text: input.svg
      };
    },

    buildWeeklySvgExport(input: {
      svg: string;
    }): { fileName: string; text: string } {
      const stamp = mainExport.buildTimestamp(new Date());
      return {
        fileName: `mikuproject-wbs-weekly-${stamp}.svg`,
        text: input.svg
      };
    },

    buildMonthlySvgZipExport(input: {
      zipBytes: Uint8Array;
    }): { fileName: string; bytes: Uint8Array } {
      const stamp = mainExport.buildTimestamp(new Date());
      return {
        fileName: `mikuproject-monthly-wbs-calendar-${stamp}.zip`,
        bytes: input.zipBytes
      };
    },

    buildMermaidExport(input: {
      mermaidText: string;
    }): { fileName: string; text: string } {
      const stamp = mainExport.buildTimestamp(new Date());
      return {
        fileName: `mikuproject-wbs-mermaid-${stamp}.mmd`,
        text: `${input.mermaidText}\n`
      };
    },

    buildWbsMarkdownExport(input: {
      markdownText: string;
    }): { fileName: string; text: string } {
      const stamp = mainExport.buildDateStamp(new Date());
      return {
        fileName: `mikuproject-wbs-${stamp}.md`,
        text: input.markdownText
      };
    }
  };

  (globalThis as typeof globalThis & {
    __mikuprojectMainExport?: typeof mainExport;
  }).__mikuprojectMainExport = mainExport;
})();
