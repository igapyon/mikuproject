/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  const mikuprojectXml = (globalThis as typeof globalThis & {
    __mikuprojectXml?: {
      buildProjectDraftRequest: (input: {
        name: string;
        plannedStart?: string;
        goal?: string;
        teamCount?: number;
        mustHavePhases?: string[];
        mustHaveMilestones?: string[];
      }) => unknown;
      importProjectDraftView: (draft: unknown) => ProjectModel;
      exportProjectOverviewView: (model: ProjectModel) => unknown;
      exportTaskEditView: (model: ProjectModel, taskUid?: string) => unknown;
      exportMermaidGantt: (model: ProjectModel) => string;
      exportPhaseDetailView: (
        model: ProjectModel,
        phaseUid?: string,
        options?: {
          mode?: "full" | "scoped";
          rootUid?: string;
          maxDepth?: number;
        }
      ) => unknown;
    };
  }).__mikuprojectXml;

  if (!mikuprojectXml) {
    throw new Error("mikuproject XML module is not loaded");
  }

  (globalThis as typeof globalThis & {
    __mikuprojectCoreApiMsprojectAi?: {
      aiViews: {
        buildProjectDraftRequest: typeof mikuprojectXml.buildProjectDraftRequest;
        importProjectDraftView: typeof mikuprojectXml.importProjectDraftView;
        exportProjectOverviewView: typeof mikuprojectXml.exportProjectOverviewView;
        exportTaskEditView: typeof mikuprojectXml.exportTaskEditView;
        exportPhaseDetailView: typeof mikuprojectXml.exportPhaseDetailView;
      };
      mermaid: {
        exportGantt: typeof mikuprojectXml.exportMermaidGantt;
      };
    };
  }).__mikuprojectCoreApiMsprojectAi = {
    aiViews: {
      buildProjectDraftRequest: mikuprojectXml.buildProjectDraftRequest,
      importProjectDraftView: mikuprojectXml.importProjectDraftView,
      exportProjectOverviewView: mikuprojectXml.exportProjectOverviewView,
      exportTaskEditView: mikuprojectXml.exportTaskEditView,
      exportPhaseDetailView: mikuprojectXml.exportPhaseDetailView
    },
    mermaid: {
      exportGantt: mikuprojectXml.exportMermaidGantt
    }
  };
})();
