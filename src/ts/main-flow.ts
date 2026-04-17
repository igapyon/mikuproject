/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  type ImportWarning = {
    message: string;
    scope?: "project" | "tasks" | "resources" | "assignments" | "calendars";
    uid?: string;
    label?: string;
  };

  type ImportChange = {
    scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
    uid: string;
    label: string;
    field: string;
    before: string | number | boolean | undefined;
    after: string | number | boolean;
  };

  const mainFlow = {
    applyModelState(input: {
      model: ProjectModel;
      issues: ValidationIssue[];
      updateSummary: (model: ProjectModel | null) => void;
      renderValidationIssues: (issues: ValidationIssue[]) => void;
      renderImportWarnings: (warnings: ImportWarning[], options?: { sourceLabel?: string }) => void;
      renderXlsxImportSummary: (changes: ImportChange[], options?: {
        sourceLabel?: string;
        warnings?: ImportWarning[];
      }) => void;
      warnings?: ImportWarning[];
      warningSourceLabel?: string;
      changes?: ImportChange[];
      changeSourceLabel?: string;
    }): void {
      input.updateSummary(input.model);
      input.renderValidationIssues(input.issues);
      input.renderImportWarnings(input.warnings || [], input.warningSourceLabel ? { sourceLabel: input.warningSourceLabel } : {});
      input.renderXlsxImportSummary(input.changes || [], input.changeSourceLabel || input.warnings
        ? {
            sourceLabel: input.changeSourceLabel,
            warnings: input.warnings
          }
        : {});
    },

    completeOutput(input: {
      setStatus: (message: string) => void;
      showToast: (message: string) => void;
      setActiveTab: (tabId: "input" | "transform" | "output", options?: { skipTransformRefresh?: boolean }) => void;
      statusMessage: string;
      toastMessage: string;
    }): void {
      input.setStatus(input.statusMessage);
      input.showToast(input.toastMessage);
      input.setActiveTab("output");
    },

    completeTransform(input: {
      setStatus: (message: string) => void;
      showToast: (message: string) => void;
      setActiveTab: (tabId: "input" | "transform" | "output", options?: { skipTransformRefresh?: boolean }) => void;
      statusMessage: string;
      toastMessage: string;
    }): void {
      input.setStatus(input.statusMessage);
      input.showToast(input.toastMessage);
      input.setActiveTab("transform", { skipTransformRefresh: true });
    }
  };

  (globalThis as typeof globalThis & {
    __mikuprojectMainFlow?: typeof mainFlow;
  }).__mikuprojectMainFlow = mainFlow;
})();
