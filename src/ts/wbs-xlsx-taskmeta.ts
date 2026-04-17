/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  type WbsXlsxTaskmetaHelper = {
    formatTaskLabel(task: TaskModel): string;
    classifyTaskKind(task: TaskModel): string;
  };

  function createWbsXlsxTaskmetaHelper(): WbsXlsxTaskmetaHelper {
    function formatTaskLabel(task: TaskModel): string {
      const prefix = task.summary ? "> " : (task.milestone ? "* " : "- ");
      return `${"  ".repeat(Math.max(0, task.outlineLevel - 1))}${prefix}${task.name}`;
    }

    function classifyTaskKind(task: TaskModel): string {
      if (task.summary) {
        return "フェーズ";
      }
      if (task.milestone) {
        return "マイル";
      }
      return "タスク";
    }

    return {
      formatTaskLabel,
      classifyTaskKind
    };
  }

  (globalThis as typeof globalThis & {
    __mikuprojectWbsXlsxTaskmeta?: {
      createWbsXlsxTaskmetaHelper: () => WbsXlsxTaskmetaHelper;
    };
  }).__mikuprojectWbsXlsxTaskmeta = {
    createWbsXlsxTaskmetaHelper
  };
})();
