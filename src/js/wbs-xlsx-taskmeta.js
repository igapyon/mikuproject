/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    function createWbsXlsxTaskmetaHelper() {
        function formatTaskLabel(task) {
            const prefix = task.summary ? "> " : (task.milestone ? "* " : "- ");
            return `${"  ".repeat(Math.max(0, task.outlineLevel - 1))}${prefix}${task.name}`;
        }
        function classifyTaskKind(task) {
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
    globalThis.__mikuprojectWbsXlsxTaskmeta = {
        createWbsXlsxTaskmetaHelper
    };
})();
