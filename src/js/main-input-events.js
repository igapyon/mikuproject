/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const mainInputEvents = {
        bind(deps) {
            deps.getImportFileInput().addEventListener("click", (event) => {
                const input = event.target;
                if (input) {
                    input.value = "";
                }
            });
            deps.getImportFileInput().addEventListener("change", async (event) => {
                const input = event.target;
                const file = (input === null || input === void 0 ? void 0 : input.files) && input.files[0];
                if (file) {
                    deps.setStatus(`${file.name} を読み込んでいます...`);
                }
                try {
                    await deps.importFromFile(file);
                }
                catch (error) {
                    console.error("[mikuproject] file import failed", error);
                    deps.setStatus(error instanceof Error ? error.message : "ファイル読込に失敗しました");
                }
                finally {
                    if (input) {
                        input.value = "";
                    }
                }
            });
            deps.getXmlInput().addEventListener("input", () => {
                deps.markXmlSourceDirty();
                deps.refreshXmlSaveState();
            });
        }
    };
    globalThis.__mikuprojectMainInputEvents = mainInputEvents;
})();
