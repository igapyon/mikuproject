/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  type MainInputEventsDeps = {
    setStatus: (message: string) => void;
    refreshXmlSaveState: () => void;
    markXmlSourceDirty: () => void;
    getImportFileInput: () => HTMLInputElement;
    getXmlInput: () => HTMLTextAreaElement;
    importFromFile: (file: File | null | undefined) => Promise<void>;
  };

  const mainInputEvents = {
    bind(deps: MainInputEventsDeps): void {
      deps.getImportFileInput().addEventListener("click", (event) => {
        const input = event.target as HTMLInputElement | null;
        if (input) {
          input.value = "";
        }
      });

      deps.getImportFileInput().addEventListener("change", async (event) => {
        const input = event.target as HTMLInputElement | null;
        const file = input?.files && input.files[0];
        if (file) {
          deps.setStatus(`${file.name} を読み込んでいます...`);
        }
        try {
          await deps.importFromFile(file);
        } catch (error) {
          console.error("[mikuproject] file import failed", error);
          deps.setStatus(error instanceof Error ? error.message : "ファイル読込に失敗しました");
        } finally {
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

  (globalThis as typeof globalThis & {
    __mikuprojectMainInputEvents?: typeof mainInputEvents;
  }).__mikuprojectMainInputEvents = mainInputEvents;
})();
