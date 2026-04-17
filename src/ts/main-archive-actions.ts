/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  const mainArchiveActions = {
    buildOutputArchive(input: {
      buildOutputArchiveEntries: () => Array<{ name: string; data: Uint8Array }>;
      formatTimestampCompact: (date: Date) => string;
      packZipEntries: (entries: Array<{ name: string; data: Uint8Array }>) => Uint8Array;
    }): { fileName: string; zipBytes: Uint8Array; entryCount: number } {
      const entries = input.buildOutputArchiveEntries();
      const stamp = input.formatTimestampCompact(new Date());
      return {
        fileName: `mikuproject-all-${stamp}.zip`,
        zipBytes: input.packZipEntries(entries),
        entryCount: entries.length
      };
    },

    downloadAllOutputs(input: {
      buildOutputArchive: () => { fileName: string; zipBytes: Uint8Array; entryCount: number };
      downloadBlob: (blob: Blob, filename: string) => void;
      setStatus: (message: string) => void;
      showToast: (message: string) => void;
      setActiveTab: () => void;
    }): void {
      const archive = input.buildOutputArchive();
      input.downloadBlob(
        new Blob([archive.zipBytes], { type: "application/zip" }),
        archive.fileName
      );
      input.setStatus(`All 出力を保存しました (${archive.entryCount} 件, ZIP)`);
      input.showToast("All を保存しました");
      input.setActiveTab();
    }
  };

  (globalThis as typeof globalThis & {
    __mikuprojectMainArchiveActions?: typeof mainArchiveActions;
  }).__mikuprojectMainArchiveActions = mainArchiveActions;
})();
