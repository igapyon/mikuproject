/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const mainArchiveActions = {
        buildOutputArchive(input) {
            const entries = input.buildOutputArchiveEntries();
            const stamp = input.formatTimestampCompact(new Date());
            return {
                fileName: `mikuproject-all-${stamp}.zip`,
                zipBytes: input.packZipEntries(entries),
                entryCount: entries.length
            };
        },
        downloadAllOutputs(input) {
            const archive = input.buildOutputArchive();
            input.downloadBlob(new Blob([archive.zipBytes], { type: "application/zip" }), archive.fileName);
            input.setStatus(`All 出力を保存しました (${archive.entryCount} 件, ZIP)`);
            input.showToast("All を保存しました");
            input.setActiveTab();
        }
    };
    globalThis.__mikuprojectMainArchiveActions = mainArchiveActions;
})();
