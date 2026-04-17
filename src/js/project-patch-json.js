/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const projectPatchJsonCore = globalThis.__mikuprojectProjectPatchJsonCore;
    if (!projectPatchJsonCore) {
        throw new Error("mikuproject Project Patch JSON core module is not loaded");
    }
    globalThis.__mikuprojectProjectPatchJson = projectPatchJsonCore;
})();
