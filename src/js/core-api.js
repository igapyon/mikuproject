/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const mikuprojectCoreApiPublic = globalThis.__mikuprojectCoreApiPublic;
    if (!mikuprojectCoreApiPublic) {
        throw new Error("mikuproject core api public module is not loaded");
    }
    globalThis.__mikuprojectCoreApi = mikuprojectCoreApiPublic;
})();
