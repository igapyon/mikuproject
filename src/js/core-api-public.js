/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const mikuprojectCoreApiRegistry = globalThis.__mikuprojectCoreApiRegistry;
    if (!mikuprojectCoreApiRegistry) {
        throw new Error("mikuproject core api registry module is not loaded");
    }
    globalThis.__mikuprojectCoreApiPublic = {
        version: 1,
        ...mikuprojectCoreApiRegistry
    };
})();
