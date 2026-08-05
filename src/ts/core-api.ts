/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  const mikuprojectCoreApiPublic = (globalThis as typeof globalThis & {
    __mikuprojectCoreApiPublic?: unknown;
  }).__mikuprojectCoreApiPublic;

  if (!mikuprojectCoreApiPublic) {
    throw new Error("mikuproject core api public module is not loaded");
  }

  const coreApi = mikuprojectCoreApiPublic;
  const globals = globalThis as typeof globalThis & {
    __mikuProjectCoreApi?: unknown;
    __mikuprojectCoreApi?: unknown;
  };

  // Canonical public API name. Keep the former name as a compatibility alias
  // for embedders that evaluate the single-file application directly.
  globals.__mikuProjectCoreApi = coreApi;
  globals.__mikuprojectCoreApi = coreApi;
})();
