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

  (globalThis as typeof globalThis & {
    __mikuprojectCoreApi?: unknown;
  }).__mikuprojectCoreApi = mikuprojectCoreApiPublic;
})();
