/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  const projectPatchJsonCore = (globalThis as typeof globalThis & {
    __mikuprojectProjectPatchJsonCore?: {
      importProjectPatchJson: (documentLike: unknown, baseModel: ProjectModel) => {
        model: ProjectModel;
        changes: ImportChange[];
        warnings: PatchWarning[];
      };
      validatePatchDocument: (documentLike: unknown) => {
        document: PatchDocument;
        warnings: PatchWarning[];
      };
    };
  }).__mikuprojectProjectPatchJsonCore;

  if (!projectPatchJsonCore) {
    throw new Error("mikuproject Project Patch JSON core module is not loaded");
  }

  (globalThis as typeof globalThis & {
    __mikuprojectProjectPatchJson?: typeof projectPatchJsonCore;
  }).__mikuprojectProjectPatchJson = projectPatchJsonCore;
})();
