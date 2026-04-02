/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  function extractLastJsonBlock(value: string): string {
    const matches = Array.from(value.matchAll(/```json\s*([\s\S]*?)```/g));
    if (matches.length > 0) {
      return matches.at(-1)?.[1]?.trim() || "";
    }
    return value.trim();
  }

  function detectJsonDocumentKind(documentLike: unknown): "workbook_json" | "project_draft_view" | "patch_json" | undefined {
    if (!documentLike || typeof documentLike !== "object") {
      return undefined;
    }
    const candidate = documentLike as {
      format?: string;
      view_type?: string;
      operations?: unknown;
    };
    if (candidate.format === "mikuproject_workbook_json") {
      return "workbook_json";
    }
    if (candidate.view_type === "project_draft_view") {
      return "project_draft_view";
    }
    if (Array.isArray(candidate.operations)) {
      return "patch_json";
    }
    return undefined;
  }

  (globalThis as typeof globalThis & {
    __mikuprojectAiJsonUtil?: {
      extractLastJsonBlock: (value: string) => string;
      detectJsonDocumentKind: (documentLike: unknown) => "workbook_json" | "project_draft_view" | "patch_json" | undefined;
    };
  }).__mikuprojectAiJsonUtil = {
    extractLastJsonBlock,
    detectJsonDocumentKind
  };
})();
