/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  const AI_JSON_SPEC_TEXT = {{AI_PROMPT_TEXT_JSON}};
  const AI_JSON_SPEC_VERSION = detectAiJsonSpecVersion(AI_JSON_SPEC_TEXT);

  function detectAiJsonSpecVersion(text: string): string {
    const match = text.match(/^- Version:\s*`([^`]+)`/m);
    return match?.[1] || "unknown";
  }

  function getAiJsonSpecText(): string {
    return AI_JSON_SPEC_TEXT;
  }

  function getAiJsonSpec(): {
    id: "miku-project-ai-json-spec";
    version: string;
    text: string;
  } {
    return {
      id: "miku-project-ai-json-spec",
      version: AI_JSON_SPEC_VERSION,
      text: AI_JSON_SPEC_TEXT
    };
  }

  (globalThis as typeof globalThis & {
    __mikuprojectAiJsonSpec?: {
      getAiJsonSpecText: () => string;
      getAiJsonSpec: () => {
        id: "miku-project-ai-json-spec";
        version: string;
        text: string;
      };
    };
  }).__mikuprojectAiJsonSpec = {
    getAiJsonSpecText,
    getAiJsonSpec
  };
})();
