/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  const mainSupport = {
    showToast(doc: Document, message: string): void {
      const toast = doc.getElementById("toast") as (HTMLElement & { show?: (text: string, duration?: number) => void }) | null;
      if (toast && typeof toast.show === "function") {
        toast.show(message, 2200);
      }
    },

    getAiPromptText(doc: Document, getSpecText?: () => string): string {
      if (typeof getSpecText === "function") {
        return getSpecText().trim();
      }
      const template = doc.getElementById("aiPromptTemplate") as HTMLTemplateElement | null;
      if (!template) {
        return "";
      }
      return (template.content?.textContent || template.textContent || "").trim();
    },

    async copyTextToClipboard(doc: Document, text: string): Promise<void> {
      if (
        typeof navigator !== "undefined" &&
        navigator.clipboard &&
        typeof navigator.clipboard.writeText === "function"
      ) {
        await navigator.clipboard.writeText(text);
        return;
      }

      const textarea = doc.createElement("textarea");
      textarea.value = text;
      textarea.setAttribute("readonly", "readonly");
      textarea.style.position = "fixed";
      textarea.style.opacity = "0";
      doc.body.appendChild(textarea);
      textarea.select();
      doc.execCommand("copy");
      doc.body.removeChild(textarea);
    },

    downloadBlob(doc: Document, blob: Blob, filename: string): void {
      const objectUrl = URL.createObjectURL(blob);
      const link = doc.createElement("a");
      link.href = objectUrl;
      link.download = filename;
      doc.body.appendChild(link);
      link.click();
      doc.body.removeChild(link);
      window.setTimeout(() => URL.revokeObjectURL(objectUrl), 0);
    }
  };

  (globalThis as typeof globalThis & {
    __mikuprojectMainSupport?: typeof mainSupport;
  }).__mikuprojectMainSupport = mainSupport;
})();
