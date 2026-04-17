/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const mainSupport = {
        showToast(doc, message) {
            const toast = doc.getElementById("toast");
            if (toast && typeof toast.show === "function") {
                toast.show(message, 2200);
            }
        },
        getAiPromptText(doc, getSpecText) {
            var _a;
            if (typeof getSpecText === "function") {
                return getSpecText().trim();
            }
            const template = doc.getElementById("aiPromptTemplate");
            if (!template) {
                return "";
            }
            return (((_a = template.content) === null || _a === void 0 ? void 0 : _a.textContent) || template.textContent || "").trim();
        },
        async copyTextToClipboard(doc, text) {
            if (typeof navigator !== "undefined" &&
                navigator.clipboard &&
                typeof navigator.clipboard.writeText === "function") {
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
        downloadBlob(doc, blob, filename) {
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
    globalThis.__mikuprojectMainSupport = mainSupport;
})();
