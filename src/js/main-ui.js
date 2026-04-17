/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const mainUi = {
        getTabButtons(doc) {
            return Array.from(doc.querySelectorAll(".md-top-tab[data-tab]"));
        },
        getTabPanels(doc) {
            return Array.from(doc.querySelectorAll(".md-tab-panel[data-tab-panel]"));
        },
        setStatus(doc, message) {
            const node = doc.getElementById("statusMessage");
            if (!(node instanceof HTMLElement)) {
                throw new Error("statusMessage element is not available");
            }
            node.textContent = message;
        },
        updateXmlSaveState(doc, input) {
            const node = doc.getElementById("xmlSaveState");
            if (!(node instanceof HTMLElement)) {
                throw new Error("xmlSaveState element is not available");
            }
            node.textContent = input.isDirty
                ? "XML 保存状態: 未保存"
                : `XML 保存状態: 保存済み (${input.lastSavedXmlStamp || "-"})`;
            node.classList.toggle("md-save-state--dirty", input.isDirty);
            node.classList.toggle("md-save-state--clean", !input.isDirty);
        },
        setActiveTab(doc, tabId) {
            for (const button of mainUi.getTabButtons(doc)) {
                const isActive = button.dataset.tab === tabId;
                button.classList.toggle("is-active", isActive);
                button.setAttribute("aria-selected", isActive ? "true" : "false");
                button.tabIndex = isActive ? 0 : -1;
            }
            for (const panel of mainUi.getTabPanels(doc)) {
                panel.hidden = panel.dataset.tabPanel !== tabId;
            }
        },
        moveTabFocus(doc, currentButton, direction) {
            const buttons = mainUi.getTabButtons(doc);
            const currentIndex = buttons.indexOf(currentButton);
            if (currentIndex < 0) {
                return null;
            }
            const nextIndex = (currentIndex + direction + buttons.length) % buttons.length;
            const nextButton = buttons[nextIndex];
            nextButton.focus();
            return nextButton;
        }
    };
    globalThis.__mikuprojectMainUi = mainUi;
})();
