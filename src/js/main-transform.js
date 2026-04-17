/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const mainTransform = {
        bindTabs(input) {
            const buttons = input.getTabButtons();
            if (buttons.length === 0) {
                return;
            }
            for (const button of buttons) {
                button.addEventListener("click", () => {
                    const tabId = button.dataset.tab;
                    if (tabId === "input" || tabId === "transform" || tabId === "output") {
                        input.setActiveTab(tabId);
                    }
                });
                button.addEventListener("keydown", (event) => {
                    if (event.key === "ArrowRight" || event.key === "ArrowDown") {
                        event.preventDefault();
                        input.moveTabFocus(button, 1);
                        return;
                    }
                    if (event.key === "ArrowLeft" || event.key === "ArrowUp") {
                        event.preventDefault();
                        input.moveTabFocus(button, -1);
                        return;
                    }
                    if (event.key === "Home") {
                        event.preventDefault();
                        buttons[0].focus();
                        input.setActiveTab("input");
                        return;
                    }
                    if (event.key === "End") {
                        event.preventDefault();
                        buttons[buttons.length - 1].focus();
                        input.setActiveTab("output");
                    }
                });
            }
            input.setActiveTab(input.currentTabId);
        },
        buildWbsOptions(input) {
            const beforeInput = input.doc.getElementById("wbsDisplayDaysBeforeInput");
            const afterInput = input.doc.getElementById("wbsDisplayDaysAfterInput");
            const beforeValue = beforeInput instanceof HTMLInputElement ? beforeInput.value : "";
            const afterValue = afterInput instanceof HTMLInputElement ? afterInput.value : "";
            return {
                holidayDates: input.collectWbsHolidayDates(input.model),
                displayDaysBeforeBaseDate: input.parseOptionalNonNegativeInteger(beforeValue),
                displayDaysAfterBaseDate: input.parseOptionalNonNegativeInteger(afterValue),
                useBusinessDaysForDisplayRange: true,
                useBusinessDaysForProgressBand: true
            };
        },
        async refreshTransformTab(input) {
            if (!input.currentModel || input.isXmlSourceDirty) {
                const xmlText = input.readXmlText().trim();
                if (!xmlText) {
                    input.setStatus("XML が空です");
                    return;
                }
                input.parseCurrentXml({ silent: true });
            }
            await input.exportCurrentMermaid({ silent: true });
        }
    };
    globalThis.__mikuprojectMainTransform = mainTransform;
})();
