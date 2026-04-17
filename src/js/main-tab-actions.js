/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const mainTabActions = {
        setActiveTab(input) {
            input.setCurrentTabId(input.tabId);
            input.setActiveTabView(input.document, input.tabId);
            if (input.tabId === "transform" && !input.skipTransformRefresh && !input.isRefreshingTransformTab) {
                void input.refreshTransformTab().catch((error) => {
                    input.setStatus(error instanceof Error ? error.message : "Transform の更新に失敗しました");
                });
            }
        },
        async refreshTransformTab(input) {
            if (input.isRefreshingTransformTab) {
                return;
            }
            input.setRefreshingTransformTab(true);
            try {
                await input.refreshTransformTabImpl({
                    currentModel: input.currentModel,
                    isXmlSourceDirty: input.isXmlSourceDirty,
                    readXmlText: input.readXmlText,
                    parseCurrentXml: input.parseCurrentXml,
                    exportCurrentMermaid: input.exportCurrentMermaid,
                    setStatus: input.setStatus
                });
            }
            finally {
                input.setRefreshingTransformTab(false);
            }
        },
        moveTabFocus(input) {
            const nextButton = input.moveTabFocusView(input.document, input.currentButton, input.direction);
            if (!nextButton) {
                return;
            }
            nextButton.focus();
            const nextTab = nextButton.dataset.tab;
            if (nextTab === "input" || nextTab === "transform" || nextTab === "output") {
                input.setActiveTab(nextTab);
            }
        },
        bindTabs(input) {
            input.bindTabsView({
                doc: input.document,
                currentTabId: input.currentTabId,
                getTabButtons: input.getTabButtons,
                setActiveTab: input.setActiveTab,
                moveTabFocus: input.moveTabFocus
            });
        }
    };
    globalThis.__mikuprojectMainTabActions = mainTabActions;
})();
