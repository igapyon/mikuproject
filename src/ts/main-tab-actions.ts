/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  const mainTabActions = {
    setActiveTab(input: {
      document: Document;
      tabId: "input" | "transform" | "output";
      skipTransformRefresh?: boolean;
      isRefreshingTransformTab: boolean;
      setCurrentTabId: (tabId: "input" | "transform" | "output") => void;
      setActiveTabView: (doc: Document, tabId: "input" | "transform" | "output") => void;
      refreshTransformTab: () => Promise<void>;
      setStatus: (message: string) => void;
    }): void {
      input.setCurrentTabId(input.tabId);
      input.setActiveTabView(input.document, input.tabId);
      if (input.tabId === "transform" && !input.skipTransformRefresh && !input.isRefreshingTransformTab) {
        void input.refreshTransformTab().catch((error) => {
          input.setStatus(error instanceof Error ? error.message : "Transform の更新に失敗しました");
        });
      }
    },

    async refreshTransformTab(input: {
      isRefreshingTransformTab: boolean;
      setRefreshingTransformTab: (value: boolean) => void;
      refreshTransformTabImpl: (input: {
        currentModel: ProjectModel | null;
        isXmlSourceDirty: boolean;
        readXmlText: () => string;
        parseCurrentXml: (options?: { silent?: boolean }) => void;
        exportCurrentMermaid: (options?: { silent?: boolean }) => Promise<void>;
        setStatus: (message: string) => void;
      }) => Promise<void>;
      currentModel: ProjectModel | null;
      isXmlSourceDirty: boolean;
      readXmlText: () => string;
      parseCurrentXml: (options?: { silent?: boolean }) => void;
      exportCurrentMermaid: (options?: { silent?: boolean }) => Promise<void>;
      setStatus: (message: string) => void;
    }): Promise<void> {
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
      } finally {
        input.setRefreshingTransformTab(false);
      }
    },

    moveTabFocus(input: {
      document: Document;
      currentButton: HTMLButtonElement;
      direction: -1 | 1;
      moveTabFocusView: (doc: Document, currentButton: HTMLButtonElement, direction: -1 | 1) => HTMLButtonElement | null;
      setActiveTab: (tabId: "input" | "transform" | "output") => void;
    }): void {
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

    bindTabs(input: {
      document: Document;
      currentTabId: "input" | "transform" | "output";
      getTabButtons: () => HTMLButtonElement[];
      bindTabsView: (input: {
        doc: Document;
        currentTabId: "input" | "transform" | "output";
        getTabButtons: () => HTMLButtonElement[];
        setActiveTab: (tabId: "input" | "transform" | "output") => void;
        moveTabFocus: (currentButton: HTMLButtonElement, direction: -1 | 1) => void;
      }) => void;
      setActiveTab: (tabId: "input" | "transform" | "output") => void;
      moveTabFocus: (currentButton: HTMLButtonElement, direction: -1 | 1) => void;
    }): void {
      input.bindTabsView({
        doc: input.document,
        currentTabId: input.currentTabId,
        getTabButtons: input.getTabButtons,
        setActiveTab: input.setActiveTab,
        moveTabFocus: input.moveTabFocus
      });
    }
  };

  (globalThis as typeof globalThis & {
    __mikuprojectMainTabActions?: typeof mainTabActions;
  }).__mikuprojectMainTabActions = mainTabActions;
})();
