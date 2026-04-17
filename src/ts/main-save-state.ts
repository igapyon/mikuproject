/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  const mainSaveState = {
    updateXmlSaveState(input: {
      document: Document;
      isDirty: boolean;
      lastSavedXmlStamp: string;
      updateXmlSaveStateView: (doc: Document, input: { isDirty: boolean; lastSavedXmlStamp: string }) => void;
    }): void {
      input.updateXmlSaveStateView(input.document, {
        isDirty: input.isDirty,
        lastSavedXmlStamp: input.lastSavedXmlStamp
      });
    },

    markXmlDirty(input: {
      document: Document;
      lastSavedXmlStamp: string;
      updateXmlSaveStateView: (doc: Document, input: { isDirty: boolean; lastSavedXmlStamp: string }) => void;
    }): void {
      this.updateXmlSaveState({
        document: input.document,
        isDirty: true,
        lastSavedXmlStamp: input.lastSavedXmlStamp,
        updateXmlSaveStateView: input.updateXmlSaveStateView
      });
    },

    markXmlSavedCurrent(input: {
      readXmlText: () => string;
      formatSaveStamp: (date: Date) => string;
      writeSavedState: (state: { lastSavedXmlText: string; lastSavedXmlStamp: string }) => void;
      document: Document;
      updateXmlSaveStateView: (doc: Document, input: { isDirty: boolean; lastSavedXmlStamp: string }) => void;
    }): void {
      const lastSavedXmlText = input.readXmlText();
      const lastSavedXmlStamp = input.formatSaveStamp(new Date());
      input.writeSavedState({ lastSavedXmlText, lastSavedXmlStamp });
      input.updateXmlSaveStateView(input.document, {
        isDirty: false,
        lastSavedXmlStamp
      });
    },

    refreshXmlSaveState(input: {
      readXmlText: () => string;
      lastSavedXmlText: string;
      lastSavedXmlStamp: string;
      document: Document;
      updateXmlSaveStateView: (doc: Document, input: { isDirty: boolean; lastSavedXmlStamp: string }) => void;
    }): void {
      input.updateXmlSaveStateView(input.document, {
        isDirty: input.readXmlText() !== input.lastSavedXmlText,
        lastSavedXmlStamp: input.lastSavedXmlStamp
      });
    }
  };

  (globalThis as typeof globalThis & {
    __mikuprojectMainSaveState?: typeof mainSaveState;
  }).__mikuprojectMainSaveState = mainSaveState;
})();
