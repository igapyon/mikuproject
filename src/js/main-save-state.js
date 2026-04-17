/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const mainSaveState = {
        updateXmlSaveState(input) {
            input.updateXmlSaveStateView(input.document, {
                isDirty: input.isDirty,
                lastSavedXmlStamp: input.lastSavedXmlStamp
            });
        },
        markXmlDirty(input) {
            this.updateXmlSaveState({
                document: input.document,
                isDirty: true,
                lastSavedXmlStamp: input.lastSavedXmlStamp,
                updateXmlSaveStateView: input.updateXmlSaveStateView
            });
        },
        markXmlSavedCurrent(input) {
            const lastSavedXmlText = input.readXmlText();
            const lastSavedXmlStamp = input.formatSaveStamp(new Date());
            input.writeSavedState({ lastSavedXmlText, lastSavedXmlStamp });
            input.updateXmlSaveStateView(input.document, {
                isDirty: false,
                lastSavedXmlStamp
            });
        },
        refreshXmlSaveState(input) {
            input.updateXmlSaveStateView(input.document, {
                isDirty: input.readXmlText() !== input.lastSavedXmlText,
                lastSavedXmlStamp: input.lastSavedXmlStamp
            });
        }
    };
    globalThis.__mikuprojectMainSaveState = mainSaveState;
})();
