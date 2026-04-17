/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const mainModel = {
        syncXmlTextFromModel(input) {
            const xmlText = input.exportMsProjectXml(input.model);
            input.writeXmlText(xmlText);
            input.clearXmlSourceDirty();
            input.refreshXmlSaveState();
            return xmlText;
        },
        ensureCurrentModel(input) {
            if (input.currentModel) {
                return input.currentModel;
            }
            const xmlText = input.readXmlText().trim();
            if (!xmlText) {
                throw new Error("内部モデルがありません");
            }
            const model = input.importMsProjectXml(xmlText);
            input.clearXmlSourceDirty();
            return model;
        },
        parseCurrentXml(input) {
            const xmlText = input.readXmlText().trim();
            if (!xmlText) {
                return {
                    model: null,
                    issues: [],
                    empty: true
                };
            }
            const model = input.importMsProjectXml(xmlText);
            input.clearXmlSourceDirty();
            return {
                model,
                issues: input.validateProjectModel(model),
                empty: false
            };
        }
    };
    globalThis.__mikuprojectMainModel = mainModel;
})();
