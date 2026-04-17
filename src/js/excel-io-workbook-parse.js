/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    function parseWorkbookEntries(entries, helpers) {
        const workbookXml = helpers.decodeRequiredEntry(entries, "xl/workbook.xml");
        const workbookRelsXml = helpers.decodeRequiredEntry(entries, "xl/_rels/workbook.xml.rels");
        const stylesXml = entries["xl/styles.xml"] ? helpers.decodeUtf8(entries["xl/styles.xml"]) : null;
        const sharedStringsXml = entries["xl/sharedStrings.xml"] ? helpers.decodeUtf8(entries["xl/sharedStrings.xml"]) : null;
        const workbookDocument = helpers.parseXmlDocument(workbookXml);
        const relationshipsDocument = helpers.parseXmlDocument(workbookRelsXml);
        const styleBook = helpers.parseStylesXml(stylesXml, {
            parseXmlDocument: helpers.parseXmlDocument,
            findDirectChild: helpers.findDirectChild,
            parseOptionalNumber: helpers.parseOptionalNumber
        });
        const sharedStrings = helpers.parseSharedStringsXml(sharedStringsXml, {
            parseXmlDocument: helpers.parseXmlDocument,
            findDirectChild: helpers.findDirectChild
        });
        const relationshipMap = new Map();
        const relationshipElements = Array.from(relationshipsDocument.getElementsByTagNameNS("http://schemas.openxmlformats.org/package/2006/relationships", "Relationship"));
        for (const relationshipElement of relationshipElements) {
            const id = relationshipElement.getAttribute("Id");
            const target = relationshipElement.getAttribute("Target");
            if (id && target) {
                relationshipMap.set(id, normalizeWorkbookTarget(target));
            }
        }
        const sheetElements = Array.from(workbookDocument.getElementsByTagNameNS("http://schemas.openxmlformats.org/spreadsheetml/2006/main", "sheet"));
        return {
            sheets: sheetElements.map((sheetElement) => {
                const name = sheetElement.getAttribute("name") || "";
                helpers.validateSheetName(name);
                const relationshipId = sheetElement.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id") || sheetElement.getAttribute("r:id");
                if (!relationshipId) {
                    throw new Error(`Worksheet relationship id is missing for sheet: ${name}`);
                }
                const target = relationshipMap.get(relationshipId);
                if (!target) {
                    throw new Error(`Worksheet relationship target is missing for sheet: ${name}`);
                }
                const worksheetXml = helpers.decodeRequiredEntry(entries, target);
                return helpers.parseWorksheetXml(name, worksheetXml, styleBook, sharedStrings, {
                    parseXmlDocument: helpers.parseXmlDocument,
                    parseOptionalNumber: helpers.parseOptionalNumber,
                    findDirectChild: helpers.findDirectChild,
                    normalizeMergedRange: helpers.normalizeMergedRange,
                    denormalizeColor: helpers.denormalizeColor
                });
            })
        };
    }
    function normalizeWorkbookTarget(target) {
        return target.startsWith("xl/") ? target : `xl/${target.replace(/^\.\//, "")}`;
    }
    globalThis.__mikuprojectExcelIoWorkbookParse = {
        parseWorkbookEntries
    };
})();
