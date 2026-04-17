/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    function createWorkbookEntries(workbook, helpers) {
        const styleBook = helpers.createStyleBook(workbook);
        const worksheetRelationships = workbook.sheets.map((sheet, index) => ({
            relationshipId: `rId${index + 1}`,
            target: `worksheets/sheet${index + 1}.xml`,
            name: sheet.name
        }));
        const worksheetEntries = workbook.sheets.map((sheet, index) => ({
            name: `xl/worksheets/sheet${index + 1}.xml`,
            data: helpers.encodeUtf8(helpers.buildWorksheetXml(sheet, styleBook, {
                resolveStyleIndex: helpers.resolveStyleIndex,
                resolveCellNumberFormat: helpers.resolveCellNumberFormat,
                escapeXml: helpers.escapeXml
            }))
        }));
        const entries = [
            {
                name: "[Content_Types].xml",
                data: helpers.encodeUtf8(helpers.buildContentTypesXml(workbook.sheets.length, styleBook.styles.length > 1))
            },
            {
                name: "_rels/.rels",
                data: helpers.encodeUtf8(helpers.buildRootRelationshipsXml())
            },
            {
                name: "xl/_rels/workbook.xml.rels",
                data: helpers.encodeUtf8(helpers.buildWorkbookRelationshipsXml(worksheetRelationships, styleBook.styles.length > 1))
            },
            {
                name: "xl/workbook.xml",
                data: helpers.encodeUtf8(helpers.buildWorkbookXml(worksheetRelationships))
            }
        ];
        if (styleBook.styles.length > 1) {
            entries.push({
                name: "xl/styles.xml",
                data: helpers.encodeUtf8(helpers.buildStylesXml(styleBook.styles))
            });
        }
        entries.push(...worksheetEntries);
        return entries;
    }
    globalThis.__mikuprojectExcelIoWorkbookBuild = {
        createWorkbookEntries
    };
})();
