/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    function buildContentTypesXml(sheetCount, includeStyles) {
        const worksheetOverrides = Array.from({ length: sheetCount }, (_unused, index) => (`<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`)).join("");
        const stylesOverride = includeStyles
            ? `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>`
            : "";
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  ${worksheetOverrides}
  ${stylesOverride}
</Types>`;
    }
    function buildRootRelationshipsXml() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
    }
    function buildWorkbookRelationshipsXml(relationships, includeStyles) {
        const worksheetNodes = relationships.map((relationship) => (`<Relationship Id="${escapeXml(relationship.relationshipId)}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="${escapeXml(relationship.target)}"/>`)).join("");
        const stylesNode = includeStyles
            ? `<Relationship Id="rId${relationships.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`
            : "";
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${worksheetNodes}
  ${stylesNode}
</Relationships>`;
    }
    function buildWorkbookXml(relationships) {
        const sheets = relationships.map((relationship, index) => (`<sheet name="${escapeXml(relationship.name)}" sheetId="${index + 1}" r:id="${escapeXml(relationship.relationshipId)}"/>`)).join("");
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>${sheets}</sheets>
</workbook>`;
    }
    function escapeXml(value) {
        return value
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&apos;");
    }
    globalThis.__mikuprojectExcelIoPackageXml = {
        buildContentTypesXml,
        buildRootRelationshipsXml,
        buildWorkbookRelationshipsXml,
        buildWorkbookXml
    };
})();
