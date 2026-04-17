/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const excelIoUtil = globalThis.__mikuprojectExcelIoUtil;
    if (!excelIoUtil) {
        throw new Error("mikuproject Excel util module is not loaded");
    }
    const excelIoZip = globalThis.__mikuprojectExcelIoZip;
    if (!excelIoZip) {
        throw new Error("mikuproject Excel ZIP module is not loaded");
    }
    const excelIoNormalize = globalThis.__mikuprojectExcelIoNormalize;
    if (!excelIoNormalize) {
        throw new Error("mikuproject Excel normalize module is not loaded");
    }
    const excelIoPackageXml = globalThis.__mikuprojectExcelIoPackageXml;
    if (!excelIoPackageXml) {
        throw new Error("mikuproject Excel package XML module is not loaded");
    }
    const excelIoWorksheetBuild = globalThis.__mikuprojectExcelIoWorksheetBuild;
    if (!excelIoWorksheetBuild) {
        throw new Error("mikuproject Excel worksheet build module is not loaded");
    }
    const excelIoStylesBuild = globalThis.__mikuprojectExcelIoStylesBuild;
    if (!excelIoStylesBuild) {
        throw new Error("mikuproject Excel styles build module is not loaded");
    }
    const excelIoStylesParse = globalThis.__mikuprojectExcelIoStylesParse;
    if (!excelIoStylesParse) {
        throw new Error("mikuproject Excel styles parse module is not loaded");
    }
    const excelIoWorksheetParse = globalThis.__mikuprojectExcelIoWorksheetParse;
    if (!excelIoWorksheetParse) {
        throw new Error("mikuproject Excel worksheet parse module is not loaded");
    }
    const excelIoWorkbookParse = globalThis.__mikuprojectExcelIoWorkbookParse;
    if (!excelIoWorkbookParse) {
        throw new Error("mikuproject Excel workbook parse module is not loaded");
    }
    const excelIoWorkbookBuild = globalThis.__mikuprojectExcelIoWorkbookBuild;
    if (!excelIoWorkbookBuild) {
        throw new Error("mikuproject Excel workbook build module is not loaded");
    }
    class XlsxWorkbookCodec {
        exportWorkbook(workbook) {
            const normalizedWorkbook = excelIoNormalize.normalizeWorkbook(workbook);
            const entries = excelIoWorkbookBuild.createWorkbookEntries(normalizedWorkbook, {
                createStyleBook: excelIoStylesBuild.createStyleBook,
                buildWorksheetXml: excelIoWorksheetBuild.buildWorksheetXml,
                resolveStyleIndex: excelIoStylesBuild.resolveStyleIndex,
                resolveCellNumberFormat: excelIoStylesBuild.resolveCellNumberFormat,
                buildContentTypesXml: excelIoPackageXml.buildContentTypesXml,
                buildRootRelationshipsXml: excelIoPackageXml.buildRootRelationshipsXml,
                buildWorkbookRelationshipsXml: excelIoPackageXml.buildWorkbookRelationshipsXml,
                buildWorkbookXml: excelIoPackageXml.buildWorkbookXml,
                buildStylesXml: excelIoStylesBuild.buildStylesXml,
                encodeUtf8: excelIoUtil.encodeUtf8,
                escapeXml: excelIoUtil.escapeXml
            });
            return excelIoZip.packZip(entries);
        }
        importWorkbook(bytes) {
            const entries = this.unpackEntries(bytes);
            return excelIoWorkbookParse.parseWorkbookEntries(entries, {
                decodeRequiredEntry: excelIoUtil.decodeRequiredEntry,
                decodeUtf8: excelIoUtil.decodeUtf8,
                parseXmlDocument: excelIoUtil.parseXmlDocument,
                findDirectChild: excelIoUtil.findDirectChild,
                parseOptionalNumber: excelIoUtil.parseOptionalNumber,
                validateSheetName: excelIoNormalize.validateSheetName,
                normalizeMergedRange: excelIoNormalize.normalizeMergedRange,
                denormalizeColor: excelIoNormalize.denormalizeColor,
                parseStylesXml: excelIoStylesParse.parseStylesXml,
                parseSharedStringsXml: excelIoWorksheetParse.parseSharedStringsXml,
                parseWorksheetXml: excelIoWorksheetParse.parseWorksheetXml
            });
        }
        async importWorkbookAsync(bytes) {
            const entries = await this.unpackEntriesAsync(bytes);
            return excelIoWorkbookParse.parseWorkbookEntries(entries, {
                decodeRequiredEntry: excelIoUtil.decodeRequiredEntry,
                decodeUtf8: excelIoUtil.decodeUtf8,
                parseXmlDocument: excelIoUtil.parseXmlDocument,
                findDirectChild: excelIoUtil.findDirectChild,
                parseOptionalNumber: excelIoUtil.parseOptionalNumber,
                validateSheetName: excelIoNormalize.validateSheetName,
                normalizeMergedRange: excelIoNormalize.normalizeMergedRange,
                denormalizeColor: excelIoNormalize.denormalizeColor,
                parseStylesXml: excelIoStylesParse.parseStylesXml,
                parseSharedStringsXml: excelIoWorksheetParse.parseSharedStringsXml,
                parseWorksheetXml: excelIoWorksheetParse.parseWorksheetXml
            });
        }
        listEntries(bytes) {
            return Object.keys(this.unpackEntries(bytes)).sort();
        }
        async listEntriesAsync(bytes) {
            const entries = await this.unpackEntriesAsync(bytes);
            return Object.keys(entries).sort();
        }
        unpackEntries(bytes) {
            return excelIoZip.unpackZip(bytes);
        }
        async unpackEntriesAsync(bytes) {
            return excelIoZip.unpackZipAsync(bytes);
        }
    }
    globalThis.__mikuprojectExcelIo = {
        XlsxWorkbookCodec
    };
})();
