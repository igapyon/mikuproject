/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    function createWbsSvgZipHelper(config) {
        const { fixedModTime, fixedModDate } = config;
        const textEncoder = new TextEncoder();
        const crc32Table = buildCrc32Table();
        function escapeXml(value) {
            return String(value || "")
                .replaceAll("&", "&amp;")
                .replaceAll("<", "&lt;")
                .replaceAll(">", "&gt;")
                .replaceAll("\"", "&quot;");
        }
        function encodeUtf8(value) {
            return textEncoder.encode(value);
        }
        function packZip(entries) {
            const localParts = [];
            const centralParts = [];
            let offset = 0;
            for (const entry of entries) {
                const nameBytes = encodeUtf8(entry.name);
                const crc32 = computeCrc32(entry.data);
                const localHeader = new Uint8Array(30 + nameBytes.length);
                const localView = new DataView(localHeader.buffer);
                localView.setUint32(0, 0x04034b50, true);
                localView.setUint16(4, 20, true);
                localView.setUint16(6, 0, true);
                localView.setUint16(8, 0, true);
                localView.setUint16(10, fixedModTime, true);
                localView.setUint16(12, fixedModDate, true);
                localView.setUint32(14, crc32, true);
                localView.setUint32(18, entry.data.byteLength, true);
                localView.setUint32(22, entry.data.byteLength, true);
                localView.setUint16(26, nameBytes.length, true);
                localView.setUint16(28, 0, true);
                localHeader.set(nameBytes, 30);
                const centralHeader = new Uint8Array(46 + nameBytes.length);
                const centralView = new DataView(centralHeader.buffer);
                centralView.setUint32(0, 0x02014b50, true);
                centralView.setUint16(4, 20, true);
                centralView.setUint16(6, 20, true);
                centralView.setUint16(8, 0, true);
                centralView.setUint16(10, 0, true);
                centralView.setUint16(12, fixedModTime, true);
                centralView.setUint16(14, fixedModDate, true);
                centralView.setUint32(16, crc32, true);
                centralView.setUint32(20, entry.data.byteLength, true);
                centralView.setUint32(24, entry.data.byteLength, true);
                centralView.setUint16(28, nameBytes.length, true);
                centralView.setUint16(30, 0, true);
                centralView.setUint16(32, 0, true);
                centralView.setUint16(34, 0, true);
                centralView.setUint16(36, 0, true);
                centralView.setUint32(38, 0, true);
                centralView.setUint32(42, offset, true);
                centralHeader.set(nameBytes, 46);
                localParts.push(localHeader, entry.data);
                centralParts.push(centralHeader);
                offset += localHeader.byteLength + entry.data.byteLength;
            }
            const centralDirectoryOffset = offset;
            const centralDirectorySize = centralParts.reduce((sum, part) => sum + part.byteLength, 0);
            const endOfCentralDirectory = new Uint8Array(22);
            const endView = new DataView(endOfCentralDirectory.buffer);
            endView.setUint32(0, 0x06054b50, true);
            endView.setUint16(4, 0, true);
            endView.setUint16(6, 0, true);
            endView.setUint16(8, entries.length, true);
            endView.setUint16(10, entries.length, true);
            endView.setUint32(12, centralDirectorySize, true);
            endView.setUint32(16, centralDirectoryOffset, true);
            endView.setUint16(20, 0, true);
            return concatUint8Arrays([...localParts, ...centralParts, endOfCentralDirectory]);
        }
        function concatUint8Arrays(parts) {
            const totalLength = parts.reduce((sum, part) => sum + part.byteLength, 0);
            const result = new Uint8Array(totalLength);
            let offset = 0;
            for (const part of parts) {
                result.set(part, offset);
                offset += part.byteLength;
            }
            return result;
        }
        function computeCrc32(bytes) {
            let crc = 0xffffffff;
            for (const byte of bytes) {
                crc = (crc >>> 8) ^ crc32Table[(crc ^ byte) & 0xff];
            }
            return (crc ^ 0xffffffff) >>> 0;
        }
        return {
            escapeXml,
            encodeUtf8,
            packZip
        };
    }
    function buildCrc32Table() {
        const table = new Uint32Array(256);
        for (let index = 0; index < 256; index += 1) {
            let value = index;
            for (let bit = 0; bit < 8; bit += 1) {
                value = (value & 1) !== 0 ? (0xedb88320 ^ (value >>> 1)) : (value >>> 1);
            }
            table[index] = value >>> 0;
        }
        return table;
    }
    globalThis.__mikuprojectWbsSvgZip = {
        createWbsSvgZipHelper
    };
})();
