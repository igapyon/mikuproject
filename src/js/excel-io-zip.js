/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const TEXT_ENCODER = new TextEncoder();
    const TEXT_DECODER = new TextDecoder();
    const CRC32_TABLE = buildCrc32Table();
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
            localView.setUint16(10, 0, true);
            localView.setUint16(12, 0, true);
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
            centralView.setUint16(12, 0, true);
            centralView.setUint16(14, 0, true);
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
    function unpackZip(bytes) {
        const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
        const endOffset = findEndOfCentralDirectoryOffset(bytes);
        const totalEntries = view.getUint16(endOffset + 10, true);
        const centralDirectoryOffset = view.getUint32(endOffset + 16, true);
        const entries = {};
        let pointer = centralDirectoryOffset;
        for (let index = 0; index < totalEntries; index += 1) {
            if (view.getUint32(pointer, true) !== 0x02014b50) {
                throw new Error("Invalid ZIP central directory header");
            }
            const compressionMethod = view.getUint16(pointer + 10, true);
            const compressedSize = view.getUint32(pointer + 20, true);
            const uncompressedSize = view.getUint32(pointer + 24, true);
            const fileNameLength = view.getUint16(pointer + 28, true);
            const extraLength = view.getUint16(pointer + 30, true);
            const commentLength = view.getUint16(pointer + 32, true);
            const localHeaderOffset = view.getUint32(pointer + 42, true);
            const fileName = decodeUtf8(bytes.subarray(pointer + 46, pointer + 46 + fileNameLength));
            const localView = new DataView(bytes.buffer, bytes.byteOffset + localHeaderOffset, bytes.byteLength - localHeaderOffset);
            if (localView.getUint32(0, true) !== 0x04034b50) {
                throw new Error(`Invalid ZIP local header for entry: ${fileName}`);
            }
            const localFileNameLength = localView.getUint16(26, true);
            const localExtraLength = localView.getUint16(28, true);
            const dataOffset = localHeaderOffset + 30 + localFileNameLength + localExtraLength;
            const data = bytes.slice(dataOffset, dataOffset + compressedSize);
            if (compressionMethod !== 0) {
                throw new Error(`Unsupported ZIP compression method for entry ${fileName}: ${compressionMethod}`);
            }
            if (compressedSize !== uncompressedSize) {
                throw new Error(`Stored ZIP entry size mismatch: ${fileName}`);
            }
            entries[fileName] = data;
            pointer += 46 + fileNameLength + extraLength + commentLength;
        }
        return entries;
    }
    async function unpackZipAsync(bytes) {
        const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
        const endOffset = findEndOfCentralDirectoryOffset(bytes);
        const totalEntries = view.getUint16(endOffset + 10, true);
        const centralDirectoryOffset = view.getUint32(endOffset + 16, true);
        const entries = {};
        let pointer = centralDirectoryOffset;
        for (let index = 0; index < totalEntries; index += 1) {
            if (view.getUint32(pointer, true) !== 0x02014b50) {
                throw new Error("Invalid ZIP central directory header");
            }
            const compressionMethod = view.getUint16(pointer + 10, true);
            const compressedSize = view.getUint32(pointer + 20, true);
            const uncompressedSize = view.getUint32(pointer + 24, true);
            const fileNameLength = view.getUint16(pointer + 28, true);
            const extraLength = view.getUint16(pointer + 30, true);
            const commentLength = view.getUint16(pointer + 32, true);
            const localHeaderOffset = view.getUint32(pointer + 42, true);
            const fileName = decodeUtf8(bytes.subarray(pointer + 46, pointer + 46 + fileNameLength));
            const localView = new DataView(bytes.buffer, bytes.byteOffset + localHeaderOffset, bytes.byteLength - localHeaderOffset);
            if (localView.getUint32(0, true) !== 0x04034b50) {
                throw new Error(`Invalid ZIP local header for entry: ${fileName}`);
            }
            const localFileNameLength = localView.getUint16(26, true);
            const localExtraLength = localView.getUint16(28, true);
            const dataOffset = localHeaderOffset + 30 + localFileNameLength + localExtraLength;
            const data = bytes.slice(dataOffset, dataOffset + compressedSize);
            if (compressionMethod === 0) {
                if (compressedSize !== uncompressedSize) {
                    throw new Error(`Stored ZIP entry size mismatch: ${fileName}`);
                }
                entries[fileName] = data;
            }
            else if (compressionMethod === 8) {
                entries[fileName] = await inflateDeflateRaw(data, uncompressedSize, fileName);
            }
            else {
                throw new Error(`Unsupported ZIP compression method for entry ${fileName}: ${compressionMethod}`);
            }
            pointer += 46 + fileNameLength + extraLength + commentLength;
        }
        return entries;
    }
    async function inflateDeflateRaw(compressed, expectedSize, fileName) {
        var _a;
        if (typeof DecompressionStream !== "function") {
            throw new Error(`ZIP deflate compression is not supported in this runtime: ${fileName}`);
        }
        const sourceStream = typeof Blob === "function" && typeof ((_a = Blob.prototype) === null || _a === void 0 ? void 0 : _a.stream) === "function"
            ? new Blob([compressed]).stream()
            : new Response(compressed).body;
        if (!sourceStream) {
            throw new Error(`ZIP deflate stream source is not available in this runtime: ${fileName}`);
        }
        const decompressedStream = sourceStream.pipeThrough(new DecompressionStream("deflate-raw"));
        const buffer = await new Response(decompressedStream).arrayBuffer();
        const inflated = new Uint8Array(buffer);
        if (inflated.byteLength !== expectedSize) {
            throw new Error(`Deflated ZIP entry size mismatch: ${fileName}`);
        }
        return inflated;
    }
    function findEndOfCentralDirectoryOffset(bytes) {
        for (let index = bytes.byteLength - 22; index >= 0; index -= 1) {
            if (bytes[index] === 0x50 &&
                bytes[index + 1] === 0x4b &&
                bytes[index + 2] === 0x05 &&
                bytes[index + 3] === 0x06) {
                return index;
            }
        }
        throw new Error("ZIP end of central directory not found");
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
            crc = (crc >>> 8) ^ CRC32_TABLE[(crc ^ byte) & 0xff];
        }
        return (crc ^ 0xffffffff) >>> 0;
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
    function encodeUtf8(value) {
        return TEXT_ENCODER.encode(value);
    }
    function decodeUtf8(bytes) {
        return TEXT_DECODER.decode(bytes);
    }
    globalThis.__mikuprojectExcelIoZip = {
        packZip,
        unpackZip,
        unpackZipAsync
    };
})();
