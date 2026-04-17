/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  const mainDownloads = {
    downloadXml(input: {
      xmlText: string;
      buildXmlExport: (input: { xmlText: string }) => { fileName: string; text: string };
      downloadBlob: (blob: Blob, filename: string) => void;
      markXmlSavedCurrent: () => void;
      completeOutput: (statusMessage: string, toastMessage: string) => void;
    }): void {
      const exported = input.buildXmlExport({ xmlText: input.xmlText });
      input.downloadBlob(
        new Blob([exported.text], { type: "application/xml;charset=utf-8" }),
        exported.fileName
      );
      input.markXmlSavedCurrent();
      input.completeOutput("XML ファイルをエクスポートしました", "XML を保存しました");
    },

    downloadDailySvg(input: {
      svg: string;
      buildDailySvgExport: (input: { svg: string }) => { fileName: string; text: string };
      downloadBlob: (blob: Blob, filename: string) => void;
      completeOutput: (statusMessage: string, toastMessage: string) => void;
    }): void {
      const exported = input.buildDailySvgExport({ svg: input.svg });
      input.downloadBlob(
        new Blob([exported.text], { type: "image/svg+xml;charset=utf-8" }),
        exported.fileName
      );
      input.completeOutput("Daily SVG を保存しました", "Daily SVG を保存しました");
    },

    downloadWeeklySvg(input: {
      svg: string;
      buildWeeklySvgExport: (input: { svg: string }) => { fileName: string; text: string };
      downloadBlob: (blob: Blob, filename: string) => void;
      completeOutput: (statusMessage: string, toastMessage: string) => void;
    }): void {
      const exported = input.buildWeeklySvgExport({ svg: input.svg });
      input.downloadBlob(
        new Blob([exported.text], { type: "image/svg+xml;charset=utf-8" }),
        exported.fileName
      );
      input.completeOutput("Weekly SVG を保存しました", "Weekly SVG を保存しました");
    },

    downloadMonthlySvgZip(input: {
      zipBytes: Uint8Array;
      buildMonthlySvgZipExport: (input: { zipBytes: Uint8Array }) => { fileName: string; bytes: Uint8Array };
      downloadBlob: (blob: Blob, filename: string) => void;
      completeOutput: (statusMessage: string, toastMessage: string) => void;
    }): void {
      const exported = input.buildMonthlySvgZipExport({ zipBytes: input.zipBytes });
      input.downloadBlob(
        new Blob([exported.bytes], { type: "application/zip" }),
        exported.fileName
      );
      input.completeOutput("月別 WBS カレンダー SVG ZIP を保存しました", "月別 WBS カレンダー SVG ZIP を保存しました");
    },

    downloadMermaid(input: {
      mermaidText: string;
      buildMermaidExport: (input: { mermaidText: string }) => { fileName: string; text: string };
      downloadBlob: (blob: Blob, filename: string) => void;
      completeOutput: (statusMessage: string, toastMessage: string) => void;
    }): void {
      const exported = input.buildMermaidExport({ mermaidText: input.mermaidText });
      input.downloadBlob(
        new Blob([exported.text], { type: "text/plain;charset=utf-8" }),
        exported.fileName
      );
      input.completeOutput("Mermaid を保存しました", "Mermaid を保存しました");
    },

    downloadWbsMarkdown(input: {
      markdownText: string;
      buildWbsMarkdownExport: (input: { markdownText: string }) => { fileName: string; text: string };
      downloadBlob: (blob: Blob, filename: string) => void;
      completeOutput: (statusMessage: string, toastMessage: string) => void;
    }): void {
      const exported = input.buildWbsMarkdownExport({ markdownText: input.markdownText });
      input.downloadBlob(
        new Blob([exported.text], { type: "text/markdown;charset=utf-8" }),
        exported.fileName
      );
      input.completeOutput("WBS Markdown を保存しました", "WBS Markdown を保存しました");
    }
  };

  (globalThis as typeof globalThis & {
    __mikuprojectMainDownloads?: typeof mainDownloads;
  }).__mikuprojectMainDownloads = mainDownloads;
})();
