// @vitest-environment jsdom

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const typesCode = readFileSync(
  path.resolve(__dirname, "../src/js/types.js"),
  "utf8"
);
const markdownEscapeCode = readFileSync(
  path.resolve(__dirname, "../src/js/markdown-escape.js"),
  "utf8"
);
const msProjectXmlCode = readFileSync(
  path.resolve(__dirname, "../src/js/msproject-xml.js"),
  "utf8"
);
const wbsMarkdownCode = readFileSync(
  path.resolve(__dirname, "../src/js/wbs-markdown.js"),
  "utf8"
);

function bootModules() {
  new Function(`${typesCode}\n${markdownEscapeCode}\n${msProjectXmlCode}\n${wbsMarkdownCode}`)();
  return {
    xml: globalThis.__mikuprojectXml,
    wbsMarkdown: globalThis.__mikuprojectWbsMarkdown
  };
}

function createDerivedTask(baseTask, overrides) {
  return {
    ...baseTask,
    ...overrides,
    milestone: false,
    summary: false,
    critical: false,
    percentWorkComplete: overrides.percentComplete,
    predecessors: [],
    extendedAttributes: [],
    baselines: [],
    timephasedData: []
  };
}

describe("mikuproject wbs markdown", () => {
  it("exports one markdown document with tree first and table after it", () => {
    const { xml, wbsMarkdown } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);

    const markdown = wbsMarkdown.exportWbsMarkdown(model, {
      useBusinessDaysForDisplayRange: true,
      useBusinessDaysForProgressBand: true
    });

    expect(markdown).toContain("# プロジェクト情報");
    expect(markdown).toContain("# WBS ツリー");
    expect(markdown).toContain("# WBS テーブル");
    expect(markdown).toContain("# サマリ");
    expect(markdown.indexOf("# WBS ツリー")).toBeLessThan(markdown.indexOf("# WBS テーブル"));
    expect(markdown.indexOf("# WBS テーブル")).toBeLessThan(markdown.indexOf("# サマリ"));
    expect(markdown).toContain("| プロジェクト名 | mikuproject開発 |");
    expect(markdown).toContain("~~~text");
    expect(markdown).toContain("1 基盤整備 (3/16 - 3/17): 100%");
    expect(markdown).toContain("┗　1.1 着手 (3/16): 100%");
    expect(markdown).toContain("┗　1.1");
    expect(markdown).toContain("| 1 | フェーズ | 1 | 基盤整備 | 2026-03-16 | 2026-03-17 |");
    expect(markdown).toContain("| WBS | 種別 | 階層 | 名称 | 開始 | 終了 | 期間 | タスク詳細 | 進捗 | 担当 | リソース | 先行 |");
  });

  it("shows notes in the tree section and summary after the table", () => {
    const { xml, wbsMarkdown } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);
    model.tasks[1].notes = "補足A\n補足B";

    const markdown = wbsMarkdown.exportWbsMarkdown(model, {
      displayDaysBeforeBaseDate: 1,
      displayDaysAfterBaseDate: 2,
      useBusinessDaysForDisplayRange: true,
      useBusinessDaysForProgressBand: true
    });

    expect(markdown).toContain("詳細: 補足A");
    expect(markdown).toContain("      補足B");
    expect(markdown).toContain("| 前日数 | 1 |");
    expect(markdown).toContain("| 後日数 | 2 |");
    expect(markdown).toContain("| 表示 | 営業日 |");
    expect(markdown).toContain("| 進捗 | 営業日 |");
  });

  it("keeps long names, deep hierarchy, and notes readable in the sample-oriented markdown", () => {
    const { xml, wbsMarkdown } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);
    const longNameTarget = model.tasks.find((task) => task.outlineNumber === "1.2");
    if (!longNameTarget) {
      throw new Error("Expected task 1.2 in sample model");
    }
    longNameTarget.name = "初期実装（MS Project XML 調査・基軸フォーマット選定・内部モデルの概要確定）";
    longNameTarget.notes = "前提と制約を整理する\n関係者レビューを実施する";
    const insertIndex = model.tasks.findIndex((task) => task.outlineNumber === "1.2");
    model.tasks.splice(
      insertIndex + 1,
      0,
      createDerivedTask(longNameTarget, {
        uid: "901",
        id: "901",
        name: "内部 JSON 形式への写像方針を確認する",
        outlineLevel: 3,
        outlineNumber: "1.2.1",
        wbs: "1.2.1",
        start: "2026-03-17T09:00:00",
        finish: "2026-03-18T18:00:00",
        duration: "PT16H0M0S",
        percentComplete: 40,
        notes: "説明責務と round-trip 観点を切り分ける"
      }),
      createDerivedTask(longNameTarget, {
        uid: "902",
        id: "902",
        name: "フィールド差分の洗い出し結果を整理する",
        outlineLevel: 4,
        outlineNumber: "1.2.1.1",
        wbs: "1.2.1.1",
        start: "2026-03-18T09:00:00",
        finish: "2026-03-18T18:00:00",
        duration: "PT8H0M0S",
        percentComplete: 10,
        notes: "XML と内部モデルの差分を箇条書きで残す"
      }),
      createDerivedTask(longNameTarget, {
        uid: "903",
        id: "903",
        name: "長い説明文の折り返しを確認するための task",
        outlineLevel: 5,
        outlineNumber: "1.2.1.1.1",
        wbs: "1.2.1.1.1",
        start: "2026-03-18T09:00:00",
        finish: "2026-03-19T18:00:00",
        duration: "PT16H0M0S",
        percentComplete: 0,
        notes: "かなり長い補足説明をここへ入れて、Markdown tree の見え方と table の見え方を同時に確認する"
      })
    );

    const markdown = wbsMarkdown.exportWbsMarkdown(model, {
      useBusinessDaysForDisplayRange: true,
      useBusinessDaysForProgressBand: true
    });

    expect(markdown).toContain("初期実装（MS Project XML 調査・基軸フォーマット選定・内部モデルの概要確定）");
    expect(markdown).toContain("　┗　1.2.1");
    expect(markdown).toContain("　　┗　1.2.1.1");
    expect(markdown).toContain("　　　┗　1.2.1.1.1");
    expect(markdown).toContain("詳細: 前提と制約を整理する");
    expect(markdown).toContain("      関係者レビューを実施する");
    expect(markdown).toContain("詳細: かなり長い補足説明をここへ入れて、Markdown tree の見え方と table の見え方を同時に確認する");
  });

  it("preserves multiline notes in tree and escapes markdown-sensitive text in table cells", () => {
    const { xml, wbsMarkdown } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);
    model.tasks[1].notes = "# 見出し風\n- 箇条書き風\n1. 番号付き風\nA | B <tag> & text";

    const markdown = wbsMarkdown.exportWbsMarkdown(model, {
      useBusinessDaysForDisplayRange: true,
      useBusinessDaysForProgressBand: true
    });

    expect(markdown).toContain("詳細: # 見出し風");
    expect(markdown).toContain("      - 箇条書き風");
    expect(markdown).toContain("      1. 番号付き風");
    expect(markdown).toContain("\\# 見出し風<br>\\- 箇条書き風<br>1\\. 番号付き風<br>A \\\\| B &lt;tag&gt; &amp; text");
  });

  it("uses a fence that does not break when tree text includes backticks", () => {
    const { xml, wbsMarkdown } = bootModules();
    const model = xml.importMsProjectXml(xml.SAMPLE_XML);
    model.tasks[1].name = "``` fenced name";
    model.tasks[1].notes = "line 1\n``` fenced note";

    const markdown = wbsMarkdown.exportWbsMarkdown(model, {
      useBusinessDaysForDisplayRange: true,
      useBusinessDaysForProgressBand: true
    });

    expect(markdown).toContain("~~~text");
    expect(markdown).toContain("``` fenced name");
    expect(markdown).toContain("``` fenced note");
  });
});
