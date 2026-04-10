/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  type ImportChange = {
    scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
    uid: string;
    label: string;
    field: string;
    before: string | number | boolean | undefined;
    after: string | number | boolean;
  };

  type ScopedWarning = {
    message: string;
    scope?: "project" | "tasks" | "resources" | "assignments" | "calendars";
    uid?: string;
    label?: string;
  };

  type AiJsonDocumentKind = "workbook_json" | "project_draft_view" | "patch_json";
  type ExternalImportFormat = "ms_project_xml" | "xlsx" | "workbook_json" | "project_draft_view" | "patch_json";
  type ExternalImportMode = "replace" | "merge" | "patch";
  type ReportDisplayRangeOptions = {
    holidayDates?: string[];
    displayDaysBeforeBaseDate?: number;
    displayDaysAfterBaseDate?: number;
    useBusinessDaysForDisplayRange?: boolean;
  };
  type ReportProgressOptions = ReportDisplayRangeOptions & {
    useBusinessDaysForProgressBand?: boolean;
  };

  const mikuprojectXml = (globalThis as typeof globalThis & {
    __mikuprojectXml?: {
      SAMPLE_XML: string;
      SAMPLE_PROJECT_DRAFT_VIEW: unknown;
      parseXmlDocument: (xmlText: string) => XMLDocument;
      importMsProjectXml: (xmlText: string) => ProjectModel;
      importCsvParentId: (csvText: string) => ProjectModel;
      exportMsProjectXml: (model: ProjectModel) => string;
      exportCsvParentId: (model: ProjectModel) => string;
      buildProjectDraftRequest: (input: {
        name: string;
        plannedStart?: string;
        goal?: string;
        teamCount?: number;
        mustHavePhases?: string[];
        mustHaveMilestones?: string[];
      }) => unknown;
      importProjectDraftView: (draft: unknown) => ProjectModel;
      exportProjectOverviewView: (model: ProjectModel) => unknown;
      exportTaskEditView: (model: ProjectModel, taskUid?: string) => unknown;
      exportPhaseDetailView: (
        model: ProjectModel,
        phaseUid?: string,
        options?: {
          mode?: "full" | "scoped";
          rootUid?: string;
          maxDepth?: number;
        }
      ) => unknown;
      normalizeProjectModel: (model: ProjectModel) => ProjectModel;
      validateProjectModel: (model: ProjectModel) => ValidationIssue[];
    };
  }).__mikuprojectXml;

  if (!mikuprojectXml) {
    throw new Error("mikuproject XML module is not loaded");
  }

  const mikuprojectAiJsonUtil = (globalThis as typeof globalThis & {
    __mikuprojectAiJsonUtil?: {
      extractLastJsonBlock: (value: string) => string;
      detectJsonDocumentKind: (documentLike: unknown) => AiJsonDocumentKind | undefined;
    };
  }).__mikuprojectAiJsonUtil;

  if (!mikuprojectAiJsonUtil) {
    throw new Error("mikuproject AI JSON util module is not loaded");
  }

  const mikuprojectAiJsonSpec = (globalThis as typeof globalThis & {
    __mikuprojectAiJsonSpec?: {
      getAiJsonSpecText: () => string;
      getAiJsonSpec: () => {
        id: "mikuproject-ai-json-spec";
        version: string;
        text: string;
      };
    };
  }).__mikuprojectAiJsonSpec;

  if (!mikuprojectAiJsonSpec) {
    throw new Error("mikuproject AI JSON spec module is not loaded");
  }

  const mikuprojectMainUtil = (globalThis as typeof globalThis & {
    __mikuprojectMainUtil?: {
      encodeUtf8: (value: string) => Uint8Array;
      packZipEntries: (entries: Array<{ name: string; data: Uint8Array }>) => Uint8Array;
    };
  }).__mikuprojectMainUtil;

  if (!mikuprojectMainUtil) {
    throw new Error("mikuproject main util module is not loaded");
  }

  const mikuprojectProjectWorkbookJson = (globalThis as typeof globalThis & {
    __mikuprojectProjectWorkbookJson?: {
      exportProjectWorkbookJson: (model: ProjectModel) => unknown;
      importProjectWorkbookJsonAsProjectModel: (documentLike: unknown) => {
        model: ProjectModel;
        warnings: Array<{ message: string }>;
      };
      importProjectWorkbookJson: (documentLike: unknown, baseModel: ProjectModel) => {
        model: ProjectModel;
        changes: ImportChange[];
        warnings: Array<{ message: string }>;
      };
      validateWorkbookJsonDocument: (documentLike: unknown) => {
        document: unknown;
        warnings: Array<{ message: string }>;
      };
    };
  }).__mikuprojectProjectWorkbookJson;

  if (!mikuprojectProjectWorkbookJson) {
    throw new Error("mikuproject Project Workbook JSON module is not loaded");
  }

  type XlsxCellLike = {
    value?: string | number | boolean;
    numberFormat?: "general" | "integer" | "decimal" | "date" | "datetime" | "percent";
    horizontalAlign?: "left" | "center" | "right";
    bold?: boolean;
    fontSize?: number;
    fillColor?: string;
    border?: "thin";
  };

  type XlsxWorkbookLike = {
    sheets: Array<{
      name: string;
      columns?: Array<{ width?: number }>;
      mergedRanges?: string[];
      dataValidations?: Array<{
        type: "list";
        sqref: string;
        formula1: string;
        allowBlank?: boolean;
      }>;
      rows: Array<{
        height?: number;
        cells: XlsxCellLike[];
      }>;
    }>;
  };

  const mikuprojectProjectXlsx = (globalThis as typeof globalThis & {
    __mikuprojectProjectXlsx?: {
      exportProjectWorkbook: (model: ProjectModel) => XlsxWorkbookLike;
      importProjectWorkbook: (workbook: XlsxWorkbookLike, baseModel: ProjectModel) => ProjectModel;
      importProjectWorkbookAsProjectModel: (workbook: XlsxWorkbookLike) => ProjectModel;
      importProjectWorkbookDetailed?: (workbook: XlsxWorkbookLike, baseModel: ProjectModel) => {
        model: ProjectModel;
        changes: ImportChange[];
      };
    };
  }).__mikuprojectProjectXlsx;

  if (!mikuprojectProjectXlsx) {
    throw new Error("mikuproject Project XLSX module is not loaded");
  }

  const mikuprojectExcelIo = (globalThis as typeof globalThis & {
    __mikuprojectExcelIo?: {
      XlsxWorkbookCodec: new () => {
        exportWorkbook: (workbook: XlsxWorkbookLike) => Uint8Array;
        importWorkbook: (bytes: Uint8Array) => XlsxWorkbookLike;
      };
    };
  }).__mikuprojectExcelIo;

  if (!mikuprojectExcelIo) {
    throw new Error("mikuproject Excel IO module is not loaded");
  }

  const mikuprojectWbsXlsx = (globalThis as typeof globalThis & {
    __mikuprojectWbsXlsx?: {
      exportWbsWorkbook: (model: ProjectModel, options?: ReportProgressOptions) => XlsxWorkbookLike;
    };
  }).__mikuprojectWbsXlsx;

  if (!mikuprojectWbsXlsx) {
    throw new Error("mikuproject WBS XLSX module is not loaded");
  }

  const mikuprojectNativeSvg = (globalThis as typeof globalThis & {
    __mikuprojectNativeSvg?: {
      exportNativeSvg: (
        model: ProjectModel,
        options?: ReportProgressOptions & { labelMode?: "near" | "list" }
      ) => string;
      exportWeeklyNativeSvg: (model: ProjectModel, options?: ReportDisplayRangeOptions) => string;
      exportMonthlyWbsCalendarSvgArchive: (model: ProjectModel) => {
        entries: Array<{ fileName: string; svg: string }>;
        zipBytes: Uint8Array;
      };
    };
  }).__mikuprojectNativeSvg;

  if (!mikuprojectNativeSvg) {
    throw new Error("mikuproject native SVG module is not loaded");
  }

  const mikuprojectWbsMarkdown = (globalThis as typeof globalThis & {
    __mikuprojectWbsMarkdown?: {
      exportWbsMarkdown: (model: ProjectModel, options?: ReportProgressOptions) => string;
    };
  }).__mikuprojectWbsMarkdown;

  if (!mikuprojectWbsMarkdown) {
    throw new Error("mikuproject WBS Markdown module is not loaded");
  }

  const mikuprojectProjectPatchJson = (globalThis as typeof globalThis & {
    __mikuprojectProjectPatchJson?: {
      importProjectPatchJson: (documentLike: unknown, baseModel: ProjectModel) => {
        model: ProjectModel;
        changes: ImportChange[];
        warnings: ScopedWarning[];
      };
      validatePatchDocument: (documentLike: unknown) => {
        document: unknown;
        warnings: ScopedWarning[];
      };
    };
  }).__mikuprojectProjectPatchJson;

  if (!mikuprojectProjectPatchJson) {
    throw new Error("mikuproject Project Patch JSON module is not loaded");
  }

  function getAiJsonSpecText(): string {
    return mikuprojectAiJsonSpec.getAiJsonSpecText();
  }

  function getAiJsonSpec(): {
    id: "mikuproject-ai-json-spec";
    version: string;
    text: string;
  } {
    return mikuprojectAiJsonSpec.getAiJsonSpec();
  }

  function detectAiJsonDocumentKind(documentLike: unknown): AiJsonDocumentKind | undefined {
    return mikuprojectAiJsonUtil.detectJsonDocumentKind(documentLike);
  }

  function parseAiJsonText(sourceText: string): {
    sourceText: string;
    jsonText: string;
    document: unknown;
    kind: AiJsonDocumentKind | undefined;
  } {
    const jsonText = mikuprojectAiJsonUtil.extractLastJsonBlock(sourceText);
    const document = JSON.parse(jsonText);
    return {
      sourceText,
      jsonText,
      document,
      kind: detectAiJsonDocumentKind(document)
    };
  }

  function importAiJsonDocument(documentLike: unknown, options: {
    baseModel?: ProjectModel;
  } = {}):
    | {
      kind: "project_draft_view";
      mode: "replace";
      model: ProjectModel;
      warnings: [];
    }
    | {
      kind: "workbook_json";
      mode: "replace";
      model: ProjectModel;
      warnings: Array<{ message: string }>;
    }
    | {
      kind: "workbook_json";
      mode: "merge";
      model: ProjectModel;
      changes: ImportChange[];
      warnings: Array<{ message: string }>;
    }
    | {
      kind: "patch_json";
      mode: "patch";
      model: ProjectModel;
      changes: ImportChange[];
      warnings: ScopedWarning[];
    } {
    const kind = detectAiJsonDocumentKind(documentLike);
    if (kind === "project_draft_view") {
      return {
        kind,
        mode: "replace",
        model: mikuprojectXml.importProjectDraftView(documentLike),
        warnings: []
      };
    }

    if (kind === "workbook_json") {
      if (options.baseModel) {
        return {
          kind,
          mode: "merge",
          ...mikuprojectProjectWorkbookJson.importProjectWorkbookJson(documentLike, options.baseModel)
        };
      }
      return {
        kind,
        mode: "replace",
        ...mikuprojectProjectWorkbookJson.importProjectWorkbookJsonAsProjectModel(documentLike)
      };
    }

    if (kind === "patch_json") {
      if (!options.baseModel) {
        throw new Error("Patch JSON の適用には baseModel が必要です");
      }
      return {
        kind,
        mode: "patch",
        ...mikuprojectProjectPatchJson.importProjectPatchJson(documentLike, options.baseModel)
      };
    }

    throw new Error("AI JSON の format / view_type を判別できません");
  }

  function importAiJsonText(sourceText: string, options: {
    baseModel?: ProjectModel;
  } = {}): {
    sourceText: string;
    jsonText: string;
    document: unknown;
    kind: AiJsonDocumentKind | undefined;
    result: ReturnType<typeof importAiJsonDocument>;
  } {
    const parsed = parseAiJsonText(sourceText);
    return {
      ...parsed,
      result: importAiJsonDocument(parsed.document, options)
    };
  }

  type ExternalImportSource =
    | { format: "ms_project_xml"; text: string }
    | { format: "xlsx"; bytes: Uint8Array }
    | { format: "workbook_json"; document: unknown }
    | { format: "project_draft_view"; document: unknown }
    | { format: "patch_json"; document: unknown };

  function importExternal(input: {
    source: ExternalImportSource;
    mode: ExternalImportMode;
    baseModel?: ProjectModel;
  }):
    | {
      kind: "ms_project_xml";
      mode: "replace";
      model: ProjectModel;
      warnings: [];
    }
    | {
      kind: "xlsx";
      mode: "replace";
      model: ProjectModel;
      warnings: [];
    }
    | {
      kind: "xlsx";
      mode: "merge";
      model: ProjectModel;
      changes: ImportChange[];
      warnings: [];
    }
    | ReturnType<typeof importAiJsonDocument> {
    const { source, mode, baseModel } = input;

    if (source.format === "ms_project_xml") {
      if (mode !== "replace") {
        throw new Error("MS Project XML は replace import のみ対応です");
      }
      return {
        kind: "ms_project_xml",
        mode,
        model: mikuprojectXml.importMsProjectXml(source.text),
        warnings: []
      };
    }

    if (source.format === "xlsx") {
      if (mode === "patch") {
        throw new Error("XLSX は replace または merge import のみ対応です");
      }
      const workbook = new mikuprojectExcelIo.XlsxWorkbookCodec().importWorkbook(source.bytes);
      if (mode === "replace") {
        return {
          kind: "xlsx",
          mode,
          model: mikuprojectProjectXlsx.importProjectWorkbookAsProjectModel(workbook),
          warnings: []
        };
      }
      if (mode === "merge") {
        if (!baseModel) {
          throw new Error("XLSX merge import には baseModel が必要です");
        }
        if (typeof mikuprojectProjectXlsx.importProjectWorkbookDetailed === "function") {
          return {
            kind: "xlsx",
            mode,
            ...mikuprojectProjectXlsx.importProjectWorkbookDetailed(workbook, baseModel),
            warnings: []
          };
        }
        return {
          kind: "xlsx",
          mode,
          model: mikuprojectProjectXlsx.importProjectWorkbook(workbook, baseModel),
          changes: [],
          warnings: []
        };
      }
    }

    if (source.format === "workbook_json") {
      if (mode === "patch") {
        throw new Error("workbook JSON は patch import に対応していません");
      }
      if (mode === "merge" && !baseModel) {
        throw new Error("workbook JSON merge import には baseModel が必要です");
      }
      return importAiJsonDocument(source.document, mode === "merge" ? { baseModel } : {});
    }

    if (source.format === "project_draft_view") {
      if (mode !== "replace") {
        throw new Error("project_draft_view は replace import のみ対応です");
      }
      return importAiJsonDocument(source.document);
    }

    if (source.format === "patch_json") {
      if (mode !== "patch") {
        throw new Error("patch JSON は patch import のみ対応です");
      }
      return importAiJsonDocument(source.document, { baseModel });
    }

    throw new Error(`未対応の import format です: ${(source as { format?: string }).format || "unknown"}`);
  }

  function exportAllReportEntries(
    model: ProjectModel,
    options: ReportProgressOptions = {}
  ): Array<{ name: string; data: Uint8Array }> {
    const codec = new mikuprojectExcelIo.XlsxWorkbookCodec();
    const monthlyArchive = mikuprojectNativeSvg.exportMonthlyWbsCalendarSvgArchive(model);
    const entries = [
      {
        name: "wbs.xlsx",
        data: codec.exportWorkbook(mikuprojectWbsXlsx.exportWbsWorkbook(model, options))
      },
      {
        name: "wbs.md",
        data: mikuprojectMainUtil.encodeUtf8(`${mikuprojectWbsMarkdown.exportWbsMarkdown(model, options)}\n`)
      },
      {
        name: "mermaid.mmd",
        data: mikuprojectMainUtil.encodeUtf8(`${mikuprojectXml.exportMermaidGantt(model)}\n`)
      },
      {
        name: "daily.svg",
        data: mikuprojectMainUtil.encodeUtf8(`${mikuprojectNativeSvg.exportNativeSvg(model, options)}\n`)
      },
      {
        name: "weekly.svg",
        data: mikuprojectMainUtil.encodeUtf8(`${mikuprojectNativeSvg.exportWeeklyNativeSvg(model, options)}\n`)
      }
    ];

    for (const entry of monthlyArchive.entries) {
      entries.push({
        name: `monthly-calendar/${entry.fileName}`,
        data: mikuprojectMainUtil.encodeUtf8(entry.svg)
      });
    }

    return entries;
  }

  (globalThis as typeof globalThis & {
    __mikuprojectCoreApi?: {
      version: 1;
      getAiJsonSpecText: () => string;
      getAiJsonSpec: () => {
        id: "mikuproject-ai-json-spec";
        version: string;
        text: string;
      };
      detectAiJsonDocumentKind: (documentLike: unknown) => AiJsonDocumentKind | undefined;
      parseAiJsonText: (sourceText: string) => {
        sourceText: string;
        jsonText: string;
        document: unknown;
        kind: AiJsonDocumentKind | undefined;
      };
      importAiJsonDocument: typeof importAiJsonDocument;
      importAiJsonText: typeof importAiJsonText;
      importExternal: typeof importExternal;
      samples: {
        getSampleXml: () => string;
        getSampleProjectDraftView: () => unknown;
      };
      projectModel: {
        normalize: (model: ProjectModel) => ProjectModel;
        validate: (model: ProjectModel) => ValidationIssue[];
      };
      msProject: {
        parseXmlDocument: (xmlText: string) => XMLDocument;
        importFromXml: (xmlText: string) => ProjectModel;
        exportToXml: (model: ProjectModel) => string;
        importFromCsvParentId: (csvText: string) => ProjectModel;
        exportToCsvParentId: (model: ProjectModel) => string;
      };
      aiViews: {
        buildProjectDraftRequest: (input: {
          name: string;
          plannedStart?: string;
          goal?: string;
          teamCount?: number;
          mustHavePhases?: string[];
          mustHaveMilestones?: string[];
        }) => unknown;
        importProjectDraftView: (draft: unknown) => ProjectModel;
        exportProjectOverviewView: (model: ProjectModel) => unknown;
        exportTaskEditView: (model: ProjectModel, taskUid?: string) => unknown;
        exportPhaseDetailView: (
          model: ProjectModel,
          phaseUid?: string,
          options?: {
            mode?: "full" | "scoped";
            rootUid?: string;
            maxDepth?: number;
          }
        ) => unknown;
      };
      workbookJson: {
        exportDocument: (model: ProjectModel) => unknown;
        validateDocument: (documentLike: unknown) => {
          document: unknown;
          warnings: Array<{ message: string }>;
        };
        importAsProjectModel: (documentLike: unknown) => {
          model: ProjectModel;
          warnings: Array<{ message: string }>;
        };
        importIntoProjectModel: (documentLike: unknown, baseModel: ProjectModel) => {
          model: ProjectModel;
          changes: ImportChange[];
          warnings: Array<{ message: string }>;
        };
      };
      xlsx: {
        decodeWorkbook: (bytes: Uint8Array) => XlsxWorkbookLike;
        encodeWorkbook: (workbook: XlsxWorkbookLike) => Uint8Array;
        exportWorkbook: (model: ProjectModel) => XlsxWorkbookLike;
        importAsProjectModel: (workbook: XlsxWorkbookLike) => ProjectModel;
        importIntoProjectModel: (workbook: XlsxWorkbookLike, baseModel: ProjectModel) => ProjectModel;
      };
      patchJson: {
        validateDocument: (documentLike: unknown) => {
          document: unknown;
          warnings: ScopedWarning[];
        };
        applyToProjectModel: (documentLike: unknown, baseModel: ProjectModel) => {
          model: ProjectModel;
          changes: ImportChange[];
          warnings: ScopedWarning[];
        };
      };
      report: {
        all: {
          export: (model: ProjectModel, options?: ReportProgressOptions) => {
            entries: Array<{ name: string; data: Uint8Array }>;
            zipBytes: Uint8Array;
          };
        };
        wbsXlsx: {
          exportWorkbook: (model: ProjectModel, options?: ReportProgressOptions) => XlsxWorkbookLike;
          exportBytes: (model: ProjectModel, options?: ReportProgressOptions) => Uint8Array;
        };
        svg: {
          exportDaily: (
            model: ProjectModel,
            options?: ReportProgressOptions & { labelMode?: "near" | "list" }
          ) => string;
          exportWeekly: (model: ProjectModel, options?: ReportDisplayRangeOptions) => string;
          exportMonthlyCalendar: (model: ProjectModel) => {
            entries: Array<{ fileName: string; svg: string }>;
            zipBytes: Uint8Array;
          };
        };
        wbsMarkdown: {
          export: (model: ProjectModel, options?: ReportProgressOptions) => string;
        };
        mermaid: {
          exportGantt: (model: ProjectModel) => string;
        };
      };
    };
  }).__mikuprojectCoreApi = {
    version: 1,
    getAiJsonSpecText,
    getAiJsonSpec,
    detectAiJsonDocumentKind,
    parseAiJsonText,
    importAiJsonDocument,
    importAiJsonText,
    importExternal,
    samples: {
      getSampleXml: () => mikuprojectXml.SAMPLE_XML,
      getSampleProjectDraftView: () => mikuprojectXml.SAMPLE_PROJECT_DRAFT_VIEW
    },
    projectModel: {
      normalize: mikuprojectXml.normalizeProjectModel,
      validate: mikuprojectXml.validateProjectModel
    },
    msProject: {
      parseXmlDocument: mikuprojectXml.parseXmlDocument,
      importFromXml: mikuprojectXml.importMsProjectXml,
      exportToXml: mikuprojectXml.exportMsProjectXml,
      importFromCsvParentId: mikuprojectXml.importCsvParentId,
      exportToCsvParentId: mikuprojectXml.exportCsvParentId
    },
    aiViews: {
      buildProjectDraftRequest: mikuprojectXml.buildProjectDraftRequest,
      importProjectDraftView: mikuprojectXml.importProjectDraftView,
      exportProjectOverviewView: mikuprojectXml.exportProjectOverviewView,
      exportTaskEditView: mikuprojectXml.exportTaskEditView,
      exportPhaseDetailView: mikuprojectXml.exportPhaseDetailView
    },
    workbookJson: {
      exportDocument: mikuprojectProjectWorkbookJson.exportProjectWorkbookJson,
      validateDocument: mikuprojectProjectWorkbookJson.validateWorkbookJsonDocument,
      importAsProjectModel: mikuprojectProjectWorkbookJson.importProjectWorkbookJsonAsProjectModel,
      importIntoProjectModel: mikuprojectProjectWorkbookJson.importProjectWorkbookJson
    },
    xlsx: {
      decodeWorkbook: (bytes: Uint8Array) => new mikuprojectExcelIo.XlsxWorkbookCodec().importWorkbook(bytes),
      encodeWorkbook: (workbook: XlsxWorkbookLike) => new mikuprojectExcelIo.XlsxWorkbookCodec().exportWorkbook(workbook),
      exportWorkbook: mikuprojectProjectXlsx.exportProjectWorkbook,
      importAsProjectModel: mikuprojectProjectXlsx.importProjectWorkbookAsProjectModel,
      importIntoProjectModel: mikuprojectProjectXlsx.importProjectWorkbook
    },
    patchJson: {
      validateDocument: mikuprojectProjectPatchJson.validatePatchDocument,
      applyToProjectModel: mikuprojectProjectPatchJson.importProjectPatchJson
    },
    report: {
      all: {
        export: (model: ProjectModel, options: ReportProgressOptions = {}) => {
          const entries = exportAllReportEntries(model, options);
          return {
            entries,
            zipBytes: mikuprojectMainUtil.packZipEntries(entries)
          };
        }
      },
      wbsXlsx: {
        exportWorkbook: mikuprojectWbsXlsx.exportWbsWorkbook,
        exportBytes: (model: ProjectModel, options: ReportProgressOptions = {}) =>
          new mikuprojectExcelIo.XlsxWorkbookCodec().exportWorkbook(mikuprojectWbsXlsx.exportWbsWorkbook(model, options))
      },
      svg: {
        exportDaily: mikuprojectNativeSvg.exportNativeSvg,
        exportWeekly: mikuprojectNativeSvg.exportWeeklyNativeSvg,
        exportMonthlyCalendar: mikuprojectNativeSvg.exportMonthlyWbsCalendarSvgArchive
      },
      wbsMarkdown: {
        export: mikuprojectWbsMarkdown.exportWbsMarkdown
      },
      mermaid: {
        exportGantt: mikuprojectXml.exportMermaidGantt
      }
    }
  };
})();
