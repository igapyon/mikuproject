/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  type ImportWarning = {
    message: string;
    scope?: "project" | "tasks" | "resources" | "assignments" | "calendars";
    uid?: string;
    label?: string;
  };

  type ImportChange = {
    scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
    uid: string;
    label: string;
    field: string;
    before: string | number | boolean | undefined;
    after: string | number | boolean;
  };

  const mainImport = {
    importProjectDraftText(input: {
      sourceText: string;
      extractLastJsonBlock: (value: string) => string;
      importProjectDraftView: (draft: unknown) => ProjectModel;
      validateProjectModel: (model: ProjectModel) => ValidationIssue[];
    }): { model: ProjectModel; issues: ValidationIssue[] } {
      const sourceText = input.sourceText.trim();
      if (!sourceText) {
        throw new Error("project_draft_view JSON を入力してください");
      }
      const draft = JSON.parse(input.extractLastJsonBlock(sourceText));
      const model = input.importProjectDraftView(draft);
      return {
        model,
        issues: input.validateProjectModel(model)
      };
    },

    importPatchJsonText(input: {
      sourceText: string;
      ensureCurrentModel: () => ProjectModel;
      extractLastJsonBlock: (value: string) => string;
      importProjectPatchJson: (documentLike: unknown, baseModel: ProjectModel) => {
        model: ProjectModel;
        changes: ImportChange[];
        warnings: ImportWarning[];
      };
      exportMsProjectXml: (model: ProjectModel) => string;
      validateProjectModel: (model: ProjectModel) => ValidationIssue[];
    }): {
      model: ProjectModel;
      issues: ValidationIssue[];
      changes: ImportChange[];
      warnings: ImportWarning[];
      regeneratedXmlText?: string;
    } {
      const sourceText = input.sourceText.trim();
      if (!sourceText) {
        throw new Error("Patch JSON を入力してください");
      }
      const documentLike = JSON.parse(input.extractLastJsonBlock(sourceText));
      const result = input.importProjectPatchJson(documentLike, input.ensureCurrentModel());
      return {
        model: result.model,
        issues: input.validateProjectModel(result.model),
        changes: result.changes,
        warnings: result.warnings,
        regeneratedXmlText: result.changes.length > 0
          ? input.exportMsProjectXml(result.model)
          : undefined
      };
    },

    detectAiEditJsonKind(input: {
      sourceText: string;
      extractLastJsonBlock: (value: string) => string;
      detectJsonDocumentKind: (documentLike: unknown) => "workbook_json" | "project_draft_view" | "patch_json" | undefined;
    }): "project_draft_view" | "patch_json" {
      const sourceText = input.sourceText.trim();
      if (!sourceText) {
        throw new Error("project_draft_view または Patch JSON を入力してください");
      }
      const documentLike = JSON.parse(input.extractLastJsonBlock(sourceText));
      const kind = input.detectJsonDocumentKind(documentLike);
      if (kind === "project_draft_view" || kind === "patch_json") {
        return kind;
      }
      throw new Error("project_draft_view または Patch JSON を入力してください");
    },

    importWorkbookJsonText(input: {
      sourceText: string;
      previousModel: ProjectModel | null;
      extractLastJsonBlock: (value: string) => string;
      importProjectWorkbookJsonAsProjectModel: (documentLike: unknown) => {
        model: ProjectModel;
        warnings: ImportWarning[];
      };
      importProjectWorkbookJson: (documentLike: unknown, baseModel: ProjectModel) => {
        model: ProjectModel;
        changes: ImportChange[];
        warnings: ImportWarning[];
      };
      validateProjectModel: (model: ProjectModel) => ValidationIssue[];
    }): {
      model: ProjectModel;
      issues: ValidationIssue[];
      changes: ImportChange[];
      warnings: ImportWarning[];
    } {
      const sourceText = input.sourceText.trim();
      if (!sourceText) {
        throw new Error("workbook JSON を入力してください");
      }
      const documentLike = JSON.parse(input.extractLastJsonBlock(sourceText));
      const replaceResult = input.importProjectWorkbookJsonAsProjectModel(documentLike);
      const diffResult = input.previousModel
        ? input.importProjectWorkbookJson(documentLike, input.previousModel)
        : { changes: [], warnings: replaceResult.warnings };
      return {
        model: replaceResult.model,
        issues: input.validateProjectModel(replaceResult.model),
        changes: diffResult.changes,
        warnings: replaceResult.warnings
      };
    },

    async importXlsxBytes(input: {
      bytes: Uint8Array;
      previousModel: ProjectModel | null;
      createWorkbookCodec: () => {
        importWorkbook: (bytes: Uint8Array) => unknown;
        importWorkbookAsync?: (bytes: Uint8Array) => Promise<unknown>;
      };
      importProjectWorkbookAsProjectModel: (workbook: unknown) => ProjectModel;
      importProjectWorkbookDetailed: (workbook: unknown, baseModel: ProjectModel) => {
        model: ProjectModel;
        changes: ImportChange[];
      };
      validateProjectModel: (model: ProjectModel) => ValidationIssue[];
    }): Promise<{
      model: ProjectModel;
      issues: ValidationIssue[];
      changes: ImportChange[];
    }> {
      const codec = input.createWorkbookCodec();
      const workbook = typeof codec.importWorkbookAsync === "function"
        ? await codec.importWorkbookAsync(input.bytes)
        : codec.importWorkbook(input.bytes);
      const model = input.importProjectWorkbookAsProjectModel(workbook);
      const diffResult = input.previousModel
        ? input.importProjectWorkbookDetailed(workbook, input.previousModel)
        : { changes: [] };
      return {
        model,
        issues: input.validateProjectModel(model),
        changes: diffResult.changes
      };
    },

    importCsvText(input: {
      csvText: string;
      importCsvParentId: (csvText: string) => ProjectModel;
      validateProjectModel: (model: ProjectModel) => ValidationIssue[];
    }): { model: ProjectModel; issues: ValidationIssue[] } {
      const csvText = input.csvText.trim();
      if (!csvText) {
        throw new Error("CSV が空です");
      }
      const model = input.importCsvParentId(csvText);
      return {
        model,
        issues: input.validateProjectModel(model)
      };
    }
  };

  (globalThis as typeof globalThis & {
    __mikuprojectMainImport?: typeof mainImport;
  }).__mikuprojectMainImport = mainImport;
})();
