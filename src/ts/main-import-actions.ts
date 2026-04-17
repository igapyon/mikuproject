/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  const mainImportActions = {
    async importProjectDraftText(input: {
      sourceText: string;
      importProjectDraftText: (input: {
        sourceText: string;
        extractLastJsonBlock: (value: string) => string;
        importProjectDraftView: (draft: unknown) => ProjectModel;
        validateProjectModel: (model: ProjectModel) => ValidationIssue[];
      }) => { model: ProjectModel; issues: ValidationIssue[] };
      extractLastJsonBlock: (value: string) => string;
      importProjectDraftView: (draft: unknown) => ProjectModel;
      validateProjectModel: (model: ProjectModel) => ValidationIssue[];
      setCurrentModel: (model: ProjectModel) => void;
      syncXmlTextFromModel: (model: ProjectModel) => string;
      applyModelState: (input: { model: ProjectModel; issues: ValidationIssue[] }) => void;
      exportCurrentMermaid: (options?: { silent?: boolean }) => Promise<void>;
      completeTransform: (statusMessage: string, toastMessage: string) => void;
    }): Promise<void> {
      const result = input.importProjectDraftText({
        sourceText: input.sourceText,
        extractLastJsonBlock: input.extractLastJsonBlock,
        importProjectDraftView: input.importProjectDraftView,
        validateProjectModel: input.validateProjectModel
      });
      input.setCurrentModel(result.model);
      input.syncXmlTextFromModel(result.model);
      input.applyModelState({ model: result.model, issues: result.issues });
      await input.exportCurrentMermaid({ silent: true });
      input.completeTransform(
        result.issues.length > 0
          ? `project_draft_view を取り込みました。検証で ${result.issues.length} 件の問題があります`
          : "project_draft_view を取り込みました",
        "project_draft_view を取り込みました"
      );
    },

    async importPatchJsonText(input: {
      sourceText: string;
      importPatchJsonText: (input: {
        sourceText: string;
        ensureCurrentModel: () => ProjectModel;
        extractLastJsonBlock: (value: string) => string;
        importProjectPatchJson: (documentLike: unknown, baseModel: ProjectModel) => {
          model: ProjectModel;
          changes: Array<{
            scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
            uid: string;
            label: string;
            field: string;
            before: string | number | boolean | undefined;
            after: string | number | boolean;
          }>;
          warnings: Array<{ message: string; scope?: "project" | "tasks" | "resources" | "assignments" | "calendars"; uid?: string; label?: string }>;
        };
        exportMsProjectXml: (model: ProjectModel) => string;
        validateProjectModel: (model: ProjectModel) => ValidationIssue[];
      }) => {
        model: ProjectModel;
        issues: ValidationIssue[];
        changes: Array<{
          scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
          uid: string;
          label: string;
          field: string;
          before: string | number | boolean | undefined;
          after: string | number | boolean;
        }>;
        warnings: Array<{ message: string; scope?: "project" | "tasks" | "resources" | "assignments" | "calendars"; uid?: string; label?: string }>;
        regeneratedXmlText?: string;
      };
      ensureCurrentModel: () => ProjectModel;
      extractLastJsonBlock: (value: string) => string;
      importProjectPatchJson: (documentLike: unknown, baseModel: ProjectModel) => {
        model: ProjectModel;
        changes: Array<{
          scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
          uid: string;
          label: string;
          field: string;
          before: string | number | boolean | undefined;
          after: string | number | boolean;
        }>;
        warnings: Array<{ message: string; scope?: "project" | "tasks" | "resources" | "assignments" | "calendars"; uid?: string; label?: string }>;
      };
      exportMsProjectXml: (model: ProjectModel) => string;
      validateProjectModel: (model: ProjectModel) => ValidationIssue[];
      setCurrentModel: (model: ProjectModel) => void;
      applyModelState: (input: {
        model: ProjectModel;
        issues: ValidationIssue[];
        warnings: Array<{ message: string; scope?: "project" | "tasks" | "resources" | "assignments" | "calendars"; uid?: string; label?: string }>;
        changes: Array<{
          scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
          uid: string;
          label: string;
          field: string;
          before: string | number | boolean | undefined;
          after: string | number | boolean;
        }>;
      }) => void;
      writeXmlText: (xmlText: string) => void;
      markXmlDirty: () => void;
      setXmlSourceDirty: (dirty: boolean) => void;
      exportCurrentMermaid: (options?: { silent?: boolean }) => Promise<void>;
      completeTransform: (statusMessage: string, toastMessage: string) => void;
    }): Promise<void> {
      const result = input.importPatchJsonText({
        sourceText: input.sourceText,
        ensureCurrentModel: input.ensureCurrentModel,
        extractLastJsonBlock: input.extractLastJsonBlock,
        importProjectPatchJson: input.importProjectPatchJson,
        exportMsProjectXml: input.exportMsProjectXml,
        validateProjectModel: input.validateProjectModel
      });
      input.setCurrentModel(result.model);
      input.applyModelState({
        model: result.model,
        issues: result.issues,
        warnings: result.warnings,
        changes: result.changes
      });
      if (result.regeneratedXmlText) {
        input.writeXmlText(result.regeneratedXmlText);
        input.markXmlDirty();
      }
      input.setXmlSourceDirty(false);
      const summaryText = result.changes.length > 0
        ? `Patch JSON を読み込んで ${result.changes.length} 件の変更を反映しました。XML は再生成済みで、必要なら XML Export で保存できます`
        : "Patch JSON に反映対象の変更はありませんでした。XML は未変更です";
      const warningText = result.warnings.length > 0 ? `。Patch JSON 取込で ${result.warnings.length} 件の warning を無視しました` : "";
      input.completeTransform(
        result.issues.length > 0 ? `${summaryText}${warningText}。検証で ${result.issues.length} 件の問題があります` : `${summaryText}${warningText}`,
        "Patch JSON を反映しました"
      );
      await input.exportCurrentMermaid({ silent: true });
    },

    async importWorkbookJsonText(input: {
      sourceText: string;
      importWorkbookJsonText: (input: {
        sourceText: string;
        previousModel: ProjectModel | null;
        extractLastJsonBlock: (value: string) => string;
        importProjectWorkbookJsonAsProjectModel: (documentLike: unknown) => {
          model: ProjectModel;
          warnings: Array<{ message: string; scope?: "project" | "tasks" | "resources" | "assignments" | "calendars"; uid?: string; label?: string }>;
        };
        importProjectWorkbookJson: (documentLike: unknown, baseModel: ProjectModel) => {
          model: ProjectModel;
          changes: Array<{
            scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
            uid: string;
            label: string;
            field: string;
            before: string | number | boolean | undefined;
            after: string | number | boolean;
          }>;
          warnings: Array<{ message: string; scope?: "project" | "tasks" | "resources" | "assignments" | "calendars"; uid?: string; label?: string }>;
        };
        validateProjectModel: (model: ProjectModel) => ValidationIssue[];
      }) => {
        model: ProjectModel;
        issues: ValidationIssue[];
        changes: Array<{
          scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
          uid: string;
          label: string;
          field: string;
          before: string | number | boolean | undefined;
          after: string | number | boolean;
        }>;
        warnings: Array<{ message: string; scope?: "project" | "tasks" | "resources" | "assignments" | "calendars"; uid?: string; label?: string }>;
      };
      previousModel: ProjectModel | null;
      extractLastJsonBlock: (value: string) => string;
      importProjectWorkbookJsonAsProjectModel: (documentLike: unknown) => {
        model: ProjectModel;
        warnings: Array<{ message: string; scope?: "project" | "tasks" | "resources" | "assignments" | "calendars"; uid?: string; label?: string }>;
      };
      importProjectWorkbookJson: (documentLike: unknown, baseModel: ProjectModel) => {
        model: ProjectModel;
        changes: Array<{
          scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
          uid: string;
          label: string;
          field: string;
          before: string | number | boolean | undefined;
          after: string | number | boolean;
        }>;
        warnings: Array<{ message: string; scope?: "project" | "tasks" | "resources" | "assignments" | "calendars"; uid?: string; label?: string }>;
      };
      validateProjectModel: (model: ProjectModel) => ValidationIssue[];
      setCurrentModel: (model: ProjectModel) => void;
      syncXmlTextFromModel: (model: ProjectModel) => string;
      applyModelState: (input: {
        model: ProjectModel;
        issues: ValidationIssue[];
        warnings: Array<{ message: string; scope?: "project" | "tasks" | "resources" | "assignments" | "calendars"; uid?: string; label?: string }>;
        changes: Array<{
          scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
          uid: string;
          label: string;
          field: string;
          before: string | number | boolean | undefined;
          after: string | number | boolean;
        }>;
      }) => void;
      exportCurrentMermaid: (options?: { silent?: boolean }) => Promise<void>;
      completeTransform: (statusMessage: string, toastMessage: string) => void;
    }): Promise<void> {
      const result = input.importWorkbookJsonText({
        sourceText: input.sourceText,
        previousModel: input.previousModel,
        extractLastJsonBlock: input.extractLastJsonBlock,
        importProjectWorkbookJsonAsProjectModel: input.importProjectWorkbookJsonAsProjectModel,
        importProjectWorkbookJson: input.importProjectWorkbookJson,
        validateProjectModel: input.validateProjectModel
      });
      input.setCurrentModel(result.model);
      input.syncXmlTextFromModel(result.model);
      input.applyModelState({
        model: result.model,
        issues: result.issues,
        warnings: result.warnings,
        changes: result.changes
      });
      const summaryText = "JSON を読み込んで project 全体を置き換えました。XML は再生成済みで、必要なら XML Export で保存できます";
      const warningText = result.warnings.length > 0 ? `。JSON 取込で ${result.warnings.length} 件の warning を無視しました` : "";
      input.completeTransform(
        result.issues.length > 0 ? `${summaryText}${warningText}。検証で ${result.issues.length} 件の問題があります` : `${summaryText}${warningText}`,
        "JSON を読み込みました"
      );
      await input.exportCurrentMermaid({ silent: true });
    },

    async importXlsxFile(input: {
      file: File;
      importXlsxBytes: (input: {
        bytes: Uint8Array;
        previousModel: ProjectModel | null;
        createWorkbookCodec: () => {
          importWorkbook: (bytes: Uint8Array) => unknown;
          importWorkbookAsync?: (bytes: Uint8Array) => Promise<unknown>;
        };
        importProjectWorkbookAsProjectModel: (workbook: unknown) => ProjectModel;
        importProjectWorkbookDetailed: (workbook: unknown, baseModel: ProjectModel) => {
          model: ProjectModel;
          changes: Array<{
            scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
            uid: string;
            label: string;
            field: string;
            before: string | number | boolean | undefined;
            after: string | number | boolean;
          }>;
        };
        validateProjectModel: (model: ProjectModel) => ValidationIssue[];
      }) => Promise<{
        model: ProjectModel;
        issues: ValidationIssue[];
        changes: Array<{
          scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
          uid: string;
          label: string;
          field: string;
          before: string | number | boolean | undefined;
          after: string | number | boolean;
        }>;
      }>;
      previousModel: ProjectModel | null;
      createWorkbookCodec: () => {
        importWorkbook: (bytes: Uint8Array) => unknown;
        importWorkbookAsync?: (bytes: Uint8Array) => Promise<unknown>;
      };
      importProjectWorkbookAsProjectModel: (workbook: unknown) => ProjectModel;
      importProjectWorkbookDetailed: (workbook: unknown, baseModel: ProjectModel) => {
        model: ProjectModel;
        changes: Array<{
          scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
          uid: string;
          label: string;
          field: string;
          before: string | number | boolean | undefined;
          after: string | number | boolean;
        }>;
      };
      validateProjectModel: (model: ProjectModel) => ValidationIssue[];
      setCurrentModel: (model: ProjectModel) => void;
      syncXmlTextFromModel: (model: ProjectModel) => string;
      updateSummary: (model: ProjectModel | null) => void;
      renderValidationIssues: (issues: ValidationIssue[]) => void;
      renderImportWarnings: (warnings: Array<{ message: string; scope?: "project" | "tasks" | "resources" | "assignments" | "calendars"; uid?: string; label?: string }>) => void;
      renderXlsxImportSummary: (changes: Array<{
        scope: "project" | "tasks" | "resources" | "assignments" | "calendars";
        uid: string;
        label: string;
        field: string;
        before: string | number | boolean | undefined;
        after: string | number | boolean;
      }>, options?: { sourceLabel?: string }) => void;
      setStatus: (message: string) => void;
      showToast: (message: string) => void;
      setActiveTab: () => void;
      exportCurrentMermaid: (options?: { silent?: boolean }) => Promise<void>;
    }): Promise<void> {
      const result = await input.importXlsxBytes({
        bytes: new Uint8Array(await input.file.arrayBuffer()),
        previousModel: input.previousModel,
        createWorkbookCodec: input.createWorkbookCodec,
        importProjectWorkbookAsProjectModel: input.importProjectWorkbookAsProjectModel,
        importProjectWorkbookDetailed: input.importProjectWorkbookDetailed,
        validateProjectModel: input.validateProjectModel
      });
      input.setCurrentModel(result.model);
      input.syncXmlTextFromModel(result.model);
      input.updateSummary(result.model);
      input.renderValidationIssues(result.issues);
      input.renderImportWarnings([]);
      input.renderXlsxImportSummary(result.changes, { sourceLabel: "XLSX Replace" });
      const summaryText = "XLSX を読み込んで project 全体を置き換えました。XML は再生成済みで、必要なら XML Export で保存できます";
      input.setStatus(result.issues.length > 0 ? `${summaryText}。検証で ${result.issues.length} 件の問題があります` : summaryText);
      input.showToast("XLSX を読み込みました");
      input.setActiveTab();
      await input.exportCurrentMermaid({ silent: true });
    },

    async importCsvFile(input: {
      file: File;
      importCsvText: (input: {
        csvText: string;
        importCsvParentId: (csvText: string) => ProjectModel;
        validateProjectModel: (model: ProjectModel) => ValidationIssue[];
      }) => { model: ProjectModel; issues: ValidationIssue[] };
      importCsvParentId: (csvText: string) => ProjectModel;
      validateProjectModel: (model: ProjectModel) => ValidationIssue[];
      setCurrentModel: (model: ProjectModel) => void;
      setXmlSourceDirty: (dirty: boolean) => void;
      updateSummary: (model: ProjectModel | null) => void;
      renderValidationIssues: (issues: ValidationIssue[]) => void;
      renderImportWarnings: (warnings: Array<{ message: string; scope?: "project" | "tasks" | "resources" | "assignments" | "calendars"; uid?: string; label?: string }>) => void;
      renderXlsxImportSummary: (changes: Array<never>) => void;
      setStatus: (message: string) => void;
      showToast: (message: string) => void;
      setActiveTab: () => void;
      exportCurrentMermaid: (options?: { silent?: boolean }) => Promise<void>;
    }): Promise<void> {
      const result = input.importCsvText({
        csvText: await input.file.text(),
        importCsvParentId: input.importCsvParentId,
        validateProjectModel: input.validateProjectModel
      });
      input.setCurrentModel(result.model);
      input.setXmlSourceDirty(false);
      input.updateSummary(result.model);
      input.renderValidationIssues(result.issues);
      input.renderImportWarnings([]);
      input.renderXlsxImportSummary([]);
      input.setStatus(
        result.issues.length > 0
          ? `CSV ファイルを読み込んで解析しました。検証で ${result.issues.length} 件の問題があります`
          : "CSV + ParentID を内部モデルへ変換しました"
      );
      input.showToast("CSV を読み込みました");
      input.setActiveTab();
      await input.exportCurrentMermaid({ silent: true });
    }
  };

  (globalThis as typeof globalThis & {
    __mikuprojectMainImportActions?: typeof mainImportActions;
  }).__mikuprojectMainImportActions = mainImportActions;
})();
