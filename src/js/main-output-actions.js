/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const mainOutputActions = {
        exportCsv(input) {
            input.syncXmlTextFromModel(input.model);
            const exported = input.buildCsvExport({
                model: input.model,
                exportCsvParentId: input.exportCsvParentId
            });
            input.downloadBlob(new Blob([exported.text], { type: "text/csv;charset=utf-8" }), exported.fileName);
            input.completeOutput("内部モデルから CSV + ParentID を生成して保存しました", "CSV を保存しました");
        },
        exportProjectOverview(input) {
            input.syncXmlTextFromModel(input.model);
            const exported = input.buildProjectOverviewExport({
                model: input.model,
                exportProjectOverviewView: input.exportProjectOverviewView
            });
            input.setOutputText(exported.text.trimEnd());
            input.downloadBlob(new Blob([exported.text], { type: "application/json;charset=utf-8" }), exported.fileName);
            input.completeOutput("project_overview_view を生成して保存しました", "project_overview_view を保存しました");
        },
        exportTaskEdit(input) {
            input.syncXmlTextFromModel(input.model);
            const exported = input.buildTaskEditExport({
                model: input.model,
                requestedTaskUid: input.requestedTaskUid,
                exportTaskEditView: input.exportTaskEditView
            });
            if (exported.resolvedTaskUid) {
                input.setResolvedTaskUid(exported.resolvedTaskUid);
            }
            input.setOutputText(exported.text.trimEnd());
            input.downloadBlob(new Blob([exported.text], { type: "application/json;charset=utf-8" }), exported.fileName);
            input.completeOutput("task_edit_view を生成して保存しました", "task_edit_view を保存しました");
        },
        exportAiProjectionBundle(input) {
            input.syncXmlTextFromModel(input.model);
            const exported = input.buildAiProjectionBundleExport({
                model: input.model,
                exportProjectOverviewView: input.exportProjectOverviewView,
                exportPhaseDetailView: input.exportPhaseDetailView,
                exportTaskEditView: input.exportTaskEditView
            });
            input.setOutputText(exported.text.trimEnd());
            input.downloadBlob(new Blob([exported.text], { type: "application/json;charset=utf-8" }), exported.fileName);
            input.completeOutput(`AI 連携用まとめ JSON を生成して保存しました (phase_detail_view full ${exported.phaseCount} 件 / task_edit_view ${exported.taskCount} 件)`, "AI 連携用まとめ JSON を保存しました");
        },
        exportPhaseDetail(input) {
            input.syncXmlTextFromModel(input.model);
            const exported = input.buildPhaseDetailExport({
                model: input.model,
                mode: input.mode,
                requestedPhaseUid: input.requestedPhaseUid,
                requestedRootUid: input.requestedRootUid,
                requestedMaxDepth: input.requestedMaxDepth,
                exportPhaseDetailView: input.exportPhaseDetailView
            });
            if (exported.resolvedPhaseUid) {
                input.setResolvedPhaseUid(exported.resolvedPhaseUid);
            }
            input.setResolvedRootUid(exported.resolvedRootUid || "");
            input.setResolvedMaxDepth(exported.resolvedMaxDepth);
            input.setOutputText(exported.text.trimEnd());
            input.downloadBlob(new Blob([exported.text], { type: "application/json;charset=utf-8" }), exported.fileName);
            input.completeOutput(`phase_detail_view (${exported.resolvedMode}) を生成して保存しました`, `phase_detail_view (${exported.resolvedMode}) を保存しました`);
        },
        exportXlsx(input) {
            input.syncXmlTextFromModel(input.model);
            const exported = input.buildXlsxExport({
                model: input.model,
                createWorkbookCodec: input.createWorkbookCodec,
                exportProjectWorkbook: input.exportProjectWorkbook
            });
            input.downloadBlob(new Blob([exported.bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), exported.fileName);
            input.completeOutput("XLSX ファイルをエクスポートしました", "XLSX を保存しました");
        },
        exportWorkbookJson(input) {
            input.syncXmlTextFromModel(input.model);
            const exported = input.buildWorkbookJsonExport({
                model: input.model,
                exportProjectWorkbookJson: input.exportProjectWorkbookJson
            });
            input.setOutputText(exported.text);
            input.downloadBlob(new Blob([`${exported.text}\n`], { type: "application/json;charset=utf-8" }), exported.fileName);
            input.completeOutput("XLSX 相当の workbook JSON を生成して保存しました", "JSON を保存しました");
        },
        exportWbsXlsx(input) {
            input.syncXmlTextFromModel(input.model);
            const exported = input.buildWbsXlsxExport({
                model: input.model,
                options: input.options,
                createWorkbookCodec: input.createWorkbookCodec,
                exportWbsWorkbook: input.exportWbsWorkbook
            });
            input.downloadBlob(new Blob([exported.bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), exported.fileName);
            const displayRangeText = input.options.displayDaysBeforeBaseDate !== undefined || input.options.displayDaysAfterBaseDate !== undefined
                ? ` / 表示期間 営業日 基準日前 ${input.options.displayDaysBeforeBaseDate || 0} 日, 基準日後 ${input.options.displayDaysAfterBaseDate || 0} 日`
                : "";
            const progressBandText = " / 進捗帯 営業日";
            input.completeOutput(`WBS XLSX ファイルをエクスポートしました${input.options.holidayDates.length > 0 ? ` (祝日 ${input.options.holidayDates.length} 件)` : ""}${displayRangeText}${progressBandText}`, "WBS XLSX を保存しました");
        }
    };
    globalThis.__mikuprojectMainOutputActions = mainOutputActions;
})();
