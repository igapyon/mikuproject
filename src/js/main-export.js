/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const mainExport = {
        buildTimestamp(date) {
            return [
                date.getFullYear(),
                String(date.getMonth() + 1).padStart(2, "0"),
                String(date.getDate()).padStart(2, "0"),
                String(date.getHours()).padStart(2, "0"),
                String(date.getMinutes()).padStart(2, "0")
            ].join("");
        },
        buildDateStamp(date) {
            return [
                date.getFullYear(),
                String(date.getMonth() + 1).padStart(2, "0"),
                String(date.getDate()).padStart(2, "0")
            ].join("");
        },
        buildCsvExport(input) {
            const stamp = mainExport.buildTimestamp(new Date());
            return {
                fileName: `mikuproject-export-${stamp}.csv`,
                text: `${input.exportCsvParentId(input.model)}\n`
            };
        },
        buildProjectOverviewExport(input) {
            return {
                fileName: "mikuproject-project-overview-view.editjson",
                text: `${JSON.stringify(input.exportProjectOverviewView(input.model), null, 2)}\n`
            };
        },
        buildTaskEditExport(input) {
            var _a;
            const view = input.exportTaskEditView(input.model, input.requestedTaskUid);
            const resolvedTaskUid = (_a = view.target_task) === null || _a === void 0 ? void 0 : _a.uid;
            const taskSuffix = resolvedTaskUid ? `-${resolvedTaskUid}` : "";
            return {
                fileName: `mikuproject-task-edit-view${taskSuffix}.editjson`,
                text: `${JSON.stringify(view, null, 2)}\n`,
                resolvedTaskUid
            };
        },
        buildAiProjectionBundleExport(input) {
            const projectOverview = input.exportProjectOverviewView(input.model);
            const phaseDetailViewsFull = (projectOverview.phases || [])
                .map((phase) => phase === null || phase === void 0 ? void 0 : phase.uid)
                .filter((uid) => Boolean(uid))
                .map((phaseUid) => input.exportPhaseDetailView(input.model, phaseUid, { mode: "full" }));
            const taskEditViewsFull = input.model.tasks
                .filter((task) => !(task.uid === "0" || task.summary))
                .map((task) => input.exportTaskEditView(input.model, task.uid));
            const bundle = {
                view_type: "ai_projection_bundle",
                project_overview_view: projectOverview,
                phase_detail_views_full: phaseDetailViewsFull,
                task_edit_views_full: taskEditViewsFull
            };
            return {
                fileName: "mikuproject-full-bundle.editjson",
                text: `${JSON.stringify(bundle, null, 2)}\n`,
                phaseCount: phaseDetailViewsFull.length,
                taskCount: taskEditViewsFull.length
            };
        },
        buildPhaseDetailExport(input) {
            var _a, _b, _c, _d;
            const view = input.exportPhaseDetailView(input.model, input.requestedPhaseUid, {
                mode: input.mode,
                rootUid: input.requestedRootUid,
                maxDepth: input.requestedMaxDepth
            });
            const resolvedPhaseUid = (_a = view.phase) === null || _a === void 0 ? void 0 : _a.uid;
            const resolvedMode = ((_b = view.scope) === null || _b === void 0 ? void 0 : _b.mode) === "scoped" ? "scoped" : "full";
            const resolvedRootUid = ((_c = view.scope) === null || _c === void 0 ? void 0 : _c.root_uid) || undefined;
            const resolvedMaxDepth = typeof ((_d = view.scope) === null || _d === void 0 ? void 0 : _d.max_depth) === "number" ? view.scope.max_depth : undefined;
            const phaseSuffix = resolvedPhaseUid ? `-${resolvedPhaseUid}` : "";
            const modeSuffix = resolvedMode === "scoped" ? "-scoped" : "-full";
            const rootSuffix = resolvedRootUid ? `-root-${resolvedRootUid}` : "";
            const depthSuffix = typeof resolvedMaxDepth === "number" ? `-depth-${resolvedMaxDepth}` : "";
            return {
                fileName: `mikuproject-phase-detail-view${phaseSuffix}${modeSuffix}${rootSuffix}${depthSuffix}.editjson`,
                text: `${JSON.stringify(view, null, 2)}\n`,
                resolvedPhaseUid,
                resolvedRootUid,
                resolvedMaxDepth,
                resolvedMode
            };
        },
        buildXlsxExport(input) {
            const stamp = mainExport.buildTimestamp(new Date());
            const codec = input.createWorkbookCodec();
            return {
                fileName: `mikuproject-export-${stamp}.xlsx`,
                bytes: codec.exportWorkbook(input.exportProjectWorkbook(input.model))
            };
        },
        buildWorkbookJsonExport(input) {
            const stamp = mainExport.buildTimestamp(new Date());
            return {
                fileName: `mikuproject-workbook-${stamp}.json`,
                text: JSON.stringify(input.exportProjectWorkbookJson(input.model), null, 2)
            };
        },
        buildWbsXlsxExport(input) {
            const stamp = mainExport.buildTimestamp(new Date());
            const codec = input.createWorkbookCodec();
            return {
                fileName: `mikuproject-wbs-${stamp}.xlsx`,
                bytes: codec.exportWorkbook(input.exportWbsWorkbook(input.model, input.options))
            };
        },
        buildXmlExport(input) {
            const stamp = mainExport.buildTimestamp(new Date());
            return {
                fileName: `mikuproject-export-${stamp}.xml`,
                text: `${input.xmlText}\n`
            };
        },
        buildDailySvgExport(input) {
            const stamp = mainExport.buildTimestamp(new Date());
            return {
                fileName: `mikuproject-wbs-daily-${stamp}.svg`,
                text: input.svg
            };
        },
        buildWeeklySvgExport(input) {
            const stamp = mainExport.buildTimestamp(new Date());
            return {
                fileName: `mikuproject-wbs-weekly-${stamp}.svg`,
                text: input.svg
            };
        },
        buildMonthlySvgZipExport(input) {
            const stamp = mainExport.buildTimestamp(new Date());
            return {
                fileName: `mikuproject-monthly-wbs-calendar-${stamp}.zip`,
                bytes: input.zipBytes
            };
        },
        buildMermaidExport(input) {
            const stamp = mainExport.buildTimestamp(new Date());
            return {
                fileName: `mikuproject-wbs-mermaid-${stamp}.mmd`,
                text: `${input.mermaidText}\n`
            };
        },
        buildWbsMarkdownExport(input) {
            const stamp = mainExport.buildDateStamp(new Date());
            return {
                fileName: `mikuproject-wbs-${stamp}.md`,
                text: input.markdownText
            };
        }
    };
    globalThis.__mikuprojectMainExport = mainExport;
})();
