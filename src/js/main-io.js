/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const mainIo = {
        buildOutputArchiveEntries(deps) {
            const xmlText = deps.syncXmlTextFromModel(deps.model);
            const now = new Date();
            const stamp = deps.formatTimestampCompact(now);
            const dateOnlyStamp = [
                now.getFullYear(),
                String(now.getMonth() + 1).padStart(2, "0"),
                String(now.getDate()).padStart(2, "0")
            ].join("");
            const codec = deps.createWorkbookCodec();
            const workbook = deps.exportProjectWorkbook(deps.model);
            const workbookJsonText = JSON.stringify(deps.exportProjectWorkbookJson(deps.model), null, 2);
            const csvText = deps.exportCsvParentId(deps.model);
            const wbsOptions = deps.buildWbsOptions(deps.model);
            const wbsWorkbook = deps.exportWbsWorkbook(deps.model, wbsOptions);
            const wbsMarkdown = deps.exportWbsMarkdown(deps.model, wbsOptions);
            const dailySvg = deps.exportNativeSvg(deps.model, wbsOptions);
            const weeklySvg = deps.exportWeeklyNativeSvg(deps.model, wbsOptions);
            const monthlyArchive = deps.exportMonthlyWbsCalendarSvgArchive(deps.model);
            const mermaidText = deps.exportMermaidGantt(deps.model);
            const projectOverview = deps.exportProjectOverviewView(deps.model);
            const phaseDetailViewsFull = (projectOverview.phases || [])
                .map((phase) => phase === null || phase === void 0 ? void 0 : phase.uid)
                .filter((uid) => Boolean(uid))
                .map((phaseUid) => deps.exportPhaseDetailView(deps.model, phaseUid, { mode: "full" }));
            const taskEditViewsFull = deps.model.tasks
                .filter((task) => !(task.uid === "0" || task.summary))
                .map((task) => deps.exportTaskEditView(deps.model, task.uid));
            const aiBundle = {
                view_type: "ai_projection_bundle",
                project_overview_view: projectOverview,
                phase_detail_views_full: phaseDetailViewsFull,
                task_edit_views_full: taskEditViewsFull
            };
            const phaseDetailFull = deps.exportPhaseDetailView(deps.model, undefined, { mode: "full" });
            const allReadmeText = [
                "mikuproject ALL ZIP",
                "",
                "GitHub: https://github.com/igapyon/mikuproject",
                "Agent Skills: https://github.com/igapyon/mikuproject-skills",
                "mikuproject is a local single-file web app that converts MS Project XML into XLSX, Markdown, SVG, Mermaid, and AI-facing JSON exports.",
                "",
                "This archive contains the main outputs generated from the current model.",
                "",
                "Files:",
                "- mikuproject-export-*.xml: regenerated MS Project XML",
                "- mikuproject-export-*.xlsx: workbook XLSX export",
                "- mikuproject-workbook-*.json: workbook JSON export",
                "- mikuproject-export-*.csv: CSV + ParentID export",
                "- mikuproject-wbs-*.xlsx: WBS workbook export",
                "- mikuproject-wbs-*.md: WBS Markdown export",
                "- mikuproject-wbs-daily-*.svg: daily WBS SVG export",
                "- mikuproject-wbs-weekly-*.svg: weekly WBS SVG export",
                "- mikuproject-wbs-mermaid-*.mmd: Mermaid gantt export",
                "- monthly-calendar/YYYY-MM.svg: month-by-month calendar SVG export",
                "- *.editjson: AI-facing projection exports"
            ].join("\n");
            const entries = [
                { name: "README.txt", data: deps.encodeUtf8(`${allReadmeText}\n`) },
                { name: `mikuproject-export-${stamp}.xml`, data: deps.encodeUtf8(`${xmlText}\n`) },
                { name: `mikuproject-export-${stamp}.xlsx`, data: codec.exportWorkbook(workbook) },
                { name: `mikuproject-workbook-${stamp}.json`, data: deps.encodeUtf8(`${workbookJsonText}\n`) },
                { name: `mikuproject-export-${stamp}.csv`, data: deps.encodeUtf8(`${csvText}\n`) },
                { name: `mikuproject-wbs-${stamp}.xlsx`, data: codec.exportWorkbook(wbsWorkbook) },
                { name: `mikuproject-wbs-${dateOnlyStamp}.md`, data: deps.encodeUtf8(`${wbsMarkdown}\n`) },
                { name: `mikuproject-wbs-daily-${stamp}.svg`, data: deps.encodeUtf8(dailySvg) },
                { name: `mikuproject-wbs-weekly-${stamp}.svg`, data: deps.encodeUtf8(weeklySvg) },
                { name: `mikuproject-wbs-mermaid-${stamp}.mmd`, data: deps.encodeUtf8(`${mermaidText}\n`) }
            ];
            for (const entry of monthlyArchive.entries) {
                entries.push({
                    name: `monthly-calendar/${entry.fileName}`,
                    data: deps.encodeUtf8(entry.svg)
                });
            }
            entries.push({ name: "mikuproject-project-overview-view.editjson", data: deps.encodeUtf8(`${JSON.stringify(projectOverview, null, 2)}\n`) }, { name: "mikuproject-full-bundle.editjson", data: deps.encodeUtf8(`${JSON.stringify(aiBundle, null, 2)}\n`) }, { name: "mikuproject-phase-detail-view-full.editjson", data: deps.encodeUtf8(`${JSON.stringify(phaseDetailFull, null, 2)}\n`) });
            return entries;
        },
        detectImportKind(input) {
            const normalizedName = input.fileName.trim().toLowerCase();
            if (normalizedName.endsWith(".xml")) {
                return "xml";
            }
            if (normalizedName.endsWith(".xlsx")) {
                return "xlsx";
            }
            if (normalizedName.endsWith(".csv")) {
                return "csv";
            }
            if (normalizedName.endsWith(".editjson")) {
                return "editjson";
            }
            if (!normalizedName.endsWith(".json")) {
                throw new Error("対応していないファイル形式です。.xml / .xlsx / .json / .editjson / .csv を指定してください");
            }
            if (typeof input.readJsonText !== "string") {
                throw new Error("JSON の判別には source text が必要です");
            }
            const documentLike = JSON.parse(input.extractLastJsonBlock(input.readJsonText));
            const kind = input.detectJsonDocumentKind(documentLike);
            if (kind === "workbook_json" || kind === "project_draft_view" || kind === "patch_json") {
                return kind;
            }
            throw new Error("JSON の format / view_type を判別できません。workbook JSON、project_draft_view、または Patch JSON を指定してください");
        },
        assertRoundTripStable(input) {
            const exportedXml = input.exportMsProjectXml(input.currentModel);
            const reparsedModel = input.importMsProjectXml(exportedXml);
            const validationIssues = input.validateProjectModel(reparsedModel);
            if (validationIssues.some((issue) => issue.level === "error")) {
                throw new Error(validationIssues.map((issue) => issue.message).join("\n"));
            }
            const normalizedLeft = JSON.stringify(input.normalizeProjectModel(input.currentModel));
            const normalizedRight = JSON.stringify(input.normalizeProjectModel(reparsedModel));
            if (normalizedLeft !== normalizedRight) {
                throw new Error("再読込後の内部モデルが一致しません");
            }
            return validationIssues;
        }
    };
    globalThis.__mikuprojectMainIo = mainIo;
})();
