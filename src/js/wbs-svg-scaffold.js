/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    function createWbsSvgScaffoldHelper() {
        function buildBaseScaffold(config) {
            const { svgWidth, svgHeight, ariaLabel, titleText, titleY, titleX, escapeXml, extraStyleLines = [], extraTextLines = [] } = config;
            return [
                `<svg xmlns="http://www.w3.org/2000/svg" width="${svgWidth}" height="${svgHeight}" viewBox="0 0 ${svgWidth} ${svgHeight}" role="img" aria-label="${escapeXml(ariaLabel)}">`,
                "<style>",
                "text { font-family: 'Hiragino Sans', 'Yu Gothic', sans-serif; fill: #1d2740; }",
                ".title { font-size: 18px; font-weight: 700; }",
                ".axis { font-size: 12px; fill: #5b6370; }",
                ".label { font-size: 13px; }",
                ".phaseLabel { font-size: 13px; font-weight: 700; }",
                ".grid { stroke: #c9d3e1; stroke-width: 1; }",
                ".today { stroke: #ff6b5a; stroke-width: 2; }",
                ".dependencyPath { fill: none; stroke: #9eb6c8; stroke-width: 1.5; stroke-linecap: round; stroke-linejoin: round; opacity: 0.95; }",
                ...extraStyleLines,
                "</style>",
                "<defs>",
                '<marker id="dependencyArrow" markerWidth="7" markerHeight="7" refX="5.5" refY="3.5" orient="auto" markerUnits="strokeWidth">',
                '<path d="M0,0 L7,3.5 L0,7 Z" fill="#9eb6c8" />',
                "</marker>",
                "</defs>",
                `<rect x="0" y="0" width="${svgWidth}" height="${svgHeight}" fill="#ffffff"/>`,
                `<text class="title" x="${titleX}" y="${titleY}" text-anchor="middle">${escapeXml(titleText)}</text>`,
                ...extraTextLines
            ];
        }
        function createDailyScaffold(config) {
            return buildBaseScaffold({
                svgWidth: config.svgWidth,
                svgHeight: config.svgHeight,
                ariaLabel: config.projectName || "Project",
                titleText: config.projectName || "-",
                titleX: config.chartOriginX + (config.chartWidth / 2),
                titleY: config.topPadding + 18,
                escapeXml: config.escapeXml
            });
        }
        function createWeeklyScaffold(config) {
            const projectName = config.projectName || "-";
            return buildBaseScaffold({
                svgWidth: config.svgWidth,
                svgHeight: config.svgHeight,
                ariaLabel: `${projectName} weekly overview`,
                titleText: `${projectName} weekly overview`,
                titleX: config.chartOriginX + (config.chartWidth / 2),
                titleY: config.topPadding + 18,
                escapeXml: config.escapeXml,
                extraStyleLines: [
                    ".meta { font-size: 12px; fill: #5b6370; }",
                    ".monthAxis { font-size: 13px; font-weight: 700; fill: #475467; }",
                    ".weekAxis { font-size: 10px; fill: #667085; }",
                    ".monthBoundary { stroke: #98a2b3; stroke-width: 1.5; }"
                ],
                extraTextLines: [
                    `<text class="meta" x="${config.chartOriginX + (config.chartWidth / 2)}" y="${config.topPadding + 40}" text-anchor="middle">${config.escapeXml(`project range ${String(config.projectStartDate || "").slice(0, 10)} - ${String(config.projectFinishDate || "").slice(0, 10)}`)}</text>`
                ]
            });
        }
        return {
            createDailyScaffold,
            createWeeklyScaffold
        };
    }
    globalThis.__mikuprojectWbsSvgScaffold = {
        createWbsSvgScaffoldHelper
    };
})();
