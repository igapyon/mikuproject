/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    function createWbsSvgViewportHelper() {
        function computePlacementBounds(labelPlacements, defaultMinX, defaultMaxX) {
            const minX = labelPlacements.reduce((min, placement) => {
                const placementMinX = placement.anchor === "start" ? placement.x : placement.x - placement.width;
                return Math.min(min, placementMinX);
            }, defaultMinX);
            const maxX = labelPlacements.reduce((max, placement) => {
                const placementMaxX = placement.anchor === "start" ? placement.x + placement.width : placement.x;
                return Math.max(max, placementMaxX);
            }, defaultMaxX);
            return { minX, maxX };
        }
        function computeDailyViewport(config) {
            const { labelPlacements, chartOriginXBase, chartWidth, rightLabelWidth, rightPadding, labelMode } = config;
            const bounds = computePlacementBounds(labelPlacements, chartOriginXBase, chartOriginXBase + chartWidth + rightLabelWidth);
            const shiftX = labelMode === "near" ? Math.max(0, -bounds.minX) : 0;
            return {
                shiftX,
                chartOriginX: chartOriginXBase + shiftX,
                svgWidth: bounds.maxX + shiftX + rightPadding
            };
        }
        function computeWeeklyViewport(config) {
            const { labelPlacements, chartOriginXBase, chartWidth, trimPadding, rightPadding } = config;
            const bounds = computePlacementBounds(labelPlacements, chartOriginXBase, chartOriginXBase + chartWidth);
            const contentMinX = Math.min(chartOriginXBase, bounds.minX);
            const contentMaxX = Math.max(chartOriginXBase + chartWidth, bounds.maxX);
            const shiftX = trimPadding - contentMinX;
            return {
                shiftX,
                chartOriginX: chartOriginXBase + shiftX,
                svgWidth: (contentMaxX - contentMinX) + (trimPadding * 2) + rightPadding
            };
        }
        return {
            computeDailyViewport,
            computeWeeklyViewport
        };
    }
    globalThis.__mikuprojectWbsSvgViewport = {
        createWbsSvgViewportHelper
    };
})();
