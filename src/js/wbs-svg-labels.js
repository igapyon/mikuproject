/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    function createWbsSvgLabelsHelper(config) {
        function formatTaskLabel(task, labelMode) {
            if (labelMode === "list") {
                return `${"　".repeat(Math.max(0, task.outlineLevel - 1))}${task.name || "-"}`;
            }
            return task.name || "-";
        }
        function estimateLabelWidth(label, isPhase) {
            const text = String(label || "").trim() || "-";
            let width = 0;
            for (const char of text) {
                if (/\s/.test(char)) {
                    width += 4;
                    continue;
                }
                if (/[\u3040-\u30ff\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff\uff01-\uff60\uffe0-\uffee]/.test(char)) {
                    width += isPhase ? 13 : 12;
                    continue;
                }
                if (/[A-Z]/.test(char)) {
                    width += 8;
                    continue;
                }
                if (/[a-z0-9]/.test(char)) {
                    width += 7;
                    continue;
                }
                if (/[\u0021-\u007e]/.test(char)) {
                    width += 6;
                    continue;
                }
                width += isPhase ? 12 : 11;
            }
            return Math.max(48, Math.ceil(width));
        }
        function buildPlacementDecision(anchor, x, textWidth, reason, geometry, placementConfig, extras = {}) {
            return {
                x,
                anchor,
                width: textWidth,
                reason,
                metrics: {
                    textWidth,
                    ...extras,
                    gap: placementConfig.gap,
                    shapeStartX: geometry.shapeStartX,
                    shapeEndX: geometry.shapeEndX,
                    shapeCenterX: geometry.shapeCenterX,
                    chartMidX: geometry.chartMidX,
                    leftRoom: geometry.leftRoom,
                    rightRoom: geometry.rightRoom,
                    chartOriginX: geometry.chartOriginX,
                    chartEndX: geometry.chartEndX
                }
            };
        }
        function buildPlacementGeometry(row, chartOriginX, chartWidth, placementConfig) {
            const shapeStartX = row.kind === "milestone"
                ? chartOriginX + (row.startIndex * placementConfig.cellWidth) + (placementConfig.cellWidth / 2) - placementConfig.milestoneHalfWidth
                : chartOriginX + (row.startIndex * placementConfig.cellWidth) + placementConfig.rangeInset;
            const shapeEndX = row.kind === "milestone"
                ? chartOriginX + (row.startIndex * placementConfig.cellWidth) + (placementConfig.cellWidth / 2) + placementConfig.milestoneHalfWidth
                : chartOriginX + (row.endIndex * placementConfig.cellWidth) + placementConfig.cellWidth - placementConfig.rangeInset;
            const chartMidX = chartOriginX + (chartWidth / 2);
            const chartEndX = chartOriginX + chartWidth;
            return {
                chartOriginX,
                shapeStartX,
                shapeEndX,
                shapeCenterX: (shapeStartX + shapeEndX) / 2,
                chartMidX,
                chartEndX,
                leftRoom: Math.max(0, (shapeStartX - placementConfig.gap) - chartOriginX),
                rightRoom: Math.max(0, chartEndX - (shapeEndX + placementConfig.gap))
            };
        }
        function resolveSideLabelPlacementDecision(row, chartOriginX, chartWidth, placementConfig) {
            const textWidth = estimateLabelWidth(row.label, row.kind === "phase");
            if (row.startIndex === null || row.endIndex === null) {
                return {
                    x: placementConfig.missingRangeX,
                    anchor: "start",
                    width: textWidth,
                    reason: placementConfig.missingRangeReason,
                    metrics: {
                        startIndex: row.startIndex,
                        endIndex: row.endIndex
                    }
                };
            }
            const geometry = buildPlacementGeometry(row, chartOriginX, chartWidth, placementConfig);
            const preferredRoom = Math.min(textWidth, placementConfig.cellWidth * placementConfig.preferredColumns);
            if (geometry.shapeCenterX >= geometry.chartMidX) {
                if (geometry.leftRoom < textWidth && geometry.rightRoom >= textWidth) {
                    return buildPlacementDecision("start", geometry.shapeEndX + placementConfig.gap, textWidth, "midpoint-right-but-left-overflows-fallback-right", geometry, placementConfig, { preferredRoom });
                }
                return buildPlacementDecision("end", geometry.shapeStartX - placementConfig.gap, textWidth, "midpoint-right-prefer-left", geometry, placementConfig, { preferredRoom });
            }
            if (geometry.rightRoom >= preferredRoom) {
                return buildPlacementDecision("start", geometry.shapeEndX + placementConfig.gap, textWidth, "enough-right-room", geometry, placementConfig, { preferredRoom });
            }
            if (geometry.leftRoom >= preferredRoom) {
                return buildPlacementDecision("end", geometry.shapeStartX - placementConfig.gap, textWidth, "enough-left-room", geometry, placementConfig, { preferredRoom });
            }
            if (geometry.rightRoom > geometry.leftRoom) {
                return buildPlacementDecision("start", geometry.shapeEndX + placementConfig.gap, textWidth, "fallback-wider-right-room", geometry, placementConfig, { preferredRoom });
            }
            return buildPlacementDecision("end", geometry.shapeStartX - placementConfig.gap, textWidth, "fallback-wider-left-room", geometry, placementConfig, { preferredRoom });
        }
        function resolveLabelPlacement(row, chartOriginX, chartWidth, labelMode) {
            const decision = resolveLabelPlacementDecision(row, chartOriginX, chartWidth, labelMode);
            return { x: decision.x, anchor: decision.anchor, width: decision.width };
        }
        function resolveLabelPlacementDecision(row, chartOriginX, chartWidth, labelMode) {
            if (labelMode === "list" || row.startIndex === null || row.endIndex === null) {
                return {
                    x: config.leftPadding + 10,
                    anchor: "start",
                    width: estimateLabelWidth(row.label, row.kind === "phase"),
                    reason: "list-mode-or-missing-range",
                    metrics: {
                        labelMode,
                        startIndex: row.startIndex,
                        endIndex: row.endIndex
                    }
                };
            }
            return resolveSideLabelPlacementDecision(row, chartOriginX, chartWidth, {
                cellWidth: config.dayWidth,
                gap: 18,
                rangeInset: 6,
                milestoneHalfWidth: 13,
                preferredColumns: 4,
                missingRangeX: config.leftPadding + 10,
                missingRangeReason: "list-mode-or-missing-range"
            });
        }
        function shouldUseDailyOnBarLabel(row, chartOriginX, chartWidth, bandLength, labelMode) {
            if (labelMode !== "near") {
                return false;
            }
            if (row.kind === "milestone") {
                return false;
            }
            if (row.startIndex === null || row.endIndex === null || bandLength <= 0) {
                return false;
            }
            if (row.startIndex === 0 && row.endIndex === bandLength - 1) {
                const fullWidthTextWidth = estimateLabelWidth(row.label, row.kind === "phase");
                const fullWidthBarWidth = Math.max(12, ((row.endIndex - row.startIndex + 1) * config.dayWidth) - 12);
                return fullWidthBarWidth >= fullWidthTextWidth + 16;
            }
            const textWidth = estimateLabelWidth(row.label, row.kind === "phase");
            const gap = 18;
            const shapeStartX = row.kind === "milestone"
                ? chartOriginX + (row.startIndex * config.dayWidth) + (config.dayWidth / 2) - 13
                : chartOriginX + (row.startIndex * config.dayWidth) + 6;
            const shapeEndX = row.kind === "milestone"
                ? chartOriginX + (row.startIndex * config.dayWidth) + (config.dayWidth / 2) + 13
                : chartOriginX + (row.endIndex * config.dayWidth) + config.dayWidth - 6;
            const barWidth = Math.max(12, ((row.endIndex - row.startIndex + 1) * config.dayWidth) - 12);
            const leftRoom = Math.max(0, shapeStartX - gap - chartOriginX);
            const rightRoom = Math.max(0, (chartOriginX + chartWidth) - (shapeEndX + gap));
            if (barWidth < textWidth + 16) {
                return false;
            }
            return leftRoom < textWidth && rightRoom < textWidth;
        }
        function resolveWeeklyLabelPlacement(row, chartOriginX, chartWidth) {
            const decision = resolveWeeklyLabelPlacementDecision(row, chartOriginX, chartWidth);
            return { x: decision.x, anchor: decision.anchor, width: decision.width };
        }
        function resolveWeeklyLabelPlacementDecision(row, chartOriginX, chartWidth) {
            return resolveSideLabelPlacementDecision(row, chartOriginX, chartWidth, {
                cellWidth: config.weekWidth,
                gap: 16,
                rangeInset: 4,
                milestoneHalfWidth: 11,
                preferredColumns: 4,
                missingRangeX: chartOriginX + 10,
                missingRangeReason: "missing-range"
            });
        }
        function buildShapeGeometry(row, chartOriginX, chartOriginY, cellWidth, rangeInset, milestoneHalfWidth) {
            if (row.startIndex === null || row.endIndex === null) {
                return null;
            }
            const startX = row.kind === "milestone"
                ? chartOriginX + (row.startIndex * cellWidth) + (cellWidth / 2) - milestoneHalfWidth
                : chartOriginX + (row.startIndex * cellWidth) + rangeInset;
            const endX = row.kind === "milestone"
                ? chartOriginX + (row.startIndex * cellWidth) + (cellWidth / 2) + milestoneHalfWidth
                : chartOriginX + (row.endIndex * cellWidth) + cellWidth - rangeInset;
            return {
                uid: row.task.uid,
                kind: row.kind,
                startX,
                endX,
                midY: chartOriginY + row.y + (config.rowHeight / 2)
            };
        }
        function buildDependencyConnectorPath(startX, startY, laneX, endX, endY, radius) {
            if (endX <= startX) {
                return `M ${startX} ${startY} L ${startX} ${startY}`;
            }
            if (startY === endY) {
                return `M ${startX} ${startY} L ${endX} ${endY}`;
            }
            const directionY = endY > startY ? 1 : -1;
            const horizontalRoom = Math.max(2, laneX - startX);
            const targetRoom = Math.max(2, endX - laneX);
            const verticalRoom = Math.max(2, Math.abs(endY - startY));
            const usableRadius = Math.min(radius, horizontalRoom * 0.8, targetRoom * 0.8, verticalRoom / 3);
            const startCurveX = startX + usableRadius;
            const startCurveY = startY + (directionY * usableRadius);
            const simpleEndCurveX = endX - usableRadius;
            if (targetRoom >= Math.max(18, usableRadius * 3)) {
                const horizontalStartX = laneX + usableRadius;
                return [
                    `M ${startX} ${startY}`,
                    `C ${startCurveX} ${startY} ${laneX} ${startY} ${laneX} ${startCurveY}`,
                    `L ${laneX} ${endY - (directionY * usableRadius)}`,
                    `C ${laneX} ${endY - (directionY * Math.max(3, usableRadius * 0.45))} ${laneX} ${endY} ${horizontalStartX} ${endY}`,
                    `L ${simpleEndCurveX} ${endY}`,
                    `L ${endX} ${endY}`
                ].join(" ");
            }
            const endCurveEntryX = endX - Math.max(usableRadius, 8);
            const midY = startY + ((endY - startY) / 2);
            const axisX = (startX + endCurveEntryX) / 2;
            const symmetricBulge = Math.min(Math.max(18, usableRadius * 2.4), Math.max(18, (endCurveEntryX - startX) * 0.95));
            const midRise = Math.min(Math.max(10, usableRadius * 1.3), Math.max(10, verticalRoom * 0.32));
            return [
                `M ${startX} ${startY}`,
                `C ${axisX + symmetricBulge} ${startY} ${axisX + symmetricBulge} ${midY - (directionY * midRise)} ${axisX} ${midY}`,
                `C ${axisX - symmetricBulge} ${midY + (directionY * midRise)} ${axisX - symmetricBulge} ${endY} ${endCurveEntryX} ${endY}`,
                `L ${endX} ${endY}`
            ].join(" ");
        }
        function buildDependencyConnectorElements(tasks, rows, chartOriginX, chartOriginY, connectorConfig) {
            const geometryByUid = new Map();
            for (const row of rows) {
                const geometry = buildShapeGeometry(row, chartOriginX, chartOriginY, connectorConfig.cellWidth, connectorConfig.rangeInset, connectorConfig.milestoneHalfWidth);
                if (geometry) {
                    geometryByUid.set(row.task.uid, geometry);
                }
            }
            const parts = [];
            let connectorIndex = 0;
            for (const task of tasks) {
                const toGeometry = geometryByUid.get(task.uid);
                if (!toGeometry) {
                    continue;
                }
                for (const predecessor of task.predecessors || []) {
                    if (predecessor.type !== undefined && predecessor.type !== 1) {
                        continue;
                    }
                    const fromGeometry = geometryByUid.get(predecessor.predecessorUid);
                    if (!fromGeometry || fromGeometry.uid === toGeometry.uid) {
                        continue;
                    }
                    const endX = toGeometry.startX - connectorConfig.targetInset;
                    const candidateLaneX = fromGeometry.endX + connectorConfig.routeOffset + ((connectorIndex % 4) * connectorConfig.routeSpacing);
                    const minLaneX = fromGeometry.endX + 4;
                    const maxLaneX = endX - 4;
                    const laneBaseX = maxLaneX <= minLaneX
                        ? (fromGeometry.endX + endX) / 2
                        : Math.max(minLaneX, Math.min(maxLaneX, candidateLaneX));
                    const path = buildDependencyConnectorPath(fromGeometry.endX, fromGeometry.midY, laneBaseX, endX, toGeometry.midY, connectorConfig.cornerRadius);
                    parts.push(`<path class="dependencyPath" d="${path}" marker-end="url(#dependencyArrow)" data-from-uid="${escapeXml(fromGeometry.uid)}" data-to-uid="${escapeXml(toGeometry.uid)}" data-link-type="FS"/>`);
                    connectorIndex += 1;
                }
            }
            return parts;
        }
        return {
            formatTaskLabel,
            resolveLabelPlacement,
            resolveWeeklyLabelPlacement,
            shouldUseDailyOnBarLabel,
            buildDependencyConnectorElements
        };
    }
    function escapeXml(value) {
        return String(value || "")
            .replaceAll("&", "&amp;")
            .replaceAll("<", "&lt;")
            .replaceAll(">", "&gt;")
            .replaceAll("\"", "&quot;");
    }
    globalThis.__mikuprojectWbsSvgLabels = {
        createWbsSvgLabelsHelper
    };
})();
