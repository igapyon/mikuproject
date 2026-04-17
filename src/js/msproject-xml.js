/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const mikuprojectMsprojectAiViews = globalThis.__mikuprojectMsprojectAiViews;
    if (!mikuprojectMsprojectAiViews) {
        throw new Error("mikuproject AI views module is not loaded");
    }
    const mikuprojectMsprojectCalendar = globalThis.__mikuprojectMsprojectCalendar;
    if (!mikuprojectMsprojectCalendar) {
        throw new Error("mikuproject calendar module is not loaded");
    }
    const mikuprojectMsprojectSamples = globalThis.__mikuprojectMsprojectSamples;
    if (!mikuprojectMsprojectSamples) {
        throw new Error("mikuproject samples module is not loaded");
    }
    const mikuprojectMsprojectCsv = globalThis.__mikuprojectMsprojectCsv;
    if (!mikuprojectMsprojectCsv) {
        throw new Error("mikuproject CSV module is not loaded");
    }
    const mikuprojectMsprojectValidateHelpers = globalThis.__mikuprojectMsprojectValidateHelpers;
    if (!mikuprojectMsprojectValidateHelpers) {
        throw new Error("mikuproject validate helpers module is not loaded");
    }
    const mikuprojectMsprojectValidate = globalThis.__mikuprojectMsprojectValidate;
    if (!mikuprojectMsprojectValidate) {
        throw new Error("mikuproject validate module is not loaded");
    }
    const mikuprojectMsprojectXmlDom = globalThis.__mikuprojectMsprojectXmlDom;
    if (!mikuprojectMsprojectXmlDom) {
        throw new Error("mikuproject XML DOM module is not loaded");
    }
    const mikuprojectMsprojectCodec = globalThis.__mikuprojectMsprojectCodec;
    if (!mikuprojectMsprojectCodec) {
        throw new Error("mikuproject codec module is not loaded");
    }
    const mikuprojectMsprojectMermaid = globalThis.__mikuprojectMsprojectMermaid;
    if (!mikuprojectMsprojectMermaid) {
        throw new Error("mikuproject mermaid module is not loaded");
    }
    const DEFAULT_PROJECT_MINUTES_PER_DAY = 480;
    const DEFAULT_PROJECT_MINUTES_PER_WEEK = 2400;
    const DEFAULT_PROJECT_DAYS_PER_MONTH = 20;
    function textContent(parent, tagName) {
        return mikuprojectMsprojectXmlDom.textContent(parent, tagName);
    }
    function parseBoolean(value) {
        return mikuprojectMsprojectXmlDom.parseBoolean(value);
    }
    function parseNumber(value, defaultValue = 0) {
        return mikuprojectMsprojectXmlDom.parseNumber(value, defaultValue);
    }
    function parseDateValue(value) {
        return mikuprojectMsprojectValidateHelpers.parseDateValue(value);
    }
    function ensureDefaultProjectCalendar(model) {
        return mikuprojectMsprojectCalendar.ensureDefaultProjectCalendar(model);
    }
    function parseWeekDays(parent) {
        return mikuprojectMsprojectXmlDom.parseWeekDays(parent);
    }
    function appendWeekDays(doc, parent, weekDays) {
        mikuprojectMsprojectXmlDom.appendWeekDays(doc, parent, weekDays);
    }
    function parseWorkingTimes(parent) {
        return mikuprojectMsprojectXmlDom.parseWorkingTimes(parent);
    }
    function appendWorkingTimes(doc, parent, workingTimes) {
        mikuprojectMsprojectXmlDom.appendWorkingTimes(doc, parent, workingTimes);
    }
    function parseOutlineCodeMasks(parent) {
        return mikuprojectMsprojectXmlDom.parseOutlineCodeMasks(parent);
    }
    function parseOutlineCodeValues(parent) {
        return mikuprojectMsprojectXmlDom.parseOutlineCodeValues(parent);
    }
    function isPlaceholderUid(value) {
        return mikuprojectMsprojectValidateHelpers.isPlaceholderUid(value);
    }
    function isUnassignedResourceUid(value) {
        return mikuprojectMsprojectValidateHelpers.isUnassignedResourceUid(value);
    }
    function describeTask(task) {
        return mikuprojectMsprojectValidateHelpers.describeTask(task);
    }
    function detectTaskOrderIssue(tasks) {
        return mikuprojectMsprojectValidateHelpers.detectTaskOrderIssue(tasks);
    }
    function describeResource(resource) {
        return mikuprojectMsprojectValidateHelpers.describeResource(resource);
    }
    function describeCalendar(calendar) {
        return mikuprojectMsprojectValidateHelpers.describeCalendar(calendar);
    }
    function describeAssignment(assignment) {
        return mikuprojectMsprojectValidateHelpers.describeAssignment(assignment);
    }
    function describeTaskRef(model, taskUid) {
        return mikuprojectMsprojectValidateHelpers.describeTaskRef(model, taskUid);
    }
    function describeResourceRef(model, resourceUid) {
        return mikuprojectMsprojectValidateHelpers.describeResourceRef(model, resourceUid);
    }
    function parseXmlDocument(xmlText) {
        return mikuprojectMsprojectXmlDom.parseXmlDocument(xmlText);
    }
    function describePredecessorType(type) {
        if (type === undefined) {
            return "default";
        }
        const typeMap = {
            0: "FF",
            1: "FS",
            2: "FF",
            3: "SF",
            4: "SS"
        };
        return typeMap[type] || `type=${type}`;
    }
    function buildProjectDraftRequest(input) {
        return mikuprojectMsprojectAiViews.buildProjectDraftRequest(input);
    }
    function importProjectDraftView(draft) {
        return mikuprojectMsprojectAiViews.importProjectDraftView({
            draft,
            normalizeProjectModel,
            ensureDefaultProjectCalendar,
            defaultProjectMinutesPerDay: DEFAULT_PROJECT_MINUTES_PER_DAY,
            defaultProjectMinutesPerWeek: DEFAULT_PROJECT_MINUTES_PER_WEEK,
            defaultProjectDaysPerMonth: DEFAULT_PROJECT_DAYS_PER_MONTH
        });
    }
    function exportProjectOverviewView(model) {
        return mikuprojectMsprojectAiViews.exportProjectOverviewView(model);
    }
    function exportPhaseDetailView(model, requestedPhaseUid, options) {
        return mikuprojectMsprojectAiViews.exportPhaseDetailView(model, requestedPhaseUid, options);
    }
    function exportTaskEditView(model, requestedTaskUid) {
        return mikuprojectMsprojectAiViews.exportTaskEditView(model, requestedTaskUid);
    }
    function exportMermaidGantt(model) {
        return mikuprojectMsprojectMermaid.exportMermaidGantt({
            model,
            describePredecessorType
        });
    }
    function exportCsvParentId(model) {
        return mikuprojectMsprojectCsv.exportCsvParentId(model);
    }
    function importCsvParentId(csvText) {
        return mikuprojectMsprojectCsv.importCsvParentId({
            csvText,
            parseNumber,
            normalizeProjectModel,
            ensureDefaultProjectCalendar
        });
    }
    function importMsProjectXml(xmlText) {
        return mikuprojectMsprojectCodec.importMsProjectXml({
            xmlText,
            parseXmlDocument,
            textContent,
            parseBoolean,
            parseNumber,
            parseOutlineCodeMasks,
            parseOutlineCodeValues,
            parseWeekDays,
            parseWorkingTimes,
            normalizeProjectModel,
            ensureDefaultProjectCalendar
        });
    }
    function appendTextElement(doc, parent, name, value) {
        mikuprojectMsprojectXmlDom.appendTextElement(doc, parent, name, value);
    }
    function formatXml(xml) {
        return mikuprojectMsprojectXmlDom.formatXml(xml);
    }
    function exportMsProjectXml(model) {
        return mikuprojectMsprojectCodec.exportMsProjectXml({
            model,
            normalizeProjectModel,
            ensureDefaultProjectCalendar,
            appendTextElement,
            appendWeekDays,
            appendWorkingTimes,
            formatXml
        });
    }
    function normalizeProjectModel(model) {
        return JSON.parse(JSON.stringify(model));
    }
    function validateProjectModel(model) {
        return mikuprojectMsprojectValidate.validateProjectModel({
            model,
            parseDateValue,
            isPlaceholderUid,
            isUnassignedResourceUid,
            describeTask,
            describeResource,
            describeCalendar,
            describeAssignment,
            describeTaskRef,
            describeResourceRef,
            detectTaskOrderIssue
        });
    }
    const SAMPLE_PROJECT_DRAFT_VIEW = mikuprojectMsprojectSamples.SAMPLE_PROJECT_DRAFT_VIEW;
    const SAMPLE_XML = mikuprojectMsprojectSamples.buildSampleXml({
        importProjectDraftView,
        exportMsProjectXml
    });
    globalThis.__mikuprojectXml = {
        SAMPLE_XML,
        SAMPLE_PROJECT_DRAFT_VIEW,
        parseXmlDocument,
        importMsProjectXml,
        importCsvParentId,
        exportMsProjectXml,
        exportMermaidGantt,
        buildProjectDraftRequest,
        importProjectDraftView,
        exportProjectOverviewView,
        exportTaskEditView,
        exportPhaseDetailView,
        exportCsvParentId,
        normalizeProjectModel,
        validateProjectModel
    };
})();
