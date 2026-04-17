/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  const mikuprojectMsprojectAiViews = (globalThis as typeof globalThis & {
    __mikuprojectMsprojectAiViews?: {
      buildProjectDraftRequest: (input: {
        name: string;
        plannedStart?: string;
        goal?: string;
        teamCount?: number;
        mustHavePhases?: string[];
        mustHaveMilestones?: string[];
      }) => unknown;
      importProjectDraftView: (input: {
        draft: unknown;
        normalizeProjectModel: (model: ProjectModel) => ProjectModel;
        ensureDefaultProjectCalendar: (model: ProjectModel) => ProjectModel;
        defaultProjectMinutesPerDay: number;
        defaultProjectMinutesPerWeek: number;
        defaultProjectDaysPerMonth: number;
      }) => ProjectModel;
      exportProjectOverviewView: (model: ProjectModel) => unknown;
      exportPhaseDetailView: (
        model: ProjectModel,
        requestedPhaseUid?: string,
        options?: {
          mode?: "full" | "scoped";
          rootUid?: string;
          maxDepth?: number;
        }
      ) => unknown;
      exportTaskEditView: (model: ProjectModel, requestedTaskUid?: string) => unknown;
    };
  }).__mikuprojectMsprojectAiViews;
  if (!mikuprojectMsprojectAiViews) {
    throw new Error("mikuproject AI views module is not loaded");
  }

  const mikuprojectMsprojectCalendar = (globalThis as typeof globalThis & {
    __mikuprojectMsprojectCalendar?: {
      ensureDefaultProjectCalendar: (model: ProjectModel) => ProjectModel;
    };
  }).__mikuprojectMsprojectCalendar;
  if (!mikuprojectMsprojectCalendar) {
    throw new Error("mikuproject calendar module is not loaded");
  }

  const mikuprojectMsprojectSamples = (globalThis as typeof globalThis & {
    __mikuprojectMsprojectSamples?: {
      SAMPLE_PROJECT_DRAFT_VIEW: unknown;
      buildSampleXml: (input: {
        importProjectDraftView: (draft: unknown) => ProjectModel;
        exportMsProjectXml: (model: ProjectModel) => string;
      }) => string;
    };
  }).__mikuprojectMsprojectSamples;
  if (!mikuprojectMsprojectSamples) {
    throw new Error("mikuproject samples module is not loaded");
  }

  const mikuprojectMsprojectCsv = (globalThis as typeof globalThis & {
    __mikuprojectMsprojectCsv?: {
      exportCsvParentId: (model: ProjectModel) => string;
      importCsvParentId: (input: {
        csvText: string;
        parseNumber: (value: string, fallback: number) => number;
        normalizeProjectModel: (model: ProjectModel) => ProjectModel;
        ensureDefaultProjectCalendar: (model: ProjectModel) => ProjectModel;
      }) => ProjectModel;
    };
  }).__mikuprojectMsprojectCsv;
  if (!mikuprojectMsprojectCsv) {
    throw new Error("mikuproject CSV module is not loaded");
  }

  const mikuprojectMsprojectValidateHelpers = (globalThis as typeof globalThis & {
    __mikuprojectMsprojectValidateHelpers?: {
      parseDateValue: (value: string | undefined) => number | null;
      isPlaceholderUid: (value: string | undefined) => boolean;
      isUnassignedResourceUid: (value: string | undefined) => boolean;
      describeTask: (task: TaskModel) => string;
      detectTaskOrderIssue: (tasks: TaskModel[]) => { previous: TaskModel; current: TaskModel } | null;
      describeResource: (resource: ResourceModel) => string;
      describeCalendar: (calendar: CalendarModel) => string;
      describeAssignment: (assignment: AssignmentModel) => string;
      describeTaskRef: (model: ProjectModel, taskUid: string | undefined) => string;
      describeResourceRef: (model: ProjectModel, resourceUid: string | undefined) => string;
    };
  }).__mikuprojectMsprojectValidateHelpers;
  if (!mikuprojectMsprojectValidateHelpers) {
    throw new Error("mikuproject validate helpers module is not loaded");
  }

  const mikuprojectMsprojectValidate = (globalThis as typeof globalThis & {
    __mikuprojectMsprojectValidate?: {
      validateProjectModel: (input: {
        model: ProjectModel;
        parseDateValue: (value: string | undefined) => number | null;
        isPlaceholderUid: (value: string | undefined) => boolean;
        isUnassignedResourceUid: (value: string | undefined) => boolean;
        describeTask: (task: TaskModel) => string;
        describeResource: (resource: ResourceModel) => string;
        describeCalendar: (calendar: CalendarModel) => string;
        describeAssignment: (assignment: AssignmentModel) => string;
        describeTaskRef: (model: ProjectModel, taskUid: string | undefined) => string;
        describeResourceRef: (model: ProjectModel, resourceUid: string | undefined) => string;
        detectTaskOrderIssue: (tasks: TaskModel[]) => { previous: TaskModel; current: TaskModel } | null;
      }) => ValidationIssue[];
    };
  }).__mikuprojectMsprojectValidate;
  if (!mikuprojectMsprojectValidate) {
    throw new Error("mikuproject validate module is not loaded");
  }

  const mikuprojectMsprojectXmlDom = (globalThis as typeof globalThis & {
    __mikuprojectMsprojectXmlDom?: {
      textContent: (parent: Element, tagName: string) => string;
      parseBoolean: (value: string) => boolean;
      parseNumber: (value: string, defaultValue?: number) => number;
      parseOutlineCodeMasks: (parent: Element) => OutlineCodeMaskModel[];
      parseOutlineCodeValues: (parent: Element) => OutlineCodeValueModel[];
      parseWeekDays: (parent: Element) => WeekDayModel[];
      appendTextElement: (doc: XMLDocument, parent: Element, name: string, value: string | number | boolean | undefined) => void;
      appendWeekDays: (doc: XMLDocument, parent: Element, weekDays: WeekDayModel[]) => void;
      parseWorkingTimes: (parent: Element) => WorkingTimeModel[];
      appendWorkingTimes: (doc: XMLDocument, parent: Element, workingTimes: WorkingTimeModel[]) => void;
      parseXmlDocument: (xmlText: string) => XMLDocument;
      formatXml: (xml: string) => string;
    };
  }).__mikuprojectMsprojectXmlDom;
  if (!mikuprojectMsprojectXmlDom) {
    throw new Error("mikuproject XML DOM module is not loaded");
  }

  const mikuprojectMsprojectCodec = (globalThis as typeof globalThis & {
    __mikuprojectMsprojectCodec?: {
      importMsProjectXml: (input: {
        xmlText: string;
        parseXmlDocument: (xmlText: string) => XMLDocument;
        textContent: (parent: Element, tagName: string) => string;
        parseBoolean: (value: string) => boolean;
        parseNumber: (value: string, defaultValue?: number) => number;
        parseOutlineCodeMasks: (parent: Element) => OutlineCodeMaskModel[];
        parseOutlineCodeValues: (parent: Element) => OutlineCodeValueModel[];
        parseWeekDays: (parent: Element) => WeekDayModel[];
        parseWorkingTimes: (parent: Element) => WorkingTimeModel[];
        normalizeProjectModel: (model: ProjectModel) => ProjectModel;
        ensureDefaultProjectCalendar: (model: ProjectModel) => ProjectModel;
      }) => ProjectModel;
      exportMsProjectXml: (input: {
        model: ProjectModel;
        normalizeProjectModel: (model: ProjectModel) => ProjectModel;
        ensureDefaultProjectCalendar: (model: ProjectModel) => ProjectModel;
        appendTextElement: (doc: XMLDocument, parent: Element, name: string, value: string | number | boolean | undefined) => void;
        appendWeekDays: (doc: XMLDocument, parent: Element, weekDays: WeekDayModel[]) => void;
        appendWorkingTimes: (doc: XMLDocument, parent: Element, workingTimes: WorkingTimeModel[]) => void;
        formatXml: (xml: string) => string;
      }) => string;
    };
  }).__mikuprojectMsprojectCodec;
  if (!mikuprojectMsprojectCodec) {
    throw new Error("mikuproject codec module is not loaded");
  }

  const mikuprojectMsprojectMermaid = (globalThis as typeof globalThis & {
    __mikuprojectMsprojectMermaid?: {
      exportMermaidGantt: (input: {
        model: ProjectModel;
        describePredecessorType: (type: number | undefined) => string;
      }) => string;
    };
  }).__mikuprojectMsprojectMermaid;
  if (!mikuprojectMsprojectMermaid) {
    throw new Error("mikuproject mermaid module is not loaded");
  }

  const DEFAULT_PROJECT_MINUTES_PER_DAY = 480;
  const DEFAULT_PROJECT_MINUTES_PER_WEEK = 2400;
  const DEFAULT_PROJECT_DAYS_PER_MONTH = 20;

  function textContent(parent: Element, tagName: string): string {
    return mikuprojectMsprojectXmlDom.textContent(parent, tagName);
  }

  function parseBoolean(value: string): boolean {
    return mikuprojectMsprojectXmlDom.parseBoolean(value);
  }

  function parseNumber(value: string, defaultValue = 0): number {
    return mikuprojectMsprojectXmlDom.parseNumber(value, defaultValue);
  }

  function parseDateValue(value: string | undefined): number | null {
    return mikuprojectMsprojectValidateHelpers.parseDateValue(value);
  }

  function ensureDefaultProjectCalendar(model: ProjectModel): ProjectModel {
    return mikuprojectMsprojectCalendar.ensureDefaultProjectCalendar(model);
  }

  function parseWeekDays(parent: Element): WeekDayModel[] {
    return mikuprojectMsprojectXmlDom.parseWeekDays(parent);
  }

  function appendWeekDays(doc: XMLDocument, parent: Element, weekDays: WeekDayModel[]): void {
    mikuprojectMsprojectXmlDom.appendWeekDays(doc, parent, weekDays);
  }

  function parseWorkingTimes(parent: Element): WorkingTimeModel[] {
    return mikuprojectMsprojectXmlDom.parseWorkingTimes(parent);
  }

  function appendWorkingTimes(doc: XMLDocument, parent: Element, workingTimes: WorkingTimeModel[]): void {
    mikuprojectMsprojectXmlDom.appendWorkingTimes(doc, parent, workingTimes);
  }

  function parseOutlineCodeMasks(parent: Element): OutlineCodeMaskModel[] {
    return mikuprojectMsprojectXmlDom.parseOutlineCodeMasks(parent);
  }

  function parseOutlineCodeValues(parent: Element): OutlineCodeValueModel[] {
    return mikuprojectMsprojectXmlDom.parseOutlineCodeValues(parent);
  }

  function isPlaceholderUid(value: string | undefined): boolean {
    return mikuprojectMsprojectValidateHelpers.isPlaceholderUid(value);
  }

  function isUnassignedResourceUid(value: string | undefined): boolean {
    return mikuprojectMsprojectValidateHelpers.isUnassignedResourceUid(value);
  }

  function describeTask(task: TaskModel): string {
    return mikuprojectMsprojectValidateHelpers.describeTask(task);
  }

  function detectTaskOrderIssue(tasks: TaskModel[]): { previous: TaskModel; current: TaskModel } | null {
    return mikuprojectMsprojectValidateHelpers.detectTaskOrderIssue(tasks);
  }

  function describeResource(resource: ResourceModel): string {
    return mikuprojectMsprojectValidateHelpers.describeResource(resource);
  }

  function describeCalendar(calendar: CalendarModel): string {
    return mikuprojectMsprojectValidateHelpers.describeCalendar(calendar);
  }

  function describeAssignment(assignment: AssignmentModel): string {
    return mikuprojectMsprojectValidateHelpers.describeAssignment(assignment);
  }

  function describeTaskRef(model: ProjectModel, taskUid: string | undefined): string {
    return mikuprojectMsprojectValidateHelpers.describeTaskRef(model, taskUid);
  }

  function describeResourceRef(model: ProjectModel, resourceUid: string | undefined): string {
    return mikuprojectMsprojectValidateHelpers.describeResourceRef(model, resourceUid);
  }

  function parseXmlDocument(xmlText: string): XMLDocument {
    return mikuprojectMsprojectXmlDom.parseXmlDocument(xmlText);
  }

  function describePredecessorType(type: number | undefined): string {
    if (type === undefined) {
      return "default";
    }
    const typeMap: Record<number, string> = {
      0: "FF",
      1: "FS",
      2: "FF",
      3: "SF",
      4: "SS"
    };
    return typeMap[type] || `type=${type}`;
  }

  function buildProjectDraftRequest(input: {
    name: string;
    plannedStart?: string;
    goal?: string;
    teamCount?: number;
    mustHavePhases?: string[];
    mustHaveMilestones?: string[];
  }) {
    return mikuprojectMsprojectAiViews.buildProjectDraftRequest(input);
  }

  function importProjectDraftView(draft: unknown): ProjectModel {
    return mikuprojectMsprojectAiViews.importProjectDraftView({
      draft,
      normalizeProjectModel,
      ensureDefaultProjectCalendar,
      defaultProjectMinutesPerDay: DEFAULT_PROJECT_MINUTES_PER_DAY,
      defaultProjectMinutesPerWeek: DEFAULT_PROJECT_MINUTES_PER_WEEK,
      defaultProjectDaysPerMonth: DEFAULT_PROJECT_DAYS_PER_MONTH
    });
  }

  function exportProjectOverviewView(model: ProjectModel) {
    return mikuprojectMsprojectAiViews.exportProjectOverviewView(model);
  }

  function exportPhaseDetailView(
    model: ProjectModel,
    requestedPhaseUid?: string,
    options?: {
      mode?: "full" | "scoped";
      rootUid?: string;
      maxDepth?: number;
    }
  ) {
    return mikuprojectMsprojectAiViews.exportPhaseDetailView(model, requestedPhaseUid, options);
  }

  function exportTaskEditView(model: ProjectModel, requestedTaskUid?: string) {
    return mikuprojectMsprojectAiViews.exportTaskEditView(model, requestedTaskUid);
  }

  function exportMermaidGantt(model: ProjectModel): string {
    return mikuprojectMsprojectMermaid.exportMermaidGantt({
      model,
      describePredecessorType
    });
  }

  function exportCsvParentId(model: ProjectModel): string {
    return mikuprojectMsprojectCsv.exportCsvParentId(model);
  }

  function importCsvParentId(csvText: string): ProjectModel {
    return mikuprojectMsprojectCsv.importCsvParentId({
      csvText,
      parseNumber,
      normalizeProjectModel,
      ensureDefaultProjectCalendar
    });
  }

  function importMsProjectXml(xmlText: string): ProjectModel {
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

  function appendTextElement(doc: XMLDocument, parent: Element, name: string, value: string | number | boolean | undefined): void {
    mikuprojectMsprojectXmlDom.appendTextElement(doc, parent, name, value);
  }

  function formatXml(xml: string): string {
    return mikuprojectMsprojectXmlDom.formatXml(xml);
  }

  function exportMsProjectXml(model: ProjectModel): string {
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

  function normalizeProjectModel(model: ProjectModel): ProjectModel {
    return JSON.parse(JSON.stringify(model)) as ProjectModel;
  }

  function validateProjectModel(model: ProjectModel): ValidationIssue[] {
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

  (globalThis as typeof globalThis & {
    __mikuprojectXml?: {
      SAMPLE_XML: string;
      SAMPLE_PROJECT_DRAFT_VIEW: unknown;
      parseXmlDocument: (xmlText: string) => XMLDocument;
      importMsProjectXml: (xmlText: string) => ProjectModel;
      importCsvParentId: (csvText: string) => ProjectModel;
      exportMsProjectXml: (model: ProjectModel) => string;
      exportMermaidGantt: (model: ProjectModel) => string;
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
      exportPhaseDetailView: (model: ProjectModel, phaseUid?: string) => unknown;
      exportCsvParentId: (model: ProjectModel) => string;
      normalizeProjectModel: (model: ProjectModel) => ProjectModel;
      validateProjectModel: (model: ProjectModel) => ValidationIssue[];
    };
  }).__mikuprojectXml = {
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
