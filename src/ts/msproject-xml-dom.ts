/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  function textContent(parent: Element, tagName: string): string {
    const element = parent.getElementsByTagName(tagName)[0];
    return String(element?.textContent || "").trim();
  }

  function parseBoolean(value: string): boolean {
    return value === "1" || value.toLowerCase() === "true";
  }

  function parseNumber(value: string, defaultValue = 0): number {
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : defaultValue;
  }

  function parseOutlineCodeMasks(parent: Element): OutlineCodeMaskModel[] {
    const masksElement = parent.getElementsByTagName("Masks")[0];
    if (!masksElement) {
      return [];
    }
    return Array.from(masksElement.children)
      .filter((child) => child.tagName === "Mask")
      .map((mask) => ({
        level: parseNumber(textContent(mask, "Level"), 0),
        mask: textContent(mask, "Mask") || undefined,
        length: textContent(mask, "Length") ? parseNumber(textContent(mask, "Length"), 0) : undefined,
        sequence: textContent(mask, "Sequence") ? parseNumber(textContent(mask, "Sequence"), 0) : undefined
      }));
  }

  function parseOutlineCodeValues(parent: Element): OutlineCodeValueModel[] {
    const valuesElement = parent.getElementsByTagName("Values")[0];
    if (!valuesElement) {
      return [];
    }
    return Array.from(valuesElement.children)
      .filter((child) => child.tagName === "Value")
      .map((value) => ({
        value: textContent(value, "Value"),
        description: textContent(value, "Description") || undefined
      }));
  }

  function parseWeekDays(parent: Element): WeekDayModel[] {
    return Array.from(parent.getElementsByTagName("WeekDays")[0]?.getElementsByTagName("WeekDay") || []).map((weekDay) => ({
      dayType: parseNumber(textContent(weekDay, "DayType"), 0),
      dayWorking: parseBoolean(textContent(weekDay, "DayWorking")),
      workingTimes: Array.from(weekDay.getElementsByTagName("WorkingTimes")[0]?.getElementsByTagName("WorkingTime") || []).map((workingTime) => ({
        fromTime: textContent(workingTime, "FromTime"),
        toTime: textContent(workingTime, "ToTime")
      }))
    }));
  }

  function appendTextElement(doc: XMLDocument, parent: Element, name: string, value: string | number | boolean | undefined): void {
    if (value === undefined || value === "") {
      return;
    }
    const element = doc.createElement(name);
    if (typeof value === "boolean") {
      element.textContent = value ? "1" : "0";
    } else {
      element.textContent = String(value);
    }
    parent.appendChild(element);
  }

  function appendWeekDays(doc: XMLDocument, parent: Element, weekDays: WeekDayModel[]): void {
    if (weekDays.length === 0) {
      return;
    }
    const weekDaysElement = doc.createElement("WeekDays");
    for (const weekDay of weekDays) {
      const weekDayElement = doc.createElement("WeekDay");
      appendTextElement(doc, weekDayElement, "DayType", weekDay.dayType);
      appendTextElement(doc, weekDayElement, "DayWorking", weekDay.dayWorking);
      if (weekDay.workingTimes.length > 0) {
        const workingTimesElement = doc.createElement("WorkingTimes");
        for (const workingTime of weekDay.workingTimes) {
          const workingTimeElement = doc.createElement("WorkingTime");
          appendTextElement(doc, workingTimeElement, "FromTime", workingTime.fromTime);
          appendTextElement(doc, workingTimeElement, "ToTime", workingTime.toTime);
          workingTimesElement.appendChild(workingTimeElement);
        }
        weekDayElement.appendChild(workingTimesElement);
      }
      weekDaysElement.appendChild(weekDayElement);
    }
    parent.appendChild(weekDaysElement);
  }

  function parseWorkingTimes(parent: Element): WorkingTimeModel[] {
    return Array.from(parent.getElementsByTagName("WorkingTimes")[0]?.getElementsByTagName("WorkingTime") || []).map((workingTime) => ({
      fromTime: textContent(workingTime, "FromTime"),
      toTime: textContent(workingTime, "ToTime")
    }));
  }

  function appendWorkingTimes(doc: XMLDocument, parent: Element, workingTimes: WorkingTimeModel[]): void {
    if (workingTimes.length === 0) {
      return;
    }
    const workingTimesElement = doc.createElement("WorkingTimes");
    for (const workingTime of workingTimes) {
      const workingTimeElement = doc.createElement("WorkingTime");
      appendTextElement(doc, workingTimeElement, "FromTime", workingTime.fromTime);
      appendTextElement(doc, workingTimeElement, "ToTime", workingTime.toTime);
      workingTimesElement.appendChild(workingTimeElement);
    }
    parent.appendChild(workingTimesElement);
  }

  function parseXmlDocument(xmlText: string): XMLDocument {
    const parser = new DOMParser();
    const xml = parser.parseFromString(xmlText, "application/xml");
    const parserError = xml.getElementsByTagName("parsererror")[0];
    if (parserError) {
      throw new Error("XML の解析に失敗しました");
    }
    return xml;
  }

  function formatXml(xml: string): string {
    const normalized = xml.replace(/>\s*</g, "><").trim();
    const tokens = normalized.replace(/></g, ">\n<").split("\n");
    let indentLevel = 0;
    const formatted: string[] = [];

    for (const rawToken of tokens) {
      const token = rawToken.trim();
      if (!token) {
        continue;
      }

      if (/^<\//.test(token)) {
        indentLevel = Math.max(indentLevel - 1, 0);
      }

      formatted.push(`${"  ".repeat(indentLevel)}${token}`);

      if (/^<[^!?/][^>]*[^/]>\s*$/.test(token) && !/<\/[^>]+>$/.test(token)) {
        indentLevel += 1;
      }
    }

    return formatted.join("\n");
  }

  globalThis.__mikuprojectMsprojectXmlDom = {
    textContent,
    parseBoolean,
    parseNumber,
    parseOutlineCodeMasks,
    parseOutlineCodeValues,
    parseWeekDays,
    appendTextElement,
    appendWeekDays,
    parseWorkingTimes,
    appendWorkingTimes,
    parseXmlDocument,
    formatXml
  };
})();
