/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
  type NativeSvgExportOptions = {
    holidayDates?: string[];
    displayDaysBeforeBaseDate?: number;
    displayDaysAfterBaseDate?: number;
    useBusinessDaysForDisplayRange?: boolean;
    useBusinessDaysForProgressBand?: boolean;
    labelMode?: "near" | "list";
  };

  type MonthlyCalendarSvgArchive = {
    entries: Array<{
      fileName: string;
      svg: string;
    }>;
    zipBytes: Uint8Array;
  };

  type NativeSvgFacade = {
    exportNativeSvg(model: ProjectModel, options?: NativeSvgExportOptions): string;
    exportWeeklyNativeSvg(model: ProjectModel, options?: NativeSvgExportOptions): string;
    exportMonthlyWbsCalendarSvgArchive(model: ProjectModel): MonthlyCalendarSvgArchive;
    collectWbsHolidayDates(model: ProjectModel): string[];
  };

  type CreateNativeSvgFacadeConfig = {
    collectWbsHolidayDates(model: ProjectModel): string[];
    exportNativeSvg(model: ProjectModel, options?: NativeSvgExportOptions): string;
    exportWeeklyNativeSvg(model: ProjectModel, options?: NativeSvgExportOptions): string;
    exportMonthlyWbsCalendarSvgArchive(model: ProjectModel): MonthlyCalendarSvgArchive;
  };

  function createNativeSvgFacade(config: CreateNativeSvgFacadeConfig): NativeSvgFacade {
    return {
      collectWbsHolidayDates: (model) => config.collectWbsHolidayDates(model),
      exportNativeSvg: (model, options) => config.exportNativeSvg(model, options),
      exportWeeklyNativeSvg: (model, options) => config.exportWeeklyNativeSvg(model, options),
      exportMonthlyWbsCalendarSvgArchive: (model) => config.exportMonthlyWbsCalendarSvgArchive(model)
    };
  }

  (globalThis as typeof globalThis & {
    __mikuprojectWbsSvgPublic?: {
      createNativeSvgFacade: (config: CreateNativeSvgFacadeConfig) => NativeSvgFacade;
    };
  }).__mikuprojectWbsSvgPublic = {
    createNativeSvgFacade
  };
})();
