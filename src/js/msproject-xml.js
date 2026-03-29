/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const SAMPLE_PROJECT_DRAFT_VIEW = {
        view_type: "project_draft_view",
        project: {
            name: "mikuproject開発",
            planned_start: "2026-03-16",
            planned_finish: "2026-04-01"
        },
        tasks: [
            {
                uid: "draft-100",
                name: "基盤整備",
                parent_uid: null,
                position: 0,
                is_summary: true,
                percent_complete: 100,
                planned_start: "2026-03-16",
                planned_finish: "2026-03-17"
            },
            {
                uid: "draft-110",
                name: "着手",
                parent_uid: "draft-100",
                position: 0,
                is_milestone: true,
                percent_complete: 100,
                planned_start: "2026-03-16",
                planned_finish: "2026-03-16"
            },
            {
                uid: "draft-120",
                name: "初期実装（MS Project XML 調査・基軸フォーマット選定・内部モデルの概要確定）",
                parent_uid: "draft-100",
                position: 1,
                percent_complete: 100,
                planned_start: "2026-03-16",
                planned_finish: "2026-03-16"
            },
            {
                uid: "draft-130",
                name: "round-trip拡張（MS Project XML → 内部JSON形式 → MS Project XML の往復対応）",
                parent_uid: "draft-100",
                position: 2,
                percent_complete: 100,
                planned_start: "2026-03-17",
                planned_finish: "2026-03-17"
            },
            {
                uid: "draft-150",
                name: "架空検討フェーズ【架空】",
                parent_uid: null,
                position: 1,
                is_summary: true,
                percent_complete: 25,
                planned_start: "2026-03-19",
                planned_finish: "2026-03-25"
            },
            {
                uid: "draft-160",
                name: "ユーザー操作フローの見直し【架空】",
                parent_uid: "draft-150",
                position: 0,
                percent_complete: 50,
                planned_start: "2026-03-19",
                planned_finish: "2026-03-23"
            },
            {
                uid: "draft-170",
                name: "画面構成の再整理【架空】",
                parent_uid: "draft-150",
                position: 1,
                planned_start: "2026-03-24",
                planned_finish: "2026-03-25"
            },
            {
                uid: "draft-200",
                name: "XLSX / UI 強化",
                parent_uid: null,
                position: 2,
                is_summary: true,
                planned_start: "2026-03-27",
                planned_finish: "2026-03-28"
            },
            {
                uid: "draft-210",
                name: "GitHub リポジトリ独立化",
                parent_uid: "draft-200",
                position: 0,
                is_milestone: true,
                planned_start: "2026-03-27",
                planned_finish: "2026-03-27"
            },
            {
                uid: "draft-220",
                name: "MS Project XML と XLSX の相互変換・round-trip実装",
                parent_uid: "draft-200",
                position: 1,
                planned_start: "2026-03-27",
                planned_finish: "2026-03-27"
            },
            {
                uid: "draft-230",
                name: "XLSXレイアウト再設計・再整理",
                parent_uid: "draft-200",
                position: 2,
                planned_start: "2026-03-28",
                planned_finish: "2026-03-28"
            },
            {
                uid: "draft-300",
                name: "リリース",
                parent_uid: null,
                position: 3,
                is_summary: true,
                planned_start: "2026-03-29",
                planned_finish: "2026-03-29"
            },
            {
                uid: "draft-310",
                name: "v1.0 リリース",
                parent_uid: "draft-300",
                position: 0,
                is_milestone: true,
                planned_start: "2026-03-29",
                planned_finish: "2026-03-29"
            }
        ]
    };
    const SAMPLE_XML = buildSampleXml();
    /*
  
    <Title>Sample Project Title</Title>
    <Company>Local HTML Tools</Company>
    <Author>Toshiki Iga</Author>
    <CreationDate>2026-03-16T08:30:00</CreationDate>
    <LastSaved>2026-03-16T09:10:00</LastSaved>
    <SaveVersion>14</SaveVersion>
    <CurrentDate>2026-03-16T09:00:00</CurrentDate>
    <StartDate>2026-03-16T09:00:00</StartDate>
    <FinishDate>2026-03-31T18:00:00</FinishDate>
    <ScheduleFromStart>1</ScheduleFromStart>
    <DefaultStartTime>09:00:00</DefaultStartTime>
    <DefaultFinishTime>18:00:00</DefaultFinishTime>
    <MinutesPerDay>480</MinutesPerDay>
    <MinutesPerWeek>2400</MinutesPerWeek>
    <DaysPerMonth>20</DaysPerMonth>
    <StatusDate>2026-03-19T09:00:00</StatusDate>
    <WeekStartDay>1</WeekStartDay>
    <WorkFormat>2</WorkFormat>
    <DurationFormat>7</DurationFormat>
    <CurrencyCode>JPY</CurrencyCode>
    <CurrencyDigits>0</CurrencyDigits>
    <CurrencySymbol>¥</CurrencySymbol>
    <CurrencySymbolPosition>0</CurrencySymbolPosition>
    <FYStartDate>2026-04-01T00:00:00</FYStartDate>
    <FiscalYearStart>1</FiscalYearStart>
    <CriticalSlackLimit>0</CriticalSlackLimit>
    <DefaultTaskType>1</DefaultTaskType>
    <DefaultFixedCostAccrual>2</DefaultFixedCostAccrual>
    <DefaultStandardRate>5000/h</DefaultStandardRate>
    <DefaultOvertimeRate>7000/h</DefaultOvertimeRate>
    <DefaultTaskEVMethod>0</DefaultTaskEVMethod>
    <NewTaskStartDate>0</NewTaskStartDate>
    <NewTasksAreManual>0</NewTasksAreManual>
    <NewTasksEffortDriven>1</NewTasksEffortDriven>
    <NewTasksEstimated>1</NewTasksEstimated>
    <ActualsInSync>0</ActualsInSync>
    <EditableActualCosts>1</EditableActualCosts>
    <HonorConstraints>1</HonorConstraints>
    <InsertedProjectsLikeSummary>1</InsertedProjectsLikeSummary>
    <MultipleCriticalPaths>0</MultipleCriticalPaths>
    <TaskUpdatesResource>1</TaskUpdatesResource>
    <UpdateManuallyScheduledTasksWhenEditingLinks>0</UpdateManuallyScheduledTasksWhenEditingLinks>
    <CalendarUID>1</CalendarUID>
    <OutlineCodes>
      <OutlineCode>
        <FieldID>188743731</FieldID>
        <FieldName>Outline Code1</FieldName>
        <Alias>Phase</Alias>
        <OnlyTableValues>1</OnlyTableValues>
        <Enterprise>0</Enterprise>
        <ResourceSubstitutionEnabled>0</ResourceSubstitutionEnabled>
        <LeafOnly>0</LeafOnly>
        <AllLevelsRequired>0</AllLevelsRequired>
        <Masks>
          <Mask>
            <Level>1</Level>
            <Mask>*</Mask>
            <Length>0</Length>
            <Sequence>0</Sequence>
          </Mask>
        </Masks>
        <Values>
          <Value>
            <Value>PLAN</Value>
            <Description>Planning</Description>
          </Value>
          <Value>
            <Value>BUILD</Value>
            <Description>Implementation</Description>
          </Value>
        </Values>
      </OutlineCode>
    </OutlineCodes>
    <WBSMasks>
      <WBSMask>
        <Level>1</Level>
        <Mask>A</Mask>
        <Length>1</Length>
        <Sequence>1</Sequence>
      </WBSMask>
      <WBSMask>
        <Level>2</Level>
        <Mask>00</Mask>
        <Length>2</Length>
        <Sequence>1</Sequence>
      </WBSMask>
    </WBSMasks>
    <ExtendedAttributes>
      <ExtendedAttribute>
        <FieldID>188743734</FieldID>
        <FieldName>Text1</FieldName>
        <Alias>Owner</Alias>
        <CalculationType>0</CalculationType>
        <RestrictValues>0</RestrictValues>
        <AppendNewValues>1</AppendNewValues>
      </ExtendedAttribute>
    </ExtendedAttributes>
    <Calendars>
      <Calendar>
        <UID>1</UID>
        <Name>Standard</Name>
        <IsBaseCalendar>1</IsBaseCalendar>
        <IsBaselineCalendar>1</IsBaselineCalendar>
        <Exceptions>
          <Exception>
            <Name>元日（公式）</Name>
            <FromDate>2026-01-01T00:00:00</FromDate>
            <ToDate>2026-01-01T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>成人の日（公式）</Name>
            <FromDate>2026-01-12T00:00:00</FromDate>
            <ToDate>2026-01-12T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>建国記念の日（公式）</Name>
            <FromDate>2026-02-11T00:00:00</FromDate>
            <ToDate>2026-02-11T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>天皇誕生日（公式）</Name>
            <FromDate>2026-02-23T00:00:00</FromDate>
            <ToDate>2026-02-23T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>春分の日（公式）</Name>
            <FromDate>2026-03-20T00:00:00</FromDate>
            <ToDate>2026-03-20T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>昭和の日（公式）</Name>
            <FromDate>2026-04-29T00:00:00</FromDate>
            <ToDate>2026-04-29T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>憲法記念日（公式）</Name>
            <FromDate>2026-05-03T00:00:00</FromDate>
            <ToDate>2026-05-03T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>みどりの日（公式）</Name>
            <FromDate>2026-05-04T00:00:00</FromDate>
            <ToDate>2026-05-04T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>こどもの日（公式）</Name>
            <FromDate>2026-05-05T00:00:00</FromDate>
            <ToDate>2026-05-05T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>休日（公式）</Name>
            <FromDate>2026-05-06T00:00:00</FromDate>
            <ToDate>2026-05-06T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>海の日（公式）</Name>
            <FromDate>2026-07-20T00:00:00</FromDate>
            <ToDate>2026-07-20T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>山の日（公式）</Name>
            <FromDate>2026-08-11T00:00:00</FromDate>
            <ToDate>2026-08-11T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>敬老の日（公式）</Name>
            <FromDate>2026-09-21T00:00:00</FromDate>
            <ToDate>2026-09-21T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>休日（公式）</Name>
            <FromDate>2026-09-22T00:00:00</FromDate>
            <ToDate>2026-09-22T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>秋分の日（公式）</Name>
            <FromDate>2026-09-23T00:00:00</FromDate>
            <ToDate>2026-09-23T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>スポーツの日（公式）</Name>
            <FromDate>2026-10-12T00:00:00</FromDate>
            <ToDate>2026-10-12T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>文化の日（公式）</Name>
            <FromDate>2026-11-03T00:00:00</FromDate>
            <ToDate>2026-11-03T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>勤労感謝の日（公式）</Name>
            <FromDate>2026-11-23T00:00:00</FromDate>
            <ToDate>2026-11-23T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>元日（公式）</Name>
            <FromDate>2027-01-01T00:00:00</FromDate>
            <ToDate>2027-01-01T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>成人の日（公式）</Name>
            <FromDate>2027-01-11T00:00:00</FromDate>
            <ToDate>2027-01-11T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>建国記念の日（公式）</Name>
            <FromDate>2027-02-11T00:00:00</FromDate>
            <ToDate>2027-02-11T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>天皇誕生日（公式）</Name>
            <FromDate>2027-02-23T00:00:00</FromDate>
            <ToDate>2027-02-23T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>春分の日（公式）</Name>
            <FromDate>2027-03-21T00:00:00</FromDate>
            <ToDate>2027-03-21T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>休日（公式）</Name>
            <FromDate>2027-03-22T00:00:00</FromDate>
            <ToDate>2027-03-22T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>昭和の日（公式）</Name>
            <FromDate>2027-04-29T00:00:00</FromDate>
            <ToDate>2027-04-29T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>憲法記念日（公式）</Name>
            <FromDate>2027-05-03T00:00:00</FromDate>
            <ToDate>2027-05-03T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>みどりの日（公式）</Name>
            <FromDate>2027-05-04T00:00:00</FromDate>
            <ToDate>2027-05-04T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>こどもの日（公式）</Name>
            <FromDate>2027-05-05T00:00:00</FromDate>
            <ToDate>2027-05-05T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>海の日（公式）</Name>
            <FromDate>2027-07-19T00:00:00</FromDate>
            <ToDate>2027-07-19T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>山の日（公式）</Name>
            <FromDate>2027-08-11T00:00:00</FromDate>
            <ToDate>2027-08-11T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>敬老の日（公式）</Name>
            <FromDate>2027-09-20T00:00:00</FromDate>
            <ToDate>2027-09-20T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>秋分の日（公式）</Name>
            <FromDate>2027-09-23T00:00:00</FromDate>
            <ToDate>2027-09-23T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>スポーツの日（公式）</Name>
            <FromDate>2027-10-11T00:00:00</FromDate>
            <ToDate>2027-10-11T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>文化の日（公式）</Name>
            <FromDate>2027-11-03T00:00:00</FromDate>
            <ToDate>2027-11-03T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>勤労感謝の日（公式）</Name>
            <FromDate>2027-11-23T00:00:00</FromDate>
            <ToDate>2027-11-23T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>元日（推定）</Name>
            <FromDate>2028-01-01T00:00:00</FromDate>
            <ToDate>2028-01-01T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>成人の日（推定）</Name>
            <FromDate>2028-01-10T00:00:00</FromDate>
            <ToDate>2028-01-10T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>建国記念の日（推定）</Name>
            <FromDate>2028-02-11T00:00:00</FromDate>
            <ToDate>2028-02-11T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>天皇誕生日（推定）</Name>
            <FromDate>2028-02-23T00:00:00</FromDate>
            <ToDate>2028-02-23T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>春分の日（推定）</Name>
            <FromDate>2028-03-20T00:00:00</FromDate>
            <ToDate>2028-03-20T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>昭和の日（推定）</Name>
            <FromDate>2028-04-29T00:00:00</FromDate>
            <ToDate>2028-04-29T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>憲法記念日（推定）</Name>
            <FromDate>2028-05-03T00:00:00</FromDate>
            <ToDate>2028-05-03T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>みどりの日（推定）</Name>
            <FromDate>2028-05-04T00:00:00</FromDate>
            <ToDate>2028-05-04T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>こどもの日（推定）</Name>
            <FromDate>2028-05-05T00:00:00</FromDate>
            <ToDate>2028-05-05T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>海の日（推定）</Name>
            <FromDate>2028-07-17T00:00:00</FromDate>
            <ToDate>2028-07-17T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>山の日（推定）</Name>
            <FromDate>2028-08-11T00:00:00</FromDate>
            <ToDate>2028-08-11T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>敬老の日（推定）</Name>
            <FromDate>2028-09-18T00:00:00</FromDate>
            <ToDate>2028-09-18T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>秋分の日（推定）</Name>
            <FromDate>2028-09-22T00:00:00</FromDate>
            <ToDate>2028-09-22T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>スポーツの日（推定）</Name>
            <FromDate>2028-10-09T00:00:00</FromDate>
            <ToDate>2028-10-09T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>文化の日（推定）</Name>
            <FromDate>2028-11-03T00:00:00</FromDate>
            <ToDate>2028-11-03T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>勤労感謝の日（推定）</Name>
            <FromDate>2028-11-23T00:00:00</FromDate>
            <ToDate>2028-11-23T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>元日（推定）</Name>
            <FromDate>2029-01-01T00:00:00</FromDate>
            <ToDate>2029-01-01T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>成人の日（推定）</Name>
            <FromDate>2029-01-08T00:00:00</FromDate>
            <ToDate>2029-01-08T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>建国記念の日（推定）</Name>
            <FromDate>2029-02-11T00:00:00</FromDate>
            <ToDate>2029-02-11T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>休日（推定）</Name>
            <FromDate>2029-02-12T00:00:00</FromDate>
            <ToDate>2029-02-12T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>天皇誕生日（推定）</Name>
            <FromDate>2029-02-23T00:00:00</FromDate>
            <ToDate>2029-02-23T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>春分の日（推定）</Name>
            <FromDate>2029-03-20T00:00:00</FromDate>
            <ToDate>2029-03-20T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>昭和の日（推定）</Name>
            <FromDate>2029-04-29T00:00:00</FromDate>
            <ToDate>2029-04-29T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>休日（推定）</Name>
            <FromDate>2029-04-30T00:00:00</FromDate>
            <ToDate>2029-04-30T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>憲法記念日（推定）</Name>
            <FromDate>2029-05-03T00:00:00</FromDate>
            <ToDate>2029-05-03T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>みどりの日（推定）</Name>
            <FromDate>2029-05-04T00:00:00</FromDate>
            <ToDate>2029-05-04T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>こどもの日（推定）</Name>
            <FromDate>2029-05-05T00:00:00</FromDate>
            <ToDate>2029-05-05T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>海の日（推定）</Name>
            <FromDate>2029-07-16T00:00:00</FromDate>
            <ToDate>2029-07-16T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>山の日（推定）</Name>
            <FromDate>2029-08-11T00:00:00</FromDate>
            <ToDate>2029-08-11T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>敬老の日（推定）</Name>
            <FromDate>2029-09-17T00:00:00</FromDate>
            <ToDate>2029-09-17T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>秋分の日（推定）</Name>
            <FromDate>2029-09-23T00:00:00</FromDate>
            <ToDate>2029-09-23T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>休日（推定）</Name>
            <FromDate>2029-09-24T00:00:00</FromDate>
            <ToDate>2029-09-24T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>スポーツの日（推定）</Name>
            <FromDate>2029-10-08T00:00:00</FromDate>
            <ToDate>2029-10-08T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>文化の日（推定）</Name>
            <FromDate>2029-11-03T00:00:00</FromDate>
            <ToDate>2029-11-03T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>勤労感謝の日（推定）</Name>
            <FromDate>2029-11-23T00:00:00</FromDate>
            <ToDate>2029-11-23T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>元日（推定）</Name>
            <FromDate>2030-01-01T00:00:00</FromDate>
            <ToDate>2030-01-01T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>成人の日（推定）</Name>
            <FromDate>2030-01-14T00:00:00</FromDate>
            <ToDate>2030-01-14T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>建国記念の日（推定）</Name>
            <FromDate>2030-02-11T00:00:00</FromDate>
            <ToDate>2030-02-11T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>天皇誕生日（推定）</Name>
            <FromDate>2030-02-23T00:00:00</FromDate>
            <ToDate>2030-02-23T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>春分の日（推定）</Name>
            <FromDate>2030-03-20T00:00:00</FromDate>
            <ToDate>2030-03-20T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>昭和の日（推定）</Name>
            <FromDate>2030-04-29T00:00:00</FromDate>
            <ToDate>2030-04-29T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>憲法記念日（推定）</Name>
            <FromDate>2030-05-03T00:00:00</FromDate>
            <ToDate>2030-05-03T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>みどりの日（推定）</Name>
            <FromDate>2030-05-04T00:00:00</FromDate>
            <ToDate>2030-05-04T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>こどもの日（推定）</Name>
            <FromDate>2030-05-05T00:00:00</FromDate>
            <ToDate>2030-05-05T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>休日（推定）</Name>
            <FromDate>2030-05-06T00:00:00</FromDate>
            <ToDate>2030-05-06T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>海の日（推定）</Name>
            <FromDate>2030-07-15T00:00:00</FromDate>
            <ToDate>2030-07-15T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>山の日（推定）</Name>
            <FromDate>2030-08-11T00:00:00</FromDate>
            <ToDate>2030-08-11T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>休日（推定）</Name>
            <FromDate>2030-08-12T00:00:00</FromDate>
            <ToDate>2030-08-12T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>敬老の日（推定）</Name>
            <FromDate>2030-09-16T00:00:00</FromDate>
            <ToDate>2030-09-16T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>秋分の日（推定）</Name>
            <FromDate>2030-09-23T00:00:00</FromDate>
            <ToDate>2030-09-23T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>スポーツの日（推定）</Name>
            <FromDate>2030-10-14T00:00:00</FromDate>
            <ToDate>2030-10-14T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>文化の日（推定）</Name>
            <FromDate>2030-11-03T00:00:00</FromDate>
            <ToDate>2030-11-03T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>休日（推定）</Name>
            <FromDate>2030-11-04T00:00:00</FromDate>
            <ToDate>2030-11-04T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
          <Exception>
            <Name>勤労感謝の日（推定）</Name>
            <FromDate>2030-11-23T00:00:00</FromDate>
            <ToDate>2030-11-23T23:59:59</ToDate>
            <DayWorking>0</DayWorking>
          </Exception>
        </Exceptions>
        <WeekDays>
          <WeekDay>
            <DayType>2</DayType>
            <DayWorking>1</DayWorking>
            <WorkingTimes>
              <WorkingTime>
                <FromTime>09:00:00</FromTime>
                <ToTime>12:00:00</ToTime>
              </WorkingTime>
              <WorkingTime>
                <FromTime>13:00:00</FromTime>
                <ToTime>18:00:00</ToTime>
              </WorkingTime>
            </WorkingTimes>
          </WeekDay>
        </WeekDays>
      </Calendar>
      <Calendar>
        <UID>2</UID>
        <Name>Development</Name>
        <IsBaseCalendar>0</IsBaseCalendar>
        <BaseCalendarUID>1</BaseCalendarUID>
        <WorkWeeks>
          <WorkWeek>
            <Name>Spring Sprint</Name>
            <FromDate>2026-03-16T00:00:00</FromDate>
            <ToDate>2026-03-31T23:59:59</ToDate>
            <WeekDays>
              <WeekDay>
                <DayType>2</DayType>
                <DayWorking>1</DayWorking>
                <WorkingTimes>
                  <WorkingTime>
                    <FromTime>10:00:00</FromTime>
                    <ToTime>18:00:00</ToTime>
                  </WorkingTime>
                </WorkingTimes>
              </WeekDay>
            </WeekDays>
          </WorkWeek>
        </WorkWeeks>
        <WeekDays>
          <WeekDay>
            <DayType>6</DayType>
            <DayWorking>1</DayWorking>
            <WorkingTimes>
              <WorkingTime>
                <FromTime>10:00:00</FromTime>
                <ToTime>15:00:00</ToTime>
              </WorkingTime>
            </WorkingTimes>
          </WeekDay>
        </WeekDays>
      </Calendar>
    </Calendars>
    <Tasks>
      <Task>
        <UID>1</UID>
        <ID>1</ID>
        <Name>Project Summary</Name>
        <OutlineLevel>1</OutlineLevel>
        <OutlineNumber>1</OutlineNumber>
        <WBS>1</WBS>
        <Type>1</Type>
        <CalendarUID>1</CalendarUID>
        <Priority>500</Priority>
        <Start>2026-03-16T09:00:00</Start>
        <Finish>2026-03-20T18:00:00</Finish>
        <Duration>PT40H0M0S</Duration>
        <StartVariance>PT0H0M0S</StartVariance>
        <FinishVariance>PT0H0M0S</FinishVariance>
        <Work>PT40H0M0S</Work>
        <WorkVariance>PT0H0M0S</WorkVariance>
        <TotalSlack>PT0H0M0S</TotalSlack>
        <FreeSlack>PT0H0M0S</FreeSlack>
        <Cost>200000</Cost>
        <ActualCost>100000</ActualCost>
        <RemainingCost>100000</RemainingCost>
        <RemainingWork>PT20H0M0S</RemainingWork>
        <ActualWork>PT20H0M0S</ActualWork>
        <Milestone>0</Milestone>
        <Summary>1</Summary>
        <Critical>0</Critical>
        <PercentComplete>50</PercentComplete>
        <PercentWorkComplete>50</PercentWorkComplete>
      </Task>
      <Task>
        <UID>2</UID>
        <ID>2</ID>
        <Name>Design</Name>
        <OutlineLevel>2</OutlineLevel>
        <OutlineNumber>1.1</OutlineNumber>
        <WBS>1.1</WBS>
        <Type>1</Type>
        <CalendarUID>1</CalendarUID>
        <Priority>500</Priority>
        <Start>2026-03-16T09:00:00</Start>
        <Finish>2026-03-17T18:00:00</Finish>
        <Duration>PT16H0M0S</Duration>
        <ActualStart>2026-03-16T09:00:00</ActualStart>
        <ActualFinish>2026-03-17T18:00:00</ActualFinish>
        <StartVariance>PT0H0M0S</StartVariance>
        <FinishVariance>PT0H0M0S</FinishVariance>
        <Work>PT16H0M0S</Work>
        <WorkVariance>PT0H0M0S</WorkVariance>
        <TotalSlack>PT0H0M0S</TotalSlack>
        <FreeSlack>PT0H0M0S</FreeSlack>
        <Cost>80000</Cost>
        <ActualCost>80000</ActualCost>
        <RemainingCost>0</RemainingCost>
        <RemainingWork>PT0H0M0S</RemainingWork>
        <ActualWork>PT16H0M0S</ActualWork>
        <Milestone>0</Milestone>
        <Summary>0</Summary>
        <Critical>0</Critical>
        <PercentComplete>100</PercentComplete>
        <PercentWorkComplete>100</PercentWorkComplete>
        <Notes>Design completed</Notes>
        <ExtendedAttribute>
          <FieldID>188743734</FieldID>
          <Value>Miku</Value>
        </ExtendedAttribute>
        <Baseline>
          <Number>0</Number>
          <Start>2026-03-16T09:00:00</Start>
          <Finish>2026-03-17T18:00:00</Finish>
          <Work>PT16H0M0S</Work>
          <Cost>80000</Cost>
        </Baseline>
        <TimephasedData>
          <Type>1</Type>
          <UID>2</UID>
          <Start>2026-03-16T09:00:00</Start>
          <Finish>2026-03-16T18:00:00</Finish>
          <Unit>2</Unit>
          <Value>PT8H0M0S</Value>
        </TimephasedData>
      </Task>
      <Task>
        <UID>3</UID>
        <ID>3</ID>
        <Name>Implementation</Name>
        <OutlineLevel>2</OutlineLevel>
        <OutlineNumber>1.2</OutlineNumber>
        <WBS>1.2</WBS>
        <Type>1</Type>
        <CalendarUID>1</CalendarUID>
        <Priority>700</Priority>
        <Start>2026-03-18T09:00:00</Start>
        <Finish>2026-03-20T18:00:00</Finish>
        <Duration>PT24H0M0S</Duration>
        <Deadline>2026-03-21T18:00:00</Deadline>
        <StartVariance>PT0H0M0S</StartVariance>
        <FinishVariance>PT0H0M0S</FinishVariance>
        <Work>PT24H0M0S</Work>
        <WorkVariance>PT0H0M0S</WorkVariance>
        <TotalSlack>PT4H0M0S</TotalSlack>
        <FreeSlack>PT2H0M0S</FreeSlack>
        <Cost>120000</Cost>
        <ActualCost>0</ActualCost>
        <RemainingCost>120000</RemainingCost>
        <RemainingWork>PT24H0M0S</RemainingWork>
        <ActualWork>PT0H0M0S</ActualWork>
        <ConstraintType>4</ConstraintType>
        <ConstraintDate>2026-03-18T09:00:00</ConstraintDate>
        <Milestone>0</Milestone>
        <Summary>0</Summary>
        <Critical>1</Critical>
        <PercentComplete>0</PercentComplete>
        <PercentWorkComplete>0</PercentWorkComplete>
        <Notes>Implementation starts after design</Notes>
        <PredecessorLink>
          <PredecessorUID>2</PredecessorUID>
          <Type>1</Type>
          <LinkLag>PT0H0M0S</LinkLag>
        </PredecessorLink>
      </Task>
    </Tasks>
    <Resources>
      <Resource>
        <UID>1</UID>
        <ID>1</ID>
        <Name>Miku</Name>
        <Type>1</Type>
        <Initials>MK</Initials>
        <Group>Engineering</Group>
        <WorkGroup>0</WorkGroup>
        <MaxUnits>1</MaxUnits>
        <CalendarUID>2</CalendarUID>
        <StandardRate>5000/h</StandardRate>
        <StandardRateFormat>2</StandardRateFormat>
        <OvertimeRate>7000/h</OvertimeRate>
        <OvertimeRateFormat>2</OvertimeRateFormat>
        <CostPerUse>1000</CostPerUse>
        <Work>PT40H0M0S</Work>
        <ActualWork>PT20H0M0S</ActualWork>
        <RemainingWork>PT20H0M0S</RemainingWork>
        <Cost>200000</Cost>
        <ActualCost>100000</ActualCost>
        <RemainingCost>100000</RemainingCost>
        <PercentWorkComplete>50</PercentWorkComplete>
        <ExtendedAttribute>
          <FieldID>188743737</FieldID>
          <Value>Platform Team</Value>
        </ExtendedAttribute>
        <Baseline>
          <Number>0</Number>
          <Start>2026-03-16T09:00:00</Start>
          <Finish>2026-03-20T18:00:00</Finish>
          <Work>PT40H0M0S</Work>
          <Cost>200000</Cost>
        </Baseline>
        <TimephasedData>
          <Type>1</Type>
          <UID>1</UID>
          <Start>2026-03-16T09:00:00</Start>
          <Finish>2026-03-16T18:00:00</Finish>
          <Unit>2</Unit>
          <Value>PT8H0M0S</Value>
        </TimephasedData>
      </Resource>
    </Resources>
    <Assignments>
      <Assignment>
        <UID>1</UID>
        <TaskUID>2</TaskUID>
        <ResourceUID>1</ResourceUID>
        <Start>2026-03-16T09:00:00</Start>
        <Finish>2026-03-17T18:00:00</Finish>
        <StartVariance>PT0H0M0S</StartVariance>
        <FinishVariance>PT0H0M0S</FinishVariance>
        <Delay>PT0H0M0S</Delay>
        <Milestone>0</Milestone>
        <WorkContour>0</WorkContour>
        <Units>1</Units>
        <Work>PT16H0M0S</Work>
        <Cost>80000</Cost>
        <ActualCost>40000</ActualCost>
        <RemainingCost>40000</RemainingCost>
        <PercentWorkComplete>50</PercentWorkComplete>
        <OvertimeWork>PT2H0M0S</OvertimeWork>
        <ActualOvertimeWork>PT1H0M0S</ActualOvertimeWork>
        <ActualWork>PT8H0M0S</ActualWork>
        <RemainingWork>PT8H0M0S</RemainingWork>
        <ExtendedAttribute>
          <FieldID>255852547</FieldID>
          <Value>Design Slot</Value>
        </ExtendedAttribute>
        <Baseline>
          <Number>0</Number>
          <Start>2026-03-16T09:00:00</Start>
          <Finish>2026-03-17T18:00:00</Finish>
          <Work>PT16H0M0S</Work>
          <Cost>80000</Cost>
        </Baseline>
        <TimephasedData>
          <Type>1</Type>
          <UID>1</UID>
          <Start>2026-03-16T09:00:00</Start>
          <Finish>2026-03-16T18:00:00</Finish>
          <Unit>2</Unit>
          <Value>PT8H0M0S</Value>
        </TimephasedData>
      </Assignment>
      <Assignment>
        <UID>2</UID>
        <TaskUID>3</TaskUID>
        <ResourceUID>1</ResourceUID>
        <Start>2026-03-18T09:00:00</Start>
        <Finish>2026-03-20T18:00:00</Finish>
        <StartVariance>PT0H0M0S</StartVariance>
        <FinishVariance>PT0H0M0S</FinishVariance>
        <Delay>PT0H0M0S</Delay>
        <Milestone>0</Milestone>
        <WorkContour>0</WorkContour>
        <Units>1</Units>
        <Work>PT24H0M0S</Work>
        <Cost>120000</Cost>
        <ActualCost>0</ActualCost>
        <RemainingCost>120000</RemainingCost>
        <PercentWorkComplete>0</PercentWorkComplete>
        <OvertimeWork>PT0H0M0S</OvertimeWork>
        <ActualOvertimeWork>PT0H0M0S</ActualOvertimeWork>
        <ActualWork>PT0H0M0S</ActualWork>
        <RemainingWork>PT24H0M0S</RemainingWork>
      </Assignment>
    </Assignments>
  
  
    */
    function buildSampleXml() {
        const sampleModel = importProjectDraftView(SAMPLE_PROJECT_DRAFT_VIEW);
        sampleModel.project.currentDate = "2026-03-23T09:00:00";
        sampleModel.project.statusDate = "2026-03-23T09:00:00";
        return exportMsProjectXml(sampleModel);
    }
    function textContent(parent, tagName) {
        const element = parent.getElementsByTagName(tagName)[0];
        return String((element === null || element === void 0 ? void 0 : element.textContent) || "").trim();
    }
    function parseBoolean(value) {
        return value === "1" || value.toLowerCase() === "true";
    }
    function parseNumber(value, defaultValue = 0) {
        const parsed = Number(value);
        return Number.isFinite(parsed) ? parsed : defaultValue;
    }
    function parseDateValue(value) {
        if (!value) {
            return null;
        }
        const timestamp = Date.parse(value);
        return Number.isFinite(timestamp) ? timestamp : null;
    }
    function parseDateOnly(value) {
        const text = String(value || "").trim().slice(0, 10);
        if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) {
            return null;
        }
        const parsed = new Date(`${text}T00:00:00`);
        return Number.isNaN(parsed.getTime()) ? null : parsed;
    }
    function formatDateOnly(value) {
        const year = value.getFullYear();
        const month = String(value.getMonth() + 1).padStart(2, "0");
        const day = String(value.getDate()).padStart(2, "0");
        return `${year}-${month}-${day}`;
    }
    function addDateDays(base, days) {
        const next = new Date(base.getTime());
        next.setDate(next.getDate() + days);
        return next;
    }
    function addDateYears(base, years) {
        const next = new Date(base.getTime());
        next.setFullYear(next.getFullYear() + years);
        return next;
    }
    function toMsDayType(value) {
        const day = value.getDay();
        return day === 0 ? 1 : day + 1;
    }
    function buildNthWeekdayOfMonth(year, monthIndex, jsWeekday, nth) {
        const first = new Date(year, monthIndex, 1);
        const offset = (jsWeekday - first.getDay() + 7) % 7;
        return new Date(year, monthIndex, 1 + offset + ((nth - 1) * 7));
    }
    function calculateVernalEquinoxDay(year) {
        return Math.floor(20.8431 + (0.242194 * (year - 1980)) - Math.floor((year - 1980) / 4));
    }
    function calculateAutumnalEquinoxDay(year) {
        return Math.floor(23.2488 + (0.242194 * (year - 1980)) - Math.floor((year - 1980) / 4));
    }
    function buildJapaneseHolidayMapForYear(year) {
        const holidays = new Map();
        const addHoliday = (date, name) => {
            holidays.set(formatDateOnly(date), name);
        };
        addHoliday(new Date(year, 0, 1), "元日");
        addHoliday(buildNthWeekdayOfMonth(year, 0, 1, 2), "成人の日");
        addHoliday(new Date(year, 1, 11), "建国記念の日");
        addHoliday(new Date(year, 1, 23), "天皇誕生日");
        addHoliday(new Date(year, 2, calculateVernalEquinoxDay(year)), "春分の日");
        addHoliday(new Date(year, 3, 29), "昭和の日");
        addHoliday(new Date(year, 4, 3), "憲法記念日");
        addHoliday(new Date(year, 4, 4), "みどりの日");
        addHoliday(new Date(year, 4, 5), "こどもの日");
        addHoliday(buildNthWeekdayOfMonth(year, 6, 1, 3), "海の日");
        addHoliday(new Date(year, 7, 11), "山の日");
        addHoliday(buildNthWeekdayOfMonth(year, 8, 1, 3), "敬老の日");
        addHoliday(new Date(year, 8, calculateAutumnalEquinoxDay(year)), "秋分の日");
        addHoliday(buildNthWeekdayOfMonth(year, 9, 1, 2), "スポーツの日");
        addHoliday(new Date(year, 10, 3), "文化の日");
        addHoliday(new Date(year, 10, 23), "勤労感謝の日");
        const baseHolidayDates = Array.from(holidays.keys()).sort();
        for (const dateText of baseHolidayDates) {
            const date = parseDateOnly(dateText);
            if (!date || date.getDay() !== 0) {
                continue;
            }
            let substitute = addDateDays(date, 1);
            while (holidays.has(formatDateOnly(substitute))) {
                substitute = addDateDays(substitute, 1);
            }
            holidays.set(formatDateOnly(substitute), "休日");
        }
        const sortedDates = Array.from(holidays.keys()).sort();
        for (let index = 0; index < sortedDates.length - 1; index += 1) {
            const current = parseDateOnly(sortedDates[index]);
            const next = parseDateOnly(sortedDates[index + 1]);
            if (!current || !next) {
                continue;
            }
            const gapDays = Math.floor((next.getTime() - current.getTime()) / 86400000);
            if (gapDays !== 2) {
                continue;
            }
            const between = addDateDays(current, 1);
            const betweenText = formatDateOnly(between);
            if (holidays.has(betweenText) || between.getDay() === 0) {
                continue;
            }
            holidays.set(betweenText, "休日");
        }
        return new Map(Array.from(holidays.entries()).sort((left, right) => left[0].localeCompare(right[0])));
    }
    function buildDefaultWorkingTimes(project) {
        const start = project.defaultStartTime || "09:00:00";
        const finish = project.defaultFinishTime || "18:00:00";
        if (start < "12:00:00" && finish > "13:00:00") {
            return [
                { fromTime: start, toTime: "12:00:00" },
                { fromTime: "13:00:00", toTime: finish }
            ];
        }
        return [{ fromTime: start, toTime: finish }];
    }
    function buildDefaultStandardWeekDays(project) {
        const workingTimes = buildDefaultWorkingTimes(project);
        return Array.from({ length: 7 }, (_, index) => {
            const dayType = index + 1;
            const dayWorking = dayType !== 1 && dayType !== 7;
            return {
                dayType,
                dayWorking,
                workingTimes: dayWorking ? workingTimes.map((item) => ({ ...item })) : []
            };
        });
    }
    function buildDefaultJapaneseHolidayExceptions(project) {
        const start = parseDateOnly(project.startDate) || parseDateOnly(project.finishDate) || new Date();
        const finishLimit = parseDateOnly(project.finishDate) || start;
        const rangeStart = start.getTime() <= finishLimit.getTime() ? start : finishLimit;
        const rangeFinish = start.getTime() <= finishLimit.getTime() ? finishLimit : start;
        const exceptions = [];
        for (let year = rangeStart.getFullYear(); year <= rangeFinish.getFullYear(); year += 1) {
            const holidays = buildJapaneseHolidayMapForYear(year);
            for (const [dateText, name] of holidays.entries()) {
                const date = parseDateOnly(dateText);
                if (!date || date.getTime() < rangeStart.getTime() || date.getTime() > rangeFinish.getTime()) {
                    continue;
                }
                exceptions.push({
                    name,
                    fromDate: `${dateText}T00:00:00`,
                    toDate: `${dateText}T23:59:59`,
                    dayWorking: false,
                    workingTimes: []
                });
            }
        }
        return exceptions;
    }
    function findFallbackCalendarUid(model) {
        var _a;
        const baseCalendar = model.calendars.find((calendar) => calendar.isBaseCalendar);
        return (baseCalendar === null || baseCalendar === void 0 ? void 0 : baseCalendar.uid) || ((_a = model.calendars[0]) === null || _a === void 0 ? void 0 : _a.uid);
    }
    function allocateDefaultCalendarUid(model) {
        const usedUids = new Set(model.calendars.map((calendar) => String(calendar.uid)));
        let candidate = 1;
        while (usedUids.has(String(candidate))) {
            candidate += 1;
        }
        return String(candidate);
    }
    function ensureDefaultProjectCalendar(model) {
        if (model.calendars.length === 0) {
            const uid = allocateDefaultCalendarUid(model);
            model.calendars.push({
                uid,
                name: "Standard",
                isBaseCalendar: true,
                isBaselineCalendar: true,
                weekDays: buildDefaultStandardWeekDays(model.project),
                exceptions: buildDefaultJapaneseHolidayExceptions(model.project),
                workWeeks: []
            });
            model.project.calendarUID = uid;
        }
        return model;
    }
    function parseWeekDays(parent) {
        var _a;
        return Array.from(((_a = parent.getElementsByTagName("WeekDays")[0]) === null || _a === void 0 ? void 0 : _a.getElementsByTagName("WeekDay")) || []).map((weekDay) => {
            var _a;
            return ({
                dayType: parseNumber(textContent(weekDay, "DayType"), 0),
                dayWorking: parseBoolean(textContent(weekDay, "DayWorking")),
                workingTimes: Array.from(((_a = weekDay.getElementsByTagName("WorkingTimes")[0]) === null || _a === void 0 ? void 0 : _a.getElementsByTagName("WorkingTime")) || []).map((workingTime) => ({
                    fromTime: textContent(workingTime, "FromTime"),
                    toTime: textContent(workingTime, "ToTime")
                }))
            });
        });
    }
    function appendWeekDays(doc, parent, weekDays) {
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
    function parseWorkingTimes(parent) {
        var _a;
        return Array.from(((_a = parent.getElementsByTagName("WorkingTimes")[0]) === null || _a === void 0 ? void 0 : _a.getElementsByTagName("WorkingTime")) || []).map((workingTime) => ({
            fromTime: textContent(workingTime, "FromTime"),
            toTime: textContent(workingTime, "ToTime")
        }));
    }
    function appendWorkingTimes(doc, parent, workingTimes) {
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
    function parseOutlineCodeMasks(parent) {
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
    function parseOutlineCodeValues(parent) {
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
    function isPlaceholderUid(value) {
        return String(value || "").trim() === "0";
    }
    function isUnassignedResourceUid(value) {
        return String(value || "").trim() === "-65535";
    }
    function describeTask(task) {
        return `UID=${task.uid}${task.name ? ` (${task.name})` : ""}`;
    }
    function describeResource(resource) {
        return `UID=${resource.uid || "(なし)"}${resource.name ? ` (${resource.name})` : ""}`;
    }
    function describeCalendar(calendar) {
        return `UID=${calendar.uid}${calendar.name ? ` (${calendar.name})` : ""}`;
    }
    function describeAssignment(assignment) {
        return `UID=${assignment.uid || "(なし)"}`;
    }
    function describeTaskRef(model, taskUid) {
        if (!taskUid) {
            return "TaskUID=(なし)";
        }
        const task = model.tasks.find((item) => item.uid === taskUid);
        return task ? `TaskUID=${taskUid}${task.name ? ` (${task.name})` : ""}` : `TaskUID=${taskUid}`;
    }
    function describeResourceRef(model, resourceUid) {
        if (!resourceUid) {
            return "ResourceUID=(なし)";
        }
        const resource = model.resources.find((item) => item.uid === resourceUid);
        return resource ? `ResourceUID=${resourceUid}${resource.name ? ` (${resource.name})` : ""}` : `ResourceUID=${resourceUid}`;
    }
    function parseXmlDocument(xmlText) {
        const parser = new DOMParser();
        const xml = parser.parseFromString(xmlText, "application/xml");
        const parserError = xml.getElementsByTagName("parsererror")[0];
        if (parserError) {
            throw new Error("XML の解析に失敗しました");
        }
        return xml;
    }
    function normalizeMermaidText(value, fallback) {
        const text = String(value || fallback).replace(/[:：#,，]/g, " ").replace(/\s+/g, " ").trim();
        return text || fallback;
    }
    function normalizeMermaidGanttLabel(value, fallback, leadingPrefix) {
        const text = normalizeMermaidText(value, fallback);
        return /^\d/.test(text) ? `${leadingPrefix} ${text}` : text;
    }
    function normalizeMermaidTaskId(value, fallback) {
        return String(value || fallback).replace(/[^A-Za-z0-9_]/g, "_");
    }
    function toMermaidDuration(duration) {
        const text = String(duration || "").trim();
        if (!text) {
            return null;
        }
        const match = /^P(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)$/.exec(text);
        if (!match) {
            return null;
        }
        const hours = Number(match[1] || 0);
        const minutes = Number(match[2] || 0);
        const seconds = Number(match[3] || 0);
        const parts = [];
        if (hours > 0) {
            parts.push(`${hours}h`);
        }
        if (minutes > 0) {
            parts.push(`${minutes}m`);
        }
        if (seconds > 0) {
            parts.push(`${seconds}s`);
        }
        return parts.length > 0 ? parts.join(" ") : null;
    }
    function formatMermaidLag(duration) {
        const short = toMermaidDuration(duration);
        if (short) {
            return short;
        }
        return String(duration || "").trim();
    }
    function isZeroDuration(duration) {
        const text = String(duration || "").trim();
        return text === "" || text === "PT0H0M0S" || text === "PT0M0S" || text === "PT0S";
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
    function formatDependencyType(type) {
        if (type === undefined) {
            return "FS";
        }
        const typeMap = {
            0: "FF",
            1: "FS",
            2: "FF",
            3: "SF",
            4: "SS"
        };
        return typeMap[type] || String(type);
    }
    function parseDurationHours(duration) {
        const text = String(duration || "").trim();
        if (!text) {
            return undefined;
        }
        const match = /^(-)?P(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)$/.exec(text);
        if (!match) {
            return undefined;
        }
        const sign = match[1] ? -1 : 1;
        const hours = Number(match[2] || 0);
        const minutes = Number(match[3] || 0);
        const seconds = Number(match[4] || 0);
        return sign * (hours + minutes / 60 + seconds / 3600);
    }
    function buildTaskParentMap(tasks) {
        const parentMap = new Map();
        const stack = [];
        for (const task of tasks) {
            while (stack.length > 0 && task.outlineLevel <= stack[stack.length - 1].outlineLevel) {
                stack.pop();
            }
            parentMap.set(task.uid, stack.length > 0 ? stack[stack.length - 1].uid : null);
            if (task.summary) {
                stack.push(task);
            }
        }
        return parentMap;
    }
    function buildTaskPositionMap(tasks, parentMap) {
        const counters = new Map();
        const positionMap = new Map();
        for (const task of tasks) {
            const parentUid = parentMap.get(task.uid) || "__root__";
            const position = counters.get(parentUid) || 0;
            positionMap.set(task.uid, position);
            counters.set(parentUid, position + 1);
        }
        return positionMap;
    }
    function collectTopLevelPhases(tasks) {
        return tasks.filter((task) => !isPlaceholderUid(task.uid) && task.summary && task.outlineLevel === 1);
    }
    function collectPhaseTaskUids(tasks, phaseUid) {
        const phaseIndex = tasks.findIndex((task) => task.uid === phaseUid);
        if (phaseIndex < 0) {
            return new Set();
        }
        const phase = tasks[phaseIndex];
        const uids = new Set();
        for (let index = phaseIndex + 1; index < tasks.length; index += 1) {
            const task = tasks[index];
            if (task.outlineLevel <= phase.outlineLevel) {
                break;
            }
            if (!isPlaceholderUid(task.uid)) {
                uids.add(task.uid);
            }
        }
        return uids;
    }
    function collectTaskSubtreeUids(tasks, rootUid, maxDepth) {
        const rootIndex = tasks.findIndex((task) => task.uid === rootUid);
        if (rootIndex < 0) {
            return new Set();
        }
        const rootTask = tasks[rootIndex];
        const uids = new Set();
        if (!isPlaceholderUid(rootTask.uid)) {
            uids.add(rootTask.uid);
        }
        for (let index = rootIndex + 1; index < tasks.length; index += 1) {
            const task = tasks[index];
            if (task.outlineLevel <= rootTask.outlineLevel) {
                break;
            }
            const relativeDepth = task.outlineLevel - rootTask.outlineLevel;
            if (typeof maxDepth === "number" && relativeDepth > maxDepth) {
                continue;
            }
            if (!isPlaceholderUid(task.uid)) {
                uids.add(task.uid);
            }
        }
        return uids;
    }
    function resolvePhaseUidForTask(taskUid, parentMap, phaseUidSet) {
        let currentUid = taskUid;
        while (currentUid) {
            if (phaseUidSet.has(currentUid)) {
                return currentUid;
            }
            currentUid = parentMap.get(currentUid) || null;
        }
        return undefined;
    }
    function buildDefaultRules(scope) {
        if (scope === "project_overview_view") {
            return {
                allow_patch_ops: ["add_task", "update_task", "move_task"],
                forbid_completed_task_changes: true,
                forbid_summary_task_direct_edit: true
            };
        }
        return {
            allow_patch_ops: ["add_task", "update_task", "move_task", "link_tasks", "unlink_tasks"],
            forbid_completed_task_changes: true,
            forbid_summary_task_direct_edit: true
        };
    }
    function toIsoLocalString(value) {
        const year = value.getFullYear();
        const month = String(value.getMonth() + 1).padStart(2, "0");
        const day = String(value.getDate()).padStart(2, "0");
        const hour = String(value.getHours()).padStart(2, "0");
        const minute = String(value.getMinutes()).padStart(2, "0");
        const second = String(value.getSeconds()).padStart(2, "0");
        return `${year}-${month}-${day}T${hour}:${minute}:${second}`;
    }
    function addHoursToDateTime(dateTime, hours) {
        const parsed = new Date(dateTime);
        if (Number.isNaN(parsed.getTime())) {
            return dateTime;
        }
        parsed.setTime(parsed.getTime() + (hours * 60 * 60 * 1000));
        return toIsoLocalString(parsed);
    }
    function isDateOnlyText(value) {
        return typeof value === "string" && /^\d{4}-\d{2}-\d{2}$/.test(value.trim());
    }
    function withTimeOnDate(dateText, timeText) {
        return `${dateText}T${timeText}`;
    }
    function buildProjectDraftRequest(input) {
        return {
            view_type: "project_draft_request",
            project: {
                name: input.name,
                planned_start: input.plannedStart || undefined
            },
            requirements: {
                goal: input.goal || undefined,
                team_count: input.teamCount,
                must_have_phases: input.mustHavePhases || [],
                must_have_milestones: input.mustHaveMilestones || []
            }
        };
    }
    function importProjectDraftView(draft) {
        var _a, _b;
        if (!draft || typeof draft !== "object") {
            throw new Error("project_draft_view がオブジェクトではありません");
        }
        const data = draft;
        if (data.view_type !== "project_draft_view") {
            throw new Error("view_type が project_draft_view ではありません");
        }
        if (!((_b = (_a = data.project) === null || _a === void 0 ? void 0 : _a.name) === null || _b === void 0 ? void 0 : _b.trim())) {
            throw new Error("project.name がありません");
        }
        const inputTasks = Array.isArray(data.tasks) ? data.tasks : [];
        const seenUids = new Set();
        for (const task of inputTasks) {
            const uid = String(task.uid || "").trim();
            if (!uid) {
                throw new Error("draft task の uid がありません");
            }
            if (seenUids.has(uid)) {
                throw new Error(`draft task の uid が重複しています: ${uid}`);
            }
            seenUids.add(uid);
            if (!String(task.name || "").trim()) {
                throw new Error(`draft task の name がありません: ${uid}`);
            }
        }
        for (const task of inputTasks) {
            const parentUid = task.parent_uid == null || task.parent_uid === "" ? null : String(task.parent_uid);
            if (parentUid && !seenUids.has(parentUid)) {
                throw new Error(`draft task の parent_uid が既存 uid を指していません: ${String(task.uid || "")} -> ${parentUid}`);
            }
        }
        const projectStart = data.project.planned_start || data.project.planned_finish || toIsoLocalString(new Date());
        const draftUidMap = new Map();
        inputTasks.forEach((task, index) => {
            draftUidMap.set(String(task.uid || "").trim(), String(index + 1));
        });
        const normalizedTasks = inputTasks.map((task, index) => ({
            uid: draftUidMap.get(String(task.uid || "").trim()) || String(index + 1),
            name: String(task.name || "").trim(),
            parentUid: task.parent_uid == null || task.parent_uid === "" ? null : (draftUidMap.get(String(task.parent_uid)) || null),
            position: typeof task.position === "number" && Number.isFinite(task.position) ? task.position : index,
            isSummary: Boolean(task.is_summary),
            isMilestone: Boolean(task.is_milestone),
            percentComplete: typeof task.percent_complete === "number" && Number.isFinite(task.percent_complete)
                ? Math.max(0, Math.min(100, task.percent_complete))
                : 0,
            plannedDuration: task.planned_duration || undefined,
            plannedDurationHours: typeof task.planned_duration_hours === "number" && Number.isFinite(task.planned_duration_hours)
                ? task.planned_duration_hours
                : undefined,
            plannedStart: task.planned_start || undefined,
            plannedFinish: task.planned_finish || undefined,
            predecessorUids: [
                ...(Array.isArray(task.predecessor_uids) ? task.predecessor_uids : []),
                ...(Array.isArray(task.predecessors)
                    ? task.predecessors.flatMap((item) => {
                        if (typeof item === "string") {
                            return [item];
                        }
                        return (item === null || item === void 0 ? void 0 : item.task_uid) ? [item.task_uid] : [];
                    })
                    : [])
            ].map((item) => draftUidMap.get(String(item)) || String(item))
        }));
        const byParent = new Map();
        for (const task of normalizedTasks) {
            const siblings = byParent.get(task.parentUid) || [];
            siblings.push(task);
            byParent.set(task.parentUid, siblings);
        }
        for (const siblings of byParent.values()) {
            siblings.sort((left, right) => left.position - right.position || left.uid.localeCompare(right.uid));
        }
        const orderedTasks = [];
        function walk(parentUid, outlinePath) {
            const siblings = byParent.get(parentUid) || [];
            siblings.forEach((task, index) => {
                const currentPath = [...outlinePath, index + 1];
                const outlineNumber = currentPath.join(".");
                let start = task.plannedStart || task.plannedFinish || projectStart;
                let finish = task.plannedFinish
                    || (typeof task.plannedDurationHours === "number" ? addHoursToDateTime(start, task.plannedDurationHours) : start);
                const dateOnlyTaskRange = !task.isMilestone
                    && isDateOnlyText(start)
                    && isDateOnlyText(finish)
                    && task.plannedDuration == null
                    && typeof task.plannedDurationHours !== "number";
                if (dateOnlyTaskRange) {
                    start = withTimeOnDate(start, "09:00:00");
                    finish = withTimeOnDate(finish, "18:00:00");
                }
                const hasChildren = (byParent.get(task.uid) || []).length > 0;
                orderedTasks.push({
                    uid: task.uid,
                    id: task.uid,
                    name: task.name,
                    outlineLevel: currentPath.length,
                    outlineNumber,
                    wbs: outlineNumber,
                    start,
                    finish,
                    duration: task.plannedDuration || (typeof task.plannedDurationHours === "number" ? `PT${task.plannedDurationHours}H` : "PT0H0M0S"),
                    milestone: task.isMilestone,
                    summary: task.isSummary || hasChildren,
                    percentComplete: task.percentComplete,
                    predecessors: task.predecessorUids.map((predecessorUid) => ({ predecessorUid })),
                    extendedAttributes: [],
                    baselines: [],
                    timephasedData: []
                });
                walk(task.uid, currentPath);
            });
        }
        walk(null, []);
        const taskFinishes = orderedTasks.map((task) => task.finish).filter(Boolean).sort();
        return ensureDefaultProjectCalendar({
            project: {
                name: data.project.name.trim(),
                startDate: projectStart,
                finishDate: data.project.planned_finish || taskFinishes.at(-1) || projectStart,
                scheduleFromStart: true,
                outlineCodes: [],
                wbsMasks: [],
                extendedAttributes: []
            },
            tasks: orderedTasks,
            resources: [],
            assignments: [],
            calendars: []
        });
    }
    function exportProjectOverviewView(model) {
        const parentMap = buildTaskParentMap(model.tasks);
        const phaseTasks = collectTopLevelPhases(model.tasks);
        const phaseUidSet = new Set(phaseTasks.map((task) => task.uid));
        const allMilestones = model.tasks.filter((task) => !isPlaceholderUid(task.uid) && task.milestone);
        const topLevelDependencyMap = new Map();
        for (const task of model.tasks) {
            const toPhaseUid = resolvePhaseUidForTask(task.uid, parentMap, phaseUidSet);
            if (!toPhaseUid) {
                continue;
            }
            for (const predecessor of task.predecessors) {
                const fromPhaseUid = resolvePhaseUidForTask(predecessor.predecessorUid, parentMap, phaseUidSet);
                if (!fromPhaseUid || fromPhaseUid === toPhaseUid) {
                    continue;
                }
                const key = `${fromPhaseUid}->${toPhaseUid}:${formatDependencyType(predecessor.type)}`;
                if (!topLevelDependencyMap.has(key)) {
                    topLevelDependencyMap.set(key, {
                        from_uid: fromPhaseUid,
                        to_uid: toPhaseUid,
                        type: formatDependencyType(predecessor.type)
                    });
                }
            }
        }
        return {
            view_type: "project_overview_view",
            project: {
                name: model.project.name,
                planned_start: model.project.startDate,
                planned_finish: model.project.finishDate,
                status_date: model.project.statusDate
            },
            summary: {
                task_count: model.tasks.filter((task) => !isPlaceholderUid(task.uid)).length,
                summary_task_count: model.tasks.filter((task) => !isPlaceholderUid(task.uid) && task.summary).length,
                milestone_count: allMilestones.length,
                max_outline_level: model.tasks.reduce((max, task) => Math.max(max, task.outlineLevel || 0), 0)
            },
            phases: phaseTasks.map((phase) => {
                const phaseTaskUids = collectPhaseTaskUids(model.tasks, phase.uid);
                const descendantTasks = model.tasks.filter((task) => phaseTaskUids.has(task.uid));
                return {
                    uid: phase.uid,
                    name: phase.name,
                    wbs: phase.wbs || phase.outlineNumber,
                    task_count: descendantTasks.length,
                    milestone_count: descendantTasks.filter((task) => task.milestone).length,
                    planned_start: phase.start,
                    planned_finish: phase.finish,
                    duration: phase.duration,
                    duration_hours: parseDurationHours(phase.duration),
                    percent_complete: phase.percentComplete,
                    sample_tasks: descendantTasks.slice(0, 3).map((task) => ({
                        uid: task.uid,
                        name: task.name
                    }))
                };
            }),
            milestones: allMilestones.map((task) => ({
                uid: task.uid,
                name: task.name,
                parent_uid: parentMap.get(task.uid),
                date: task.finish || task.start
            })),
            top_level_dependencies: Array.from(topLevelDependencyMap.values()),
            rules: buildDefaultRules("project_overview_view")
        };
    }
    function exportPhaseDetailView(model, requestedPhaseUid, options) {
        var _a;
        const phaseTasks = collectTopLevelPhases(model.tasks);
        if (phaseTasks.length === 0) {
            throw new Error("phase が見つかりません");
        }
        const phase = requestedPhaseUid
            ? phaseTasks.find((task) => task.uid === requestedPhaseUid)
            : phaseTasks[0];
        if (!phase) {
            throw new Error(`phase が見つかりません: ${requestedPhaseUid}`);
        }
        const parentMap = buildTaskParentMap(model.tasks);
        const positionMap = buildTaskPositionMap(model.tasks, parentMap);
        const phaseTaskUids = collectPhaseTaskUids(model.tasks, phase.uid);
        const phaseTasksOnly = model.tasks.filter((task) => phaseTaskUids.has(task.uid));
        const mode = (options === null || options === void 0 ? void 0 : options.mode) === "scoped" ? "scoped" : "full";
        const rootUid = mode === "scoped" ? ((_a = options === null || options === void 0 ? void 0 : options.rootUid) === null || _a === void 0 ? void 0 : _a.trim()) || undefined : undefined;
        const maxDepth = mode === "scoped" && typeof (options === null || options === void 0 ? void 0 : options.maxDepth) === "number" && Number.isFinite(options.maxDepth) && options.maxDepth >= 0
            ? Math.floor(options.maxDepth)
            : undefined;
        let scopedTaskUids = phaseTaskUids;
        if (rootUid) {
            const rootTask = phaseTasksOnly.find((task) => task.uid === rootUid);
            if (!rootTask) {
                throw new Error(`phase 配下に root_uid が見つかりません: ${rootUid}`);
            }
            scopedTaskUids = collectTaskSubtreeUids(phaseTasksOnly, rootUid, maxDepth);
        }
        const descendantTasks = phaseTasksOnly.filter((task) => scopedTaskUids.has(task.uid));
        return {
            view_type: "phase_detail_view",
            project: {
                name: model.project.name,
                planned_start: model.project.startDate,
                planned_finish: model.project.finishDate
            },
            phase: {
                uid: phase.uid,
                name: phase.name,
                wbs: phase.wbs || phase.outlineNumber,
                planned_start: phase.start,
                planned_finish: phase.finish,
                task_count: descendantTasks.length,
                milestone_count: descendantTasks.filter((task) => task.milestone).length,
                percent_complete: phase.percentComplete
            },
            scope: {
                mode,
                root_uid: rootUid || null,
                max_depth: maxDepth !== null && maxDepth !== void 0 ? maxDepth : null
            },
            tasks: descendantTasks.map((task) => {
                var _a;
                return ({
                    uid: task.uid,
                    name: task.name,
                    parent_uid: parentMap.get(task.uid),
                    position: (_a = positionMap.get(task.uid)) !== null && _a !== void 0 ? _a : 0,
                    is_summary: task.summary,
                    is_milestone: task.milestone,
                    planned_duration: task.duration,
                    planned_duration_hours: parseDurationHours(task.duration),
                    planned_start: task.start,
                    planned_finish: task.finish,
                    percent_complete: task.percentComplete,
                    predecessor_uids: task.predecessors.map((item) => item.predecessorUid)
                });
            }),
            milestones: descendantTasks.filter((task) => task.milestone).map((task) => ({
                uid: task.uid,
                name: task.name,
                date: task.finish || task.start
            })),
            dependency_summary: descendantTasks.flatMap((task) => task.predecessors
                .filter((predecessor) => scopedTaskUids.has(predecessor.predecessorUid))
                .map((predecessor) => {
                var _a;
                return ({
                    from_uid: predecessor.predecessorUid,
                    to_uid: task.uid,
                    type: formatDependencyType(predecessor.type),
                    lag: predecessor.linkLag || "PT0H0M0S",
                    lag_hours: (_a = parseDurationHours(predecessor.linkLag || "PT0H0M0S")) !== null && _a !== void 0 ? _a : 0
                });
            })),
            rules: buildDefaultRules("phase_detail_view")
        };
    }
    function buildTaskSectionMap(tasks, projectName) {
        const sectionMap = new Map();
        const summaryStack = [];
        for (const task of tasks) {
            while (summaryStack.length > 0 && task.outlineLevel <= summaryStack[summaryStack.length - 1].outlineLevel) {
                summaryStack.pop();
            }
            if (task.summary) {
                summaryStack.push(task);
                continue;
            }
            const sectionName = summaryStack.length > 0
                ? normalizeMermaidGanttLabel(summaryStack[summaryStack.length - 1].name, "Summary", "Section")
                : normalizeMermaidGanttLabel(projectName, "Tasks", "Section");
            sectionMap.set(task.uid, sectionName);
        }
        return sectionMap;
    }
    function exportMermaidGantt(model) {
        const lines = [
            "gantt",
            `  title ${normalizeMermaidGanttLabel(model.project.name, "Project", "Project")}`,
            "  dateFormat YYYY-MM-DDTHH:mm:ss",
            "  axisFormat %m/%d"
        ];
        const sectionMap = buildTaskSectionMap(model.tasks, model.project.name);
        const taskNameMap = new Map(model.tasks.map((task) => [
            task.uid,
            normalizeMermaidGanttLabel(task.name, `Task ${task.uid}`, "Task")
        ]));
        const exportedTasks = model.tasks.filter((task) => !task.summary && task.start && task.finish);
        let currentSection = "";
        for (const task of exportedTasks) {
            const section = sectionMap.get(task.uid) || "Tasks";
            if (section !== currentSection) {
                currentSection = section;
                lines.push(`  section ${section}`);
            }
            const tags = [];
            if (task.critical) {
                tags.push("crit");
            }
            if (task.milestone) {
                tags.push("milestone");
            }
            else if (task.percentComplete >= 100) {
                tags.push("done");
            }
            else if (task.percentComplete > 0) {
                tags.push("active");
            }
            const taskId = `task_${normalizeMermaidTaskId(task.uid || task.id || "x", "x")}`;
            const singlePredecessor = task.predecessors.length === 1 ? task.predecessors[0] : null;
            const nativeDependencyTarget = singlePredecessor
                ? `task_${normalizeMermaidTaskId(singlePredecessor.predecessorUid, "x")}`
                : null;
            const nativeDuration = !task.milestone ? toMermaidDuration(task.duration) : null;
            const useNativeDependency = Boolean(singlePredecessor
                && nativeDependencyTarget
                && nativeDuration
                && isZeroDuration(singlePredecessor.linkLag)
                && (singlePredecessor.type === undefined || singlePredecessor.type === 1));
            const fields = useNativeDependency
                ? [...tags, taskId, `after ${nativeDependencyTarget}`, nativeDuration]
                : [...tags, taskId, task.start, task.finish].filter(Boolean);
            lines.push(`  ${normalizeMermaidGanttLabel(task.name, `Task ${task.uid}`, "Task")} :${fields.join(", ")}`);
            for (const predecessor of task.predecessors) {
                const predecessorTaskId = `task_${normalizeMermaidTaskId(predecessor.predecessorUid, "x")}`;
                const predecessorName = taskNameMap.get(predecessor.predecessorUid) || `Task ${predecessor.predecessorUid}`;
                if (useNativeDependency && predecessorTaskId === nativeDependencyTarget) {
                    lines.push(`  %% dependency(native): ${task.name || taskId} after ${predecessorName} (${taskId} after ${predecessorTaskId})`);
                    continue;
                }
                const details = [
                    `type=${describePredecessorType(predecessor.type)}`,
                    !isZeroDuration(predecessor.linkLag) ? `lag=${formatMermaidLag(predecessor.linkLag)}` : ""
                ].filter(Boolean).join(", ");
                lines.push(`  %% dependency: ${task.name || taskId} after ${predecessorName}${details ? ` (${details})` : ""} [${taskId} after ${predecessorTaskId}]`);
                if (!isZeroDuration(predecessor.linkLag)) {
                    lines.push(`  %% dependency(pseudo): ${task.name || taskId} ~= after ${predecessorName} + ${formatMermaidLag(predecessor.linkLag)}`);
                }
            }
            if (task.predecessors.length > 1) {
                lines.push(`  %% dependency(note): ${task.name || taskId} has multiple predecessors`);
            }
            else if (singlePredecessor && !useNativeDependency) {
                const reasons = [
                    !isZeroDuration(singlePredecessor.linkLag) ? `lag=${formatMermaidLag(singlePredecessor.linkLag)}` : "",
                    singlePredecessor.type !== undefined && singlePredecessor.type !== 1 ? `type=${describePredecessorType(singlePredecessor.type)}` : "",
                    !nativeDuration && !task.milestone ? `duration=${task.duration || "(empty)"}` : ""
                ].filter(Boolean).join(", ");
                if (reasons) {
                    lines.push(`  %% dependency(note): ${task.name || taskId} kept as comment because ${reasons}`);
                }
            }
        }
        if (exportedTasks.length === 0) {
            lines.push("  section Tasks");
            lines.push("  No tasks :milestone, empty_0, 1970-01-01T00:00:00, 1970-01-01T00:00:00");
        }
        return `${lines.join("\n")}\n`;
    }
    function buildTaskParentUidMap(tasks) {
        const parentMap = new Map();
        const stack = [];
        for (const task of tasks) {
            while (stack.length > 0 && task.outlineLevel <= stack[stack.length - 1].outlineLevel) {
                stack.pop();
            }
            const parent = stack[stack.length - 1];
            if (parent) {
                parentMap.set(task.uid, parent.uid);
            }
            stack.push(task);
        }
        return parentMap;
    }
    function escapeCsvCell(value) {
        const text = String(value !== null && value !== void 0 ? value : "");
        if (/[",\n]/.test(text)) {
            return `"${text.replace(/"/g, "\"\"")}"`;
        }
        return text;
    }
    function exportCsvParentId(model) {
        const header = ["ID", "ParentID", "WBS", "Name", "Start", "Finish", "PredecessorID", "Resource", "PercentComplete", "PercentWorkComplete", "Milestone", "Summary", "Critical", "Type", "Priority", "Work", "CalendarUID", "ConstraintType", "ConstraintDate", "Deadline", "Notes"];
        const parentMap = buildTaskParentUidMap(model.tasks);
        const resourceMap = new Map(model.resources.map((resource) => [resource.uid, resource.name]));
        const assignmentMap = new Map();
        for (const assignment of model.assignments) {
            const resourceName = resourceMap.get(assignment.resourceUid);
            if (!resourceName) {
                continue;
            }
            const names = assignmentMap.get(assignment.taskUid) || [];
            if (!names.includes(resourceName)) {
                names.push(resourceName);
            }
            assignmentMap.set(assignment.taskUid, names);
        }
        const rows = model.tasks.map((task) => {
            var _a, _b, _c, _d;
            return [
                task.uid,
                parentMap.get(task.uid) || "",
                task.wbs || task.outlineNumber || "",
                task.name,
                task.start || "",
                task.finish || "",
                task.predecessors.map((item) => item.predecessorUid).join("|"),
                (assignmentMap.get(task.uid) || []).join("|"),
                task.percentComplete,
                (_a = task.percentWorkComplete) !== null && _a !== void 0 ? _a : "",
                task.milestone ? 1 : 0,
                task.summary ? 1 : 0,
                task.critical === undefined ? "" : (task.critical ? 1 : 0),
                (_b = task.type) !== null && _b !== void 0 ? _b : "",
                (_c = task.priority) !== null && _c !== void 0 ? _c : "",
                task.work || "",
                task.calendarUID || "",
                (_d = task.constraintType) !== null && _d !== void 0 ? _d : "",
                task.constraintDate || "",
                task.deadline || "",
                task.notes || ""
            ];
        });
        return [header, ...rows].map((row) => row.map((cell) => escapeCsvCell(cell)).join(",")).join("\n") + "\n";
    }
    function parseCsvRows(csvText) {
        const rows = [];
        let row = [];
        let cell = "";
        let inQuotes = false;
        for (let index = 0; index < csvText.length; index += 1) {
            const char = csvText[index];
            const next = csvText[index + 1];
            if (char === "\"") {
                if (inQuotes && next === "\"") {
                    cell += "\"";
                    index += 1;
                }
                else {
                    inQuotes = !inQuotes;
                }
                continue;
            }
            if (!inQuotes && char === ",") {
                row.push(cell);
                cell = "";
                continue;
            }
            if (!inQuotes && (char === "\n" || char === "\r")) {
                if (char === "\r" && next === "\n") {
                    index += 1;
                }
                row.push(cell);
                rows.push(row);
                row = [];
                cell = "";
                continue;
            }
            cell += char;
        }
        if (cell.length > 0 || row.length > 0) {
            row.push(cell);
            rows.push(row);
        }
        return rows.filter((item) => item.some((cellValue) => String(cellValue).trim() !== ""));
    }
    function parseCsvMultiValueCell(value) {
        const normalized = String(value || "").trim();
        if (!normalized) {
            return [];
        }
        const items = normalized
            .split(/[|;,、]/)
            .map((item) => item.trim())
            .filter(Boolean);
        return Array.from(new Set(items));
    }
    function parseCsvBooleanCell(value, fallback) {
        const normalized = String(value || "").trim().toLowerCase();
        if (!normalized) {
            return fallback;
        }
        if (["1", "true", "yes", "y", "on"].includes(normalized)) {
            return true;
        }
        if (["0", "false", "no", "n", "off"].includes(normalized)) {
            return false;
        }
        return fallback;
    }
    function importCsvParentId(csvText) {
        const rows = parseCsvRows(csvText.trim());
        if (rows.length === 0) {
            throw new Error("CSV が空です");
        }
        const header = rows[0].map((item) => item.trim());
        const requiredColumns = ["ID", "ParentID", "Name"];
        for (const requiredColumn of requiredColumns) {
            if (!header.includes(requiredColumn)) {
                throw new Error(`CSV に必須列がありません: ${requiredColumn}`);
            }
        }
        const columnIndex = (name) => header.indexOf(name);
        const entries = rows.slice(1).map((row) => ({
            id: String(row[columnIndex("ID")] || "").trim(),
            parentId: String(row[columnIndex("ParentID")] || "").trim(),
            wbs: String((columnIndex("WBS") >= 0 ? row[columnIndex("WBS")] : "") || "").trim(),
            name: String(row[columnIndex("Name")] || "").trim(),
            start: String((columnIndex("Start") >= 0 ? row[columnIndex("Start")] : "") || "").trim(),
            finish: String((columnIndex("Finish") >= 0 ? row[columnIndex("Finish")] : "") || "").trim(),
            predecessorId: String((columnIndex("PredecessorID") >= 0 ? row[columnIndex("PredecessorID")] : "") || "").trim(),
            resource: String((columnIndex("Resource") >= 0 ? row[columnIndex("Resource")] : "") || "").trim(),
            percentComplete: parseNumber(String((columnIndex("PercentComplete") >= 0 ? row[columnIndex("PercentComplete")] : "0") || "0").trim(), 0),
            percentWorkComplete: columnIndex("PercentWorkComplete") >= 0 && String(row[columnIndex("PercentWorkComplete")] || "").trim()
                ? parseNumber(String(row[columnIndex("PercentWorkComplete")] || "").trim(), 0)
                : undefined,
            milestone: parseCsvBooleanCell(String((columnIndex("Milestone") >= 0 ? row[columnIndex("Milestone")] : "") || "").trim(), false),
            summary: columnIndex("Summary") >= 0 && String(row[columnIndex("Summary")] || "").trim()
                ? parseCsvBooleanCell(String(row[columnIndex("Summary")] || "").trim(), false)
                : undefined,
            critical: columnIndex("Critical") >= 0 && String(row[columnIndex("Critical")] || "").trim()
                ? parseCsvBooleanCell(String(row[columnIndex("Critical")] || "").trim(), false)
                : undefined,
            type: columnIndex("Type") >= 0 && String(row[columnIndex("Type")] || "").trim()
                ? parseNumber(String(row[columnIndex("Type")] || "").trim(), 0)
                : undefined,
            priority: columnIndex("Priority") >= 0 && String(row[columnIndex("Priority")] || "").trim()
                ? parseNumber(String(row[columnIndex("Priority")] || "").trim(), 0)
                : undefined,
            work: String((columnIndex("Work") >= 0 ? row[columnIndex("Work")] : "") || "").trim(),
            calendarUID: String((columnIndex("CalendarUID") >= 0 ? row[columnIndex("CalendarUID")] : "") || "").trim(),
            constraintType: columnIndex("ConstraintType") >= 0 && String(row[columnIndex("ConstraintType")] || "").trim()
                ? parseNumber(String(row[columnIndex("ConstraintType")] || "").trim(), 0)
                : undefined,
            constraintDate: String((columnIndex("ConstraintDate") >= 0 ? row[columnIndex("ConstraintDate")] : "") || "").trim(),
            deadline: String((columnIndex("Deadline") >= 0 ? row[columnIndex("Deadline")] : "") || "").trim(),
            notes: String((columnIndex("Notes") >= 0 ? row[columnIndex("Notes")] : "") || "").trim(),
            children: []
        })).filter((entry) => entry.id);
        const seenIds = new Set();
        for (const entry of entries) {
            if (seenIds.has(entry.id)) {
                throw new Error(`CSV の ID が重複しています: ${entry.id}`);
            }
            seenIds.add(entry.id);
            if (!entry.name) {
                throw new Error(`CSV の Name が空です: ID=${entry.id}`);
            }
            if (entry.parentId && entry.parentId === entry.id) {
                throw new Error(`CSV の ParentID が自身を指しています: ID=${entry.id}`);
            }
        }
        const entryMap = new Map(entries.map((entry) => [entry.id, entry]));
        for (const entry of entries) {
            if (entry.parentId && !entryMap.has(entry.parentId)) {
                throw new Error(`CSV の ParentID が既存 ID を指していません: ID=${entry.id}, ParentID=${entry.parentId}`);
            }
        }
        const visiting = new Set();
        const visited = new Set();
        function checkCycle(entry) {
            if (visited.has(entry.id)) {
                return;
            }
            if (visiting.has(entry.id)) {
                throw new Error(`CSV の ParentID が循環しています: ID=${entry.id}`);
            }
            visiting.add(entry.id);
            if (entry.parentId) {
                const parent = entryMap.get(entry.parentId);
                if (parent) {
                    checkCycle(parent);
                }
            }
            visiting.delete(entry.id);
            visited.add(entry.id);
        }
        entries.forEach((entry) => checkCycle(entry));
        const roots = [];
        for (const entry of entries) {
            const parent = entry.parentId ? entryMap.get(entry.parentId) : undefined;
            if (parent) {
                parent.children.push(entry);
            }
            else {
                roots.push(entry);
            }
        }
        const tasks = [];
        function walk(entry, outlinePath) {
            var _a;
            const children = entry.children;
            let start = entry.start;
            let finish = entry.finish;
            if ((!start || !finish) && children.length > 0) {
                const childStarts = children.map((child) => child.start).filter(Boolean).sort();
                const childFinishes = children.map((child) => child.finish).filter(Boolean).sort();
                start = start || childStarts[0] || "";
                finish = finish || childFinishes.at(-1) || "";
            }
            const outlineNumber = outlinePath.join(".");
            tasks.push({
                uid: entry.id,
                id: entry.id,
                name: entry.name,
                outlineLevel: outlinePath.length,
                outlineNumber,
                wbs: entry.wbs || outlineNumber,
                type: entry.type,
                priority: entry.priority,
                work: entry.work || undefined,
                calendarUID: entry.calendarUID || undefined,
                start,
                finish,
                duration: "PT0H0M0S",
                milestone: entry.milestone || Boolean(start && finish && start === finish),
                summary: (_a = entry.summary) !== null && _a !== void 0 ? _a : (children.length > 0),
                critical: entry.critical,
                percentComplete: Math.max(0, Math.min(100, entry.percentComplete)),
                percentWorkComplete: entry.percentWorkComplete !== undefined ? Math.max(0, Math.min(100, entry.percentWorkComplete)) : undefined,
                constraintType: entry.constraintType,
                constraintDate: entry.constraintDate || undefined,
                deadline: entry.deadline || undefined,
                notes: entry.notes || undefined,
                predecessors: parseCsvMultiValueCell(entry.predecessorId).map((item) => ({ predecessorUid: item })),
                extendedAttributes: [],
                baselines: [],
                timephasedData: []
            });
            children.forEach((child, index) => walk(child, [...outlinePath, index + 1]));
        }
        roots.forEach((root, index) => walk(root, [index + 1]));
        const resourceNames = Array.from(new Set(entries.flatMap((entry) => parseCsvMultiValueCell(entry.resource))));
        const resources = resourceNames.map((name, index) => ({
            uid: String(index + 1),
            id: String(index + 1),
            name,
            extendedAttributes: [],
            baselines: [],
            timephasedData: []
        }));
        const resourceUidByName = new Map(resources.map((resource) => [resource.name, resource.uid]));
        let assignmentUid = 1;
        const taskByUid = new Map(tasks.map((task) => [task.uid, task]));
        const assignments = entries.flatMap((entry) => {
            const task = taskByUid.get(entry.id);
            if (!task) {
                return [];
            }
            return parseCsvMultiValueCell(entry.resource).map((name) => ({
                uid: String(assignmentUid++),
                taskUid: entry.id,
                resourceUid: resourceUidByName.get(name) || "",
                start: task.start || undefined,
                finish: task.finish || undefined,
                percentWorkComplete: task.percentComplete,
                extendedAttributes: [],
                baselines: [],
                timephasedData: []
            }));
        });
        const taskStarts = tasks.map((task) => task.start).filter(Boolean).sort();
        const taskFinishes = tasks.map((task) => task.finish).filter(Boolean).sort();
        return ensureDefaultProjectCalendar({
            project: {
                name: "CSV Imported Project",
                startDate: taskStarts[0] || "",
                finishDate: taskFinishes.at(-1) || "",
                scheduleFromStart: true,
                outlineCodes: [],
                wbsMasks: [],
                extendedAttributes: []
            },
            tasks,
            resources,
            assignments,
            calendars: []
        });
    }
    function importMsProjectXml(xmlText) {
        var _a, _b, _c, _d, _e, _f, _g;
        const xml = parseXmlDocument(xmlText);
        const projectElement = xml.documentElement;
        const calendars = Array.from(((_a = projectElement.getElementsByTagName("Calendars")[0]) === null || _a === void 0 ? void 0 : _a.getElementsByTagName("Calendar")) || []);
        const tasks = Array.from(((_b = projectElement.getElementsByTagName("Tasks")[0]) === null || _b === void 0 ? void 0 : _b.getElementsByTagName("Task")) || []);
        const resources = Array.from(((_c = projectElement.getElementsByTagName("Resources")[0]) === null || _c === void 0 ? void 0 : _c.getElementsByTagName("Resource")) || []);
        const assignments = Array.from(((_d = projectElement.getElementsByTagName("Assignments")[0]) === null || _d === void 0 ? void 0 : _d.getElementsByTagName("Assignment")) || []);
        return ensureDefaultProjectCalendar({
            project: {
                name: textContent(projectElement, "Name"),
                title: textContent(projectElement, "Title") || undefined,
                author: textContent(projectElement, "Author") || undefined,
                company: textContent(projectElement, "Company") || undefined,
                creationDate: textContent(projectElement, "CreationDate") || undefined,
                lastSaved: textContent(projectElement, "LastSaved") || undefined,
                saveVersion: textContent(projectElement, "SaveVersion") ? parseNumber(textContent(projectElement, "SaveVersion"), 0) : undefined,
                currentDate: textContent(projectElement, "CurrentDate") || undefined,
                startDate: textContent(projectElement, "StartDate"),
                finishDate: textContent(projectElement, "FinishDate"),
                scheduleFromStart: parseBoolean(textContent(projectElement, "ScheduleFromStart")),
                defaultStartTime: textContent(projectElement, "DefaultStartTime") || undefined,
                defaultFinishTime: textContent(projectElement, "DefaultFinishTime") || undefined,
                minutesPerDay: textContent(projectElement, "MinutesPerDay") ? parseNumber(textContent(projectElement, "MinutesPerDay"), 0) : undefined,
                minutesPerWeek: textContent(projectElement, "MinutesPerWeek") ? parseNumber(textContent(projectElement, "MinutesPerWeek"), 0) : undefined,
                daysPerMonth: textContent(projectElement, "DaysPerMonth") ? parseNumber(textContent(projectElement, "DaysPerMonth"), 0) : undefined,
                statusDate: textContent(projectElement, "StatusDate") || undefined,
                weekStartDay: textContent(projectElement, "WeekStartDay") ? parseNumber(textContent(projectElement, "WeekStartDay"), 0) : undefined,
                workFormat: textContent(projectElement, "WorkFormat") ? parseNumber(textContent(projectElement, "WorkFormat"), 0) : undefined,
                durationFormat: textContent(projectElement, "DurationFormat") ? parseNumber(textContent(projectElement, "DurationFormat"), 0) : undefined,
                currencyCode: textContent(projectElement, "CurrencyCode") || undefined,
                currencyDigits: textContent(projectElement, "CurrencyDigits") ? parseNumber(textContent(projectElement, "CurrencyDigits"), 0) : undefined,
                currencySymbol: textContent(projectElement, "CurrencySymbol") || undefined,
                currencySymbolPosition: textContent(projectElement, "CurrencySymbolPosition") ? parseNumber(textContent(projectElement, "CurrencySymbolPosition"), 0) : undefined,
                fyStartDate: textContent(projectElement, "FYStartDate") || undefined,
                fiscalYearStart: textContent(projectElement, "FiscalYearStart") ? parseBoolean(textContent(projectElement, "FiscalYearStart")) : undefined,
                criticalSlackLimit: textContent(projectElement, "CriticalSlackLimit") ? parseNumber(textContent(projectElement, "CriticalSlackLimit"), 0) : undefined,
                defaultTaskType: textContent(projectElement, "DefaultTaskType") ? parseNumber(textContent(projectElement, "DefaultTaskType"), 0) : undefined,
                defaultFixedCostAccrual: textContent(projectElement, "DefaultFixedCostAccrual") ? parseNumber(textContent(projectElement, "DefaultFixedCostAccrual"), 0) : undefined,
                defaultStandardRate: textContent(projectElement, "DefaultStandardRate") || undefined,
                defaultOvertimeRate: textContent(projectElement, "DefaultOvertimeRate") || undefined,
                defaultTaskEVMethod: textContent(projectElement, "DefaultTaskEVMethod") ? parseNumber(textContent(projectElement, "DefaultTaskEVMethod"), 0) : undefined,
                newTaskStartDate: textContent(projectElement, "NewTaskStartDate") ? parseNumber(textContent(projectElement, "NewTaskStartDate"), 0) : undefined,
                newTasksAreManual: textContent(projectElement, "NewTasksAreManual") ? parseBoolean(textContent(projectElement, "NewTasksAreManual")) : undefined,
                newTasksEffortDriven: textContent(projectElement, "NewTasksEffortDriven") ? parseBoolean(textContent(projectElement, "NewTasksEffortDriven")) : undefined,
                newTasksEstimated: textContent(projectElement, "NewTasksEstimated") ? parseBoolean(textContent(projectElement, "NewTasksEstimated")) : undefined,
                actualsInSync: textContent(projectElement, "ActualsInSync") ? parseBoolean(textContent(projectElement, "ActualsInSync")) : undefined,
                editableActualCosts: textContent(projectElement, "EditableActualCosts") ? parseBoolean(textContent(projectElement, "EditableActualCosts")) : undefined,
                honorConstraints: textContent(projectElement, "HonorConstraints") ? parseBoolean(textContent(projectElement, "HonorConstraints")) : undefined,
                insertedProjectsLikeSummary: textContent(projectElement, "InsertedProjectsLikeSummary") ? parseBoolean(textContent(projectElement, "InsertedProjectsLikeSummary")) : undefined,
                multipleCriticalPaths: textContent(projectElement, "MultipleCriticalPaths") ? parseBoolean(textContent(projectElement, "MultipleCriticalPaths")) : undefined,
                taskUpdatesResource: textContent(projectElement, "TaskUpdatesResource") ? parseBoolean(textContent(projectElement, "TaskUpdatesResource")) : undefined,
                updateManuallyScheduledTasksWhenEditingLinks: textContent(projectElement, "UpdateManuallyScheduledTasksWhenEditingLinks") ? parseBoolean(textContent(projectElement, "UpdateManuallyScheduledTasksWhenEditingLinks")) : undefined,
                calendarUID: textContent(projectElement, "CalendarUID") || undefined,
                outlineCodes: Array.from(((_e = projectElement.getElementsByTagName("OutlineCodes")[0]) === null || _e === void 0 ? void 0 : _e.getElementsByTagName("OutlineCode")) || []).map((outlineCode) => ({
                    fieldID: textContent(outlineCode, "FieldID") || undefined,
                    fieldName: textContent(outlineCode, "FieldName") || undefined,
                    alias: textContent(outlineCode, "Alias") || undefined,
                    onlyTableValues: textContent(outlineCode, "OnlyTableValues") ? parseBoolean(textContent(outlineCode, "OnlyTableValues")) : undefined,
                    enterprise: textContent(outlineCode, "Enterprise") ? parseBoolean(textContent(outlineCode, "Enterprise")) : undefined,
                    resourceSubstitutionEnabled: textContent(outlineCode, "ResourceSubstitutionEnabled") ? parseBoolean(textContent(outlineCode, "ResourceSubstitutionEnabled")) : undefined,
                    leafOnly: textContent(outlineCode, "LeafOnly") ? parseBoolean(textContent(outlineCode, "LeafOnly")) : undefined,
                    allLevelsRequired: textContent(outlineCode, "AllLevelsRequired") ? parseBoolean(textContent(outlineCode, "AllLevelsRequired")) : undefined,
                    masks: parseOutlineCodeMasks(outlineCode),
                    values: parseOutlineCodeValues(outlineCode)
                })),
                wbsMasks: Array.from(((_f = projectElement.getElementsByTagName("WBSMasks")[0]) === null || _f === void 0 ? void 0 : _f.getElementsByTagName("WBSMask")) || []).map((wbsMask) => ({
                    level: parseNumber(textContent(wbsMask, "Level"), 0),
                    mask: textContent(wbsMask, "Mask") || undefined,
                    length: textContent(wbsMask, "Length") ? parseNumber(textContent(wbsMask, "Length"), 0) : undefined,
                    sequence: textContent(wbsMask, "Sequence") ? parseNumber(textContent(wbsMask, "Sequence"), 0) : undefined
                })),
                extendedAttributes: Array.from(((_g = projectElement.getElementsByTagName("ExtendedAttributes")[0]) === null || _g === void 0 ? void 0 : _g.getElementsByTagName("ExtendedAttribute")) || []).map((attribute) => ({
                    fieldID: textContent(attribute, "FieldID") || undefined,
                    fieldName: textContent(attribute, "FieldName") || undefined,
                    alias: textContent(attribute, "Alias") || undefined,
                    calculationType: textContent(attribute, "CalculationType") ? parseNumber(textContent(attribute, "CalculationType"), 0) : undefined,
                    restrictValues: textContent(attribute, "RestrictValues") ? parseBoolean(textContent(attribute, "RestrictValues")) : undefined,
                    appendNewValues: textContent(attribute, "AppendNewValues") ? parseBoolean(textContent(attribute, "AppendNewValues")) : undefined
                }))
            },
            calendars: calendars.map((calendar) => {
                var _a, _b;
                return ({
                    uid: textContent(calendar, "UID"),
                    name: textContent(calendar, "Name"),
                    isBaseCalendar: parseBoolean(textContent(calendar, "IsBaseCalendar")),
                    isBaselineCalendar: textContent(calendar, "IsBaselineCalendar") ? parseBoolean(textContent(calendar, "IsBaselineCalendar")) : undefined,
                    baseCalendarUID: textContent(calendar, "BaseCalendarUID") || undefined,
                    weekDays: parseWeekDays(calendar),
                    exceptions: Array.from(((_a = calendar.getElementsByTagName("Exceptions")[0]) === null || _a === void 0 ? void 0 : _a.getElementsByTagName("Exception")) || []).map((exception) => ({
                        name: textContent(exception, "Name") || undefined,
                        fromDate: textContent(exception, "FromDate") || undefined,
                        toDate: textContent(exception, "ToDate") || undefined,
                        dayWorking: textContent(exception, "DayWorking") ? parseBoolean(textContent(exception, "DayWorking")) : undefined,
                        workingTimes: parseWorkingTimes(exception)
                    })),
                    workWeeks: Array.from(((_b = calendar.getElementsByTagName("WorkWeeks")[0]) === null || _b === void 0 ? void 0 : _b.getElementsByTagName("WorkWeek")) || []).map((workWeek) => ({
                        name: textContent(workWeek, "Name") || undefined,
                        fromDate: textContent(workWeek, "FromDate") || undefined,
                        toDate: textContent(workWeek, "ToDate") || undefined,
                        weekDays: parseWeekDays(workWeek)
                    }))
                });
            }),
            tasks: tasks.map((task) => ({
                uid: textContent(task, "UID"),
                id: textContent(task, "ID"),
                name: textContent(task, "Name"),
                outlineLevel: parseNumber(textContent(task, "OutlineLevel"), 1),
                outlineNumber: textContent(task, "OutlineNumber"),
                wbs: textContent(task, "WBS") || undefined,
                type: textContent(task, "Type") ? parseNumber(textContent(task, "Type"), 0) : undefined,
                calendarUID: textContent(task, "CalendarUID") || undefined,
                priority: textContent(task, "Priority") ? parseNumber(textContent(task, "Priority"), 0) : undefined,
                start: textContent(task, "Start"),
                finish: textContent(task, "Finish"),
                duration: textContent(task, "Duration"),
                actualStart: textContent(task, "ActualStart") || undefined,
                actualFinish: textContent(task, "ActualFinish") || undefined,
                deadline: textContent(task, "Deadline") || undefined,
                startVariance: textContent(task, "StartVariance") || undefined,
                finishVariance: textContent(task, "FinishVariance") || undefined,
                work: textContent(task, "Work") || undefined,
                workVariance: textContent(task, "WorkVariance") || undefined,
                totalSlack: textContent(task, "TotalSlack") || undefined,
                freeSlack: textContent(task, "FreeSlack") || undefined,
                cost: textContent(task, "Cost") ? parseNumber(textContent(task, "Cost"), 0) : undefined,
                actualCost: textContent(task, "ActualCost") ? parseNumber(textContent(task, "ActualCost"), 0) : undefined,
                remainingCost: textContent(task, "RemainingCost") ? parseNumber(textContent(task, "RemainingCost"), 0) : undefined,
                remainingWork: textContent(task, "RemainingWork") || undefined,
                actualWork: textContent(task, "ActualWork") || undefined,
                milestone: parseBoolean(textContent(task, "Milestone")),
                summary: parseBoolean(textContent(task, "Summary")),
                critical: textContent(task, "Critical") ? parseBoolean(textContent(task, "Critical")) : undefined,
                percentComplete: parseNumber(textContent(task, "PercentComplete"), 0),
                percentWorkComplete: textContent(task, "PercentWorkComplete") ? parseNumber(textContent(task, "PercentWorkComplete"), 0) : undefined,
                notes: textContent(task, "Notes") || undefined,
                constraintType: textContent(task, "ConstraintType") ? parseNumber(textContent(task, "ConstraintType"), 0) : undefined,
                constraintDate: textContent(task, "ConstraintDate") || undefined,
                extendedAttributes: Array.from(task.getElementsByTagName("ExtendedAttribute")).map((attribute) => ({
                    fieldID: textContent(attribute, "FieldID") || undefined,
                    value: textContent(attribute, "Value") || undefined
                })),
                baselines: Array.from(task.getElementsByTagName("Baseline")).map((baseline) => ({
                    number: textContent(baseline, "Number") ? parseNumber(textContent(baseline, "Number"), 0) : undefined,
                    start: textContent(baseline, "Start") || undefined,
                    finish: textContent(baseline, "Finish") || undefined,
                    work: textContent(baseline, "Work") || undefined,
                    cost: textContent(baseline, "Cost") ? parseNumber(textContent(baseline, "Cost"), 0) : undefined
                })),
                timephasedData: Array.from(task.getElementsByTagName("TimephasedData")).map((timephasedData) => ({
                    type: textContent(timephasedData, "Type") ? parseNumber(textContent(timephasedData, "Type"), 0) : undefined,
                    uid: textContent(timephasedData, "UID") || undefined,
                    start: textContent(timephasedData, "Start") || undefined,
                    finish: textContent(timephasedData, "Finish") || undefined,
                    unit: textContent(timephasedData, "Unit") ? parseNumber(textContent(timephasedData, "Unit"), 0) : undefined,
                    value: textContent(timephasedData, "Value") || undefined
                })),
                predecessors: Array.from(task.getElementsByTagName("PredecessorLink")).map((link) => ({
                    predecessorUid: textContent(link, "PredecessorUID"),
                    type: parseNumber(textContent(link, "Type"), 0),
                    linkLag: textContent(link, "LinkLag") || undefined
                }))
            })),
            resources: resources.map((resource) => ({
                uid: textContent(resource, "UID"),
                id: textContent(resource, "ID"),
                name: textContent(resource, "Name"),
                type: parseNumber(textContent(resource, "Type"), 0),
                initials: textContent(resource, "Initials") || undefined,
                group: textContent(resource, "Group") || undefined,
                workGroup: textContent(resource, "WorkGroup") ? parseNumber(textContent(resource, "WorkGroup"), 0) : undefined,
                maxUnits: textContent(resource, "MaxUnits") ? parseNumber(textContent(resource, "MaxUnits"), 0) : undefined,
                calendarUID: textContent(resource, "CalendarUID") || undefined,
                standardRate: textContent(resource, "StandardRate") || undefined,
                standardRateFormat: textContent(resource, "StandardRateFormat") ? parseNumber(textContent(resource, "StandardRateFormat"), 0) : undefined,
                overtimeRate: textContent(resource, "OvertimeRate") || undefined,
                overtimeRateFormat: textContent(resource, "OvertimeRateFormat") ? parseNumber(textContent(resource, "OvertimeRateFormat"), 0) : undefined,
                costPerUse: textContent(resource, "CostPerUse") ? parseNumber(textContent(resource, "CostPerUse"), 0) : undefined,
                work: textContent(resource, "Work") || undefined,
                actualWork: textContent(resource, "ActualWork") || undefined,
                remainingWork: textContent(resource, "RemainingWork") || undefined,
                cost: textContent(resource, "Cost") ? parseNumber(textContent(resource, "Cost"), 0) : undefined,
                actualCost: textContent(resource, "ActualCost") ? parseNumber(textContent(resource, "ActualCost"), 0) : undefined,
                remainingCost: textContent(resource, "RemainingCost") ? parseNumber(textContent(resource, "RemainingCost"), 0) : undefined,
                percentWorkComplete: textContent(resource, "PercentWorkComplete") ? parseNumber(textContent(resource, "PercentWorkComplete"), 0) : undefined,
                extendedAttributes: Array.from(resource.getElementsByTagName("ExtendedAttribute")).map((attribute) => ({
                    fieldID: textContent(attribute, "FieldID") || undefined,
                    value: textContent(attribute, "Value") || undefined
                })),
                baselines: Array.from(resource.getElementsByTagName("Baseline")).map((baseline) => ({
                    number: textContent(baseline, "Number") ? parseNumber(textContent(baseline, "Number"), 0) : undefined,
                    start: textContent(baseline, "Start") || undefined,
                    finish: textContent(baseline, "Finish") || undefined,
                    work: textContent(baseline, "Work") || undefined,
                    cost: textContent(baseline, "Cost") ? parseNumber(textContent(baseline, "Cost"), 0) : undefined
                })),
                timephasedData: Array.from(resource.getElementsByTagName("TimephasedData")).map((timephasedData) => ({
                    type: textContent(timephasedData, "Type") ? parseNumber(textContent(timephasedData, "Type"), 0) : undefined,
                    uid: textContent(timephasedData, "UID") || undefined,
                    start: textContent(timephasedData, "Start") || undefined,
                    finish: textContent(timephasedData, "Finish") || undefined,
                    unit: textContent(timephasedData, "Unit") ? parseNumber(textContent(timephasedData, "Unit"), 0) : undefined,
                    value: textContent(timephasedData, "Value") || undefined
                }))
            })),
            assignments: assignments.map((assignment) => ({
                uid: textContent(assignment, "UID"),
                taskUid: textContent(assignment, "TaskUID"),
                resourceUid: textContent(assignment, "ResourceUID"),
                start: textContent(assignment, "Start") || undefined,
                finish: textContent(assignment, "Finish") || undefined,
                startVariance: textContent(assignment, "StartVariance") || undefined,
                finishVariance: textContent(assignment, "FinishVariance") || undefined,
                delay: textContent(assignment, "Delay") || undefined,
                milestone: textContent(assignment, "Milestone") ? parseBoolean(textContent(assignment, "Milestone")) : undefined,
                workContour: textContent(assignment, "WorkContour") ? parseNumber(textContent(assignment, "WorkContour"), 0) : undefined,
                units: parseNumber(textContent(assignment, "Units"), 0),
                work: textContent(assignment, "Work") || undefined,
                cost: textContent(assignment, "Cost") ? parseNumber(textContent(assignment, "Cost"), 0) : undefined,
                actualCost: textContent(assignment, "ActualCost") ? parseNumber(textContent(assignment, "ActualCost"), 0) : undefined,
                remainingCost: textContent(assignment, "RemainingCost") ? parseNumber(textContent(assignment, "RemainingCost"), 0) : undefined,
                percentWorkComplete: textContent(assignment, "PercentWorkComplete") ? parseNumber(textContent(assignment, "PercentWorkComplete"), 0) : undefined,
                overtimeWork: textContent(assignment, "OvertimeWork") || undefined,
                actualOvertimeWork: textContent(assignment, "ActualOvertimeWork") || undefined,
                actualWork: textContent(assignment, "ActualWork") || undefined,
                remainingWork: textContent(assignment, "RemainingWork") || undefined,
                extendedAttributes: Array.from(assignment.getElementsByTagName("ExtendedAttribute")).map((attribute) => ({
                    fieldID: textContent(attribute, "FieldID") || undefined,
                    value: textContent(attribute, "Value") || undefined
                })),
                baselines: Array.from(assignment.getElementsByTagName("Baseline")).map((baseline) => ({
                    number: textContent(baseline, "Number") ? parseNumber(textContent(baseline, "Number"), 0) : undefined,
                    start: textContent(baseline, "Start") || undefined,
                    finish: textContent(baseline, "Finish") || undefined,
                    work: textContent(baseline, "Work") || undefined,
                    cost: textContent(baseline, "Cost") ? parseNumber(textContent(baseline, "Cost"), 0) : undefined
                })),
                timephasedData: Array.from(assignment.getElementsByTagName("TimephasedData")).map((timephasedData) => ({
                    type: textContent(timephasedData, "Type") ? parseNumber(textContent(timephasedData, "Type"), 0) : undefined,
                    uid: textContent(timephasedData, "UID") || undefined,
                    start: textContent(timephasedData, "Start") || undefined,
                    finish: textContent(timephasedData, "Finish") || undefined,
                    unit: textContent(timephasedData, "Unit") ? parseNumber(textContent(timephasedData, "Unit"), 0) : undefined,
                    value: textContent(timephasedData, "Value") || undefined
                }))
            }))
        });
    }
    function appendTextElement(doc, parent, name, value) {
        if (value === undefined || value === "") {
            return;
        }
        const element = doc.createElement(name);
        if (typeof value === "boolean") {
            element.textContent = value ? "1" : "0";
        }
        else {
            element.textContent = String(value);
        }
        parent.appendChild(element);
    }
    function formatXml(xml) {
        const normalized = xml.replace(/>\s*</g, "><").trim();
        const tokens = normalized.replace(/></g, ">\n<").split("\n");
        let indentLevel = 0;
        const formatted = [];
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
    function exportMsProjectXml(model) {
        const normalizedModel = ensureDefaultProjectCalendar(normalizeProjectModel(model));
        const doc = document.implementation.createDocument("", "Project", null);
        const project = doc.documentElement;
        project.setAttribute("xmlns", "http://schemas.microsoft.com/project");
        appendTextElement(doc, project, "Name", normalizedModel.project.name);
        appendTextElement(doc, project, "Title", normalizedModel.project.title);
        appendTextElement(doc, project, "Company", normalizedModel.project.company);
        appendTextElement(doc, project, "Author", normalizedModel.project.author);
        appendTextElement(doc, project, "CreationDate", normalizedModel.project.creationDate);
        appendTextElement(doc, project, "LastSaved", normalizedModel.project.lastSaved);
        appendTextElement(doc, project, "SaveVersion", normalizedModel.project.saveVersion);
        appendTextElement(doc, project, "CurrentDate", normalizedModel.project.currentDate);
        appendTextElement(doc, project, "StartDate", normalizedModel.project.startDate);
        appendTextElement(doc, project, "FinishDate", normalizedModel.project.finishDate);
        appendTextElement(doc, project, "ScheduleFromStart", normalizedModel.project.scheduleFromStart);
        appendTextElement(doc, project, "DefaultStartTime", normalizedModel.project.defaultStartTime);
        appendTextElement(doc, project, "DefaultFinishTime", normalizedModel.project.defaultFinishTime);
        appendTextElement(doc, project, "MinutesPerDay", normalizedModel.project.minutesPerDay);
        appendTextElement(doc, project, "MinutesPerWeek", normalizedModel.project.minutesPerWeek);
        appendTextElement(doc, project, "DaysPerMonth", normalizedModel.project.daysPerMonth);
        appendTextElement(doc, project, "StatusDate", normalizedModel.project.statusDate);
        appendTextElement(doc, project, "WeekStartDay", normalizedModel.project.weekStartDay);
        appendTextElement(doc, project, "WorkFormat", normalizedModel.project.workFormat);
        appendTextElement(doc, project, "DurationFormat", normalizedModel.project.durationFormat);
        appendTextElement(doc, project, "CurrencyCode", normalizedModel.project.currencyCode);
        appendTextElement(doc, project, "CurrencyDigits", normalizedModel.project.currencyDigits);
        appendTextElement(doc, project, "CurrencySymbol", normalizedModel.project.currencySymbol);
        appendTextElement(doc, project, "CurrencySymbolPosition", normalizedModel.project.currencySymbolPosition);
        appendTextElement(doc, project, "FYStartDate", normalizedModel.project.fyStartDate);
        appendTextElement(doc, project, "FiscalYearStart", normalizedModel.project.fiscalYearStart);
        appendTextElement(doc, project, "CriticalSlackLimit", normalizedModel.project.criticalSlackLimit);
        appendTextElement(doc, project, "DefaultTaskType", normalizedModel.project.defaultTaskType);
        appendTextElement(doc, project, "DefaultFixedCostAccrual", normalizedModel.project.defaultFixedCostAccrual);
        appendTextElement(doc, project, "DefaultStandardRate", normalizedModel.project.defaultStandardRate);
        appendTextElement(doc, project, "DefaultOvertimeRate", normalizedModel.project.defaultOvertimeRate);
        appendTextElement(doc, project, "DefaultTaskEVMethod", normalizedModel.project.defaultTaskEVMethod);
        appendTextElement(doc, project, "NewTaskStartDate", normalizedModel.project.newTaskStartDate);
        appendTextElement(doc, project, "NewTasksAreManual", normalizedModel.project.newTasksAreManual);
        appendTextElement(doc, project, "NewTasksEffortDriven", normalizedModel.project.newTasksEffortDriven);
        appendTextElement(doc, project, "NewTasksEstimated", normalizedModel.project.newTasksEstimated);
        appendTextElement(doc, project, "ActualsInSync", normalizedModel.project.actualsInSync);
        appendTextElement(doc, project, "EditableActualCosts", normalizedModel.project.editableActualCosts);
        appendTextElement(doc, project, "HonorConstraints", normalizedModel.project.honorConstraints);
        appendTextElement(doc, project, "InsertedProjectsLikeSummary", normalizedModel.project.insertedProjectsLikeSummary);
        appendTextElement(doc, project, "MultipleCriticalPaths", normalizedModel.project.multipleCriticalPaths);
        appendTextElement(doc, project, "TaskUpdatesResource", normalizedModel.project.taskUpdatesResource);
        appendTextElement(doc, project, "UpdateManuallyScheduledTasksWhenEditingLinks", normalizedModel.project.updateManuallyScheduledTasksWhenEditingLinks);
        appendTextElement(doc, project, "CalendarUID", normalizedModel.project.calendarUID);
        if (normalizedModel.project.outlineCodes.length > 0) {
            const outlineCodesElement = doc.createElement("OutlineCodes");
            for (const outlineCode of normalizedModel.project.outlineCodes) {
                const outlineCodeElement = doc.createElement("OutlineCode");
                appendTextElement(doc, outlineCodeElement, "FieldID", outlineCode.fieldID);
                appendTextElement(doc, outlineCodeElement, "FieldName", outlineCode.fieldName);
                appendTextElement(doc, outlineCodeElement, "Alias", outlineCode.alias);
                appendTextElement(doc, outlineCodeElement, "OnlyTableValues", outlineCode.onlyTableValues);
                appendTextElement(doc, outlineCodeElement, "Enterprise", outlineCode.enterprise);
                appendTextElement(doc, outlineCodeElement, "ResourceSubstitutionEnabled", outlineCode.resourceSubstitutionEnabled);
                appendTextElement(doc, outlineCodeElement, "LeafOnly", outlineCode.leafOnly);
                appendTextElement(doc, outlineCodeElement, "AllLevelsRequired", outlineCode.allLevelsRequired);
                if (outlineCode.masks.length > 0) {
                    const masksElement = doc.createElement("Masks");
                    for (const mask of outlineCode.masks) {
                        const maskElement = doc.createElement("Mask");
                        appendTextElement(doc, maskElement, "Level", mask.level);
                        appendTextElement(doc, maskElement, "Mask", mask.mask);
                        appendTextElement(doc, maskElement, "Length", mask.length);
                        appendTextElement(doc, maskElement, "Sequence", mask.sequence);
                        masksElement.appendChild(maskElement);
                    }
                    outlineCodeElement.appendChild(masksElement);
                }
                if (outlineCode.values.length > 0) {
                    const valuesElement = doc.createElement("Values");
                    for (const value of outlineCode.values) {
                        const valueElement = doc.createElement("Value");
                        appendTextElement(doc, valueElement, "Value", value.value);
                        appendTextElement(doc, valueElement, "Description", value.description);
                        valuesElement.appendChild(valueElement);
                    }
                    outlineCodeElement.appendChild(valuesElement);
                }
                outlineCodesElement.appendChild(outlineCodeElement);
            }
            project.appendChild(outlineCodesElement);
        }
        if (normalizedModel.project.wbsMasks.length > 0) {
            const wbsMasksElement = doc.createElement("WBSMasks");
            for (const wbsMask of normalizedModel.project.wbsMasks) {
                const wbsMaskElement = doc.createElement("WBSMask");
                appendTextElement(doc, wbsMaskElement, "Level", wbsMask.level);
                appendTextElement(doc, wbsMaskElement, "Mask", wbsMask.mask);
                appendTextElement(doc, wbsMaskElement, "Length", wbsMask.length);
                appendTextElement(doc, wbsMaskElement, "Sequence", wbsMask.sequence);
                wbsMasksElement.appendChild(wbsMaskElement);
            }
            project.appendChild(wbsMasksElement);
        }
        if (normalizedModel.project.extendedAttributes.length > 0) {
            const extendedAttributesElement = doc.createElement("ExtendedAttributes");
            for (const attribute of normalizedModel.project.extendedAttributes) {
                const extendedAttributeElement = doc.createElement("ExtendedAttribute");
                appendTextElement(doc, extendedAttributeElement, "FieldID", attribute.fieldID);
                appendTextElement(doc, extendedAttributeElement, "FieldName", attribute.fieldName);
                appendTextElement(doc, extendedAttributeElement, "Alias", attribute.alias);
                appendTextElement(doc, extendedAttributeElement, "CalculationType", attribute.calculationType);
                appendTextElement(doc, extendedAttributeElement, "RestrictValues", attribute.restrictValues);
                appendTextElement(doc, extendedAttributeElement, "AppendNewValues", attribute.appendNewValues);
                extendedAttributesElement.appendChild(extendedAttributeElement);
            }
            project.appendChild(extendedAttributesElement);
        }
        const calendarsElement = doc.createElement("Calendars");
        for (const calendar of normalizedModel.calendars) {
            const calendarElement = doc.createElement("Calendar");
            appendTextElement(doc, calendarElement, "UID", calendar.uid);
            appendTextElement(doc, calendarElement, "Name", calendar.name);
            appendTextElement(doc, calendarElement, "IsBaseCalendar", calendar.isBaseCalendar);
            appendTextElement(doc, calendarElement, "IsBaselineCalendar", calendar.isBaselineCalendar);
            appendTextElement(doc, calendarElement, "BaseCalendarUID", calendar.baseCalendarUID);
            if (calendar.exceptions.length > 0) {
                const exceptionsElement = doc.createElement("Exceptions");
                for (const exception of calendar.exceptions) {
                    const exceptionElement = doc.createElement("Exception");
                    appendTextElement(doc, exceptionElement, "Name", exception.name);
                    appendTextElement(doc, exceptionElement, "FromDate", exception.fromDate);
                    appendTextElement(doc, exceptionElement, "ToDate", exception.toDate);
                    appendTextElement(doc, exceptionElement, "DayWorking", exception.dayWorking);
                    appendWorkingTimes(doc, exceptionElement, exception.workingTimes);
                    exceptionsElement.appendChild(exceptionElement);
                }
                calendarElement.appendChild(exceptionsElement);
            }
            if (calendar.workWeeks.length > 0) {
                const workWeeksElement = doc.createElement("WorkWeeks");
                for (const workWeek of calendar.workWeeks) {
                    const workWeekElement = doc.createElement("WorkWeek");
                    appendTextElement(doc, workWeekElement, "Name", workWeek.name);
                    appendTextElement(doc, workWeekElement, "FromDate", workWeek.fromDate);
                    appendTextElement(doc, workWeekElement, "ToDate", workWeek.toDate);
                    appendWeekDays(doc, workWeekElement, workWeek.weekDays);
                    workWeeksElement.appendChild(workWeekElement);
                }
                calendarElement.appendChild(workWeeksElement);
            }
            appendWeekDays(doc, calendarElement, calendar.weekDays);
            calendarsElement.appendChild(calendarElement);
        }
        project.appendChild(calendarsElement);
        const tasksElement = doc.createElement("Tasks");
        for (const task of normalizedModel.tasks) {
            const taskElement = doc.createElement("Task");
            appendTextElement(doc, taskElement, "UID", task.uid);
            appendTextElement(doc, taskElement, "ID", task.id);
            appendTextElement(doc, taskElement, "Name", task.name);
            appendTextElement(doc, taskElement, "OutlineLevel", task.outlineLevel);
            appendTextElement(doc, taskElement, "OutlineNumber", task.outlineNumber);
            appendTextElement(doc, taskElement, "WBS", task.wbs);
            appendTextElement(doc, taskElement, "Type", task.type);
            appendTextElement(doc, taskElement, "CalendarUID", task.calendarUID);
            appendTextElement(doc, taskElement, "Priority", task.priority);
            appendTextElement(doc, taskElement, "Start", task.start);
            appendTextElement(doc, taskElement, "Finish", task.finish);
            appendTextElement(doc, taskElement, "Duration", task.duration);
            appendTextElement(doc, taskElement, "ActualStart", task.actualStart);
            appendTextElement(doc, taskElement, "ActualFinish", task.actualFinish);
            appendTextElement(doc, taskElement, "Deadline", task.deadline);
            appendTextElement(doc, taskElement, "StartVariance", task.startVariance);
            appendTextElement(doc, taskElement, "FinishVariance", task.finishVariance);
            appendTextElement(doc, taskElement, "Work", task.work);
            appendTextElement(doc, taskElement, "WorkVariance", task.workVariance);
            appendTextElement(doc, taskElement, "TotalSlack", task.totalSlack);
            appendTextElement(doc, taskElement, "FreeSlack", task.freeSlack);
            appendTextElement(doc, taskElement, "Cost", task.cost);
            appendTextElement(doc, taskElement, "ActualCost", task.actualCost);
            appendTextElement(doc, taskElement, "RemainingCost", task.remainingCost);
            appendTextElement(doc, taskElement, "RemainingWork", task.remainingWork);
            appendTextElement(doc, taskElement, "ActualWork", task.actualWork);
            appendTextElement(doc, taskElement, "ConstraintType", task.constraintType);
            appendTextElement(doc, taskElement, "ConstraintDate", task.constraintDate);
            appendTextElement(doc, taskElement, "Milestone", task.milestone);
            appendTextElement(doc, taskElement, "Summary", task.summary);
            appendTextElement(doc, taskElement, "Critical", task.critical);
            appendTextElement(doc, taskElement, "PercentComplete", task.percentComplete);
            appendTextElement(doc, taskElement, "PercentWorkComplete", task.percentWorkComplete);
            appendTextElement(doc, taskElement, "Notes", task.notes);
            for (const attribute of task.extendedAttributes) {
                const extendedAttributeElement = doc.createElement("ExtendedAttribute");
                appendTextElement(doc, extendedAttributeElement, "FieldID", attribute.fieldID);
                appendTextElement(doc, extendedAttributeElement, "Value", attribute.value);
                taskElement.appendChild(extendedAttributeElement);
            }
            for (const baseline of task.baselines) {
                const baselineElement = doc.createElement("Baseline");
                appendTextElement(doc, baselineElement, "Number", baseline.number);
                appendTextElement(doc, baselineElement, "Start", baseline.start);
                appendTextElement(doc, baselineElement, "Finish", baseline.finish);
                appendTextElement(doc, baselineElement, "Work", baseline.work);
                appendTextElement(doc, baselineElement, "Cost", baseline.cost);
                taskElement.appendChild(baselineElement);
            }
            for (const timephasedData of task.timephasedData) {
                const timephasedDataElement = doc.createElement("TimephasedData");
                appendTextElement(doc, timephasedDataElement, "Type", timephasedData.type);
                appendTextElement(doc, timephasedDataElement, "UID", timephasedData.uid);
                appendTextElement(doc, timephasedDataElement, "Start", timephasedData.start);
                appendTextElement(doc, timephasedDataElement, "Finish", timephasedData.finish);
                appendTextElement(doc, timephasedDataElement, "Unit", timephasedData.unit);
                appendTextElement(doc, timephasedDataElement, "Value", timephasedData.value);
                taskElement.appendChild(timephasedDataElement);
            }
            for (const predecessor of task.predecessors) {
                const predecessorElement = doc.createElement("PredecessorLink");
                appendTextElement(doc, predecessorElement, "PredecessorUID", predecessor.predecessorUid);
                appendTextElement(doc, predecessorElement, "Type", predecessor.type);
                appendTextElement(doc, predecessorElement, "LinkLag", predecessor.linkLag);
                taskElement.appendChild(predecessorElement);
            }
            tasksElement.appendChild(taskElement);
        }
        project.appendChild(tasksElement);
        const resourcesElement = doc.createElement("Resources");
        for (const resource of normalizedModel.resources) {
            const resourceElement = doc.createElement("Resource");
            appendTextElement(doc, resourceElement, "UID", resource.uid);
            appendTextElement(doc, resourceElement, "ID", resource.id);
            appendTextElement(doc, resourceElement, "Name", resource.name);
            appendTextElement(doc, resourceElement, "Type", resource.type);
            appendTextElement(doc, resourceElement, "Initials", resource.initials);
            appendTextElement(doc, resourceElement, "Group", resource.group);
            appendTextElement(doc, resourceElement, "WorkGroup", resource.workGroup);
            appendTextElement(doc, resourceElement, "MaxUnits", resource.maxUnits);
            appendTextElement(doc, resourceElement, "CalendarUID", resource.calendarUID);
            appendTextElement(doc, resourceElement, "StandardRate", resource.standardRate);
            appendTextElement(doc, resourceElement, "StandardRateFormat", resource.standardRateFormat);
            appendTextElement(doc, resourceElement, "OvertimeRate", resource.overtimeRate);
            appendTextElement(doc, resourceElement, "OvertimeRateFormat", resource.overtimeRateFormat);
            appendTextElement(doc, resourceElement, "CostPerUse", resource.costPerUse);
            appendTextElement(doc, resourceElement, "Work", resource.work);
            appendTextElement(doc, resourceElement, "ActualWork", resource.actualWork);
            appendTextElement(doc, resourceElement, "RemainingWork", resource.remainingWork);
            appendTextElement(doc, resourceElement, "Cost", resource.cost);
            appendTextElement(doc, resourceElement, "ActualCost", resource.actualCost);
            appendTextElement(doc, resourceElement, "RemainingCost", resource.remainingCost);
            appendTextElement(doc, resourceElement, "PercentWorkComplete", resource.percentWorkComplete);
            for (const attribute of resource.extendedAttributes) {
                const extendedAttributeElement = doc.createElement("ExtendedAttribute");
                appendTextElement(doc, extendedAttributeElement, "FieldID", attribute.fieldID);
                appendTextElement(doc, extendedAttributeElement, "Value", attribute.value);
                resourceElement.appendChild(extendedAttributeElement);
            }
            for (const baseline of resource.baselines) {
                const baselineElement = doc.createElement("Baseline");
                appendTextElement(doc, baselineElement, "Number", baseline.number);
                appendTextElement(doc, baselineElement, "Start", baseline.start);
                appendTextElement(doc, baselineElement, "Finish", baseline.finish);
                appendTextElement(doc, baselineElement, "Work", baseline.work);
                appendTextElement(doc, baselineElement, "Cost", baseline.cost);
                resourceElement.appendChild(baselineElement);
            }
            for (const timephasedData of resource.timephasedData) {
                const timephasedDataElement = doc.createElement("TimephasedData");
                appendTextElement(doc, timephasedDataElement, "Type", timephasedData.type);
                appendTextElement(doc, timephasedDataElement, "UID", timephasedData.uid);
                appendTextElement(doc, timephasedDataElement, "Start", timephasedData.start);
                appendTextElement(doc, timephasedDataElement, "Finish", timephasedData.finish);
                appendTextElement(doc, timephasedDataElement, "Unit", timephasedData.unit);
                appendTextElement(doc, timephasedDataElement, "Value", timephasedData.value);
                resourceElement.appendChild(timephasedDataElement);
            }
            resourcesElement.appendChild(resourceElement);
        }
        project.appendChild(resourcesElement);
        const assignmentsElement = doc.createElement("Assignments");
        for (const assignment of normalizedModel.assignments) {
            const assignmentElement = doc.createElement("Assignment");
            appendTextElement(doc, assignmentElement, "UID", assignment.uid);
            appendTextElement(doc, assignmentElement, "TaskUID", assignment.taskUid);
            appendTextElement(doc, assignmentElement, "ResourceUID", assignment.resourceUid);
            appendTextElement(doc, assignmentElement, "Start", assignment.start);
            appendTextElement(doc, assignmentElement, "Finish", assignment.finish);
            appendTextElement(doc, assignmentElement, "StartVariance", assignment.startVariance);
            appendTextElement(doc, assignmentElement, "FinishVariance", assignment.finishVariance);
            appendTextElement(doc, assignmentElement, "Delay", assignment.delay);
            appendTextElement(doc, assignmentElement, "Milestone", assignment.milestone);
            appendTextElement(doc, assignmentElement, "WorkContour", assignment.workContour);
            appendTextElement(doc, assignmentElement, "Units", assignment.units);
            appendTextElement(doc, assignmentElement, "Work", assignment.work);
            appendTextElement(doc, assignmentElement, "Cost", assignment.cost);
            appendTextElement(doc, assignmentElement, "ActualCost", assignment.actualCost);
            appendTextElement(doc, assignmentElement, "RemainingCost", assignment.remainingCost);
            appendTextElement(doc, assignmentElement, "PercentWorkComplete", assignment.percentWorkComplete);
            appendTextElement(doc, assignmentElement, "OvertimeWork", assignment.overtimeWork);
            appendTextElement(doc, assignmentElement, "ActualOvertimeWork", assignment.actualOvertimeWork);
            appendTextElement(doc, assignmentElement, "ActualWork", assignment.actualWork);
            appendTextElement(doc, assignmentElement, "RemainingWork", assignment.remainingWork);
            for (const attribute of assignment.extendedAttributes) {
                const extendedAttributeElement = doc.createElement("ExtendedAttribute");
                appendTextElement(doc, extendedAttributeElement, "FieldID", attribute.fieldID);
                appendTextElement(doc, extendedAttributeElement, "Value", attribute.value);
                assignmentElement.appendChild(extendedAttributeElement);
            }
            for (const baseline of assignment.baselines) {
                const baselineElement = doc.createElement("Baseline");
                appendTextElement(doc, baselineElement, "Number", baseline.number);
                appendTextElement(doc, baselineElement, "Start", baseline.start);
                appendTextElement(doc, baselineElement, "Finish", baseline.finish);
                appendTextElement(doc, baselineElement, "Work", baseline.work);
                appendTextElement(doc, baselineElement, "Cost", baseline.cost);
                assignmentElement.appendChild(baselineElement);
            }
            for (const timephasedData of assignment.timephasedData) {
                const timephasedDataElement = doc.createElement("TimephasedData");
                appendTextElement(doc, timephasedDataElement, "Type", timephasedData.type);
                appendTextElement(doc, timephasedDataElement, "UID", timephasedData.uid);
                appendTextElement(doc, timephasedDataElement, "Start", timephasedData.start);
                appendTextElement(doc, timephasedDataElement, "Finish", timephasedData.finish);
                appendTextElement(doc, timephasedDataElement, "Unit", timephasedData.unit);
                appendTextElement(doc, timephasedDataElement, "Value", timephasedData.value);
                assignmentElement.appendChild(timephasedDataElement);
            }
            assignmentsElement.appendChild(assignmentElement);
        }
        project.appendChild(assignmentsElement);
        const serializer = new XMLSerializer();
        const serialized = serializer.serializeToString(doc);
        return `<?xml version="1.0" encoding="UTF-8"?>\n${formatXml(serialized)}\n`;
    }
    function normalizeProjectModel(model) {
        return JSON.parse(JSON.stringify(model));
    }
    function validateProjectModel(model) {
        const issues = [];
        const taskUidSet = new Set();
        const taskIdSet = new Set();
        const resourceUidSet = new Set();
        const calendarUidSet = new Set();
        if (!model.project.name) {
            issues.push({ level: "warning", scope: "project", message: "Project Name が空です" });
        }
        if (model.project.saveVersion !== undefined && model.project.saveVersion < 0) {
            issues.push({ level: "warning", scope: "project", message: "Project SaveVersion は 0 以上が望ましいです" });
        }
        if (!model.project.startDate) {
            issues.push({ level: "warning", scope: "project", message: "Project StartDate が空です" });
        }
        if (!model.project.finishDate) {
            issues.push({ level: "warning", scope: "project", message: "Project FinishDate が空です" });
        }
        if (model.project.minutesPerDay !== undefined && model.project.minutesPerDay <= 0) {
            issues.push({ level: "warning", scope: "project", message: "Project MinutesPerDay は正の値が望ましいです" });
        }
        if (model.project.minutesPerWeek !== undefined && model.project.minutesPerWeek <= 0) {
            issues.push({ level: "warning", scope: "project", message: "Project MinutesPerWeek は正の値が望ましいです" });
        }
        if (model.project.daysPerMonth !== undefined && model.project.daysPerMonth <= 0) {
            issues.push({ level: "warning", scope: "project", message: "Project DaysPerMonth は正の値が望ましいです" });
        }
        if (model.project.weekStartDay !== undefined && (model.project.weekStartDay < 1 || model.project.weekStartDay > 7)) {
            issues.push({ level: "warning", scope: "project", message: "Project WeekStartDay は 1..7 が望ましいです" });
        }
        if (model.project.workFormat !== undefined && model.project.workFormat < 0) {
            issues.push({ level: "warning", scope: "project", message: "Project WorkFormat は 0 以上が望ましいです" });
        }
        if (model.project.durationFormat !== undefined && model.project.durationFormat < 0) {
            issues.push({ level: "warning", scope: "project", message: "Project DurationFormat は 0 以上が望ましいです" });
        }
        if (model.project.currencyDigits !== undefined && model.project.currencyDigits < 0) {
            issues.push({ level: "warning", scope: "project", message: "Project CurrencyDigits は 0 以上が望ましいです" });
        }
        if (model.project.currencySymbolPosition !== undefined && model.project.currencySymbolPosition < 0) {
            issues.push({ level: "warning", scope: "project", message: "Project CurrencySymbolPosition は 0 以上が望ましいです" });
        }
        if (model.project.fyStartDate !== undefined && !parseDateValue(model.project.fyStartDate)) {
            issues.push({ level: "warning", scope: "project", message: "Project FYStartDate の日付形式が解釈できません" });
        }
        if (model.project.criticalSlackLimit !== undefined && model.project.criticalSlackLimit < 0) {
            issues.push({ level: "warning", scope: "project", message: "Project CriticalSlackLimit は 0 以上が望ましいです" });
        }
        if (model.project.defaultTaskType !== undefined && model.project.defaultTaskType < 0) {
            issues.push({ level: "warning", scope: "project", message: "Project DefaultTaskType は 0 以上が望ましいです" });
        }
        if (model.project.defaultFixedCostAccrual !== undefined && model.project.defaultFixedCostAccrual < 0) {
            issues.push({ level: "warning", scope: "project", message: "Project DefaultFixedCostAccrual は 0 以上が望ましいです" });
        }
        if (model.project.defaultTaskEVMethod !== undefined && model.project.defaultTaskEVMethod < 0) {
            issues.push({ level: "warning", scope: "project", message: "Project DefaultTaskEVMethod は 0 以上が望ましいです" });
        }
        if (model.project.newTaskStartDate !== undefined && model.project.newTaskStartDate < 0) {
            issues.push({ level: "warning", scope: "project", message: "Project NewTaskStartDate は 0 以上が望ましいです" });
        }
        for (const outlineCode of model.project.outlineCodes) {
            if (!outlineCode.fieldID && !outlineCode.fieldName) {
                issues.push({ level: "warning", scope: "project", message: "Project OutlineCode は FieldID または FieldName を持つことが望ましいです" });
            }
            for (const mask of outlineCode.masks) {
                if (mask.level < 1) {
                    issues.push({ level: "warning", scope: "project", message: "Project OutlineCode Mask Level は 1 以上が望ましいです" });
                }
            }
        }
        for (const wbsMask of model.project.wbsMasks) {
            if (wbsMask.level < 1) {
                issues.push({ level: "warning", scope: "project", message: "Project WBSMask Level は 1 以上が望ましいです" });
            }
        }
        for (const attribute of model.project.extendedAttributes) {
            if (!attribute.fieldID && !attribute.fieldName) {
                issues.push({ level: "warning", scope: "project", message: "Project ExtendedAttribute は FieldID または FieldName を持つことが望ましいです" });
            }
            if (attribute.calculationType !== undefined && attribute.calculationType < 0) {
                issues.push({ level: "warning", scope: "project", message: "Project ExtendedAttribute CalculationType は 0 以上が望ましいです" });
            }
        }
        for (const calendar of model.calendars) {
            if (!calendar.uid) {
                issues.push({ level: "error", scope: "calendars", message: "Calendar UID が空です" });
            }
            if (calendar.isBaselineCalendar !== undefined && !calendar.isBaseCalendar && calendar.isBaselineCalendar) {
                issues.push({
                    level: "warning",
                    scope: "calendars",
                    message: `Calendar IsBaselineCalendar は通常 BaseCalendar と整合していることが望ましいです: ${describeCalendar(calendar)}`
                });
            }
            if (calendarUidSet.has(calendar.uid)) {
                issues.push({ level: "error", scope: "calendars", message: `Calendar UID が重複しています: ${calendar.uid}` });
            }
            calendarUidSet.add(calendar.uid);
            for (const weekDay of calendar.weekDays) {
                if (weekDay.dayType < 1 || weekDay.dayType > 7) {
                    issues.push({
                        level: "warning",
                        scope: "calendars",
                        message: `Calendar WeekDay DayType が 1..7 の範囲外です: ${describeCalendar(calendar)}`
                    });
                }
                for (const workingTime of weekDay.workingTimes) {
                    if (!workingTime.fromTime || !workingTime.toTime) {
                        issues.push({
                            level: "warning",
                            scope: "calendars",
                            message: `Calendar WorkingTime の時刻が不足しています: ${describeCalendar(calendar)}`
                        });
                    }
                }
            }
            for (const exception of calendar.exceptions) {
                const exceptionFrom = parseDateValue(exception.fromDate);
                const exceptionTo = parseDateValue(exception.toDate);
                if (exceptionFrom !== null && exceptionTo !== null && exceptionFrom > exceptionTo) {
                    issues.push({
                        level: "warning",
                        scope: "calendars",
                        message: `Calendar Exception FromDate が ToDate より後です: ${describeCalendar(calendar)}`
                    });
                }
                for (const workingTime of exception.workingTimes) {
                    if (!workingTime.fromTime || !workingTime.toTime) {
                        issues.push({
                            level: "warning",
                            scope: "calendars",
                            message: `Calendar Exception WorkingTime の時刻が不足しています: ${describeCalendar(calendar)}`
                        });
                    }
                }
            }
            for (const workWeek of calendar.workWeeks) {
                const workWeekFrom = parseDateValue(workWeek.fromDate);
                const workWeekTo = parseDateValue(workWeek.toDate);
                if (workWeekFrom !== null && workWeekTo !== null && workWeekFrom > workWeekTo) {
                    issues.push({
                        level: "warning",
                        scope: "calendars",
                        message: `Calendar WorkWeek FromDate が ToDate より後です: ${describeCalendar(calendar)}`
                    });
                }
                for (const weekDay of workWeek.weekDays) {
                    if (weekDay.dayType < 1 || weekDay.dayType > 7) {
                        issues.push({
                            level: "warning",
                            scope: "calendars",
                            message: `Calendar WorkWeek DayType が 1..7 の範囲外です: ${describeCalendar(calendar)}`
                        });
                    }
                }
            }
        }
        if (model.project.calendarUID && !calendarUidSet.has(model.project.calendarUID)) {
            issues.push({
                level: "error",
                scope: "project",
                message: `Project CalendarUID が既存 Calendar を指していません: ${model.project.calendarUID}`
            });
        }
        for (const calendar of model.calendars) {
            if (calendar.baseCalendarUID && !calendarUidSet.has(calendar.baseCalendarUID)) {
                issues.push({
                    level: "warning",
                    scope: "calendars",
                    message: `Calendar BaseCalendarUID が既存 Calendar を指していません: ${describeCalendar(calendar)}`
                });
            }
            if (calendar.baseCalendarUID && calendar.baseCalendarUID === calendar.uid) {
                issues.push({
                    level: "warning",
                    scope: "calendars",
                    message: `Calendar BaseCalendarUID が自身を指しています: ${describeCalendar(calendar)}`
                });
            }
        }
        for (const task of model.tasks) {
            if (!task.uid) {
                issues.push({ level: "error", scope: "tasks", message: "Task UID が空です" });
            }
            if (!task.id) {
                issues.push({ level: "error", scope: "tasks", message: `Task ID が空です: ${task.name || "(無名)"}` });
            }
            if (!task.name) {
                if (!isPlaceholderUid(task.uid)) {
                    issues.push({ level: "warning", scope: "tasks", message: `Task Name が空です: ${describeTask(task)}` });
                }
            }
            if (taskIdSet.has(task.id)) {
                issues.push({ level: "error", scope: "tasks", message: `Task ID が重複しています: ${task.id}` });
            }
            taskIdSet.add(task.id);
            if (!task.start) {
                issues.push({ level: "warning", scope: "tasks", message: `Task Start が空です: ${describeTask(task)}` });
            }
            if (!task.finish) {
                issues.push({ level: "warning", scope: "tasks", message: `Task Finish が空です: ${describeTask(task)}` });
            }
            if (task.outlineLevel < 1 && !isPlaceholderUid(task.uid)) {
                issues.push({ level: "error", scope: "tasks", message: `Task OutlineLevel が不正です: ${describeTask(task)}` });
            }
            if (task.calendarUID && !calendarUidSet.has(task.calendarUID)) {
                issues.push({
                    level: "warning",
                    scope: "tasks",
                    message: `Task CalendarUID が既存 Calendar を指していません: ${describeTask(task)}`
                });
            }
            if (task.outlineNumber && !isPlaceholderUid(task.uid)) {
                const outlineParts = task.outlineNumber.split(".").filter(Boolean);
                if (outlineParts.length !== task.outlineLevel) {
                    issues.push({
                        level: "warning",
                        scope: "tasks",
                        message: `Task OutlineNumber と OutlineLevel の整合が取れていません: ${describeTask(task)}`
                    });
                }
            }
            if (task.percentComplete < 0 || task.percentComplete > 100) {
                issues.push({
                    level: "warning",
                    scope: "tasks",
                    message: `Task PercentComplete が 0..100 の範囲外です: ${describeTask(task)}`
                });
            }
            if (task.percentWorkComplete !== undefined &&
                (task.percentWorkComplete < 0 || task.percentWorkComplete > 100)) {
                issues.push({
                    level: "warning",
                    scope: "tasks",
                    message: `Task PercentWorkComplete が 0..100 の範囲外です: ${describeTask(task)}`
                });
            }
            const taskStart = parseDateValue(task.start);
            const taskFinish = parseDateValue(task.finish);
            const taskActualStart = parseDateValue(task.actualStart);
            const taskActualFinish = parseDateValue(task.actualFinish);
            const taskDeadline = parseDateValue(task.deadline);
            if (taskStart !== null && taskFinish !== null && taskStart > taskFinish) {
                issues.push({
                    level: "warning",
                    scope: "tasks",
                    message: `Task Start が Finish より後です: ${describeTask(task)}`
                });
            }
            if (taskFinish !== null && taskDeadline !== null && taskFinish > taskDeadline) {
                issues.push({
                    level: "warning",
                    scope: "tasks",
                    message: `Task Finish が Deadline より後です: ${describeTask(task)}`
                });
            }
            if (taskActualStart !== null && taskActualFinish !== null && taskActualStart > taskActualFinish) {
                issues.push({
                    level: "warning",
                    scope: "tasks",
                    message: `Task ActualStart が ActualFinish より後です: ${describeTask(task)}`
                });
            }
            if (taskUidSet.has(task.uid)) {
                issues.push({ level: "error", scope: "tasks", message: `Task UID が重複しています: ${task.uid}` });
            }
            for (const attribute of task.extendedAttributes) {
                if (!attribute.fieldID) {
                    issues.push({ level: "warning", scope: "tasks", message: `Task ExtendedAttribute に FieldID がありません: ${describeTask(task)}` });
                }
            }
            for (const baseline of task.baselines) {
                if (baseline.number !== undefined && baseline.number < 0) {
                    issues.push({ level: "warning", scope: "tasks", message: `Task Baseline Number は 0 以上が望ましいです: ${describeTask(task)}` });
                }
                const baselineStart = parseDateValue(baseline.start);
                const baselineFinish = parseDateValue(baseline.finish);
                if (baselineStart !== null && baselineFinish !== null && baselineStart > baselineFinish) {
                    issues.push({ level: "warning", scope: "tasks", message: `Task Baseline Start が Finish より後です: ${describeTask(task)}` });
                }
            }
            for (const timephasedData of task.timephasedData) {
                if (timephasedData.type !== undefined && timephasedData.type < 0) {
                    issues.push({ level: "warning", scope: "tasks", message: `Task TimephasedData Type は 0 以上が望ましいです: ${describeTask(task)}` });
                }
                const timephasedStart = parseDateValue(timephasedData.start);
                const timephasedFinish = parseDateValue(timephasedData.finish);
                if (timephasedStart !== null && timephasedFinish !== null && timephasedStart > timephasedFinish) {
                    issues.push({ level: "warning", scope: "tasks", message: `Task TimephasedData Start が Finish より後です: ${describeTask(task)}` });
                }
            }
            taskUidSet.add(task.uid);
            if (task.priority !== undefined && (task.priority < 0 || task.priority > 1000)) {
                issues.push({
                    level: "warning",
                    scope: "tasks",
                    message: `Task Priority が 0..1000 の範囲外です: ${describeTask(task)}`
                });
            }
            if (task.cost !== undefined && task.cost < 0) {
                issues.push({ level: "warning", scope: "tasks", message: `Task Cost が負値です: ${describeTask(task)}` });
            }
            if (task.actualCost !== undefined && task.actualCost < 0) {
                issues.push({ level: "warning", scope: "tasks", message: `Task ActualCost が負値です: ${describeTask(task)}` });
            }
            if (task.remainingCost !== undefined && task.remainingCost < 0) {
                issues.push({ level: "warning", scope: "tasks", message: `Task RemainingCost が負値です: ${describeTask(task)}` });
            }
        }
        for (const resource of model.resources) {
            if (!resource.uid) {
                issues.push({ level: "error", scope: "resources", message: "Resource UID が空です" });
            }
            if (!resource.name) {
                if (!isPlaceholderUid(resource.uid)) {
                    issues.push({ level: "warning", scope: "resources", message: `Resource Name が空です: ${describeResource(resource)}` });
                }
            }
            if (resourceUidSet.has(resource.uid)) {
                issues.push({ level: "error", scope: "resources", message: `Resource UID が重複しています: ${resource.uid}` });
            }
            resourceUidSet.add(resource.uid);
            if (resource.calendarUID && !calendarUidSet.has(resource.calendarUID)) {
                issues.push({
                    level: "warning",
                    scope: "resources",
                    message: `Resource CalendarUID が既存 Calendar を指していません: ${describeResource(resource)}`
                });
            }
            if (resource.workGroup !== undefined && resource.workGroup < 0) {
                issues.push({
                    level: "warning",
                    scope: "resources",
                    message: `Resource WorkGroup は 0 以上が望ましいです: ${describeResource(resource)}`
                });
            }
            if (resource.overtimeRateFormat !== undefined && resource.overtimeRateFormat < 0) {
                issues.push({
                    level: "warning",
                    scope: "resources",
                    message: `Resource OvertimeRateFormat は 0 以上が望ましいです: ${describeResource(resource)}`
                });
            }
            if (resource.cost !== undefined && resource.cost < 0) {
                issues.push({ level: "warning", scope: "resources", message: `Resource Cost が負値です: ${describeResource(resource)}` });
            }
            if (resource.actualCost !== undefined && resource.actualCost < 0) {
                issues.push({ level: "warning", scope: "resources", message: `Resource ActualCost が負値です: ${describeResource(resource)}` });
            }
            if (resource.remainingCost !== undefined && resource.remainingCost < 0) {
                issues.push({ level: "warning", scope: "resources", message: `Resource RemainingCost が負値です: ${describeResource(resource)}` });
            }
            if (resource.percentWorkComplete !== undefined &&
                (resource.percentWorkComplete < 0 || resource.percentWorkComplete > 100)) {
                issues.push({
                    level: "warning",
                    scope: "resources",
                    message: `Resource PercentWorkComplete が 0..100 の範囲外です: ${describeResource(resource)}`
                });
            }
            for (const attribute of resource.extendedAttributes) {
                if (!attribute.fieldID) {
                    issues.push({ level: "warning", scope: "resources", message: `Resource ExtendedAttribute に FieldID がありません: ${describeResource(resource)}` });
                }
            }
            for (const baseline of resource.baselines) {
                if (baseline.number !== undefined && baseline.number < 0) {
                    issues.push({
                        level: "warning",
                        scope: "resources",
                        message: `Resource Baseline Number は 0 以上が望ましいです: ${describeResource(resource)}`
                    });
                }
                const baselineStart = parseDateValue(baseline.start);
                const baselineFinish = parseDateValue(baseline.finish);
                if (baselineStart !== null && baselineFinish !== null && baselineStart > baselineFinish) {
                    issues.push({
                        level: "warning",
                        scope: "resources",
                        message: `Resource Baseline Start が Finish より後です: ${describeResource(resource)}`
                    });
                }
            }
            for (const timephasedData of resource.timephasedData) {
                if (timephasedData.type !== undefined && timephasedData.type < 0) {
                    issues.push({
                        level: "warning",
                        scope: "resources",
                        message: `Resource TimephasedData Type は 0 以上が望ましいです: ${describeResource(resource)}`
                    });
                }
                const timephasedStart = parseDateValue(timephasedData.start);
                const timephasedFinish = parseDateValue(timephasedData.finish);
                if (timephasedStart !== null && timephasedFinish !== null && timephasedStart > timephasedFinish) {
                    issues.push({
                        level: "warning",
                        scope: "resources",
                        message: `Resource TimephasedData Start が Finish より後です: ${describeResource(resource)}`
                    });
                }
            }
        }
        for (const task of model.tasks) {
            for (const predecessor of task.predecessors) {
                if (!taskUidSet.has(predecessor.predecessorUid)) {
                    issues.push({
                        level: "error",
                        scope: "tasks",
                        message: `PredecessorUID が既存 Task を指していません: ${describeTask(task)}, ${describeTaskRef(model, predecessor.predecessorUid)}`
                    });
                }
            }
        }
        for (const assignment of model.assignments) {
            if (!assignment.uid) {
                issues.push({ level: "warning", scope: "assignments", message: "Assignment UID が空です" });
            }
            if (!taskUidSet.has(assignment.taskUid)) {
                issues.push({
                    level: "error",
                    scope: "assignments",
                    message: `Assignment TaskUID が既存 Task を指していません: ${describeAssignment(assignment)}, ${describeTaskRef(model, assignment.taskUid)}`
                });
            }
            if (!resourceUidSet.has(assignment.resourceUid) && !isUnassignedResourceUid(assignment.resourceUid)) {
                issues.push({
                    level: "error",
                    scope: "assignments",
                    message: `Assignment ResourceUID が既存 Resource を指していません: ${describeAssignment(assignment)}, ${describeTaskRef(model, assignment.taskUid)}, ${describeResourceRef(model, assignment.resourceUid)}`
                });
            }
            if (!assignment.start) {
                issues.push({ level: "warning", scope: "assignments", message: `Assignment Start が空です: ${describeAssignment(assignment)}` });
            }
            if (!assignment.finish) {
                issues.push({ level: "warning", scope: "assignments", message: `Assignment Finish が空です: ${describeAssignment(assignment)}` });
            }
            const assignmentStart = parseDateValue(assignment.start);
            const assignmentFinish = parseDateValue(assignment.finish);
            if (assignmentStart !== null && assignmentFinish !== null && assignmentStart > assignmentFinish) {
                issues.push({
                    level: "warning",
                    scope: "assignments",
                    message: `Assignment Start が Finish より後です: ${describeAssignment(assignment)}`
                });
            }
            if (assignment.units !== undefined && assignment.units < 0) {
                issues.push({
                    level: "warning",
                    scope: "assignments",
                    message: `Assignment Units が負値です: ${describeAssignment(assignment)}`
                });
            }
            if (assignment.cost !== undefined && assignment.cost < 0) {
                issues.push({
                    level: "warning",
                    scope: "assignments",
                    message: `Assignment Cost が負値です: ${describeAssignment(assignment)}`
                });
            }
            if (assignment.actualCost !== undefined && assignment.actualCost < 0) {
                issues.push({
                    level: "warning",
                    scope: "assignments",
                    message: `Assignment ActualCost が負値です: ${describeAssignment(assignment)}`
                });
            }
            if (assignment.remainingCost !== undefined && assignment.remainingCost < 0) {
                issues.push({
                    level: "warning",
                    scope: "assignments",
                    message: `Assignment RemainingCost が負値です: ${describeAssignment(assignment)}`
                });
            }
            if (assignment.percentWorkComplete !== undefined &&
                (assignment.percentWorkComplete < 0 || assignment.percentWorkComplete > 100)) {
                issues.push({
                    level: "warning",
                    scope: "assignments",
                    message: `Assignment PercentWorkComplete が 0..100 の範囲外です: ${describeAssignment(assignment)}`
                });
            }
            if (assignment.overtimeWork !== undefined && !assignment.overtimeWork) {
                issues.push({
                    level: "warning",
                    scope: "assignments",
                    message: `Assignment OvertimeWork が空です: ${describeAssignment(assignment)}`
                });
            }
            if (assignment.actualOvertimeWork !== undefined && !assignment.actualOvertimeWork) {
                issues.push({
                    level: "warning",
                    scope: "assignments",
                    message: `Assignment ActualOvertimeWork が空です: ${describeAssignment(assignment)}`
                });
            }
            if (assignment.workContour !== undefined && assignment.workContour < 0) {
                issues.push({
                    level: "warning",
                    scope: "assignments",
                    message: `Assignment WorkContour は 0 以上が望ましいです: ${describeAssignment(assignment)}`
                });
            }
            if (assignment.startVariance !== undefined && !assignment.startVariance) {
                issues.push({
                    level: "warning",
                    scope: "assignments",
                    message: `Assignment StartVariance が空です: ${describeAssignment(assignment)}`
                });
            }
            for (const attribute of assignment.extendedAttributes) {
                if (!attribute.fieldID) {
                    issues.push({
                        level: "warning",
                        scope: "assignments",
                        message: `Assignment ExtendedAttribute に FieldID がありません: ${describeAssignment(assignment)}`
                    });
                }
            }
            for (const baseline of assignment.baselines) {
                if (baseline.number !== undefined && baseline.number < 0) {
                    issues.push({
                        level: "warning",
                        scope: "assignments",
                        message: `Assignment Baseline Number は 0 以上が望ましいです: ${describeAssignment(assignment)}`
                    });
                }
                const baselineStart = parseDateValue(baseline.start);
                const baselineFinish = parseDateValue(baseline.finish);
                if (baselineStart !== null && baselineFinish !== null && baselineStart > baselineFinish) {
                    issues.push({
                        level: "warning",
                        scope: "assignments",
                        message: `Assignment Baseline Start が Finish より後です: ${describeAssignment(assignment)}`
                    });
                }
            }
            for (const timephasedData of assignment.timephasedData) {
                if (timephasedData.type !== undefined && timephasedData.type < 0) {
                    issues.push({
                        level: "warning",
                        scope: "assignments",
                        message: `Assignment TimephasedData Type は 0 以上が望ましいです: ${describeAssignment(assignment)}`
                    });
                }
                const timephasedStart = parseDateValue(timephasedData.start);
                const timephasedFinish = parseDateValue(timephasedData.finish);
                if (timephasedStart !== null && timephasedFinish !== null && timephasedStart > timephasedFinish) {
                    issues.push({
                        level: "warning",
                        scope: "assignments",
                        message: `Assignment TimephasedData Start が Finish より後です: ${describeAssignment(assignment)}`
                    });
                }
            }
            if (assignment.finishVariance !== undefined && !assignment.finishVariance) {
                issues.push({
                    level: "warning",
                    scope: "assignments",
                    message: `Assignment FinishVariance が空です: ${describeAssignment(assignment)}`
                });
            }
        }
        return issues;
    }
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
        exportPhaseDetailView,
        exportCsvParentId,
        normalizeProjectModel,
        validateProjectModel
    };
})();
