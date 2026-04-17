(() => {
    const DEFAULT_PROJECT_MINUTES_PER_DAY = 480;
    const DEFAULT_PROJECT_MINUTES_PER_WEEK = 2400;
    const DEFAULT_PROJECT_DAYS_PER_MONTH = 20;
    const SAMPLE_PROJECT_DRAFT_VIEW = {
        view_type: "project_draft_view",
        project: {
            name: "mikuproject開発",
            planned_start: "2026-03-16",
            planned_finish: "2026-04-01",
            schedule_from_start: true,
            minutes_per_day: DEFAULT_PROJECT_MINUTES_PER_DAY,
            minutes_per_week: DEFAULT_PROJECT_MINUTES_PER_WEEK,
            days_per_month: DEFAULT_PROJECT_DAYS_PER_MONTH
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
        ],
        resources: [
            {
                uid: "res-1",
                name: "Mikuku",
                initials: "M",
                group: "Development",
                max_units: 1,
                calendar_uid: "1"
            }
        ],
        assignments: [
            {
                uid: "asg-1",
                task_uid: "draft-120",
                resource_uid: "res-1",
                start: "2026-03-16T09:00:00",
                finish: "2026-03-16T18:00:00",
                units: 1,
                work: "PT8H0M0S",
                percent_work_complete: 100
            },
            {
                uid: "asg-2",
                task_uid: "draft-130",
                resource_uid: "res-1",
                start: "2026-03-17T09:00:00",
                finish: "2026-03-17T18:00:00",
                units: 1,
                work: "PT8H0M0S",
                percent_work_complete: 100
            }
        ]
    };
    globalThis.__mikuprojectMsprojectSamples = {
        SAMPLE_PROJECT_DRAFT_VIEW,
        buildSampleXml(input) {
            const sampleModel = input.importProjectDraftView(SAMPLE_PROJECT_DRAFT_VIEW);
            sampleModel.project.currentDate = "2026-03-23T09:00:00";
            sampleModel.project.statusDate = "2026-03-23T09:00:00";
            return input.exportMsProjectXml(sampleModel);
        }
    };
})();
