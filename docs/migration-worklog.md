---
title: mikuproject miku-soft standardization worklog
description: Post-standardization alignment record for the historical combined mikuproject repository.
topics:
  - mikuproject
  - miku-soft
  - migration
  - release
  - web-app
category: workflow
status: active
audience:
  - maintainer
  - developer
  - agent
created: 2026-08-05
updated: 2026-08-05
sources:
  - type: upstream-doc
    role: primary
    url: https://github.com/igapyon/igapyon-agent-skills/tree/devel/skills/igapyon-miku-soft-developer
    label: current miku-soft shared policy
    checked: 2026-08-05
  - type: manual-verification
    role: verification
    label: npm run test:fast and bundled CLI version/help smoke
    checked: 2026-08-05
---

# mikuproject miku-soft Standardization Worklog

## 2026-08-05 Baseline

`mikuproject` predates the current miku-soft shared standards. It remains a historical combined repository: the `10 Main Application` core and CLI, plus the `11 Web App` source, generated HTML, browser JavaScript, and `lht-cmn` live together.

Verified before alignment work:

- `npm run test:fast`: 12 test files and 210 tests passed.
- `npm run build:cli-bundle`: generated the Node CLI runtime and source archive.
- The generated CLI runtime ran `--version` and `--help` successfully.

## Applied Local Alignment

- Added the repository-local entry point [miku-soft-reference.md](miku-soft-reference.md).
- Added `dist/` to the Node/Web local build-output ignore rules.
- Documented the CLI runtime contract in `--help` and fixed it with CLI tests.
- Replaced locale-sensitive ordering in source-archive, calendar, and AI-view paths with explicit UTF-16 code-unit ordering.

## Historical Reference Snapshots

The following pre-standardization documents remain temporarily as history and must not be treated as current shared policy:

- `docs/miku-soft-00-overview-design-v20260427.md`
- `docs/miku-soft-10-mainapp-design-v20260501.md`
- `docs/miku-soft-20-javaapp-design-v20260501.md`
- `docs/miku-soft-30-straight-conversion-v20260425.md`
- `docs/miku-soft-40-agentskills-design-v20260501.md`
- `docs/miku-soft-50-mcp-design-v20260501.md`

Before moving or removing them, inspect inbound links from published documentation, Agent Skills, and sibling repositories. Preserve only a clearly labelled historical record or project-specific decisions; do not maintain another copy of the shared standard.

## Decisions Required Before the Next Migration

### Node, CI, and Release Policy

The repository currently has a release-only workflow using GitHub Release publication and Node 20. The target Node support range, CI matrix, release build runtime, Action update/pinning policy, and checksum policy must be decided together with the shared miku-soft release profile.

Do not silently adopt Node 22/24 or change the release trigger until that policy is approved. After approval, align `package.json` runtime metadata, normal CI, release workflow, staged-asset checks, and runtime smoke checks as one compatibility change.

### Main Application Rename

The canonical Main Application name is `miku-project`. The package name, canonical CLI command, generated single-file app, CLI runtime/source archive, release asset names, repository URL, and current maintenance documents use that name.

The migration keeps these intentionally bounded compatibility contracts:

- `mikuproject` remains an npm CLI alias for `miku-project`.
- `mikuproject.html` is a lightweight redirect to `miku-project.html`.
- `globalThis.__mikuprojectCoreApi` remains an alias of the canonical `globalThis.__mikuProjectCoreApi`.
- `mikuproject_workbook_json` remains the existing JSON format identifier; changing a persisted data format is out of scope for a product-name migration.

Internal module globals and test file identifiers continue to use their established `__mikuproject...` / `mikuproject-...` names where they are not published contracts. Historical articles and imported miku-soft standard documents are retained as historical records rather than mass-renamed.

### Web Separation

Issue [#123](https://github.com/igapyon/miku-project/issues/123) establishes `miku-project` as the canonical Main Application name and explicitly excludes Web separation from the same rename change. The GitHub repository has been renamed to `igapyon/miku-project`; this repository retains the current combined layout until the follow-up Web migration begins. Web separation requires a human-created `miku-project-web` repository and GitHub Pages / Release decisions. Establish and verify that repository first, using the renamed main application's core API or a documented runtime artifact. Only then may the renamed main application repository relinquish Web-only source, generated HTML, `lht-cmn`, browser tests, and tracked generated browser JavaScript.

The current combined layout is retained until that separate migration is approved and complete.
