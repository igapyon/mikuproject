---
title: miku-project miku-soft reference
description: Current shared miku-soft references used to maintain this historical combined main-application and Web-App repository.
topics:
  - miku-project
  - miku-soft
  - maintenance
  - architecture
category: reference
status: stable
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
    label: igapyon-miku-soft-developer shared skill
    checked: 2026-08-05
  - type: local-file
    role: supporting
    path: docs/architecture.md
    label: miku-project product architecture
    checked: 2026-08-05
---

# miku-project miku-soft Reference

The shared [igapyon-miku-soft-developer](https://github.com/igapyon/igapyon-agent-skills/tree/devel/skills/igapyon-miku-soft-developer) skill is the current source for miku-soft policy. This repository uses its existing-project maintenance, Node main-application, and Web App references.

`miku-project` is a historical combined repository: it owns the `10 Main Application` product core, CLI, and runtime bundle as well as the colocated `11 Web App` source and generated single-file HTML. The semantic center remains `MS Project XML` and `ProjectModel`; the browser UI and CLI use the same core API contract.

## Reference Checkpoint

- Checked: 2026-08-05
- Installed-skill commit: unavailable; the installed skill directory is not a Git worktree.
- Main workflow: existing-project maintenance
- Relevant layers: `10 Main Application` and historical combined `11 Web App`

## Repository-Specific Status

- `src/ts/` is the source of truth. The tracked `src/js/`, `index.html`, and `miku-project.html` files are intentional generated browser artifacts for this still-combined repository and must be rebuilt, not hand-edited.
- `bundle/miku-project.mjs` is the intentionally single-file Node CLI runtime consumed by downstream Agent Skills. It is regenerated with `npm run build:cli-bundle`.
- The versioned `docs/miku-soft-*-design-v*.md` files are historical snapshots from before the shared reference policy. They are not the current miku-soft source of truth. Their eventual archival or removal is tracked in [migration-worklog.md](migration-worklog.md) after inbound-link review.

For product-specific architecture and command details, use [architecture.md](architecture.md), [development.md](development.md), and [miku-project-ai-json-spec.md](miku-project-ai-json-spec.md).
