# BimTasksV2 User Guide — Design Document

**Created:** 2026-03-10
**Purpose:** Personal developer reference guide (HTML) covering architecture and all commands
**Audience:** You (Assaf) — no installation docs, no sales pitch, just functionality + honest notes

---

## Decisions Made

| Decision | Choice | Reasoning |
|----------|--------|-----------|
| Organization | By ribbon panel | Matches Revit UI — find button, find docs |
| Status badges | ✅ Solid / 🟡 Sketchy / ⛔ Broken / 🔧 Needs Improvement | Honest assessment per command |
| Issue granularity | Command-level + architecture-level | Per-command issues inline, systemic issues in final chapter |
| Build strategy | Multi-file → join at end | Easier to iterate per chapter, merge when done |
| Visual style | Matches BIM Israel DB guide sample | Same CSS framework, Mermaid diagrams, callouts, cards |

---

## File Structure

```
docs/guide/
├── styles.css              — shared CSS (Inter + JetBrains Mono, cards, callouts, badges)
├── 00-cover-toc.html       — cover page + table of contents
├── 01-architecture.html    — bird's eye: ALC isolation, command flow, event system, key services
├── 02-walls.html           — 10 commands (Sync Phases, Copy Linked, Height, Split, Overlaps, Cladding x3, Trim Corners, Window Families)
├── 03-structure.html       — 4 commands (Beams To Ground, Join Walls & Floors, Split Floor, Join Beams to Walls)
├── 04-schedules.html       — 3 commands (Export Schedule, Edit In Excel, Import Key Schedules)
├── 05-boq.html             — 5 commands (Setup, Schedules, AutoFill, Calc Price, Seifei Choze)
├── 06-tools.html           — 11+ commands (Filter Tree, Calc Area, Uniformat, Copy From Link, Dockable Panel, Floating Toolbar, Export JSON, Voice, Floor From DWG x2, Filter Legend)
└── 07-issues.html          — cross-cutting architectural debt and improvement ideas
```

Final step: merge all into single `guide.html` with print-to-PDF button.

---

## Per-Command Card Template

Each command gets:
1. **Header**: command name + ribbon button label + status badge
2. **What it does**: 2-4 sentences, functional description
3. **Workflow**: numbered steps — what the user does + what the code does behind the scenes
4. **Diagram**: Mermaid flow (only for multi-step commands like SplitWall, BOQ pipeline)
5. **Key files**: handler, helpers, views involved
6. **Known issues / improvements**: issue cards (only if applicable)

---

## Architecture Section Content (01)

- **Bird's eye diagram**: Revit → BimTasksApp (default ALC) → reflection → BimTasksBootstrapper (isolated ALC) → Prism container → services + handlers
- **Command execution flow**: Ribbon click → proxy ExternalCommand → reflection invoke → Handler.Execute(UIApplication)
- **Initialization pipeline**: OnStartup → Initialize → RegisterDockablePane → OnFirstIdle (3-phase boot)
- **Event system**: Prism EventAggregator for UI coordination (toolbar toggle, panel switch, filter tree reset)
- **Key services**: RevitContextService, CommandDispatcherService, VoskVoiceService, ScheduleExcelExportService
- **NO Prism/Unity internals** — just "what talks to what"

---

## Build Order (Roadmap)

| Phase | Deliverable | Status |
|-------|-------------|--------|
| 1 | `styles.css` + `00-cover-toc.html` | ⬜ Not started |
| 2 | `01-architecture.html` | ⬜ Not started |
| 3 | `02-walls.html` | ⬜ Not started |
| 4 | `03-structure.html` | ⬜ Not started |
| 5 | `04-schedules.html` | ⬜ Not started |
| 6 | `05-boq.html` | ⬜ Not started |
| 7 | `06-tools.html` | ⬜ Not started |
| 8 | `07-issues.html` | ⬜ Not started |
| 9 | Merge into single `guide.html` | ⬜ Not started |

**Rule:** Complete and review one phase before starting the next. Each phase requires reading the actual handler code — no guessing from file names.
