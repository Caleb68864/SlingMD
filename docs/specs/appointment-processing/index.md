---
type: phase-spec-index
master_spec: "docs/specs/2026-03-13-appointment-processing.md"
date: 2026-03-13
sub_specs: 9
---

# SlingMD Appointment Processing -- Phase Specs

Refined from [2026-03-13-appointment-processing.md](../2026-03-13-appointment-processing.md).

| Sub-Spec | Title | Dependencies | Phase Spec |
|----------|-------|--------------|------------|
| 1 | ObsidianSettings Extensions + AppointmentProcessor Core | none | [sub-spec-1-settings-processor-core.md](sub-spec-1-settings-processor-core.md) |
| 2 | ContactService Meeting Extensions + TemplateService Appointment Contexts | 1 | [sub-spec-2-contact-template-extensions.md](sub-spec-2-contact-template-extensions.md) |
| 3 | Recurring Meeting Threading | 1 | [sub-spec-3-recurring-meeting-threading.md](sub-spec-3-recurring-meeting-threading.md) |
| 4 | Companion Meeting Notes | 1, 2 | [sub-spec-4-companion-meeting-notes.md](sub-spec-4-companion-meeting-notes.md) |
| 5 | Bulk "Save Today's Appointments" | 1 | [sub-spec-5-bulk-save-today.md](sub-spec-5-bulk-save-today.md) |
| 6 | Ribbon Extensions | 1, 5 | [sub-spec-6-ribbon-extensions.md](sub-spec-6-ribbon-extensions.md) |
| 7 | Tabbed SettingsForm Rewrite | 1 | [sub-spec-7-tabbed-settings-form.md](sub-spec-7-tabbed-settings-form.md) |
| 8 | Task Creation Integration for Appointments | 1 | [sub-spec-8-task-creation.md](sub-spec-8-task-creation.md) |
| 9 | Tests | 1, 2, 3, 4, 5, 8 | [sub-spec-9-tests.md](sub-spec-9-tests.md) |

## Execution Order

**Parallelizable after Sub-Spec 1:**
- Sub-specs 2, 3, 5, 7, 8 can run in parallel (all depend only on sub-spec 1)

**Sequential dependencies:**
- Sub-spec 4 requires sub-specs 1 + 2
- Sub-spec 6 requires sub-specs 1 + 5
- Sub-spec 9 requires all others

## Execution

Run `/forge-run docs/specs/appointment-processing/` to execute all phase specs.
Run `/forge-run docs/specs/appointment-processing/ --sub N` to execute a single sub-spec.
