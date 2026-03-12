---
date: 2026-03-12
title: "SlingMD Reliability Hardening"
client: Open Source
project: SlingMD
repo: SlingMD
author: Codex
quality_score:
  outcome: 5
  scope: 5
  decision_guidance: 4
  edges: 5
  criteria: 5
  decomposition: 4
  total: 28
status: ready
tags:
  - spec
  - slingmd
  - reliability
  - error-handling
  - testing
---

# SlingMD Reliability Hardening

## Outcome

When complete, SlingMD handles malformed settings, missing vault folders, export-time failures, metadata edge cases, and attachment path variations without derailing the Outlook session or producing misleading follow-up behavior. Successful exports should continue to look and behave the same as they do today except where a fix is required to prevent incorrect failure handling.

## Intent

**Trade-off hierarchy:**
1. Preserve current user-visible functionality over opportunistic cleanup
2. Fail safely and predictably over attempting to continue after a broken export state
3. Targeted hardening over broad refactors
4. Testable guardrails over silent best-effort behavior

**Decision boundaries:**
- Prefer explicit guards, fallbacks, and state resets over changing normal export semantics.
- Keep fixes localized to the services that already own the relevant responsibility unless shared behavior is the root cause.
- Preserve existing note shapes, thread organization, and task behavior unless the current implementation is internally inconsistent or corrupts output.

**Escalation triggers:**
- Stop and ask if fixing a reliability issue requires changing the saved note schema or user-visible thread naming conventions beyond compatibility repairs.
- Stop and ask if robust attachment-link handling would require a user-visible link-format change for existing notes.
- Stop and ask if automated verification requires replacing or materially restructuring the current VSTO build/test setup.

## Context

This spec is derived from a repository review performed on 2026-03-12 against the Outlook VSTO add-in in `SlingMD.Outlook/` and its tests in `SlingMD.Tests/`.

Current repo context:
- `CLAUDE.md` defines a service-oriented architecture centered on `EmailProcessor`, plus conventions around explicit typing and exception handling.
- `docs/specs/2026-03-12-slingmd-typed-template-system.md` shows the current spec style used in this repo.
- No `docs/intent.md` exists, so this spec provides its own execution guidance.
- Existing tests cover `ObsidianSettings`, `FileService`, `ContactService`, and `TaskService`, but there is no meaningful coverage for `EmailProcessor`, `ThreadService`, or `AttachmentService`.
- Local verification in this environment is currently blocked because `dotnet test SlingMD.Tests\SlingMD.Tests.csproj` fails before build due to a missing `Microsoft.VisualStudio.Tools.Office.targets` installation.

Key findings driving this work:
- Startup can fail on a corrupt `ObsidianSettings.json`.
- Export failures can still fall through into contact creation and Obsidian launch.
- Canceling task options can disable task creation for the rest of the Outlook session.
- First-run vault states can fail during duplicate-cache population.
- Thread discovery parses a different date shape than the exporter writes.
- Frontmatter generation is not YAML-safe for quoted or multiline metadata.
- Attachment links do not account for non-local storage modes.
- Parts of the settings form appear to have handlers defined but not wired.

## Requirements

1. A malformed or partially corrupt settings file must not prevent the add-in from starting.
2. Export processing must stop cleanly after a fatal error and must not proceed into contacts or Obsidian launch for that failed export.
3. Task creation state must be scoped to the current export attempt so one canceled task dialog does not affect later exports.
4. First-run or missing vault folders must be treated as valid empty states rather than hard failures.
5. Thread discovery must parse metadata in the same shape the exporter writes today, or the exporter and parser must be aligned on one safe format.
6. Frontmatter generation must remain valid for subjects, names, and other metadata containing quotes, backslashes, or line breaks.
7. Attachment links must resolve correctly for same-folder, subfolder-per-note, and centralized-storage modes.
8. The settings form must wire the existing controls that are intended to be interactive.
9. Automated coverage must be expanded around the newly hardened behaviors, and stale expectations must be corrected.
10. Verification guidance must document the current VSTO/Office prerequisite for running the test suite in a compatible environment.

## Sub-Specs

### Sub-Spec 1: Startup and Export Flow Safeguards

**Scope:** Harden startup settings loading, per-run task state, and the top-level export control flow so broken state does not leak into later actions.

**Files:**
- `SlingMD.Outlook/ThisAddIn.cs`
- `SlingMD.Outlook/Models/ObsidianSettings.cs`
- `SlingMD.Outlook/Services/EmailProcessor.cs`
- `SlingMD.Outlook/Services/TaskService.cs`

**Acceptance Criteria:**
- `[BEHAVIORAL]` A corrupt or partially invalid `ObsidianSettings.json` does not prevent Outlook add-in startup; SlingMD loads with an explicit safe fallback path instead of crashing startup.
- `[BEHAVIORAL]` If email export hits a fatal processing error before completion, SlingMD does not continue into contact creation or Obsidian launch for that export attempt.
- `[BEHAVIORAL]` Canceling the task-options dialog affects only the current export attempt; a later successful export can still create Obsidian and/or Outlook tasks according to settings.
- `[STRUCTURAL]` Task creation state is reset or reinitialized per export attempt rather than persisting a prior canceled state indefinitely.

**Dependencies:** none

### Sub-Spec 2: File, Metadata, Attachment, and UI Hardening

**Scope:** Align metadata parsing and writing, protect first-run filesystem flows, produce valid frontmatter for real-world metadata, fix attachment-link generation, and wire intended settings-form controls.

**Files:**
- `SlingMD.Outlook/Services/EmailProcessor.cs`
- `SlingMD.Outlook/Services/ThreadService.cs`
- `SlingMD.Outlook/Services/TemplateService.cs`
- `SlingMD.Outlook/Services/AttachmentService.cs`
- `SlingMD.Outlook/Forms/SettingsForm.cs`

**Acceptance Criteria:**
- `[BEHAVIORAL]` Exporting into a fresh vault state with a missing inbox directory succeeds without throwing during duplicate detection or cache initialization.
- `[BEHAVIORAL]` Thread discovery and export metadata use compatible date parsing rules so earliest-thread reconstruction works with notes generated by the current exporter.
- `[BEHAVIORAL]` Frontmatter produced for email metadata remains valid and parseable when source fields contain double quotes, backslashes, or embedded newlines.
- `[BEHAVIORAL]` Attachment links inserted into notes resolve correctly for same-folder, per-note-subfolder, and centralized attachment storage modes.
- `[STRUCTURAL]` The intended interactive controls in `SettingsForm` are explicitly wired to their handlers rather than relying on unreachable methods.

**Dependencies:** Sub-Spec 1

### Sub-Spec 3: Coverage, Regression Tests, and Verification Guidance

**Scope:** Add or update tests for the hardened behaviors, fix stale expectations, and document the environment needed to execute the suite.

**Files:**
- `SlingMD.Tests/Models/ObsidianSettingsTests.cs`
- `SlingMD.Tests/Services/TaskServiceTests.cs`
- `SlingMD.Tests/Services/ContactServiceTests.cs`
- New or updated tests under `SlingMD.Tests/Services/` covering thread, attachment, template, and export-flow edge cases
- `README.md`

**Acceptance Criteria:**
- `[STRUCTURAL]` Automated tests exist for corrupt-settings fallback, task-state reset after cancel, missing inbox handling, thread date compatibility, frontmatter escaping, and attachment-link path generation at the service level where feasible.
- `[STRUCTURAL]` The stale contact-note expectation is updated so the suite matches the current `## Email History` output unless the implementation intentionally restores the previous heading.
- `[MECHANICAL]` `dotnet test SlingMD.Tests\SlingMD.Tests.csproj` passes on a machine with the required VSTO Office targets installed.
- `[STRUCTURAL]` `README.md` documents the prerequisite Office and VSTO tooling needed to run build and test commands successfully.

**Dependencies:** Sub-Spec 2

## Edge Cases

- `ObsidianSettings.json` exists but contains malformed JSON, truncated content, or type-mismatched values for only some properties.
- The configured vault path exists but the inbox folder does not yet exist because the user is exporting for the first time.
- An email subject, sender, or recipient name contains quotes, escaping characters, or line breaks that would invalidate YAML when inserted naively.
- A threaded note created before the fix uses the current exported date shape and must still be detected correctly during thread reconstruction.
- A user cancels task options on one export and immediately exports another email in the same Outlook session.
- Attachments are stored outside the note directory, requiring generated links to include a relative or otherwise correct path instead of only the filename.
- The settings form is opened by a user relying on browse or pattern-management controls that are visually present but currently may not be wired.

## Out of Scope

- Adding new end-user features unrelated to hardening or error handling
- Redesigning the output format of email, thread, contact, or task notes beyond compatibility-safe fixes
- Replacing the VSTO build and test model or migrating the project away from Outlook Office tooling
- Refactoring the entire exporter into new abstractions unless needed to land one of the listed hardening fixes
- Expanding into a broader UX redesign of the settings or progress dialogs

## Constraints

### Musts

- Preserve current success-path behavior wherever the current implementation is already correct.
- Keep fixes compatible with existing saved settings files and existing generated notes.
- Keep error handling explicit enough that failed exports do not silently trigger unrelated follow-up work.
- Add automated coverage for the highest-risk regressions introduced by the hardening work.

### Must-Nots

- Must not turn a recoverable settings or folder issue into an Outlook-startup crash.
- Must not leave task creation disabled across the session because of one canceled dialog.
- Must not emit invalid frontmatter for common real-world email metadata.
- Must not assume attachment links can always be represented by bare filenames.

### Preferences

- Prefer defensive checks and localized state management over large structural changes.
- Prefer shared parsing or escaping helpers over copy-pasted one-off fixes when the same metadata is read and written in multiple places.
- Prefer tests at the service boundary where Outlook interop can be avoided or isolated.

### Escalation Triggers

- Ask before intentionally changing note-frontmatter semantics in a way that could break downstream Obsidian or Dataview queries.
- Ask before introducing a new dependency solely to solve YAML escaping or path-generation concerns.
- Ask before reworking the project structure or test runner to compensate for missing local Office targets.

## Verification

1. Replace the saved settings file with malformed JSON and start the add-in. Confirm SlingMD initializes and surfaces a safe fallback behavior rather than failing startup.
2. Trigger an export failure in the main processing path and confirm SlingMD does not continue into contact prompts or Obsidian launch for that failed export.
3. Cancel the task-options dialog on one export, then export a second email with task creation enabled. Confirm the second export can still create the expected task output.
4. Export an email into a fresh vault state where the inbox folder does not yet exist. Confirm duplicate detection and cache logic treat the folder as empty and the note is saved successfully.
5. Export or simulate notes whose metadata contains quotes and multiline values, then inspect the resulting frontmatter to confirm it remains parseable and stable.
6. Export attachments using each storage mode and confirm the generated note links resolve to the saved files from within Obsidian-compatible paths.
7. Open the settings form and exercise the visible browse and pattern-management controls to confirm their handlers fire as intended.
8. Run `dotnet test SlingMD.Tests\SlingMD.Tests.csproj` on a machine with the required Office and VSTO targets installed and confirm the hardening-focused suite passes.

## Phase Specs

Refined by `/forge-prep` on 2026-03-12.

| Sub-Spec | Phase Spec |
|----------|------------|
| 1. Startup and Export Flow Safeguards | `docs/specs/slingmd-reliability-hardening/sub-spec-1-startup-and-export-flow-safeguards.md` |
| 2. File, Metadata, Attachment, and UI Hardening | `docs/specs/slingmd-reliability-hardening/sub-spec-2-file-metadata-attachment-and-ui-hardening.md` |
| 3. Coverage, Regression Tests, and Verification Guidance | `docs/specs/slingmd-reliability-hardening/sub-spec-3-coverage-regression-tests-and-verification-guidance.md` |

Index: `docs/specs/slingmd-reliability-hardening/index.md`

