---
date: 2026-03-12
title: "SlingMD Typed Template System"
client: Open Source
project: SlingMD
repo: SlingMD
author: Caleb Bennett
quality_score:
  outcome: 5
  scope: 4
  decision_guidance: 5
  edges: 4
  criteria: 5
  decomposition: 4
  total: 27
status: ready
tags:
  - spec
  - slingmd
  - templates
  - obsidian
  - dataview
---

# SlingMD Typed Template System

## Outcome

When complete, SlingMD can render email notes, contact notes, thread notes, and inline Obsidian task text from first-class templates instead of hardcoded string assembly. A user can point SlingMD at template files and filename formats, customize frontmatter and note body for Dataview-heavy workflows, and still get the current behavior when no custom templates are configured.

## Intent

**Trade-off hierarchy:**
1. Backward compatibility over template power
2. First-class template architecture over more one-off settings
3. Predictable fallback behavior over strict template validation
4. Incremental delivery over trying to add standalone task-note support in the same pass

**Decision boundaries:**
- Prefer a typed template context and explicit render paths over generic free-form string munging.
- Preserve current default output semantics whenever a user has not configured custom templates.
- Treat "task templating" in this phase as customization of the inline Obsidian task text SlingMD already generates, not creation of a separate task note type.

**Escalation triggers:**
- Stop and ask if the implementation would require changing the existing task model from inline tasks to standalone task notes.
- Stop and ask if supporting third-party templating engines or arbitrary code execution becomes necessary to satisfy the issue.
- Stop and ask if template-driven filenames would break duplicate detection, thread ordering, or attachment storage guarantees.

## Context

GitHub issue #8 requests support for Dataview-friendly templates so users can customize frontmatter, body, and filenames for Email, Contact, and Tasks. Doug also asked whether the VSTO project must be built and published in Visual Studio rather than VS Code.

Current repo findings:
- `SlingMD.Outlook/Services/TemplateService.cs` only provides file loading, naive `{{token}}` replacement, YAML frontmatter construction, and a default thread-note template.
- `SlingMD.Outlook/Services/EmailProcessor.cs` still assembles email note content inline and only uses `TemplateService` for frontmatter plus the thread summary note.
- `SlingMD.Outlook/Services/ContactService.cs` still assembles contact note content inline.
- `SlingMD.Outlook/Services/TaskService.cs` generates a hardcoded inline Obsidian task line and Outlook task items; there is no standalone task note concept today.
- `SlingMD.Outlook/Models/ObsidianSettings.cs` and `SlingMD.Outlook/Forms/SettingsForm.cs` already expose several formatting settings, so template configuration belongs in the same settings surface.
- `SlingMD.Outlook/SlingMD.Outlook.csproj` imports `Microsoft.VisualStudio.Tools.Office.targets`, so build and publish remain tied to a Visual Studio Office tooling environment even if edits happen in VS Code.

Relevant vault context:
- `[[SlingMD]]` exists in the personal vault as the project note.
- No project-level `docs/intent.md` exists in this repo, so intent is defined by this spec.

## Requirements

1. Users can configure a templates folder plus template filenames for email, contact, task, and thread rendering.
2. Users can configure filename format settings for the rendered note types that need files on disk.
3. SlingMD keeps working for existing users with no new settings present in their saved JSON.
4. Email note rendering uses a typed template context rather than hardcoded inline assembly.
5. Contact note rendering uses the same template architecture rather than hardcoded inline assembly.
6. Thread note rendering stays template-driven and is brought under the same typed rendering model.
7. Task templating in this phase applies to the inline Obsidian task text SlingMD inserts into email notes.
8. Missing, unreadable, or incomplete custom templates fall back to built-in defaults instead of failing note export.
9. Rendered filenames are sanitized and remain compatible with duplicate detection, thread grouping, and attachment handling.
10. Tests cover template loading, fallback behavior, token replacement, and filename rendering.
11. Documentation clarifies the new template feature and the Visual Studio requirement for building or publishing the VSTO add-in.

## Sub-Specs

### Sub-Spec 1: Template Settings and Configuration Surface

**Scope:** Add template-related settings, persistence, validation, and UI so users can configure the feature without editing code.

**Files:**
- `SlingMD.Outlook/Models/ObsidianSettings.cs`
- `SlingMD.Outlook/Forms/SettingsForm.cs`

**Acceptance Criteria:**
- `[STRUCTURAL]` `ObsidianSettings` contains template folder, per-template filename settings, and per-note filename format settings with defaults that preserve current behavior.
- `[STRUCTURAL]` Settings load/save logic accepts older `ObsidianSettings.json` files that do not contain the new template fields.
- `[BEHAVIORAL]` A user who upgrades without changing settings still gets the same email, contact, thread, and task output they got before this feature.
- `[HUMAN REVIEW]` The settings UI makes it clear that task templating customizes inline Obsidian task text only, not standalone task notes.

**Dependencies:** none

### Sub-Spec 2: Typed Template Renderer and Default Templates

**Scope:** Replace ad hoc template handling with a typed rendering model that can render frontmatter, body, and filename outputs for each supported note type, with safe defaults.

**Files:**
- `SlingMD.Outlook/Services/TemplateService.cs`
- `SlingMD.Outlook/Templates/ThreadNoteTemplate.md`
- New default templates for email, contact, and task rendering under `SlingMD.Outlook/Templates/`
- `SlingMD.Outlook/SlingMD.Outlook.csproj`

**Acceptance Criteria:**
- `[STRUCTURAL]` `TemplateService` exposes typed rendering paths for email, contact, task, and thread contexts instead of relying only on generic string replacement.
- `[STRUCTURAL]` Default template files for every supported render type are included in the project output alongside the existing thread template.
- `[BEHAVIORAL]` If a configured custom template file is missing or unreadable, SlingMD falls back to the built-in default for that render type and still completes the export.
- `[BEHAVIORAL]` Filename rendering produces sanitized, deterministic names that do not include invalid path characters even when token values contain them.

**Dependencies:** Sub-Spec 1

### Sub-Spec 3: Integrate Template Rendering into Email, Contact, and Task Flows

**Scope:** Refactor the note generation flows to use the typed template renderer, preserve current operational behavior, add tests, and document the VSTO build/publish constraint.

**Files:**
- `SlingMD.Outlook/Services/EmailProcessor.cs`
- `SlingMD.Outlook/Services/ContactService.cs`
- `SlingMD.Outlook/Services/TaskService.cs`
- `SlingMD.Tests/Services/*`
- `README.md`

**Acceptance Criteria:**
- `[BEHAVIORAL]` Email export uses the typed template renderer for frontmatter, body, and filename generation while preserving duplicate detection, threading, and attachment handling.
- `[BEHAVIORAL]` Contact note creation uses the typed template renderer and still produces a usable default Dataview-oriented contact note when no custom template is configured.
- `[BEHAVIORAL]` Inline Obsidian task generation uses the typed template renderer and does not introduce standalone task-note creation in this phase.
- `[MECHANICAL]` `dotnet test SlingMD.Tests\\SlingMD.Tests.csproj` passes in an environment with the required VSTO/Office build targets installed.
- `[STRUCTURAL]` `README.md` documents template customization points and explicitly states that editing can happen in VS Code but building/publishing the VSTO add-in requires Visual Studio Office tooling.

**Dependencies:** Sub-Spec 2

## Edge Cases

- Existing users have older settings files with none of the new template properties. Load must succeed and default behavior must remain intact.
- A user points SlingMD at a template folder that exists on one machine but not another. Export should fall back cleanly rather than hard-failing.
- A rendered filename contains invalid characters, excessive length, or values that would collide with existing thread naming rules. Sanitization must happen after template rendering, not before.
- A custom email template omits fields that duplicate detection or thread grouping depend on. Required operational metadata must still be emitted by default templates, and any future support for omitting them should be treated as an escalation point rather than guessed.
- Contact and thread templates must remain Dataview-friendly out of the box even when users do not provide custom templates.
- Task templates must not imply support for separate task notes; documentation and UI copy need to avoid that ambiguity.

## Out of Scope

- Creating standalone task notes or a new task-note storage location
- Supporting third-party template engines, script execution, or arbitrary Dataview code generation helpers
- Redesigning attachment storage, thread grouping, or duplicate detection algorithms beyond what is required to preserve compatibility with templated filenames
- Replacing the VSTO build and publish model with a VS Code-only toolchain

## Constraints

### Musts

- Preserve backward compatibility for users who never touch the new template settings.
- Keep template rendering deterministic and file-system safe.
- Keep thread-note support working with existing user override behavior.
- Add automated coverage for renderer fallback and filename generation.

### Must-Nots

- Must not silently convert the task feature into a standalone task-note system.
- Must not require users to author custom template files just to keep SlingMD working.
- Must not break duplicate detection, attachment linking, or thread note updates.

### Preferences

- Prefer explicit template contexts over unstructured string dictionaries when the output shape is known.
- Prefer repo-shipped default templates that mirror current output so the feature feels additive, not disruptive.
- Prefer small, localized refactors over broad architectural churn unrelated to template rendering.

### Escalation Triggers

- Ask before allowing templates to suppress metadata fields that the rest of the system depends on.
- Ask before adding scripting, conditionals, or a nontrivial mini-language to templates.
- Ask before changing file naming semantics in a way that would alter user-visible thread organization.

## Verification

1. Configure no template settings and export an email, a contact, and a threaded email. Confirm the produced notes match current behavior closely enough that existing users would not notice a regression.
2. Configure custom email and contact templates plus custom filename formats. Export representative messages and confirm frontmatter, body, and filenames reflect the custom templates.
3. Configure a missing custom template path and confirm SlingMD falls back to built-in defaults without aborting the export.
4. Export a threaded email with attachments while using custom filename formats. Confirm duplicate detection, thread summary updates, attachment links, and note locations still work.
5. Run `dotnet test SlingMD.Tests\\SlingMD.Tests.csproj` in a machine that has the VSTO Office targets installed.
6. Review the settings UI and README text to confirm users will understand the distinction between inline task templating and standalone task-note support, and will understand the Visual Studio requirement for build/publish.
