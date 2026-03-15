# Phase Specs Index

**Run ID:** 2026-03-13-1430-appointment-contacts-and-contact-slinging
**Generated:** 2026-03-13
**Sub-specs:** 6

## Dependency Graph

```
Sub-Spec 1 (Appointment Contact Linking) — no dependencies
Sub-Spec 2 (ContactTemplateContext Extensions) — no dependencies
Sub-Spec 4 (Settings) — no dependencies
Sub-Spec 3 (ContactService Extensions) — depends on Sub-Spec 2
Sub-Spec 5 (ContactProcessor) — depends on Sub-Specs 2, 3
Sub-Spec 6 (Contact Slinging UI) — depends on Sub-Spec 5
```

## Execution Order

**Parallel Wave 1:** Sub-Specs 1, 2, 4 (no dependencies)
**Sequential Wave 2:** Sub-Spec 3 (after Sub-Spec 2)
**Sequential Wave 3:** Sub-Spec 5 (after Sub-Specs 2, 3)
**Sequential Wave 4:** Sub-Spec 6 (after Sub-Spec 5)

## Phase Spec Files

| # | Title | Dependencies | File |
|---|-------|-------------|------|
| 1 | Appointment Contact Linking | None | [sub-spec-1-appointment-contact-linking.md](sub-spec-1-appointment-contact-linking.md) |
| 2 | ContactTemplateContext & TemplateService Extensions | None | [sub-spec-2-contact-template-context-extensions.md](sub-spec-2-contact-template-context-extensions.md) |
| 3 | ContactService Extensions | 2 | [sub-spec-3-contact-service-extensions.md](sub-spec-3-contact-service-extensions.md) |
| 4 | Settings — ContactNoteIncludeDetails | None | [sub-spec-4-settings-contact-note-include-details.md](sub-spec-4-settings-contact-note-include-details.md) |
| 5 | ContactProcessor Orchestrator | 2, 3 | [sub-spec-5-contact-processor.md](sub-spec-5-contact-processor.md) |
| 6 | Contact Slinging UI (Ribbon + ThisAddIn Dispatch) | 5 | [sub-spec-6-contact-slinging-ui.md](sub-spec-6-contact-slinging-ui.md) |

## Build & Test Commands

```bash
# Build entire solution
dotnet build SlingMD.sln --configuration Release

# Run all tests
dotnet test SlingMD.Tests\SlingMD.Tests.csproj

# Verify EmailProcessor untouched
git diff main -- SlingMD.Outlook/Services/EmailProcessor.cs
```
