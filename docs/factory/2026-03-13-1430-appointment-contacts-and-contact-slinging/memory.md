# Factory Memory: 2026-03-13-1430-appointment-contacts-and-contact-slinging

This file accumulates context across all factory stages. Each stage agent reads it before starting and appends findings.

## Project Conventions (from CLAUDE.md)

- .NET Framework 4.7.2 C# Outlook VSTO add-in
- Service-oriented architecture with EmailProcessor as orchestrator
- Services: FileService, ThreadService, TaskService, ContactService, TemplateService, StatusService
- PascalCase for classes/methods/properties, camelCase for variables/parameters, _camelCase for private fields
- Explicit typing over var
- Braces on new lines, 4-space indentation
- Services must have 'Service' suffix, Interfaces must start with 'I' prefix
- Use fully qualified System.Exception, catch specific exceptions
- Build: dotnet build SlingMD.sln --configuration Release
- Test: dotnet test SlingMD.Tests\SlingMD.Tests.csproj
- xUnit + Moq for testing

## Brain Dump Input

I want the appointments module to check the vault for contacts and link them in the note if possible, also have the option to offer to create them like the email module does. Also if it makes sense we might need to start splitting code into modules that do each type of thing. But don't touch the bullet proof and tested code for email slinging if you don't need to. Next thing I want to do is add contact slinging — either single contacts or all contacts from an email or all contacts in your address book.

## Stage Outputs

(Populated by stage agents as they complete)

### Stage 1: Brainstorm (Design)

**Output:** `design.md` — full design document covering all three features.

**Key findings from codebase analysis:**

- `AppointmentProcessor` (1000+ lines) already extracts attendee names and writes `[[Name]]` wiki-links in frontmatter, but never calls `ContactService.ContactExists()` or `CreateContactNote()`. The contact linking gap is approximately 30 lines of code to add post-export.
- `EmailProcessor` contact resolution flow is at lines 466-516: collects names, checks `ManagedContactNoteExists()` / `ContactExists()`, refreshes managed notes, shows `ContactConfirmationDialog` for new contacts. This exact pattern will be replicated in `AppointmentProcessor`.
- `ContactService` already has all vault lookup methods needed (`ContactExists`, `ManagedContactNoteExists`, `CreateContactNote`). New methods needed only for rich `ContactItem` data extraction.
- Both processors duplicate ~14 compiled regex fields and a `CleanSubject()` method. Extraction deferred to avoid touching EmailProcessor.
- `ThisAddIn.ProcessSelection()` already dispatches on `MailItem` vs `AppointmentItem` — adding `ContactItem` detection is straightforward.
- `SlingRibbon.xml` has three groups (Email, Settings, Appointments). Adding a Contacts group follows the same pattern.
- Test pattern: structural tests for processors (no COM mocking), filesystem-based tests for services using temp directories.
- Settings model uses JSON serialization with `NormalizeLoadedSettings()` for backward compatibility — new settings must be added there too.

**Architecture decision:** No shared base class or interface extraction. EmailProcessor remains completely untouched. Duplication is accepted as the safer trade-off.

### Stage 2: Forge (Spec)

**Output:** `spec.md` — full agent-executable spec with 6 sub-specs.

**Quality Score:** 28/30 (walk-away quality)

| Dimension | Score |
|-----------|-------|
| outcome_clarity | 5 |
| scope_boundaries | 5 |
| decision_guidance | 4 |
| edge_coverage | 4 |
| acceptance_criteria | 5 |
| decomposition | 5 |

**Sub-spec count:** 6
**Refinement passes:** 0 (scored 28 on first pass, above 24 threshold)

**Sub-spec dependency graph:**
```
Sub-Spec 1 (Appointment Contact Linking) — no dependencies
Sub-Spec 2 (ContactTemplateContext Extensions) — no dependencies
Sub-Spec 4 (Settings) — no dependencies
Sub-Spec 3 (ContactService Extensions) — depends on Sub-Spec 2
Sub-Spec 5 (ContactProcessor) — depends on Sub-Specs 2, 3
Sub-Spec 6 (Contact Slinging UI) — depends on Sub-Spec 5
```

**Key decisions made during spec writing:**
- Attendee names from AppointmentProcessor are `[[Name]]` formatted — spec notes bracket stripping is required.
- Contact resolution insertion point identified: after line 566 (`coreExportSucceeded` check), before line 568 (Outlook task creation).
- `RenderRichContactContent` as a new method rather than modifying existing `RenderContactContent` — safer for existing contact note flow.
- From-email ContactItem resolution (`Recipient.AddressEntry.GetContact()`) explicitly deferred to out-of-scope.
- Birthday sentinel year 4501 edge case identified from Outlook COM behavior.

**Codebase findings from exploration:**
- `ContactConfirmationDialog` takes `List<string>` of contact names, returns `SelectedContacts` — reusable as-is.
- `ContactTemplateContext` currently has 5 properties — will grow to 13 with rich fields.
- `SettingsForm` already has a Contacts tab (line 136) — checkbox addition is straightforward.
- SlingRibbon.xml has 3 groups — adding 4th (Contacts) follows same pattern.
- Test pattern: `AppointmentProcessorTests` uses structural-only tests (no COM mocking), temp directory with cleanup.

### Stage 3: Prep (Phase Specs)

**Output:** `phase-specs/` directory with 7 files (6 phase specs + index).

**Artifacts produced:** 7 files
- `phase-specs/index.md` — dependency graph and execution order
- `phase-specs/sub-spec-1-appointment-contact-linking.md`
- `phase-specs/sub-spec-2-contact-template-context-extensions.md`
- `phase-specs/sub-spec-3-contact-service-extensions.md`
- `phase-specs/sub-spec-4-settings-contact-note-include-details.md`
- `phase-specs/sub-spec-5-contact-processor.md`
- `phase-specs/sub-spec-6-contact-slinging-ui.md`

**Sub-specs refined:** 6 (all from spec.md)

**Interface contracts identified:**
- Sub-Spec 2 provides: `ContactTemplateContext` rich fields, `RenderRichContactContent()`, `GetDefaultRichContactTemplate()`
- Sub-Spec 3 provides: `ExtractContactData(ContactItem)`, `CreateRichContactNote(ContactTemplateContext)`
- Sub-Spec 4 provides: `ObsidianSettings.ContactNoteIncludeDetails` (bool, default true)
- Sub-Spec 5 provides: `ContactProcessor` class, `ContactProcessingResult` enum, `ProcessContact()`, `ProcessAddressBook()`
- Sub-Spec 6 provides: Ribbon buttons, ThisAddIn dispatch methods

**Codebase patterns found:**
- AppointmentProcessor constructor pattern (lines 73-82): settings -> FileService -> TemplateService -> services
- EmailProcessor contact resolution (lines 466-516): dedup/sort/check managed/check exists/dialog
- COM property read pattern (AppointmentProcessor lines 184-203): individual try/catch per property
- Bulk error collection: `_bulkErrors.Add()` + `GetBulkErrors()` clear-on-read
- SettingsForm tab layout: `TableLayoutPanel` with `cRow` counter, `ColumnStyles 35%/65%`
- Test helpers: `TestFileService`, `TestTemplateService`, `ObsidianSettingsTestable` subclasses
- Template rendering: `LoadConfiguredTemplate()` -> `BuildMetadataReplacements()` -> `AddReplacement()` -> `ProcessTemplate()`

**Build/test commands:**
- Build: `dotnet build SlingMD.sln --configuration Release`
- Test: `dotnet test SlingMD.Tests\SlingMD.Tests.csproj`
- EmailProcessor guard: `git diff main -- SlingMD.Outlook/Services/EmailProcessor.cs`

**Files analyzed:**
- `SlingMD.Outlook/Services/AppointmentProcessor.cs` (1000+ lines, insertion point at line 566)
- `SlingMD.Outlook/Services/ContactService.cs` (534 lines, 2 new methods needed)
- `SlingMD.Outlook/Services/TemplateService.cs` (720 lines, ContactTemplateContext at lines 29-37)
- `SlingMD.Outlook/Services/EmailProcessor.cs` (contact pattern at lines 466-516, DO NOT MODIFY)
- `SlingMD.Outlook/Models/ObsidianSettings.cs` (494 lines, NormalizeLoadedSettings at line 404)
- `SlingMD.Outlook/ThisAddIn.cs` (330 lines, ProcessSelection at line 96)
- `SlingMD.Outlook/Ribbon/SlingRibbon.xml` (31 lines, 3 groups)
- `SlingMD.Outlook/Ribbon/SlingRibbon.cs` (152 lines, 3 callbacks)
- `SlingMD.Outlook/Forms/SettingsForm.cs` (Contacts tab at lines 349-388)
- 12 existing test files analyzed for patterns

### Stage 4: Run

- Sub-specs executed: 6
- Results: 6 PASS, 0 PARTIAL, 0 FAIL
- Waves: 4 (Wave 1: sub-specs 1,2,4 parallel; Wave 2: sub-spec 3; Wave 3: sub-spec 5; Wave 4: sub-spec 6)
- Files changed:
  - Modified: `SlingMD.Outlook/Services/AppointmentProcessor.cs` (contact linking block)
  - Modified: `SlingMD.Outlook/Services/TemplateService.cs` (ContactTemplateContext extensions, RenderRichContactContent)
  - Modified: `SlingMD.Outlook/Services/ContactService.cs` (ExtractContactData, CreateRichContactNote)
  - Modified: `SlingMD.Outlook/Models/ObsidianSettings.cs` (ContactNoteIncludeDetails)
  - Modified: `SlingMD.Outlook/Forms/SettingsForm.cs` (checkbox)
  - Created: `SlingMD.Outlook/Services/ContactProcessor.cs` (new orchestrator)
  - Modified: `SlingMD.Outlook/Ribbon/SlingRibbon.xml` (ContactsGroup)
  - Modified: `SlingMD.Outlook/Ribbon/SlingRibbon.cs` (callbacks)
  - Modified: `SlingMD.Outlook/ThisAddIn.cs` (dispatch, ProcessSelectedContact, SlingAllContacts)
  - Modified: `SlingMD.Outlook/SlingMD.Outlook.csproj` (ContactProcessor.cs include)
  - Created: `SlingMD.Tests/Services/ContactProcessorTests.cs` (3 tests)
  - Modified: `SlingMD.Tests/Services/AppointmentProcessorTests.cs` (1 test)
  - Modified: `SlingMD.Tests/Services/TemplateServiceTests.cs` (3 tests)
  - Modified: `SlingMD.Tests/Services/ContactServiceTests.cs` (2 tests)
  - Modified: `SlingMD.Tests/Models/ObsidianSettingsTests.cs` (3 tests)
  - Modified: `SlingMD.Tests/SlingMD.Tests.csproj` (ContactProcessorTests include)
- Build: PASS (0 errors, 0 warnings)
- Tests: 87/87 PASS
- EmailProcessor: untouched (git diff empty)
- Issues: none

### Stage 5: Verify

- Verdict: **PASS**
- Acceptance criteria: 35/36 PASS, 0 FAIL, 1 NEEDS_REVIEW (human review — ribbon visual layout)
- Code quality: 3 SUGGESTION-level findings, 0 CRITICAL, 0 IMPORTANT
- Integration: All API contracts match, no missing stubs, no unresolved imports, dependency order correct
- EmailProcessor: confirmed untouched (git diff empty)
- Report: `verify-report.md`

**Suggestion-level findings (non-blocking):**
1. Unused variable `cleanName` in `ContactService.ExtractContactData` (line 539) — `fileNameNoExtension` on line 540 duplicates the same call
2. `ContactProcessor.ProcessAddressBook` uses Logger for progress instead of StatusService (spec mentioned StatusService but Logger is acceptable)
3. `NormalizeLoadedSettings` does not explicitly handle `ContactNoteIncludeDetails` — intentional per Stage 3 decision (C# default of `true` handles upgrade case)

## Issues Log

(Failures, retries, and decisions logged here)
