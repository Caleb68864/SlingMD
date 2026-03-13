---
date: 2026-03-13
title: "SlingMD Appointment Processing"
client: Open Source
project: SlingMD
repo: SlingMD
author: Caleb Bennett
quality_score:
  outcome: 5
  scope: 4
  decision_guidance: 5
  edges: 5
  criteria: 4
  decomposition: 4
  total: 27
status: executed
executed: 2026-03-13
result: 9/9 sub-specs passed
tags:
  - spec
  - slingmd
  - appointments
  - outlook
  - obsidian
---

# SlingMD Appointment Processing

## Outcome

When complete, SlingMD can export Outlook appointments to Obsidian as markdown notes with full parity to the email pipeline: HTML-to-markdown conversion, customizable templates, attachment handling, contact creation, task creation, recurring meeting threading, and companion meeting notes. Users can sling individual appointments (from Explorer or an open appointment inspector), bulk-save today's appointments, and configure all appointment behavior through a new tabbed settings UI.

## Intent

**Trade-off hierarchy (highest priority first):**
1. Match EmailProcessor patterns -- mirror how EmailProcessor orchestrates services, handles COM objects, and structures notes. Consistency with the battle-tested email pipeline is the default.
2. Additive-only changes to shared services -- extend ContactService, TemplateService, etc. without modifying existing method signatures or behavior.
3. Working end-to-end over perfect -- get the pipeline functional, accept rough edges that can be polished later.

**Decision boundaries:**
- **Decide autonomously:** Template format choices, frontmatter field names, filename sanitization, error message wording, test structure
- **Escalate (stop and ask):** Any change that modifies existing EmailProcessor behavior, any refactoring of shared service method signatures, any architectural pattern that diverges significantly from how EmailProcessor does it

## Context

SlingMD is a .NET Framework 4.7.2 C# Outlook VSTO add-in that exports emails to Obsidian as markdown. The existing EmailProcessor (~1100 lines) orchestrates FileService, TemplateService, ThreadService, TaskService, ContactService, AttachmentService, and StatusService. It takes `ObsidianSettings` in its constructor and creates all service instances internally.

The design doc (`docs/plans/2026-03-13-appointment-processing-design.md`) specifies **Approach A: Dedicated AppointmentProcessor** as a parallel peer to EmailProcessor, reusing all shared services with zero changes to the email pipeline.

A fork (dlboutwe/SlingMD) provides a working reference for appointment processing that can accelerate development.

Key existing patterns to follow:
- EmailProcessor constructor takes `ObsidianSettings`, creates services internally
- `ThisAddIn` holds processor instances as private fields, routes by selection type
- SlingRibbon implements `IRibbonExtensibility`, returns XML from `GetCustomUI()`
- TemplateService uses `{{key}}` placeholder replacement with typed context classes
- ContactService resolves SMTP addresses and builds wikilink lists filtered by `OlMailRecipientType`
- ObsidianSettings serializes to/from JSON in AppData via `Load()`/`Save()`
- SettingsForm already uses a TabControl
- Templates live in `SlingMD.Outlook/Templates/` as embedded .md files

Existing specs in this project:
- `2026-03-12-slingmd-typed-template-system.md` (ready) -- may overlap with template work here; coordinate
- `2026-03-12-slingmd-reliability-hardening.md` (executed) -- hardening already applied

## Requirements

1. AppointmentProcessor must process a single Outlook AppointmentItem into a markdown note with frontmatter (title, organizer, attendees by role, location, start/end, recurrence, tags, attachments)
2. HTML body of appointments must be converted to markdown using the same pipeline as email
3. Recurring meeting instances must be grouped into thread folders via ThreadService when GroupRecurringMeetings is enabled
4. Companion meeting notes stubs must be created with backlinks to the appointment note when CreateMeetingNotes is enabled
5. Bulk "Save Today's Appointments" must iterate all accounts' calendar folders with DASL filtering, suppress dialogs, and show summary
6. ContactService must resolve SMTP addresses from Recipient objects and build wikilinks filtered by OlMeetingRecipientType (Required, Optional, Resource)
7. TemplateService must support AppointmentTemplateContext and MeetingNoteTemplateContext with default templates
8. ObsidianSettings must include all appointment-specific properties (AppointmentsFolder, title format, tags, meeting notes, recurring grouping, cancelled filtering, task creation mode)
9. SlingRibbon must add an "Appointments" group with "Save Today's Appointments" button and an inspector-level Sling button via `GetCustomUI("Microsoft.Outlook.Appointment")`
10. SettingsForm must be a tabbed interface with tabs: General, Email, Appointments, Contacts, Tasks, Threading, Attachments, Developer
11. Appointments must support task creation (None/Obsidian/Outlook/Both) via existing TaskService
12. Duplicate detection must use GlobalAppointmentID as cache key
13. `.ics` attachments must be filtered out by default
14. Cancelled appointments must be filterable via SaveCancelledAppointments setting (default: false)
15. All new components must have xUnit tests with Moq for COM object mocking

## Sub-Specs

### Sub-Spec 1: ObsidianSettings Extensions + AppointmentProcessor Core

**Scope:** Add appointment-specific properties to ObsidianSettings. Create AppointmentProcessor as a sibling to EmailProcessor with single-appointment processing.

**Files:**
- `SlingMD.Outlook/Models/ObsidianSettings.cs` -- add new properties, `GetAppointmentsPath()` method, update validation
- `SlingMD.Outlook/Services/AppointmentProcessor.cs` -- NEW, ~800-1000 lines

**Implementation Details:**

New ObsidianSettings properties:
- `AppointmentsFolder` (string, default: `"Appointments"`)
- `AppointmentNoteTitleFormat` (string, default: `"{Date} - {Subject}"`)
- `AppointmentNoteTitleMaxLength` (int, default: 50)
- `AppointmentDefaultNoteTags` (List<string>, default: `["Appointment"]`)
- `AppointmentSaveAttachments` (bool, default: true)
- `CreateMeetingNotes` (bool, default: true)
- `MeetingNoteTemplate` (string, optional custom template path)
- `GroupRecurringMeetings` (bool, default: true)
- `SaveCancelledAppointments` (bool, default: false)
- `AppointmentTaskCreation` (string enum: "None"/"Obsidian"/"Outlook"/"Both", default: "None")

New method: `GetAppointmentsPath()` -- returns `Path.Combine(VaultBasePath, VaultName, AppointmentsFolder)`

AppointmentProcessor pattern (mirror EmailProcessor):
- Constructor: `public AppointmentProcessor(ObsidianSettings settings)` -- creates FileService, TemplateService, ThreadService, TaskService, ContactService, AttachmentService, StatusService internally
- Main method: `public async Task ProcessAppointment(AppointmentItem appointment, bool bulkMode = false, CancellationToken cancellationToken = default)`
- Extract metadata: Subject, HTMLBody, Location, Start, End, GetOrganizer(), Recipients (by OlMeetingRecipientType), RecurrenceState, GlobalAppointmentID, Attachments
- Clean subject using same compiled regex patterns as EmailProcessor
- Apply AppointmentNoteTitleFormat with `{Date}`, `{Subject}`, `{Sender}` placeholders
- Duplicate detection via `ConcurrentDictionary<string, byte>` keyed on GlobalAppointmentID
- Build frontmatter dictionary, convert HTMLBody to markdown, render via TemplateService
- Write note via FileService
- Process attachments via AttachmentService (filter out `.ics` by extension)
- COM object cleanup matching EmailProcessor patterns

**Acceptance Criteria:**
- `[MECHANICAL]` `dotnet build SlingMD.sln --configuration Release` compiles clean
- `[STRUCTURAL]` `AppointmentProcessor.cs` exists in `SlingMD.Outlook/Services/` with `ProcessAppointment()` method
- `[STRUCTURAL]` ObsidianSettings has all 10 new properties with defaults, `GetAppointmentsPath()` method
- `[BEHAVIORAL]` ObsidianSettings round-trips new properties through Save/Load without data loss

**Dependencies:** none

---

### Sub-Spec 2: ContactService Meeting Extensions + TemplateService Appointment Contexts

**Scope:** Extend ContactService with meeting-role-aware methods. Add AppointmentTemplateContext and MeetingNoteTemplateContext to TemplateService with default templates.

**Files:**
- `SlingMD.Outlook/Services/ContactService.cs` -- add new methods (additive only)
- `SlingMD.Outlook/Services/TemplateService.cs` -- add new context classes and render methods
- `SlingMD.Outlook/Templates/AppointmentTemplate.md` -- NEW default template
- `SlingMD.Outlook/Templates/MeetingNoteTemplate.md` -- NEW default template

**Implementation Details:**

ContactService new methods (additive, do not modify existing methods):
- `GetSMTPEmailAddress(Recipient recipient)` -- resolve SMTP from a Recipient object (similar pattern to existing `GetSenderEmail(MailItem)` but for Recipient)
- `BuildLinkedNames(Recipients recipients, params OlMeetingRecipientType[] types)` -- overload that filters by meeting role (olRequired, olOptional, olResource)
- `BuildEmailList(Recipients recipients, IEnumerable<OlMeetingRecipientType> types)` -- emails by meeting role
- `GetMeetingResourceData(Recipients recipients)` -- extract conference rooms/equipment from olResource recipients

TemplateService additions:
- `AppointmentTemplateContext` nested class: organizer, attendees, optionalAttendees, resources, recurrence, startDateTime, endDateTime, location, plus inherited fields (metadata, title, body, taskBlock)
- `MeetingNoteTemplateContext` nested class: appointmentTitle, appointmentLink, organizer, attendees, date, location
- `RenderAppointmentContent(AppointmentTemplateContext context)` -- renders appointment note
- `RenderMeetingNoteContent(MeetingNoteTemplateContext context)` -- renders companion meeting note stub
- `GetDefaultAppointmentTemplate()` / `GetDefaultMeetingNoteTemplate()` -- embedded defaults

Default AppointmentTemplate.md should include: frontmatter section, attendee list, body content, task block -- matching the style of EmailTemplate.md.

Default MeetingNoteTemplate.md should include: frontmatter with backlink to appointment note, attendee list, empty sections for Agenda, Notes, Action Items.

**Acceptance Criteria:**
- `[MECHANICAL]` Build succeeds with no errors
- `[STRUCTURAL]` ContactService has 4 new methods accepting `OlMeetingRecipientType` parameters
- `[STRUCTURAL]` TemplateService has `AppointmentTemplateContext`, `MeetingNoteTemplateContext` classes and `RenderAppointmentContent()`, `RenderMeetingNoteContent()` methods
- `[STRUCTURAL]` `AppointmentTemplate.md` and `MeetingNoteTemplate.md` exist in `Templates/`

**Dependencies:** Sub-Spec 1

---

### Sub-Spec 3: Recurring Meeting Threading

**Scope:** Wire AppointmentProcessor to use ThreadService for recurring meeting instances. Series folder keyed by RecurrencePattern + cleaned subject. Instance notes date-stamped inside folder.

**Files:**
- `SlingMD.Outlook/Services/AppointmentProcessor.cs` -- add recurring detection and ThreadService routing

**Implementation Details:**

In `ProcessAppointment()`, after metadata extraction:
1. Check `RecurrenceState` -- if `olApptOccurrence` and `GroupRecurringMeetings` enabled:
   - Generate thread folder name from `RecurrencePattern.PatternStartDate` + cleaned subject
   - Use ThreadService to create/get thread folder under AppointmentsFolder
   - Write instance note inside thread folder with date-stamped filename (e.g., `2026-03-13 - Weekly Standup.md`)
   - Update thread summary note (0-threadname.md) with timeline entry
2. If not recurring or GroupRecurringMeetings disabled: write flat to AppointmentsFolder
3. If user slings a series master (RecurrenceState == olApptMaster): show warning via MessageBox, offer to process next upcoming instance only

Edge cases:
- Always check RecurrenceState before accessing RecurrencePattern (COM can throw)
- Deleted instances from series: catch COMException from GetOccurrence(), skip in bulk mode
- Filename collision for same-day instances: append time suffix (`_HHmm`)

**Acceptance Criteria:**
- `[MECHANICAL]` Build succeeds
- `[STRUCTURAL]` AppointmentProcessor checks `RecurrenceState` and routes to ThreadService when recurring
- `[BEHAVIORAL]` Recurring instance creates note inside thread folder, not flat in AppointmentsFolder

**Dependencies:** Sub-Spec 1

---

### Sub-Spec 4: Companion Meeting Notes

**Scope:** Generate meeting notes stubs with backlinks when CreateMeetingNotes is enabled. Never overwrite existing user content.

**Files:**
- `SlingMD.Outlook/Services/AppointmentProcessor.cs` -- add meeting notes generation step

**Implementation Details:**

After writing the appointment note (step 10 in design data flow):
1. If `CreateMeetingNotes` is enabled:
   - Generate meeting note filename: `{Date} - {Subject} - Meeting Notes.md`
   - Check if file already exists -- if yes, skip creation but ensure appointment note frontmatter links to it
   - If doesn't exist: render via `TemplateService.RenderMeetingNoteContent()` with backlink to appointment note
   - Write meeting note adjacent to appointment note (same folder, or inside thread folder if recurring)
2. If custom `MeetingNoteTemplate` path is set in settings, load that template instead of default

**Acceptance Criteria:**
- `[MECHANICAL]` Build succeeds
- `[STRUCTURAL]` AppointmentProcessor has meeting notes generation logic gated by `CreateMeetingNotes` setting
- `[BEHAVIORAL]` Processing an appointment with CreateMeetingNotes=true creates two files: appointment note + meeting notes stub
- `[BEHAVIORAL]` Processing again does NOT overwrite existing meeting notes file

**Dependencies:** Sub-Spec 1, Sub-Spec 2

---

### Sub-Spec 5: Bulk "Save Today's Appointments"

**Scope:** Implement SaveTodaysAppointments() in ThisAddIn and bulk mode in AppointmentProcessor. DASL filtering, cancelled filtering, error collection, summary dialog.

**Files:**
- `SlingMD.Outlook/ThisAddIn.cs` -- add `SaveTodaysAppointments()` method
- `SlingMD.Outlook/Services/AppointmentProcessor.cs` -- ensure bulkMode flag works (suppress dialogs, suppress Obsidian launch, collect errors)

**Implementation Details:**

ThisAddIn.SaveTodaysAppointments():
1. Iterate all Outlook accounts via `Application.Session.Accounts`
2. For each account, get default Calendar folder
3. Set `Items.IncludeRecurrences = true`
4. DASL Restrict to today's date range: `[Start] >= '{today 00:00}' AND [End] <= '{today+1 00:00}'`
5. Filter out cancelled appointments if `SaveCancelledAppointments == false` (check `MeetingStatus == olMeetingCanceled`)
6. Process each appointment via `_appointmentProcessor.ProcessAppointment(appt, bulkMode: true)`
7. Catch COMException per-account, continue with remaining accounts
8. Show summary: "Saved X/Y appointments. Z skipped (duplicates/cancelled). Errors: N" with optional detail view
9. Optional single Obsidian launch to AppointmentsFolder after all processing

AppointmentProcessor bulk mode behavior:
- Suppress TaskOptionsForm dialog
- Suppress CountdownForm
- Suppress individual Obsidian launches
- Collect errors in a `List<string>` instead of showing MessageBox per error
- Return processing result (success/skip/error) for summary counting

**Acceptance Criteria:**
- `[MECHANICAL]` Build succeeds
- `[STRUCTURAL]` ThisAddIn has `SaveTodaysAppointments()` method
- `[STRUCTURAL]` AppointmentProcessor.ProcessAppointment has `bulkMode` parameter
- `[BEHAVIORAL]` Bulk mode suppresses all dialogs and collects errors for summary

**Dependencies:** Sub-Spec 1

---

### Sub-Spec 6: Ribbon Extensions

**Scope:** Add "Appointments" group to the ribbon with "Save Today's Appointments" button. Add inspector-level Sling button for open appointments. Update Explorer Sling button to handle both MailItem and AppointmentItem.

**Files:**
- `SlingMD.Outlook/Ribbon/SlingRibbon.xml` -- add Appointments group, inspector customUI
- `SlingMD.Outlook/Ribbon/SlingRibbon.cs` -- add button click handlers, `GetCustomUI` routing
- `SlingMD.Outlook/ThisAddIn.cs` -- add `ProcessSelection()` replacing `ProcessSelectedEmail()`, add `ProcessCurrentAppointment()`

**Implementation Details:**

SlingRibbon.xml changes:
- Add "Appointments" group with "Save Today's Appointments" button (imageMso appropriate calendar icon)
- The existing "Email" group Sling button should work for both emails and appointments (routing handled in ThisAddIn)

SlingRibbon.cs changes:
- `GetCustomUI(string ribbonID)`: if ribbonID is `"Microsoft.Outlook.Appointment"`, return inspector ribbon XML with a Sling button
- `OnSaveTodaysClick(IRibbonControl)` -- calls `_addIn.SaveTodaysAppointments()`
- `OnInspectorSlingClick(IRibbonControl)` -- calls `_addIn.ProcessCurrentAppointment()`

ThisAddIn changes:
- Add `_appointmentProcessor` field (initialized alongside `_emailProcessor`)
- `ProcessSelection()` replaces `ProcessSelectedEmail()`:
  - Get selected item from Explorer
  - If MailItem: route to `_emailProcessor.ProcessEmail()`
  - If AppointmentItem: route to `_appointmentProcessor.ProcessAppointment()`
  - Else: show "Please select an email or appointment"
- `ProcessCurrentAppointment()`:
  - Get AppointmentItem from active Inspector
  - Check `appointment.Saved` -- if unsaved, prompt "Save changes before slinging?"
  - Route to `_appointmentProcessor.ProcessAppointment()`
- Update `OnSlingButtonClick` to call `ProcessSelection()` instead of `ProcessSelectedEmail()`

**Acceptance Criteria:**
- `[MECHANICAL]` Build succeeds
- `[STRUCTURAL]` SlingRibbon.xml has "Appointments" group with "Save Today's Appointments" button
- `[STRUCTURAL]` SlingRibbon.cs handles `GetCustomUI("Microsoft.Outlook.Appointment")` for inspector
- `[STRUCTURAL]` ThisAddIn has `ProcessSelection()`, `ProcessCurrentAppointment()`, `SaveTodaysAppointments()`, and `_appointmentProcessor` field

**Dependencies:** Sub-Spec 1, Sub-Spec 5

---

### Sub-Spec 7: Tabbed SettingsForm Rewrite

**Scope:** Convert SettingsForm to a tabbed interface organizing all settings (existing + new appointment settings) into logical tabs.

**Files:**
- `SlingMD.Outlook/Forms/SettingsForm.cs` -- rewrite as tabbed form
- `SlingMD.Outlook/Forms/SettingsForm.Designer.cs` -- updated designer file

**Implementation Details:**

Note: SettingsForm already uses a TabControl. This sub-spec reorganizes existing tabs and adds the Appointments tab.

Tab structure:
- **General**: Vault path, vault name, launch Obsidian, countdown delay, templates folder
- **Email**: Inbox folder, note title format (`NoteTitleFormat`), title max length, include date in title, default note tags, subject cleanup patterns
- **Appointments** (NEW): Appointments folder, note title format, max length, default tags, save attachments, create meeting notes, meeting note template path, group recurring meetings, save cancelled appointments, task creation mode (dropdown: None/Obsidian/Outlook/Both)
- **Contacts**: Contacts folder, enable contact saving, vault-wide search
- **Tasks**: Create Obsidian task, create Outlook task, ask for dates, default due days, default reminder days/hour, task tags
- **Threading**: Group email threads, thread debug settings
- **Attachments**: Save inline images, save all attachments, use Obsidian wikilinks, attachment storage mode, attachments folder
- **Developer**: Show development settings, show thread debug, other advanced options

Each tab binds to ObsidianSettings properties. On Save, validate all tabs, persist to settings. Preserve existing behavior -- no functional changes to how settings work, just reorganize the UI.

**Acceptance Criteria:**
- `[MECHANICAL]` Build succeeds
- `[STRUCTURAL]` SettingsForm has 8 tabs: General, Email, Appointments, Contacts, Tasks, Threading, Attachments, Developer
- `[BEHAVIORAL]` All existing email settings still save/load correctly after UI rewrite
- `[BEHAVIORAL]` New appointment settings save/load correctly
- `[HUMAN REVIEW]` Tab organization is logical and settings are easy to find

**Dependencies:** Sub-Spec 1

---

### Sub-Spec 8: Task Creation Integration for Appointments

**Scope:** Wire AppointmentProcessor to use TaskService for follow-up task creation based on AppointmentTaskCreation setting.

**Files:**
- `SlingMD.Outlook/Services/AppointmentProcessor.cs` -- add task creation step

**Implementation Details:**

In `ProcessAppointment()`, after contact processing (step 14 in design data flow):
1. If `AppointmentTaskCreation != "None"`:
   - If not bulkMode: show TaskOptionsForm to get due/reminder dates
   - If bulkMode: use default due/reminder values from settings
   - If `AppointmentTaskCreation` includes "Obsidian": call TaskService to create Obsidian task line in the appointment note
   - If `AppointmentTaskCreation` includes "Outlook": call TaskService to create Outlook task item
2. Task context: link to appointment note, use appointment subject as task name, use AppointmentDefaultNoteTags

Mirror the exact pattern from EmailProcessor's task creation flow.

**Acceptance Criteria:**
- `[MECHANICAL]` Build succeeds
- `[STRUCTURAL]` AppointmentProcessor has task creation logic gated by `AppointmentTaskCreation` setting
- `[BEHAVIORAL]` Setting AppointmentTaskCreation to "Obsidian" creates a task line in the note
- `[BEHAVIORAL]` Bulk mode skips TaskOptionsForm and uses defaults

**Dependencies:** Sub-Spec 1

---

### Sub-Spec 9: Tests

**Scope:** xUnit tests with Moq for all new components. Mock COM objects (AppointmentItem, Recipients, Recipient).

**Files:**
- `SlingMD.Tests/Services/AppointmentProcessorTests.cs` -- NEW
- `SlingMD.Tests/Services/ContactServiceMeetingTests.cs` -- NEW
- `SlingMD.Tests/Services/TemplateServiceAppointmentTests.cs` -- NEW
- `SlingMD.Tests/Models/ObsidianSettingsAppointmentTests.cs` -- NEW

**Implementation Details:**

Test categories:

**ObsidianSettings tests:**
- Default values for all 10 new properties
- Round-trip serialization (Save + Load preserves all appointment settings)
- `GetAppointmentsPath()` returns correct combined path
- Validation of AppointmentNoteTitleMaxLength range
- AppointmentTaskCreation enum values

**AppointmentProcessor tests:**
- Single appointment processing creates markdown file with expected frontmatter
- Duplicate detection skips already-processed GlobalAppointmentID
- Subject cleaning applies same regex patterns as email
- `.ics` attachments are filtered out
- BulkMode flag suppresses Obsidian launch
- Cancelled appointment filtered when SaveCancelledAppointments=false
- Filename collision appends time suffix

**ContactService meeting tests:**
- `GetSMTPEmailAddress()` resolves from Recipient object
- `BuildLinkedNames()` filters by OlMeetingRecipientType correctly
- `BuildEmailList()` returns emails for specified meeting roles
- `GetMeetingResourceData()` extracts resource recipients

**TemplateService appointment tests:**
- `RenderAppointmentContent()` produces expected markdown with all context fields
- `RenderMeetingNoteContent()` produces stub with backlink
- Default appointment template loads successfully
- Default meeting note template loads successfully

COM mocking pattern: Use Moq to create `Mock<AppointmentItem>`, `Mock<Recipients>`, `Mock<Recipient>` with SetupGet for properties. Follow existing test patterns in the project.

**Acceptance Criteria:**
- `[MECHANICAL]` `dotnet test SlingMD.Tests\SlingMD.Tests.csproj` passes all new tests
- `[STRUCTURAL]` 4 new test files exist in SlingMD.Tests
- `[BEHAVIORAL]` Tests cover: settings defaults, settings serialization, appointment processing, duplicate detection, contact meeting methods, template rendering

**Dependencies:** Sub-Spec 1, Sub-Spec 2, Sub-Spec 3, Sub-Spec 4, Sub-Spec 5, Sub-Spec 8

---

## Edge Cases

1. **COM object access failures** -- Wrap each AppointmentItem property access in try/catch, degrade gracefully (e.g., "Untitled Appointment" if Subject fails). Match EmailProcessor's COM cleanup patterns.
2. **Recurring instance deleted from series** -- `GetOccurrence()` throws COMException. Catch and skip in bulk mode, show error in single mode.
3. **Series master sling** -- User slings the master recurring appointment, not an instance. Warn and offer to process next upcoming instance only.
4. **Duplicate detection race** -- Use GlobalAppointmentID as cache key (not filename). For filename collisions on same-day recurring instances, append `_HHmm` time suffix.
5. **Meeting notes stub conflict** -- Never overwrite existing meeting notes file. If file exists, just ensure appointment note frontmatter links to it.
6. **Attachment access failure** -- Filter out `.ics` by default. Wrap each `SaveAsFile` in try/catch (AttachmentService already does this). Log failures but don't block note creation.
7. **Calendar folder access failure (bulk)** -- Catch COMException per-account in SaveTodaysAppointments(). Report which accounts succeeded/failed in summary.
8. **Inspector unsaved changes** -- Check `appointment.Saved` before processing. If unsaved: prompt "Save appointment changes before slinging?"
9. **Empty appointment body** -- Handle null/empty HTMLBody gracefully, produce note with frontmatter only.
10. **No attendees** -- Handle appointments with no recipients (personal events). Omit attendee frontmatter fields.

## Out of Scope

- Refactoring EmailProcessor or extracting a shared base class (Approach B from design doc)
- Item adapter abstraction layer (Approach C from design doc)
- Calendar sync or two-way appointment updates
- Processing other Outlook item types (tasks, contacts, journal entries)
- Obsidian plugin development
- Cloud/sync features
- Mobile/web interface

## Constraints

### Musts
- .NET Framework 4.7.2 (VSTO requirement, cannot upgrade)
- Follow CLAUDE.md code style: PascalCase, explicit typing, braces on new lines, 4-space indent
- Use fully qualified `System.Exception`, catch specific exceptions
- Services must have 'Service' suffix
- COM objects must be properly released to prevent memory leaks
- AppointmentProcessor must be a sibling to EmailProcessor, not a subclass or refactor

### Must-Nots
- Must NOT modify existing EmailProcessor behavior or method signatures
- Must NOT modify existing ContactService/TemplateService method signatures (additive only)
- Must NOT break existing email processing pipeline
- Must NOT overwrite user-authored meeting notes content
- Must NOT store unfiltered `.ics` attachments by default

### Preferences
- Prefer matching EmailProcessor patterns over cleaner alternatives
- Prefer additive changes over refactoring shared code
- Prefer explicit COM object cleanup over relying on GC
- Prefer DASL filtering over client-side filtering for bulk performance

### Escalation Triggers
- Any change that would modify EmailProcessor.cs
- Any modification to existing shared service method signatures
- Any architectural pattern that significantly diverges from EmailProcessor
- Adding new NuGet dependencies
- Changes to the VSTO startup/shutdown lifecycle

## Verification

**End-to-end verification:**
1. `[MECHANICAL]` `dotnet build SlingMD.sln --configuration Release` -- zero errors
2. `[BEHAVIORAL]` `dotnet test SlingMD.Tests\SlingMD.Tests.csproj` -- all tests pass
3. `[STRUCTURAL]` New files exist: `AppointmentProcessor.cs`, `AppointmentTemplate.md`, `MeetingNoteTemplate.md`, 4 test files
4. `[STRUCTURAL]` ObsidianSettings has 10 new appointment properties with correct defaults
5. `[STRUCTURAL]` SettingsForm has 8 tabs including Appointments
6. `[STRUCTURAL]` SlingRibbon.xml has Appointments group
7. `[HUMAN REVIEW]` Install add-in in Outlook, sling an appointment, verify markdown note appears in vault with correct frontmatter, attendee wikilinks, and converted body

## Phase Specs

Refined by `/forge-prep` on 2026-03-13.

| Sub-Spec | Phase Spec |
|----------|------------|
| 1. ObsidianSettings Extensions + AppointmentProcessor Core | `docs/specs/appointment-processing/sub-spec-1-settings-processor-core.md` |
| 2. ContactService Meeting Extensions + TemplateService Appointment Contexts | `docs/specs/appointment-processing/sub-spec-2-contact-template-extensions.md` |
| 3. Recurring Meeting Threading | `docs/specs/appointment-processing/sub-spec-3-recurring-meeting-threading.md` |
| 4. Companion Meeting Notes | `docs/specs/appointment-processing/sub-spec-4-companion-meeting-notes.md` |
| 5. Bulk "Save Today's Appointments" | `docs/specs/appointment-processing/sub-spec-5-bulk-save-today.md` |
| 6. Ribbon Extensions | `docs/specs/appointment-processing/sub-spec-6-ribbon-extensions.md` |
| 7. Tabbed SettingsForm Rewrite | `docs/specs/appointment-processing/sub-spec-7-tabbed-settings-form.md` |
| 8. Task Creation Integration for Appointments | `docs/specs/appointment-processing/sub-spec-8-task-creation.md` |
| 9. Tests | `docs/specs/appointment-processing/sub-spec-9-tests.md` |

Index: `docs/specs/appointment-processing/index.md`
