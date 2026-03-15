---
date: 2026-03-13
topic: "Appointment Processing for SlingMD"
author: Caleb Bennett
status: draft
tags:
  - design
  - appointment-processing
---

# Appointment Processing for SlingMD -- Design

## Summary

Add full appointment/calendar processing to SlingMD, allowing users to sling Outlook appointments to Obsidian as markdown notes -- both individually (from Explorer or an open appointment) and in bulk ("Save Today's Appointments"). The feature achieves full parity with email processing: HTML-to-markdown conversion, customizable templates, attachment handling, contact creation, task creation, recurring meeting threading, and companion meeting notes. A new tabbed settings UI organizes the growing configuration surface.

## Approach Selected

**Approach A: Dedicated AppointmentProcessor** -- a parallel peer to EmailProcessor that reuses all existing shared services. Chosen for zero regression risk to the battle-tested email pipeline, clean separation of concerns, and the fork (dlboutwe/SlingMD) providing a working reference to accelerate development.

## Architecture

```
+-------------------------------------------------------------+
|                    Outlook Ribbon                             |
|  +----------+  +------------------+  +-------------------+   |
|  |  Sling   |  | Save Today's     |  |   Configure       |   |
|  |  Button   |  | Appointments     |  |   Button          |   |
|  +----+-----+  +--------+---------+  +-------------------+   |
+-------+-----------------+------------------------------------+
        |                 |
        v                 v
+--------------------------------------+
|           ThisAddIn                   |
|  ProcessSelection()                  |
|    +- MailItem? --> EmailProcessor    |
|    +- AppointmentItem? -->           |
|         AppointmentProcessor         |
|  ProcessCurrentAppointment()         |
|    +- Inspector button -->           |
|         AppointmentProcessor         |
|  SaveTodaysAppointments()            |
|    +- bulk --> AppointmentProcessor  |
+----------+---------------+-----------+
           |               |
    +------v------+  +-----v--------------+
    | Email       |  | Appointment         |
    | Processor   |  | Processor (NEW)     |
    | (existing)  |  |                     |
    +------+------+  +-----+--------------+
           |               |
           v               v
+--------------------------------------+
|        Shared Services Layer         |
|  FileService    | TemplateService    |
|  ContactService | AttachmentService  |
|  ThreadService  | TaskService        |
|  StatusService  |                    |
+--------------------------------------+
           |
           v
+--------------------------------------+
|        ObsidianSettings              |
|  (+ new appointment properties)      |
+--------------------------------------+
```

Key architectural decisions:
- AppointmentProcessor is a sibling to EmailProcessor, not a subclass
- Both processors consume the same shared services layer
- ThisAddIn routes based on selected item type (MailItem vs AppointmentItem)
- Inspector-level Sling button via `GetCustomUI("Microsoft.Outlook.Appointment")` allows slinging from an open appointment
- Recurring meeting series reuse ThreadService infrastructure (series folder + instance notes)

## Components

### AppointmentProcessor (NEW)
**Owns:** Orchestrating the full appointment-to-Obsidian pipeline
**Does NOT own:** File I/O, template rendering, contact management, attachment storage

Responsibilities:
- Extract appointment metadata (organizer, attendees by role, resources, recurrence info, start/end times, location)
- Determine if this is a recurring instance and route to threading if so
- Build frontmatter dictionary with meeting-specific fields
- Build note content via TemplateService (with HTML-to-markdown conversion, same as email)
- Cache-based duplicate detection keyed on GlobalAppointmentID
- Trigger companion meeting notes stub creation (if enabled)
- Handle bulk mode (suppress Obsidian launch, suppress dialogs, collect errors for summary)
- Invoke TaskService for follow-up task creation (if enabled per settings)

### ContactService (EXTENDED -- additive only)
New methods:
- `GetSMTPEmailAddress(Recipient)` -- resolve SMTP from a Recipient object
- `BuildLinkedNames(Recipients, OlMeetingRecipientType[])` -- filter by meeting role
- `BuildEmailList(Recipients, IEnumerable<OlMeetingRecipientType>)` -- emails by meeting role
- `GetMeetingResourceData(Recipients)` -- extract conference rooms/equipment from olResource recipients

### TemplateService (EXTENDED)
New additions:
- `AppointmentTemplateContext` -- context class with: organizer, attendees, optional attendees, resources, recurrence, start/end, location
- `MeetingNoteTemplateContext` -- context for companion meeting notes stub
- `RenderAppointmentContent()` / `RenderMeetingNoteContent()` -- typed render methods
- Default appointment template and meeting note template (loadable/overridable from templates folder, same pattern as email)

### ThreadService (REUSED for recurring meetings)
No changes needed. Existing capability, new usage:
- Recurring meeting series --> thread folder (keyed by RecurrencePattern.PatternStartDate + cleaned subject)
- Each instance --> note inside the folder with date-stamped filename
- Thread summary note --> timeline of all captured instances with Dataview queries

### AttachmentService (REUSED as-is)
- All 3 storage modes (SameAsNote, SubfolderPerNote, Centralized) work unchanged
- `.ics` attachments filtered out by default

### ObsidianSettings (EXTENDED)
New properties:
- `AppointmentsFolder` (default: `"Appointments"`)
- `AppointmentNoteTitleFormat` (default: `"{Date} - {Subject}"`)
- `AppointmentNoteTitleMaxLength` (default: 50)
- `AppointmentDefaultNoteTags` (default: `["Appointment"]`)
- `AppointmentSaveAttachments` (default: true)
- `CreateMeetingNotes` (default: true)
- `MeetingNoteTemplate` (optional custom template path)
- `GroupRecurringMeetings` (default: true)
- `SaveCancelledAppointments` (default: false)
- `AppointmentTaskCreation` (enum: None, Obsidian, Outlook, Both -- default: None)

New method: `GetAppointmentsPath()`

### SlingRibbon (EXTENDED)
- New "Appointments" group with "Save Today's Appointments" button
- Inspector-level Sling button via `GetCustomUI("Microsoft.Outlook.Appointment")`
- Existing Explorer Sling button works for appointments via ProcessSelection()

### ThisAddIn (EXTENDED)
- New `_appointmentProcessor` field
- `ProcessSelection()` replaces `ProcessSelectedEmail()` -- detects MailItem vs AppointmentItem
- `ProcessCurrentAppointment()` for inspector-level button
- `SaveTodaysAppointments()` -- bulk action with DASL filter, handles cancelled appointment filtering per settings

### Settings UI (REWRITTEN as tabbed form)
Convert existing SettingsForm to a tabbed interface:
- **General tab**: Vault path, vault name, launch Obsidian, countdown settings
- **Email tab**: Inbox folder, note title format, title max length, default tags, subject cleanup patterns
- **Appointments tab**: Appointments folder, note title format, max length, save attachments, default tags, create meeting notes, group recurring, save cancelled, task creation mode
- **Contacts tab**: Contacts folder, enable contact saving, vault-wide search
- **Tasks tab**: Task defaults, Obsidian task format, Outlook task creation
- **Threading tab**: Group email threads, thread settings
- **Attachments tab**: Storage mode, attachment settings
- **Developer tab**: Debug/advanced settings

## Data Flow

### Single Appointment Flow
1. User clicks Sling on selected/open appointment
2. ThisAddIn routes to AppointmentProcessor.ProcessAppointment(appointment)
3. Extract metadata from AppointmentItem COM object (subject, body, location, organizer, recipients by role, start/end, recurrence state, attachments)
4. Clean subject, apply AppointmentNoteTitleFormat with {Date}/{Subject}/{Sender} placeholders, truncate
5. If recurring + GroupRecurringMeetings: use ThreadService for series folder, else flat in AppointmentsFolder
6. Duplicate detection via GlobalAppointmentID cache (same pattern as EmailProcessor)
7. Build frontmatter (title, organizer as wikilink, attendees/optional/resources as wikilinks, emails, location, start/end, dailyNoteLink, recurrence, tags, attachment links)
8. Convert HTMLBody to markdown (same as email pipeline), render via TemplateService
9. Write appointment note via FileService
10. If CreateMeetingNotes and stub doesn't exist: write companion meeting notes with backlinks
11. If recurring + GroupRecurringMeetings: update thread summary note
12. Process attachments via AttachmentService (filter out .ics)
13. Contact processing: collect names, filter new, show ContactConfirmationDialog, create notes
14. If AppointmentTaskCreation != None: show TaskOptionsForm, create tasks
15. If not bulk mode: launch Obsidian

### Bulk "Save Today's Appointments" Flow
1. Iterate all accounts' calendar folders
2. DASL Restrict to today's date range with IncludeRecurrences = true
3. Filter out cancelled appointments if SaveCancelledAppointments = false (check appointment.MeetingStatus)
4. Process each via AppointmentProcessor with bulkMode = true (suppresses dialogs + launch)
5. Show summary: "Saved X/Y appointments. Z skipped (duplicates/cancelled)."
6. Optional single Obsidian launch to Appointments folder

### Key Data Transformations
| Source (COM) | Destination (Markdown) |
|---|---|
| appointment.Subject | frontmatter `title`, cleaned filename |
| appointment.GetOrganizer() | frontmatter `organizer: "[[Name]]"` |
| Recipients (olRequired) | frontmatter `attendees: ["[[Name]]", ...]` |
| Recipients (olOptional) | frontmatter `optional: ["[[Name]]", ...]` |
| Recipients (olResource) | frontmatter `resources: ["[[Room]]", ...]` |
| appointment.Location | frontmatter `location` |
| appointment.Start/End | frontmatter `startDateTime/endDateTime` |
| appointment.HTMLBody | note body (HTML --> markdown) |
| RecurrencePattern | frontmatter `recurrence` + thread folder grouping |
| appointment.Attachments | wikilinks in frontmatter + saved files |
| GlobalAppointmentID | duplicate detection key |

## Error Handling

### COM Object Access Failures
- Wrap each property access in try/catch, degrade gracefully (e.g., "Untitled Appointment" if Subject fails)
- Recurring instances accessed via GetOccurrence() can throw if deleted from series -- catch and skip in bulk mode

### Recurring Meeting Edge Cases
- Always check RecurrenceState before accessing RecurrencePattern
- Bulk mode DASL filter returns individual occurrences, not series master
- Single-sling of series master: warn user, offer to process next upcoming instance only

### Duplicate Detection Races
- Use GlobalAppointmentID as cache key (not filename)
- For filename collisions, append time suffix (_HHmm)

### Meeting Notes Stub Conflicts
- Check if file exists before creating -- never overwrite user content
- If exists, just ensure the appointment note's frontmatter links to it

### Attachment Access
- Filter out .ics attachments by default
- Wrap each SaveAsFile in try/catch (AttachmentService already does this)
- Log failures but don't block note creation

### Calendar Folder Access (Bulk Mode)
- Catch COMException per-account in SaveTodaysAppointments()
- Report which accounts succeeded/failed in summary

### Inspector Button Edge Case
- Check appointment.Saved property before processing
- If unsaved changes: prompt "Save appointment changes before slinging?"

### Error Surfacing
- Single mode: MessageBox.Show() for blocking errors
- Bulk mode: Collect errors silently, show summary at end with optional detail view
- Never throw unhandled exceptions -- catch at top level per code style guidelines

## Open Questions

Resolved during design:
1. **HTML conversion** --> Yes, convert HTMLBody to markdown same as email pipeline
2. **Template customization** --> Full template support, same pattern as email (loadable from templates folder)
3. **Cancelled events** --> New setting `SaveCancelledAppointments` (default: false), filter by MeetingStatus in bulk mode
4. **Task creation** --> New setting `AppointmentTaskCreation` with enum (None/Obsidian/Outlook/Both), integrates with existing TaskService
5. **Settings UI** --> Convert to tabbed form, organize all settings into logical tabs (General, Email, Appointments, Contacts, Tasks, Threading, Attachments, Developer)

## Approaches Considered

### Approach A: Dedicated AppointmentProcessor (SELECTED)
Parallel peer to EmailProcessor reusing shared services. Zero regression risk, clean separation, fork provides working reference. Some orchestration duplication accepted as manageable trade-off.

### Approach B: Shared Base Processor + Specialization
Extract common logic into base OutlookItemProcessor class. More DRY but requires refactoring 988 lines of battle-tested EmailProcessor code with regression risk. Better suited if 3+ item types planned.

### Approach C: Composition Pipeline with Item Adapters
IOutlookItemAdapter interface normalizing COM objects, composable processing steps. Most flexible but over-engineered for 2 item types. COM interop differences make adapter layer leaky.

## Next Steps
- [ ] Turn this design into a Forge spec (`/forge docs/plans/2026-03-13-appointment-processing-design.md`)
- [ ] Phase 1: ObsidianSettings extensions + AppointmentProcessor core (single appointment flow)
- [ ] Phase 2: ContactService meeting extensions + TemplateService appointment contexts
- [ ] Phase 3: Recurring meeting threading via ThreadService
- [ ] Phase 4: Companion meeting notes generation
- [ ] Phase 5: Bulk "Save Today's Appointments" with DASL filtering
- [ ] Phase 6: Ribbon extensions (Appointments group + Inspector button)
- [ ] Phase 7: Tabbed SettingsForm rewrite
- [ ] Phase 8: Task creation integration for appointments
- [ ] Phase 9: Tests for all new components
