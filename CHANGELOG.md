# Changelog

All notable changes to SlingMD are documented in this file.

## [1.1.0.1] - 2026-03-13

### Added
- **Appointment Processing** — full pipeline for exporting Outlook calendar appointments to Obsidian markdown notes, mirroring the existing email export flow
- **Appointment Ribbon Integration** — "Save Today's Appointments" button in the Sling ribbon for bulk-exporting all of today's calendar items; "Sling" button in the appointment inspector for single-item export
- **Recurring Meeting Threading** — recurring meeting instances are automatically grouped into thread folders with summary notes, reusing the email threading pattern
- **Companion Meeting Notes** — optional blank meeting note created alongside the appointment note for capturing real-time meeting notes, linked bidirectionally
- **Appointment Task Creation** — configurable task creation for appointments (None / Obsidian / Outlook / Both), using the same TaskService pipeline as emails
- **Appointment Templates** — dedicated `AppointmentTemplate.md` and `MeetingNoteTemplate.md` with full template variable support (attendees, location, recurrence, resources, etc.)
- **ContactService Meeting Extensions** — `BuildLinkedNames`, `BuildEmailList`, and `GetMeetingResourceData` methods for extracting attendee information from appointments
- **TemplateService Appointment Support** — `AppointmentTemplateContext` (17 properties) and `MeetingNoteTemplateContext` for rendering appointment and meeting note content
- **Tabbed Settings Form** — settings dialog reorganized from a single scrollable form into 8 focused tabs (General, Email, Appointments, Contacts, Tasks, Threading, Attachments, Developer)
- 10 new appointment-related settings: `AppointmentsFolder`, `AppointmentNoteTitleFormat`, `AppointmentNoteTitleMaxLength`, `AppointmentDefaultNoteTags`, `AppointmentSaveAttachments`, `CreateMeetingNotes`, `MeetingNoteTemplate`, `GroupRecurringMeetings`, `SaveCancelledAppointments`, `AppointmentTaskCreation`
- 13 new tests across 4 test files covering appointment settings validation, processor logic, contact meeting methods, and template rendering (75/75 total tests passing)

### Fixed
- Settings form last row on every tab stretched to fill remaining space — added explicit `AutoSize` RowStyles with a `Percent` filler row to all tab `TableLayoutPanel` layouts
- Settings form displayed the default Windows icon — now loads `SlingMD.ico` from embedded resources for the title bar

## [1.0.0.124] - 2026-03-13

### Fixed
- Corrupt or malformed `ObsidianSettings.json` no longer crashes the add-in on startup; safe defaults are loaded instead ([#6](https://github.com/Caleb68864/SlingMD/issues/6))
- Settings now persist correctly across Outlook restarts and Sling operations ([#6](https://github.com/Caleb68864/SlingMD/issues/6))
- Fatal export errors no longer fall through into contact creation or Obsidian launch
- Canceling the task-options dialog no longer disables task creation for the rest of the Outlook session
- Thread summary notes (0-file) DataviewJS query now correctly scopes to the thread folder instead of producing a `SyntaxError: Invalid or unexpected token`
- Thread discovery now parses both second-precision (`yyyy-MM-dd HH:mm:ss`) and legacy minute-precision (`yyyy-MM-dd HH:mm`) date formats
- Frontmatter is now YAML-safe when email metadata contains double quotes, backslashes, or embedded newlines
- Attachment links now resolve correctly for same-folder, per-note-subfolder, and centralized attachment storage modes
- Exporting into a fresh vault with a missing inbox folder no longer throws during duplicate detection or cache initialization
- Contact notes now use `## Communication History` heading with a working DataviewJS query for email history ([#4](https://github.com/Caleb68864/SlingMD/issues/4))

### Added
- Customizable markdown templates for email notes, contact notes, task lines, and thread summaries via Settings ([#8](https://github.com/Caleb68864/SlingMD/issues/8))
- Regression test coverage for corrupt-settings fallback, task-state reset, missing inbox handling, thread date compatibility, frontmatter escaping, and attachment-link generation
- VSTO build/test prerequisite documentation in README

## [1.0.0.121] - 2025-12-15

### Fixed
- Outlook settings persistence improvements
- Contact history heading alignment

## [1.0.0.44] - 2025-03-15

### Added
- Automatic email thread detection and organization
- Thread summary pages with timeline views
- Configurable subject cleanup patterns
- Thread folder creation for related emails
- Participant tracking in thread summaries
- Dataview integration for thread visualization

### Improved
- Email relationship detection
- Thread navigation with bidirectional links

## [1.0.0.14] - 2025-02-01

### Added
- Follow-up task creation in Obsidian notes
- Follow-up task creation in Outlook
- Configurable due dates and reminder times
- Task options dialog for custom timing

## [1.0.0.8] - 2025-01-15

### Added
- Initial release
- Email to Obsidian note conversion
- Email metadata preservation
- Obsidian vault configuration
- Launch delay settings
