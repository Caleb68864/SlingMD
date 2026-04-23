# Changelog

All notable changes to SlingMD are documented in this file.

## [Unreleased]

### Added

#### Searchable settings help + hover tooltips
- **New `Help` button** in the Settings dialog footer opens a pop-out `HelpForm` with a live search box, a TreeView grouped by tab, and a detail pane that formats Title / Summary / Description / Default / Tokens / Examples per entry. Uses the SlingMD icon in the title bar; `Esc` clears the search, `Down` focuses the tree.
- **ⓘ indicator glyph** on every labelled setting (and many checkboxes) as a visual signal that hover-help exists.
- **Rich hover tooltips** on every setting pulled from a single `SettingsHelp` registry — so the tooltip and the manual entry are guaranteed to stay in sync. Tooltip display time bumped from 5 s → 30 s to give time to read the Tokens/Examples tables.
- **`SettingsHelp` registry** (`Forms/SettingsHelp.cs`) — one `HelpEntry` per setting with Title, Summary, Description, optional Tokens dictionary, optional Examples list, Default. Adding a new setting means one new entry + one `BindHelp(id, label, ...)` call and it appears in both surfaces automatically.

#### Complete Thread icon
- Programmatically-composed button icon (Sling logo + green-check badge) rendered via `getImage` rather than an unreliable `imageMso` reference. Cached at startup and disposed with the ribbon; falls back to the bare Sling logo if compositing fails.

#### Maintainability review (sprint)
- **Continuous integration** — `.github/workflows/ci.yml` runs MSBuild + vstest on `windows-latest` for every push and PR to `main`.
- **AutoSlingService eligibility is now testable** — extracted `AutoSlingService.EvaluateEligibility(enableAutoSling, isShuttingDown, isAlreadyProcessed, currentUserAddress, snapshot, rules, decisionEngine)` static method returning an `AutoSlingEligibility` enum (`Disabled` / `AlreadyProcessed` / `ShuttingDown` / `SelfSend` / `NoMatch` / `Sling`). 10 new unit tests cover the guard chain.
- **`SubjectFilenameCleaner`** (`Services/SubjectFilenameCleaner.cs`) — composes the three-step subject pipeline (`SubjectCleanerService.Clean` → `FilenameSubjectNormalizer.Normalize` → `FileService.CleanFileName`) that used to be duplicated inline as private `CleanSubject` helpers in both `EmailProcessor` and `AppointmentProcessor`. 6 new tests.
- **`NoteTitleBuilder.BuildTrimmed`** — wraps `Build` + trailing-dash-and-space strip, so both processors can stop re-implementing the `TrailingDashSpaceRegex` post-pass. 3 new tests.
- **`Infrastructure/MapiPropertyTags.cs`** — single home for the four MAPI property-tag URIs (`PrSmtpAddress`, `PrConversationIndex`, `PrAttachContentId`, `PrInternetMessageId`) previously scattered as duplicate `const string`s across `ContactService`, `ThreadService`, `AttachmentService`, and `EmailProcessor`.
- **`IClock` now threads through every date-touching service** — `EmailProcessor` (cache TTL), `ContactService` (contact-note `created`), `AppointmentProcessor` (COM fallback), `AttachmentService` (year-month folder). Completes the testability story advertised by the previous refactor. Constructors accept an optional `IClock clock = null` that defaults to `SystemClock`.
- **`EmailProcessor` internal injection-seam constructor** accepting `FileService`, `TemplateService`, `ThreadService`, `TaskService`, `ContactService`, `AttachmentService`, `IClock`. Any null argument falls back to production wiring; tests can now substitute a single collaborator without subclassing the orchestrator.
- **`FlagMonitorService` and `FolderMonitorService` implement `IDisposable`**, delegating to their existing `Stop`/`StopWatching` methods so they can participate in `using` blocks and signal ownership of Outlook COM handles to static analyzers.
- **`TemplateService.LoadTemplate` is mtime-cached** — keyed on `(path, lastWriteTimeUtc)`, invalidates automatically on user edits to templates. Replaces `File.ReadAllText` on every render call (4+ reads per sling → at most 4 reads per template file).
- **Round-trip test** for `ContactLinkFormat` / `EmailDateFormat` / `ContactDateFormat` / `AppointmentDateFormat` proving each new customization setting survives `Save` → `Load`, plus a normalization test confirming blank values restore defaults.
- **Named constants** in `EmailProcessor`: `CacheTtlMinutes`, `MaxFileWaitAttempts`, `InitialFileWaitDelayMs` replace three unrelated literal `5`s and a bare `50`.
- **Logging on COM-read failures in `ContactService`** — 15 previously silent `catch (System.Exception) { }` blocks now emit `Logger.Warning` with the field name, making it possible to diagnose "my phone numbers stopped exporting" without attaching a debugger.

### Changed
- **`Services/Formatting/*` helpers marked `internal`** — all 14 pure helpers (`ContactLinkFormatter`, `ContactNameParser`, `DateFormatter`, `EmailAddressParser`, `FileNameSanitizer`, `FilenameSubjectNormalizer`, `FrontmatterReader`, `LegacyFilenameStripper`, `MarkdownSectionFinder`, `NoteTitleBuilder`, `SubjectCleanerService`, `TemplatePathResolver`, `ThreadIdHasher`, `UniqueFilenameResolver`) accurately reflect their project-internal scope now. Tests continue to compile via `InternalsVisibleTo`.
- `ContactService.GetShortName` → `GetFilenameSafeShortName` so it no longer name-collides with `ContactName.ShortName` (which is a semantic short name, not a filename-safe one).

### Fixed
- **`Complete Thread` returned every email in the inbox** instead of just the current thread — `ThreadCompletionService.CollectMissingFromFolder` had lost its per-item conversation-ID check. `FindMissingEmails` now requires a `Func<MailItem, string> getConversationId` delegate (the caller passes `ThreadService.GetConversationId`) and each item's computed ID is compared against the target thread.
- **`Complete Thread` ribbon button rendered without an icon** — `imageMso="ConversationSettings"` and `"TaskMarkComplete"` both failed to resolve on real Outlook installs. Replaced with a custom composited Sling + green-check badge delivered via `getImage`.
- **Signing key is no longer committed to the repo** — `*.pfx` added to `.gitignore` and the `SlingMD.Outlook_TemporaryKey.pfx` previously tracked in `SlingMD.Outlook/` was removed from tracking. Regenerate a fresh key locally via *Visual Studio → Project → Signing*.
- **Help button alignment** — `FlowLayoutPanel.WrapContents` set to `false` in the Settings footer so Save / Cancel / Help stay on one row, and `Help`'s `Margin` matches Save/Cancel's default padding so all three buttons share the same vertical baseline.

---

### Earlier on this branch

#### Added — customization settings (user-facing)
- **`ContactLinkFormat`** — format string that controls how `{{to}}`, `{{from}}`, `{{cc}}` recipients and appointment organizer/attendees are rendered. Default `"[[{FullName}]]"` preserves today's wikilink output; users can switch to `"@{FirstName}{LastName}"` for `@JohnSmith` At-People plugin mentions, `"@{FirstInitial}{LastInitial}"` for initials, etc. Tokens supported: `{FullName}`, `{FirstName}`, `{LastName}`, `{MiddleName}`, `{Suffix}`, `{DisplayName}`, `{ShortName}`, `{Email}`, `{FirstInitial}`, `{LastInitial}`.
- **`EmailDateFormat`** — .NET format string for `{{timestamp}}` and the email `date` frontmatter field. Default `"yyyy-MM-dd HH:mm:ss"`.
- **`ContactDateFormat`** — .NET format string for the `{{created}}` placeholder in contact notes. Default `"yyyy-MM-dd"`.
- **`AppointmentDateFormat`** — .NET format string for appointment `{{startDateTime}}`/`{{endDateTime}}` metadata and the body `**Start:**`/`**End:**` lines. Default `"yyyy-MM-dd HH:mm"`.
- **`FilenameSubjectPatterns`** — ordered list of regex find/replace rules applied after subject cleanup to canonicalize filenames (e.g. `"Re: Re: foo"` → `"Re_foo"`). Ships with the 11 rules SlingMD has always used; clearing the list restores the built-in defaults so users can't accidentally regress baseline behavior.
- **Settings UI** — new textboxes in the Contacts, Email, and Appointments tabs of the Settings dialog for `ContactLinkFormat`, `ContactDateFormat`, `EmailDateFormat`, and `AppointmentDateFormat`.

#### New contact template placeholders (SS-05)
- `{{firstName}}`, `{{lastName}}`, `{{middleName}}`, `{{suffix}}`, `{{fullName}}`, `{{displayName}}` now populate from a parsed `ContactName` so contact templates can render name parts individually.

#### Auto-sling, thread completion, date-range appointments (earlier in the branch)
- Folder-watching auto-sling with rule engine (sender / domain / category matchers) and a watched-folder list.
- Manual **Complete Thread** command that backfills missing emails in a conversation from the current folder and Sent Items.
- Custom **Save Date Range** dialog for exporting appointments in an arbitrary window alongside the existing "Save Today's Appointments".
- `Sent to Obsidian` Outlook category auto-applied to slung emails for visual tracking.

### Fixed
- `TemplateService.RenderContactContent` previously skipped the user's `ContactTemplate.md` when `_settings.ContactTemplateFile` equaled the default filename — the gate was inverted. Removed. The bundled `Templates/ContactTemplate.md` is updated to match the rich in-memory default so existing-default users see no visible change.
- `TemplateService.BuildFrontMatter` produced unquoted, seconds-less `date: 2026-04-21 14:05` for any `DateTime` frontmatter value (bypassing YAML escaping and the rest of the customization story). Now routes through `DateFormatter` with `EmailDateFormat`, properly quoted and escaped.
- `Helpers/FileHelper.cs` — deleted. It was a dead static helper with its own divergent `CleanFileName` (used `-` instead of `_`) that had no call sites anywhere in the codebase; risked confusion if anyone re-wired it.

### Changed — testability refactor
Thirteen pure, unit-testable helpers were extracted from the Outlook-coupled services so formatting and decision logic can be exercised without Outlook installed. Every orchestrator now delegates through these helpers:

| Helper (`Services/Formatting/` unless noted) | Replaces inline logic in | Tests |
|---|---|---|
| `DateFormatter` | `EmailProcessor`, `AppointmentProcessor`, `ContactService`, `TemplateService` | 8 |
| `SubjectCleanerService` | `EmailProcessor.CleanSubject`, `AppointmentProcessor.CleanSubject`, subject-cleanup legacy-pattern migration | 16 |
| `ContactNameParser` | New — drives `ContactLinkFormat` token resolution | 12 |
| `ContactLinkFormatter` | `ContactService.BuildLinkedNames` (×2), `EmailProcessor` sender link, `AppointmentProcessor` organizer/attendees | 15 |
| `NoteTitleBuilder` | Inline title-format construction in `EmailProcessor` and `AppointmentProcessor` | 17 |
| `ThreadIdHasher` | `ThreadService.GetConversationId` two MD5 branches | 14 (incl. 5 golden-hash pins locking historical thread IDs) |
| `FrontmatterReader` | 5 inline regex parses in `ThreadService.FindExistingThread` and `ThreadCompletionService` | 10 |
| `FileNameSanitizer` | `FileService.CleanFileName` post-cleanup pass (invalid chars, prefix strip, separator collapse) | 15 |
| `FilenameSubjectNormalizer` | Duplicated 15-line Re_/Fw_ collapse + ColonSpace block in both processors | 12 |
| `TemplatePathResolver` | `TemplateService.BuildTemplateCandidatePaths` (4-location search) | 9 |
| `EmailAddressParser` | `AutoSlingService.ExtractDomain` + `ContactNameParser.ExtractLocalPart` | 13 |
| `MarkdownSectionFinder` | `ContactService.FindSectionStart` | 11 |
| `UniqueFilenameResolver` | Duplicated `while (File.Exists) { _N++ }` loop in `AttachmentService` and `AppointmentProcessor` | 10 |
| `LegacyFilenameStripper` | 3 inline regexes in `ThreadService.ResuffixThreadNotes` | 13 |
| `ReminderDueDateCalculator` (`Services/`) | `TaskService` due/reminder math, `AppointmentProcessor` appointment-task reminder | 8 + 5 clock-injected integration tests |
| `SlingDecisionEngine` (`Services/`) | `AutoSlingService` (swapped direct `RuleEngine.ShouldAutoSling` for matched-rule-returning helper) | 10 |
| `IClock` / `SystemClock` (`Infrastructure/`) | `DateTime.Now` in `TaskService` and `AppointmentProcessor` — now mockable via `FakeClock` | — |

#### Supporting infrastructure
- `ObsidianSettings.FilenameSubjectPatterns` default list is mirrored in `FilenameSubjectNormalizer.BuiltInDefaults`, with a pin test asserting the two stay in sync.
- `RuleEngine` gained `Match(...)` that returns the matched rule; `ShouldAutoSling` is now a thin wrapper over `Match(...) != null`.
- New DTOs under `Models/`: `ContactName`, `MailItemSnapshot`, `SlingDecision`, `TaskDueDates`, `TaskDueSettings`, `FilenameSubjectRule`.

#### Test coverage
- Total unit-test count grew from 163 → **356** (+193) across 15 new test files.
- All existing orchestrator tests (`EmailProcessorTests`, `ContactProcessorTests`, `AppointmentProcessorTests`, `TaskServiceTests`, `ThreadServiceTests`, `FlagMonitorServiceTests`, `FolderMonitorServiceTests`) continue to pass without modification.

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
