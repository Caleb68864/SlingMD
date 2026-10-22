# Changelog

All notable changes to SlingMD are documented in this file.

## [1.2.5.0] - Unreleased

> **Not yet built or tested on Windows.** Parts of this release were written on a machine that
> can't compile the VSTO project. Build, run the tests and try it in Outlook before publishing.

### Added

#### Copy emails for an AI session
- **New `Copy Email` menu** on the Sling ribbon. It copies the selected emails, or their complete
  threads, to the clipboard as plain text, with or without a metadata header per email. When the
  text is too large for the clipboard, SlingMD offers to save it as a UTF-8 text file instead.

#### Better email bodies
- **HTML emails are converted to Markdown** — headings, lists, links, bold/italic and quotes come
  through instead of Outlook's plain-text rendering. Tracking pixels are dropped.
- **Inline images are saved and linked.** Images embedded in the body (screenshots, pasted
  pictures, signature logos) are saved with the attachments and the `cid:` references in the note
  point at the saved files.

#### New options
- **Always create a thread folder** (Threading tab, off by default) — files every email into a
  thread folder on its first sling, so a note never moves when a reply arrives.
- **Reuse duplicate attachments** (Attachments tab, off by default) — a file identical to one
  already saved (e.g. a signature logo) is stored once in `Attachments/Shared` and linked from there.
- **Put note in its attachment folder** (Attachments tab, off by default, *Subfolder per note* mode
  only) — the note and its files live together in one folder.
- **Copy note path to clipboard** (General tab, **on** by default) — after a manual sling the saved
  note's full path is on the clipboard.

#### Feedback
- **Duplicate slings say so.** Slinging an email that's already in the vault shows a toast instead
  of appearing to do nothing.
- **Successful slings are confirmed** with a toast showing where the note was saved, when Launch
  Obsidian is off.

### Changed
- **Long paths are shortened automatically.** Outlook isn't long-path aware, so the Windows
  `LongPathsEnabled` registry switch doesn't help the add-in; deep vaults and long subjects failed
  with *"The specified path, file name, or both are too long"*, mostly during Complete Thread.
  SlingMD now keeps every file path under 240 characters and every folder under 230 by shortening
  thread folder names and note names at a word break. The date prefix on thread notes is kept, and
  names that already fit are unchanged. Temporary thread note names use a 12-character hash
  instead of the full Message-ID.
- **Complete Thread no longer prompts for every email.** It now runs like Sling Multiple: no
  task-date dialog, contact prompt, countdown or Obsidian launch per email (tasks use your default
  days). Obsidian opens once at the end, on the thread's summary note.
- **Sling with several items selected** now hands off to Sling Multiple. It used to export only the
  first item, silently.
- **The thread summary note (`0-….md`) is written once and never rewritten.** Its content renders
  live through Dataview, so rewriting it only wiped anything you'd added. Changes to the thread
  template now apply to new threads only.
- **`Save Inline Images` is on by default for new installs.** Upgraded installs keep their saved
  value (off), so turn it on under *Attachments* to get images in your notes.
- **Thread folders reuse the existing folder** even when an older email of the conversation is still
  loose in the inbox, and new folders are named from the conversation topic rather than the first
  note's title. Subfolders without a `0-` summary note (batch-sling folders) are no longer mistaken
  for threads.
- **Standalone notes are written under the name they keep inside a thread**, so links to a note
  survive a later move into a thread folder. Moved notes bring their attachments with them, gain
  the `threadNote` property, and have their links rewritten.
- **Obsidian opens notes by their path from the vault root**, so notes in thread folders open
  correctly when the Inbox isn't the vault root.

### Fixed
- **Subjects with " - " no longer lose the separator.** A former default subject-cleanup pattern
  (`[-_\s]{2,}`) removed it and fused the words on either side ("Renewal Meeting - BOT DOC" became
  "Renewal MeetingBOT DOC"). It's removed from existing settings automatically on load.
- **A different email with the same subject, sender and minute no longer overwrites a note**; the
  new note gets a `_N` suffix.
- **A colliding note is never deleted when a note moves into a thread folder.**
- **A contact note that can't be read (locked, open in another app) is no longer overwritten.**
- **Contact aliases no longer add a UTF-8 BOM** that showed as a stray character in Obsidian.
- **One locked or unreadable inbox note no longer aborts an export** whose note was already written.
- **Background slings no longer pop a modal error per email** when the vault is offline; they log
  and toast instead. A failed auto-sling is also retried later instead of being marked done.
- **Obsidian opens the right file** when a thread rename can't track the new note.
- **Invalid custom date formats** in appointment and contact notes fall back to the defaults instead
  of aborting the export.
- **COM handles are released reliably on shutdown** and when monitored folders share a name.

## [1.2.4.0-alpha] - 2026-08-06

> **Alpha release — lightly tested.** The features below are verified by unit tests and code
> review, but have not been exercised against a live Outlook install. Expect rough edges, and
> keep a backup of your vault. Prefer [1.2.3.0](https://github.com/Caleb68864/SlingMD/releases/tag/v1.2.3.0)
> if you need a stable build.

> **You must uninstall 1.2.3.0 before installing this build.** The signing certificate changed
> (`CN=AzureAD\CalebBennett` → `CN=SlingMD Development`), because the previous one expired on
> 2026-03-22. ClickOnce treats a different publisher as a different product, so it will refuse to
> update in place. Remove SlingMD via *Settings → Apps* first, then run `setup.exe`. Your settings
> live in AppData and are not affected.

### Added

#### Choose a destination folder when slinging a single email
- **New `Ask Where to Sling` setting** (Email tab, **off** by default). When enabled, slinging one
  email opens a folder picker listing the subfolders under your Inbox folder, where you can also
  type a new one to create. *Skip* writes to the Inbox folder as usual; *Cancel* aborts the sling.
  Previously only a multi-email batch sling could choose a destination — a single sling always went
  to the configured Inbox folder. Resolves [#14](https://github.com/Caleb68864/SlingMD/issues/14).
- The chosen folder is applied to an isolated copy of your settings, so background and auto-slings
  running at the same time keep writing to the real Inbox.

### Fixed

- **Folder picker no longer crashes when the vault path is unconfigured.** A blank Inbox path
  reached `Path.GetFullPath("")`, which threw straight out of the OK handler. The picker now
  disables folder creation and explains why.
- **Folder picker no longer creates vault folders next to `OUTLOOK.EXE`.** When `VaultBasePath` was
  unset the Inbox path resolved to a *relative* path, silently creating folders in Outlook's working
  directory with no error. The picker now requires an absolute, well-formed path.
- **Empty folders are no longer left behind by a sling that writes nothing.** The picker creates the
  destination folder on OK, but a sling can still stop early — slinging an already-slung email hits
  the duplicate check and writes no note. Such a folder is now removed, but only if this run created
  it and only while it is still empty, so a folder you picked is never touched.

## [1.2.3.0] - 2026-07-20

- **Settings persistence fixed** ([#13](https://github.com/Caleb68864/SlingMD/issues/13)) — settings on the Contacts, Tasks, Threading, and Developer tabs now survive an Outlook restart. The Settings dialog no longer closes when a save fails; it reports the specific error instead
- **Attachments fixed** ([#12](https://github.com/Caleb68864/SlingMD/issues/12)) — real attachments are now saved under default settings
- **Atomic settings writes** — settings are written to a temp file and swapped into place, so an interrupted write can no longer truncate the file and silently reset every setting
- **Numeric settings clamped on load**, so an out-of-range value can no longer prevent the Settings dialog from opening
- **Resource leaks fixed** — Outlook COM objects (Explorer, Selection, MailItem, Session) and GDI bitmaps are now released across the ribbon, auto-sling, and bulk paths
- **Hardening sweep** — file I/O and path safety, exception and null safety in the export path, background-monitor thread safety, and input validation
- **Release packaging verifies binaries** — the packaging script rebuilds and refuses to publish unless the shipped DLL matches the freshly built one byte-for-byte

## [1.2.2.3-prerelease] - 2026-05-06

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

## [1.1.0.7]

- **Contact Slinging** — export single contacts or your entire address book to Obsidian with rich detail notes (phone, email, company, address, birthday)
- **Appointment Contact Linking** — after exporting an appointment, attendees are checked against vault contacts; existing managed notes are refreshed, new contacts offered via dialog
- **Unified Sling Button** — main Sling button now handles emails, appointments, and contacts; dispatches based on selected item type
- **Obsidian Launch on Contact Export** — single-contact export now launches Obsidian for immediate feedback, matching email behavior
- **Unified Contact Template System** — one `RenderContactContent()` path for all contacts; custom templates get all 13 placeholders regardless of source (email or slung)
- **ContactNoteIncludeDetails Setting** — toggle `## Contact Details` section on/off for slung contacts in Settings > Contacts
- **"Sling All Contacts" Ribbon Button** — bulk-export entire address book with progress tracking and summary dialog
- **Contact Confirmation Dialog Clarity** — Cancel button renamed to "Skip Creating Contacts" for clearer intent
- **Custom Templates Documentation** — README now documents all template types, available placeholders, and includes example custom contact template
- New `ContactProcessor` orchestrator, `ExtractContactData()`, `CreateContactNote(ContactTemplateContext)` overload
- 12 new tests (87/87 total passing)

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
