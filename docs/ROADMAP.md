# SlingMD Roadmap

Open issues found in reviews that have not been fixed yet. Paths are relative to the repository
root; line numbers are as of the hardening pass on 2026-11-09 (branch `harden/2026-11-09`) and
will drift. Move an item to `CHANGELOG.md` when it ships.

## Engine / front end

Goal (owner, 2026-11-16): everything that does not need Windows or Outlook lives in SlingMD.Core
and is tested on Linux; SlingMD.Outlook becomes a thin front end. The project keeps its name and
the namespaces stay `SlingMD.Outlook.*`.

**The rule.** SlingMD.Core has no `Microsoft.Office.Interop`, no WinForms (`System.Windows.Forms`,
`System.Drawing`), no COM (`Marshal`, `COMException`). SlingMD.Outlook holds only COM extraction
(reading `MailItem`/`AppointmentItem`/`ContactItem` into plain data), the ribbon, the forms, the
monitors (folder, flag, auto-sling event hooks), and the composition root in `ThisAddIn`. Anything
the engine needs from the front end (a prompt, a confirmation, a toast, saving an attachment
through COM, opening Obsidian) goes through a small interface defined in Core and implemented in
Outlook.

**Per file in SlingMD.Outlook** (move = no Office/WinForms/COM, moves as is; split = engine to
Core, COM/UI part stays; stays = front end):

| File | Target |
| --- | --- |
| `Forms/*` (all 16 files) | stays (WinForms) |
| `Forms/SettingsHelp.cs` | move later: help text only, but read by `SettingsForm`/`HelpForm`; no gain until a non-WinForms front end exists |
| `Helpers/SafeComAction.cs` | stays (COM) |
| `Infrastructure/IClock.cs`, `SystemClock.cs` | move |
| `Infrastructure/MapiPropertyTags.cs` | move (string constants only) |
| `Models/AttachmentSignals.cs`, `Models/ClipboardEmail.cs` | move |
| `Properties/*`, `Ribbon/*`, `ThisAddIn*.cs` | stays (VSTO, ribbon, composition) |
| `Services/TemplateService.cs` | move |
| `Services/SubjectFilenameCleaner.cs` | move (with FileService) |
| `Services/FileService.cs` | move; the "could not launch Obsidian" `MessageBox` goes behind `IUserNotifier` (Core), implemented by `MessageBoxUserNotifier` (Outlook). `Process.Start` is in netstandard2.0. |
| `Services/ClipboardEmailService.cs`, `ContactInteractionMode.cs` | move (formatting only / enum) |
| `Services/StatusService.cs` | done: moved to Core over an `IProgressView` (Core); `ProgressFormView` (Outlook) wraps the `ProgressForm` |
| `Services/NotificationService.cs` | done: moved to Core; takes an `IToastNotifier` (`ThisAddIn` passes `ToastFormNotifier`), pinned by `NotificationServiceTests` |
| `Services/SupportService.cs` | stays (MessageBox); the engine calls the front end through `IUserNotifier` and `IToastNotifier` (Core, implemented by `MessageBoxUserNotifier` and `ToastFormNotifier`) |
| `Services/EmailProcessor.cs` | split (first of the big four, see below) |
| `Services/AppointmentProcessor.cs` | split (second) |
| `Services/ThreadService.cs` | done: moved to Core on `EmailData`; `GetConversationId` stays in Outlook as `ConversationIdReader` |
| `Services/ContactService.cs` | split (fourth): contact notes, links and the contact index are file work; `GetSenderEmail`/`BuildEmailList` (COM recipients) stay as a reader |
| `Services/ContactProcessor.cs` | split: `ContactItem` reader stays, note writing moves |
| `Services/TaskService.cs` | split: Obsidian task text moves; Outlook `TaskItem` creation stays |
| `Services/AttachmentService.cs` | split: naming, folders, inline/real decision (`ShouldSave` on `AttachmentSignals`) move; `Attachment.SaveAsFile` and the MAPI reads stay behind an attachment-saver interface |
| `Services/ThreadCompletionService.cs` | split: matching moves, the Outlook search stays |
| `Services/SmtpAddressResolver.cs`, `CalendarFolderService.cs` | stays (COM readers) |
| `Services/AutoSlingService.cs`, `FlagMonitorService.cs`, `FolderMonitorService.cs` | stays (monitors); their decisions already use `RuleEngine`/`SlingDecisionEngine` in Core |

**Done so far** (`refactor/engine-frontend-split`): the "move" rows above, apart from
`SettingsHelp`, are in SlingMD.Core, with their tests linked into SlingMD.Core.Tests (572 on
Linux). `StatusService` did not move then: it creates a `ProgressForm`, so it needed the interface
split (done since, see below).
Needs Windows: `TemplateService` looks for templates beside `Assembly.GetExecutingAssembly()`,
which is now `SlingMD.Core.dll` instead of `SlingMD.Outlook.dll`; both are built into the same
folder, and `AppDomain.BaseDirectory` is tried first anyway, but check a published install still
finds the `Templates/` copies. ~~The `IUserNotifier` warning when Obsidian cannot be
launched is not covered by a test.~~ Done on `refactor/thread-core-and-fixes`: `FileService` takes an
internal URI-opener seam (default: `Process.Start` with `UseShellExecute = true`, which is what .NET
Framework's `Process.Start(string)` did; on .NET Core it would have tried to run the URI as a file),
and `FileServiceLaunchTests` (7, Linux) pin the URI, the exact warning for each absorbed failure
type, and that other exceptions propagate.

**EmailProcessor, first cut** (same branch). `EmailData` (Core, `Models/`) holds what the note is
built from: subject, sender name and SMTP address, received time, HTML and plain bodies, EntryID,
Internet message id, first To recipient, and each recipient's name and SMTP address.
`OutlookEmailReader` (Outlook) fills it, with the COM reads and releases moved there unchanged
(`ExtractEmailUniqueIds`, `GetFirstRecipient`, the contact-collection recipient loop).
`EmailNoteEngine` (Core) holds the note-building code moved out of `EmailProcessor` unchanged:
title, base file name, front matter (`BuildMetadata`), body conversion and cleanup, note rendering
and the vault-relative path. `EmailProcessor` now reads once (`ReadRecipients` where the contact
loop was, `Read` inside the existing try, so a failed read is reported as before) and calls the
engine; its internal `BuildEmailMetadata`/`CleanNoteBody`/`GetVaultRelativePath` delegate, so
`EmailProcessorTests` are unchanged. Pinned by `EmailNoteEngineTests` (11, Linux): exact title,
file names, front-matter keys and order, body cleanup flag, and the exact note text for both a
token-dump template and the built-in default template.
Left for EmailProcessor:
  - The duplicate check moved to Core as `ProcessedEmailCache` (code unchanged; EmailProcessor
    keeps one static instance, as it had static fields, and passes its clock's time), pinned by
    `ProcessedEmailCacheTests` (6, Linux). The tests were written after the move, against the same
    code; nothing tested the cache before.
  - The rest of `ProcessEmailWithResult` (thread folder moves and renames, attachments, contacts,
    tasks, launch/toast) is still in Outlook. Next: `WaitForFileAvailability` (pure file work) to
    Core (done: `FileAvailabilityWaiter.WaitAsync`, code unchanged, the delay injectable for tests,
    pinned by `FileAvailabilityWaiterTests`, 5, Linux), then the thread steps behind an interface
    until ThreadService is split.
  - Toasts and progress window (done): `IToastNotifier` and `IProgressView` in Core;
    `ToastFormNotifier` and `ProgressFormView` in Outlook wrap `ToastForm.ShowToast` and
    `ProgressForm`. `StatusService` moved to Core and takes an `IProgressView` (null = silent, as
    `silent: true` was); `EmailProcessor` and `AppointmentProcessor` pass `ProgressFormView.Open()`.
    `EmailProcessor` takes both as optional arguments of its test constructor (defaults: the add-in
    wrappers), and its three toasts go through `IToastNotifier` with the same text and gating.
    Pinned by `StatusServiceTests` (5, Linux). `EmailProcessor` still uses WinForms directly for
    nothing but `Clipboard.SetText` (copy note path) and the `MessageBoxUserNotifier` it hands
    its default `FileService`.
  - Error boxes and countdown (done): `IEmailPrompts` in Core (`ShowError(message)`,
    `ShowLaunchCountdown(seconds)`); `DialogEmailPrompts` in Outlook wraps the `MessageBox`
    ("SlingMD Error", OK, error icon) and `CountdownForm`. `EmailProcessor`'s three error boxes and
    the pre-launch countdown go through it with the same text and gating; it is an optional
    `prompts` argument of the test constructor (default: `DialogEmailPrompts`). Only a construction
    test pins it (Mono): the calls sit inside `ProcessEmailWithResult`, which needs a COM `MailItem`.
    Task dates (done): `IEmailPrompts.PromptTaskOptions(defaults)` returns a `TaskDateChoice` (Core
    model: due days, reminder days, reminder hour, relative flag) or null on cancel;
    `DialogEmailPrompts` wraps `TaskOptionsForm` and EmailProcessor initialises or disables tasks
    from it as before. Contacts (done): `IEmailPrompts.ConfirmNewContacts(names)` returns the
    ticked names (empty on cancel), over `ContactConfirmationDialog`.
    Appointments (done): `AppointmentProcessor` takes an optional `prompts` constructor argument
    (default: `DialogEmailPrompts`) and its task-dates prompt and new-contact confirmation go
    through `PromptTaskOptions`/`ConfirmNewContacts`, same defaults and cancel handling; a
    construction test pins it (Mono). Still direct there: the master-note Yes/No and warning
    `MessageBox`es, the contacts error box and `CountdownForm`.
    Remaining: the clipboard copy has no interface yet. No test drives
    the prompts through `ProcessEmailWithResult` (needs a `MailItem` double).
    Needs Windows: the VS build (`DialogEmailPrompts.cs` added to `SlingMD.Outlook.csproj`) and a
    sling that fails, and one with the countdown on, to see the same dialogs.
    Needs Windows: the VS build (two new files in `SlingMD.Outlook.csproj`, `StatusService.cs`
    removed from it) and a sling and an appointment save to see the progress window and toasts.
  - Side-effect interfaces not yet introduced: the clipboard copy, attachment saving
    (`AttachmentService.ProcessAttachments(MailItem, ...)`), and the Outlook task (`TaskService`).
  - To/Cc linked names and address lists still come from `ContactService.BuildLinkedNames`/
    `BuildEmailList` on the COM `Recipients`; `EmailData.Recipients` has no To/Cc flag yet. Adding
    it means one more COM read (`Recipient.Type`) in the contact loop.
  - `GetConversationId`, `GetThreadFolderName` and `GetThreadNoteName` still take the `MailItem`.
  - Needs Windows: the add-in has not been built with VS MSBuild since this change (the Mono
    harness compiles the same sources minus `ThisAddIn` and the ribbon), and a real sling should be
    checked in Outlook, since `ProcessEmailWithResult` itself has no automated test.

**AppointmentProcessor, first cut** (`refactor/appointment-note-engine`, done). `AppointmentData`
(Core, `Models/`) holds what the note is built from: subject, body, location, start/end, organizer
name and address, GlobalAppointmentID, recurrence (new Core enum `AppointmentRecurrence`, mapped
from `OlRecurrenceState`), calendar name, and required/optional/resource attendee names.
`OutlookAppointmentReader` (Outlook) fills it, with the COM reads and releases moved there unchanged
(the per-property reads, `GetCalendarName`, the organizer `AddressEntry` and the recipient loop).
`AppointmentNoteEngine` (Core) holds the note-building code moved out of `AppointmentProcessor`
unchanged: person links, title, `ShouldGroupAsRecurring`, `DescribeRecurrence`, front matter
(the meeting-URL property folded in, still the last key), custom-template check and note rendering.
`AppointmentProcessor` reads once and calls the engine; its `internal DescribeRecurrence` delegates,
so `AppointmentProcessorTests` are unchanged. Pinned by `AppointmentNoteEngineTests` (19, Linux):
title, recurrence mapping and grouping, front-matter keys and order, contact link formats, the exact
built-in note text, and every custom-template token. `AppointmentProcessorExportTests` (3, Mono;
VSTest) drive `ProcessAppointment` end to end in bulk mode with a Moq `AppointmentItem` and pin the
exact note and meeting-note files; they passed on the old code and pass on the new. The double has
no organizer or recipients, because the reader releases those with `Marshal.ReleaseComObject`,
which throws on a non-COM object.
Windows CI on PR #26 (`979fc52`) failed
`ProcessAppointment_RecurringOccurrence_WritesIntoSeriesFolder` with `MissingMethodException` for
`_AppointmentItem.GetRecurrencePattern()`: the test assembly's embedded interop type lacked that
member, so the Moq proxy did not implement it (Mono tolerated it). `b5f1d9c` removed the call; the
test double now also sets up `Parent`, so every member the processor reads is referenced. Not yet
re-run on Windows.
Left for AppointmentProcessor:
  - ~~The companion meeting note and recurring-folder summary are still built in the processor.~~
    Done on `refactor/thread-core-and-fixes`: `AppointmentNoteEngine.BuildMeetingNoteContent`,
    `MeetingNoteFileName` and `BuildRecurringSummaryContent` (code moved unchanged); the processor
    keeps the exists check and the write. Pinned by four more `AppointmentNoteEngineTests` (exact
    meeting-note text). The contact list and the Outlook task are still in the processor.
  - ~~`GetRecurringThreadFolderName` still calls `GetRecurrencePattern()` and ignores the result.~~
    Done: the read is gone (the folder name is the cleaned subject, as before).
    `AppointmentProcessorExportTests.ProcessAppointment_RecurringOccurrence_DoesNotReadTheRecurrencePattern`
    fails on the old code; `..._RecurringSubjectEndingInDot_FolderHasNoTrailingDot` pins the name.
  - The master-note Yes/No and warning `MessageBox`es, the contacts error box and `CountdownForm`
    (see the prompts item above), and the early `MeetingStatus`/`RecurrenceState` reads before the
    series-master guard.
  - Needs Windows: the VS build (`OutlookAppointmentReader.cs` added to `SlingMD.Outlook.csproj`,
    two new test files in `SlingMD.Tests.csproj`) and saving a real appointment, a recurring
    occurrence and one with attendees in Outlook.

**Order for the big four**, each pinned by Linux characterization tests of the engine before the
Outlook class is cut down:

1. **EmailProcessor.** `EmailData` (plain values read from a `MailItem`) in Core,
   `OutlookEmailReader` in Outlook, and the note-building pipeline as a Core engine. This is the
   pattern for the others. Its ThreadService/ContactService/TaskService calls are passed in as
   interfaces until those are split.
2. ~~**AppointmentProcessor**: same shape with `AppointmentData`; it shares the template, file and
   contact pieces, so it goes once EmailProcessor shows the pattern works.~~ First cut done (see
   above); the leftovers listed there remain.
3. ~~**ThreadService**: its file-system half (folders, summary note, resuffixing, `FindExistingThread`)
   is already COM-free apart from the `MailItem` parameters; switch those to `EmailData`.~~ Done on
   `refactor/thread-core-and-fixes`: `ThreadService` is in SlingMD.Core (same name and namespace,
   `git mv` so history follows). `GetThreadFolderName`, `GetThreadNoteName` and `UpdateThreadNote`
   take `EmailData`, which gained `ConversationTopic` (read by `OutlookEmailReader`). The two
   `FindExistingThread` message boxes (thread-debug details, search error) go through
   `IUserNotifier`, which gained `ShowInformation`; the constructor takes an optional notifier and
   all three add-in call sites pass `MessageBoxUserNotifier`, so the same boxes show.
   `GetConversationId` and its PropertyAccessor reads moved unchanged to `ConversationIdReader`
   (Outlook); `ThreadServiceConversationIdTests` now construct that (call sites only).
   `ThreadServiceTests` (whole file) is linked into SlingMD.Core.Tests, and
   `ThreadServiceEmailDataTests` (12) pin the names and summary note: their expected values were
   taken from the old `MailItem` methods under Mono with Moq doubles of the same input.
   Found while doing it (fixed later, see the bug hunt below): an email whose conversation topic
   is empty (not null) got an empty summary title.
   Windows CI (run 44972012621) failed
   `UpdateThreadNote_NoTopic_TitleFromSubjectAndVaultRelativeFolder`: the built-in thread template
   is a C# verbatim string, so the summary note's line endings are those of the checked-out source
   (CRLF on Windows under `* text=auto`). The test now compares with `\n` only, as the other
   note-text tests do. Left as is in the product: Obsidian reads either, but the line endings of
   SlingMD's notes depend on how the build machine checked out the source, and email notes mix
   CRLF front matter (`AppendLine`) with an LF body (the HTML converter emits `\n`).
   Needs Windows: the VS build (`ThreadService.cs` removed from and `ConversationIdReader.cs`
   added to `SlingMD.Outlook.csproj`), and a sling into an existing thread.
4. **ContactService**: last, because both processors call it; the COM recipient reads move into the
   readers and the rest goes to Core.

## Fork features

- ~~**steveosullivan/main `4b4026e` ("SteveO custom build")**~~ Brought in on
  `refactor/thread-core-and-fixes` (owner-approved), re-implemented in SlingMD.Core rather than
  cherry-picked (the fork is based on `702110a`, before the Core split):
  - Custom note filename (`{Subject} - {SenderFull} - MM-dd-yyyy`, hard-coded in the fork): now the
    existing **Email Filename Format** with a new `{SenderFull}` token and a new
    `EmailFilenameDateFormat` setting ("Filename {Date} Format", default `yyyy-MM-dd`) for `{Date}`.
    `EmailNoteEngine.BuildBaseFileName` takes the display name (`EmailData.SenderName`). The fork
    cleaned each part separately because it thought cleanup would strip " - "; it does not
    (`BuildBaseFileName_SenderFullToken_UsesTheDisplayName`).
  - Trimmed front matter (fork removed `threadId`, `internetMessageId`, `entryId` outright):
    `IncludeEmailIdProperties` and `IncludeThreadIdProperty`, both default on. `threadId` is kept
    while `GroupEmailThreads` is on, since `FindExistingThread` matches on it (the fork's build broke
    thread grouping). Without the ids, `ProcessedEmailCache`'s scan of the inbox can't recognise an
    email slung before a restart, and `MoveToThreadFolder` adds no id suffix; both are in the help
    text. Keys are added conditionally rather than removed, because a `Dictionary` reuses a removed
    slot and would move later keys up (pinned by `BuildMetadata_AllTrimmed_LaterKeysStayInOrder`).
  - HTML email bodies as Markdown: already on `main` since `58fc88d` (1.2.5.0) in
    `HtmlToMarkdownConverter`, which also unwraps Safe Links, handles quoted replies and skips the
    same tracking pixels as the fork; nothing taken.
  - Not taken: the fork's build tweaks (its own ClickOnce certificate thumbprint, and taking
    SlingMD.Tests out of the solution build). Machine-specific, and would stop CI running the tests.
  - Tests: `EmailNoteEngineFilenameAndPropertiesTests` (17, Linux). Settings UI (`SettingsForm`
    Email tab, `SettingsHelp`) compiled under Mono only. Needs Windows: open Settings, change each
    of the three settings, save, reopen, and sling an email to see the filename and properties.

## Bug hunt, 2026-11-21 (`refactor/thread-core-and-fixes`)

Looked at in SlingMD.Core: file-name sanitising, front-matter escaping, unicode and long subjects,
date and time-zone handling. Each fix has a test that fails on the old code.

- ~~**Culture-dependent dates.**~~ Fixed: bare `DateTime.ToString("yyyy-MM-dd")` (and `"HH:mm:ss"`,
  `"_HHmm"`, `"yyyy-MM"`) used the current culture, so under th-TH (Thai Buddhist calendar) titles,
  `{Date}`, appointment file names, task dates (Tasks plugin), attachment month folders and contact
  birthdays got 2569, and fi-FI turned `:` into `.` in the copied-email timestamp. All now pass
  `InvariantCulture`, as the front-matter dates and `GetThreadNoteName` already did.
  `CultureInvariantDatesTests` (5, Linux), `TaskServiceClockTests.GenerateObsidianTask_ThaiCulture_*`
  and `AppointmentProcessorExportTests.ProcessAppointment_ThaiCulture_*` (Mono) fail on the old code.
  `EmailProcessor`'s file timestamp, `ContactService`'s birthday and `AttachmentService`'s month
  folder got the same change without a test of their own (they need a `MailItem` or `ContactItem`).
  Left alone: UI text (`ThreadCompletionDialog`, the batch summary, the Outlook task body), log file
  names, and the diagnostics export file name.
- ~~**Truncation through a surrogate pair.**~~ Fixed: `PathLengthBudget.Fit`,
  `NoteTitleBuilder.Build`, `ThreadService`'s 47-character subject cut, `ContactLinkFormatter`'s
  `{FirstInitial}`/`{LastInitial}` and `ContactService.GetFilenameSafeShortName` cut by UTF-16
  unit, so an emoji at the limit left a lone surrogate in the name, which the UTF-8 note text then
  wrote as U+FFFD. New `SurrogateSafe.Truncate` (Core) steps back one unit instead.
  `SurrogatePairTruncationTests` (5, Linux; the 4 at a cut fail on the old code) and
  `ContactServiceTests.GetFilenameSafeShortName_EmojiAtTheCuts_*` (Mono, fails on the old code).
  Not changed: other `Substring` calls cut at found indexes (separators, `@`, front-matter
  markers), which cannot split a pair.
- ~~**Front-matter escaping let non-printable characters through.**~~ Fixed:
  `TemplateService.EscapeYamlDoubleQuotedScalar` escaped C0 controls and DEL but not the C1
  controls U+0080-U+009F, U+FFFE/U+FFFF or unpaired surrogates, all outside YAML's printable set;
  js-yaml (Obsidian) then rejects the whole front matter. They are now `\xNN`/`\uNNNN` escapes.
  `TemplateServiceTests.EscapeYamlDoubleQuotedScalar_EscapesCharactersYamlCannotHold` fails on the
  old code; `..._KeepsPrintableUnicode` pins that accents, CJK, emoji, NBSP and U+FEFF pass through.
  Left: `ClipboardEmailService.EscapeYaml` (copied text for AI sessions, not parsed by Obsidian)
  and `ContactAliasWriter.FormatYamlValue` (aliases from contact names) have their own simpler
  escaping; worth sharing the one escaper if either is ever parsed as YAML.
- ~~**Recurring-series summary never listed anything.**~~ Fixed: `WriteRecurringThreadSummary`
  passed the absolute folder path to the thread template's `dv.pages('"{{folderPath}}"')`
  (Dataview wants it vault-relative; the email path already converts), the instance notes had no
  `threadId` for the query's filter, and the summary was rewritten on every instance, wiping user
  edits. Now: `AppointmentNoteEngine.BuildRecurringSummaryContent` makes the folder vault-relative,
  grouped instances get `threadId` (folder name, last key), and an existing summary is kept.
  `AppointmentProcessorExportTests` (summary path, second occurrence keeps edits, exact recurring
  note) fail on the old code. Left: the summary still uses the email thread template ("Email
  Timeline", `type: email-thread`, sorted by `date`, which appointment notes don't have); an
  appointment-specific summary template would read better. Old summaries are not migrated.
- ~~**Empty conversation topic gave a blank thread summary title.**~~ Fixed: `UpdateThreadNote`
  used `??`, so `""` won over the subject (the folder name used `IsNullOrEmpty`).
  `ThreadServiceEmailDataTests.UpdateThreadNote_EmptyTopic_TitleFromSubject` fails on the old code.
- ~~**Reserved device names with an extension.**~~ Fixed: `FileNameSanitizer` and
  `FolderNameSanitizer` caught only exact names ("CON"), but Windows also reserves them before an
  extension ("NUL.txt", "Con.Ed bill.md": the part before the first dot, trailing spaces dropped).
  Shared `FileNameSanitizer.IsReservedDeviceName`; six new sanitizer test cases fail on the old
  code. Windows 11 relaxed some of these rules; the `_` prefix is harmless there. Not covered:
  `COM0`/`LPT0` and the superscript-digit forms (`COM¹`), which Windows also reserves.

## Correctness

Done on `fix/folder-names-thread-ids` (2026-11-16); see `CHANGELOG.md`:

- ~~**Folder names ending in a dot or space.**~~ `FolderNameSanitizer`
  (`SlingMD.Core/Services/Formatting/FolderNameSanitizer.cs`) trims trailing dots/spaces and guards
  empty and reserved results. Used by `ThreadService.GetThreadFolderName`,
  `AppointmentProcessor.GetRecurringThreadFolderName`, `AttachmentService.GetAttachmentFolder`
  (per-note subfolder), and wherever `ThreadService` re-derives a note's folder from its name
  (`RelocateNoteCompanions`, `IsNoteFolder`). `FileNameSanitizer` still keeps the dot for names.
  `BatchFolderPickerForm` already rejected a typed name ending in a dot. Tests:
  `SlingMD.Core.Tests/.../FolderNameSanitizerTests.cs` (Linux), and
  `AttachmentServiceTests.*EndingInDot*`, `ThreadServiceTests.MoveToThreadFolder_GroupedNoteNameEndingInDot_*`
  (VSTest).
- ~~**Random thread id as a last resort.**~~ `ThreadService.GetConversationId` now falls back to
  `FallbackThreadId.FromIdentity` (SHA-1 of the Internet message id, else the EntryID). Random only
  when neither can be read. Tests: `SlingMD.Core.Tests/.../FallbackThreadIdTests.cs`.

Found and left:

- ~~**Old notes still link to dotted folder names.**~~ Done on `fix/dotted-folder-links`
  (2026-11-16): a one-shot **Repair Folder Links** ribbon button (Settings group) rather than
  rewriting on the next write, because such links can sit in any note (email notes, the thread
  summary, the user's own notes) and the dotted folders came from four different write paths.
  `DottedFolderLinkRewriter` (`SlingMD.Core/Services/Formatting/`) rewrites only link targets
  (wikilinks, embeds, Markdown links/images, `<...>` and URL-encoded targets; not aliases,
  headings, titles, URLs, code spans or fenced blocks), and only when the target is broken and the
  repaired one exists. `DottedFolderLinkRepairService` (`SlingMD.Core/Services/`) walks the vault
  (skipping dot-folders), checks existence by exact folder names relative to the note and the vault
  root (Windows would open `Jr.\file.pdf` as `Jr\file.pdf`), and writes through a temp file,
  keeping any BOM. `ThisAddIn.RepairDottedFolderLinks` is the thin confirm/summary wrapper. Tests:
  `DottedFolderLinkRewriterTests` (30), `DottedFolderLinkRepairServiceTests` (6), both linked into
  SlingMD.Core.Tests. The ribbon button and dialog are not covered by tests; check them in Outlook.
  Left: a thread summary's Dataview `FROM "...Jr."` query is not a link and is not rewritten, and
  wikilinks Obsidian would only resolve by path suffix are not checked (note folder and vault root
  only).
- **An email with no topic, no conversation index and an empty subject gets an empty thread id**
  (`ThreadIdHasher.Hash("")` returns `""`), so all such emails share one. The new identity hash
  could be used there too, but that changes grouping for subjectless emails; not done without a
  decision.
- ~~**`GetConversationId` itself has no test.**~~ Done on `refactor/finish-core-split`:
  `SlingMD.Tests/Services/ThreadServiceConversationIdTests.cs` (9 tests) mocks `MailItem` and
  `PropertyAccessor` with Moq. A probe showed Moq 4.18.4 proxies the embedded Outlook interop types
  under .NET 10 and Mono, so no seam was needed; the tests passed under Mono but have not run under
  VSTest on .NET Framework. `ReadInternetMessageId` now releases the accessor only when
  `Marshal.IsComObject` is true, since `ReleaseComObject` throws on a test double. The same approach
  should work for the recipient-address fallbacks below.
- ~~**`MapiPropertyTags.PrConversationIndex` is probably the wrong type.**~~ Fixed on
  `refactor/core-helpers-and-conv-index` (owner decision, new slings only): the tag is now the
  binary `0x00710102`, so `GetConversationId`'s strategy 2 gets the `byte[]` it expects for emails
  with no conversation topic. The step now reads the index through `ReadConversationIndex`, which
  releases the `PropertyAccessor` (shared `ReleaseAccessor`, as `ReadInternetMessageId` does). Test:
  `ThreadServiceConversationIdTests.GetConversationId_NoTopic_ReadsTheConversationIndexAsBinary`
  (fails with the old tag; passed under Mono, needs the VSTest run). Not checked against a real
  Outlook store that `GetProperty` returns `byte[]` for this tag.
- **Thread ids changed for emails with no conversation topic** (from the fix above). They used to
  get the subject hash; new slings now get the conversation-index id, so a new sling of a
  conversation slung before may land in a new thread folder beside the old one. Old notes are not
  migrated; a fix-up would have to find notes whose `threadId` is the subject hash and either
  rewrite it or teach the thread lookup to try both ids.
- **The conversation-index id uses only 10 bytes of the header**: the reserved byte, the 5-byte
  timestamp and 4 of the 16 GUID bytes (`Take(22)` then the first 20 hex digits). Two conversations
  started within a second or two of each other with equal leading GUID bytes would share an id. Unlikely;
  left as is so the id shape matches what the code always meant to produce.

## Error handling and logging

The four items first listed here (AppendToFile, the recipient-lookup fallbacks, silent toast and
icon failures, ContactAliasWriter's catches) were fixed on `harden/2026-11-09`; see `CHANGELOG.md`.
Left over from that work:

- **Some best-effort catches are still broad on purpose.** The toast catches
  (`SlingMD.Core/Services/NotificationService.cs`, `Forms/ToastForm.cs`, `EmailProcessor.cs` near the export-failure toast)
  and the ribbon icon loaders (`Ribbon/SlingRibbon.cs`) catch `System.Exception`, now with a log
  line and a comment. Narrow them only once the real failure types are known from logs.
- ~~**Other bare catches not yet reviewed.**~~ Reviewed on `refactor/finish-core-split`; no bare
  `catch` is left in these files:
  - `SlingMD.Core/Services/Formatting/ContactLinkFormatter.cs` (`Resolve`): narrowed to file-system
    failures (`IOException`, `UnauthorizedAccessException`, `ArgumentException`,
    `NotSupportedException`, `SecurityException`) and logged; anything else from the callback now
    propagates. Safe because both shipped callbacks (`ContactService.NoteFileExists`) catch their own
    errors. Three tests in `ContactLinkFormatterTests`.
  - `Services/EmailProcessor.cs` (subject read for the fuzzy-match log): narrowed to `COMException`
    and `InvalidComObjectException`, and logged. Anything else reaches the existing "Error processing
    contacts" handler. Not tested (needs a mail double driven through `ProcessEmail`).
  - `Helpers/Logger.cs` (now in SlingMD.Core): `Log()` stays broad and silent, written as
    `catch (System.Exception)` with the reason; there is nowhere to report to. `CleanOldLogs` stays
    broad but now logs, and catches I/O and permission errors per file so one locked log doesn't stop
    the rest. Also found: the constructor's `Directory.CreateDirectory` was unguarded inside a
    `Lazy<T>`, which caches the exception, so an unwritable AppData would have made every
    `Logger.Instance` call throw. It is now guarded. Not tested (singleton bound to AppData).
- **No tests for the recipient-address fallbacks.** `SmtpAddressResolver.ForRecipient` and both
  `ContactService.BuildEmailList` overloads take Outlook COM objects. Moq can mock the embedded
  interop types (see `ThreadServiceConversationIdTests`), so a Moq test of `Recipient`
  (PropertyAccessor throws, Address throws) should work; watch for `Marshal.ReleaseComObject` calls
  on the doubles, which throw `ArgumentException`.

## Build and tests

- ~~**Tests can only run on Windows with Visual Studio.**~~ Partly done on
  `fix/folder-names-thread-ids`: the pure helpers moved to `SlingMD.Core` (`net472;netstandard2.0`)
  and `SlingMD.Core.Tests` (`net10.0`, outside `SlingMD.sln`) runs them with `dotnet test` on Linux.
  On `refactor/finish-core-split` (2026-11-16) `ObsidianSettings` (plus the models it serialises)
  and `Logger` moved as well, bringing `BodyCleanerService`, `SubjectCleanerService`,
  `ThreadIdHasher`, `FilenameSubjectNormalizer`, `TemplatePathResolver` and `ContactAliasWriter`;
  their tests and the `Models/` settings tests are linked in, 428 tests on Linux. The `core-tests`
  CI job on `ubuntu-latest` passed on PR #16. On `refactor/core-helpers-and-conv-index` `RuleEngine`,
  `SlingDecisionEngine`, `ReminderDueDateCalculator`, `BatchSlingReport` and `BoundedHashSet` moved
  too, with the five DTOs they use (`MailItemSnapshot`, `SlingDecision`, `TaskDueSettings`,
  `TaskDueDates`, `EmailSlingResult`); their five test files are linked in (485 tests on Linux; 521 after `fix/dotted-folder-links`).
  Still Windows-only:
  - The rest of SlingMD.Tests: services and forms that take Outlook or WinForms types.
  - The add-in build with SlingMD.Core referenced has not been built with VS MSBuild, nor
    published; check that the ClickOnce payload includes `SlingMD.Core.dll` (and still exactly one
    `Newtonsoft.Json.dll`, now that SlingMD.Core references it by PackageReference).
- **Tests added on 2026-11-09 have not been run.** `TemplateServiceTests` (control-character
  escaping, invariant numbers, `{{<list>Yaml}}` escaping) and `FileServiceTests`
  (`BuildObsidianOpenUri`, `EnsureDirectoryExists` on an invalid path, `AppendToFile`) were written
  on Linux. The
  escaping and URI logic was checked in a throwaway .NET 10 harness, but the tests themselves need a
  run under VSTest.
  On 2026-11-16 the whole SlingMD.Tests suite was compiled against the add-in sources (minus
  `ThisAddIn` and the ribbon) as net472 and run under Mono's xunit console on Linux: 721 tests, 43
  failures, the same 43 that fail on `main` there, all from `C:\` paths or WinForms under Mono
  (`BuildObsidianOpenUri_AbsolutePathInsideVault_IsMadeVaultRelative` is one of them). The other
  2026-11-09 tests passed. That is evidence, not a substitute for the VSTest run.
- ~~**Path tests fail off Windows.**~~ Done on `test/portable-paths` (2026-11-16, test code only).
  With that Mono run (net472 build of SlingMD.Tests, `xvfb-run -a mono xunit.console.exe
  SlingMD.Tests.dll -parallel none`), `main` had 17 failures. 14 were path tests that fed literal
  `C:\` paths to `AttachmentService`, `FileService.BuildObsidianOpenUri` and
  `EmailProcessor.GetVaultRelativePath` (12 in `AttachmentServiceTests`, one each in the other two),
  and one was `TemplateServiceTests.LoadTemplate_CachesContentUntilFileMtimeChanges`, which restored a
  sub-second mtime that Mono truncates. They now build paths with `Path.Combine(Path.GetTempPath(), ...)`
  and a whole-second mtime; the expected values did not change. Mono: 17 failures before, 2 after
  (734 tests). None of these test SlingMD.Core classes, so nothing new is linked into
  SlingMD.Core.Tests (still 485, all passing). Not yet run under VSTest on Windows.
  `AttachmentServiceTests.GetAttachmentFolder_SubfolderPerNote_NoteNameEndingInDot_DropsTheDot` still
  uses a `C:\` literal (it passes under Mono; left alone because the dotted-folder work touches it).
- **WinForms tests are Windows-only.** The five classes in `SlingMD.Tests/Forms/` carry
  `[Trait("Category", "WindowsOnly")]` (46 tests). Two of them fail under Mono:
  `SettingsFormPersistenceTests.InvalidField_DoesNotCloseDialogAndDoesNotPersist` (Mono WinForms) and
  `BatchFolderPickerFormTests.UnsafeFolderName_IsRejected("sub\\nested")` (`\` is not a separator
  on Linux). A Mono run with `-notrait "Category=WindowsOnly"` has 688 tests and 0 failures. The
  VSTest CI job still runs them all.
