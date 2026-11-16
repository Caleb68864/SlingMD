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
| `Services/StatusService.cs` | split: it owns a `ProgressForm`; an `IStatusReporter` interface goes to Core, the form-backed implementation stays |
| `Services/NotificationService.cs`, `SupportService.cs` | stays (toasts, MessageBox); the engine calls them through `IUserNotifier` |
| `Services/EmailProcessor.cs` | split (first of the big four, see below) |
| `Services/AppointmentProcessor.cs` | split (second) |
| `Services/ThreadService.cs` | split (third): thread folder/note naming, `FindExistingThread`, summary note, resuffixing are file work; `GetConversationId` and the `MailItem` overloads stay as a reader |
| `Services/ContactService.cs` | split (fourth): contact notes, links and the contact index are file work; `GetSenderEmail`/`BuildEmailList` (COM recipients) stay as a reader |
| `Services/ContactProcessor.cs` | split: `ContactItem` reader stays, note writing moves |
| `Services/TaskService.cs` | split: Obsidian task text moves; Outlook `TaskItem` creation stays |
| `Services/AttachmentService.cs` | split: naming, folders, inline/real decision (`ShouldSave` on `AttachmentSignals`) move; `Attachment.SaveAsFile` and the MAPI reads stay behind an attachment-saver interface |
| `Services/ThreadCompletionService.cs` | split: matching moves, the Outlook search stays |
| `Services/SmtpAddressResolver.cs`, `CalendarFolderService.cs` | stays (COM readers) |
| `Services/AutoSlingService.cs`, `FlagMonitorService.cs`, `FolderMonitorService.cs` | stays (monitors); their decisions already use `RuleEngine`/`SlingDecisionEngine` in Core |

**Order for the big four**, each pinned by Linux characterization tests of the engine before the
Outlook class is cut down:

1. **EmailProcessor.** `EmailData` (plain values read from a `MailItem`) in Core,
   `OutlookEmailReader` in Outlook, and the note-building pipeline as a Core engine. This is the
   pattern for the others. Its ThreadService/ContactService/TaskService calls are passed in as
   interfaces until those are split.
2. **AppointmentProcessor**: same shape with `AppointmentData`; it shares the template, file and
   contact pieces, so it goes once EmailProcessor shows the pattern works.
3. **ThreadService**: its file-system half (folders, summary note, resuffixing, `FindExistingThread`)
   is already COM-free apart from the `MailItem` parameters; switch those to `EmailData`.
4. **ContactService**: last, because both processors call it; the COM recipient reads move into the
   readers and the rest goes to Core.

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
  (`NotificationService.cs`, `Forms/ToastForm.cs`, `EmailProcessor.cs` near the export-failure toast)
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
