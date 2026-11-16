# SlingMD Roadmap

Open issues found in reviews that have not been fixed yet. Paths are relative to the repository
root; line numbers are as of the hardening pass on 2026-11-09 (branch `harden/2026-11-09`) and
will drift. Move an item to `CHANGELOG.md` when it ships.

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

- **Old notes still link to dotted folder names.** Notes written before the fix may link to
  `...Jr./file.pdf` while the folder on disk is `...Jr`. Not repaired; a one-off vault fix-up would
  need to find and rewrite those links.
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
  `TaskDueDates`, `EmailSlingResult`); their five test files are linked in (485 tests on Linux).
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
