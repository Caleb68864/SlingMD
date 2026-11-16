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
- **`MapiPropertyTags.PrConversationIndex` is probably the wrong type.** PR_CONVERSATION_INDEX is
  binary (`0x00710102`), but the tag asks for `0x0071001F` (Unicode string), so `GetProperty`
  likely never returns a `byte[]` and `GetConversationId`'s strategy 2 never fires: emails with no
  topic go straight to the subject hash. Not confirmed against a real Outlook store. Fixing it would
  change the thread id of such emails, so not done without a decision. Strategy 2 also never
  releases the `PropertyAccessor` it reads.

## Error handling and logging

The four items first listed here (AppendToFile, the recipient-lookup fallbacks, silent toast and
icon failures, ContactAliasWriter's catches) were fixed on `harden/2026-11-09`; see `CHANGELOG.md`.
Left over from that work:

- **Some best-effort catches are still broad on purpose.** The toast catches
  (`NotificationService.cs`, `Forms/ToastForm.cs`, `EmailProcessor.cs` near the export-failure toast)
  and the ribbon icon loaders (`Ribbon/SlingRibbon.cs`) catch `System.Exception`, now with a log
  line and a comment. Narrow them only once the real failure types are known from logs.
- **Other bare catches not yet reviewed:** `Services/ThreadService.cs` is done, but
  `Services/Formatting/ContactLinkFormatter.cs:100`, `Services/EmailProcessor.cs:667` (subject read
  for logging) and `Helpers/Logger.cs` (94, 119; the logger must not throw, so these may stay) were
  not looked at.
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
  CI job on `ubuntu-latest` passed on PR #16. Still Windows-only:
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
