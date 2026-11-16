# SlingMD Roadmap

Open issues found in reviews that have not been fixed yet. Paths are relative to the repository
root; line numbers are as of the hardening pass on 2026-11-09 (branch `harden/2026-11-09`) and
will drift. Move an item to `CHANGELOG.md` when it ships.

## Correctness

- **Folder names ending in a dot or space.** `FileNameSanitizer.Sanitize` deliberately keeps a
  trailing dot (`SlingMD.Outlook/Services/Formatting/FileNameSanitizer.cs:65`) because the value
  doubles as a contact name. That is safe for notes (`.md` follows) but not for folders built from
  the same value (thread folders, attachment subfolders): Windows silently drops trailing dots and
  spaces from a directory name, so the folder on disk and the path SlingMD links to can differ.
  Needs a separate folder-name sanitising step and a test.
- **Random thread id as a last resort.** `ThreadService.GetConversationId`
  (`SlingMD.Outlook/Services/ThreadService.cs:98`) still returns a random id when the topic and
  subject both fail, so a retried sling of the same email lands in a new folder. Hashing the
  EntryID or Internet message id would be stable. (The PR_CONVERSATION_INDEX failure case now falls
  through to the subject; only a failure of everything is left.)

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
  `ContactService.BuildEmailList` overloads take Outlook COM objects. A Moq test of `Recipient`
  (PropertyAccessor throws, Address throws) is worth trying, but whether Moq can mock the
  embedded interop types has not been checked.

## Build and tests

- ~~**Tests can only run on Windows with Visual Studio.**~~ Partly done on
  `fix/folder-names-thread-ids`: the pure helpers moved to `SlingMD.Core` (`net472;netstandard2.0`)
  and `SlingMD.Core.Tests` (`net10.0`, outside `SlingMD.sln`) runs them with `dotnet test` on Linux,
  including 19 linked `SlingMD.Tests/Services/Formatting/*Tests.cs` files. A `core-tests` CI job on
  `ubuntu-latest` was added but has not run yet. Still Windows-only:
  - The rest of SlingMD.Tests. Left in SlingMD.Outlook because they take `ObsidianSettings` or
    `Logger`: `BodyCleanerService`, `SubjectCleanerService` (and so `ThreadIdHasher`),
    `FilenameSubjectNormalizer`, `TemplatePathResolver`, `ContactAliasWriter`. Moving
    `ObsidianSettings` (Newtonsoft only) and `Logger` to Core would bring these, plus
    `RuleEngine`, `SlingDecisionEngine`, `ReminderDueDateCalculator`, `BatchSlingReport` and
    `Helpers/BoundedHashSet`, which look pure but were not checked.
  - The add-in build with SlingMD.Core referenced has not been built with VS MSBuild, nor
    published; check that the ClickOnce payload includes `SlingMD.Core.dll`.
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
