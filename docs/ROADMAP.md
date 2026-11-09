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

- **`FileService.AppendToFile` is neither atomic nor guarded**
  (`SlingMD.Outlook/Services/FileService.cs:97`), unlike `WriteUtf8File`. An interrupted append can
  leave a half-written note, and IO errors go straight to the caller. Check each caller's handling
  before changing it.
- **Bare `catch` around recipient address lookups.** `ContactService.cs` (~266 and ~379) and
  `SmtpAddressResolver.ForRecipient` (`SlingMD.Outlook/Services/SmtpAddressResolver.cs:47`) catch
  everything, and the fallback (`recipient.Address`) is itself a COM call that can throw from
  inside the catch. Narrow to `COMException` and guard the fallback.
- **Silent failures without logging.** `NotificationService.cs` (31, 54), `Forms/ToastForm.cs:208`,
  `EmailProcessor.cs:619` and the icon loaders in `Ribbon/SlingRibbon.cs` (43, 102) swallow errors
  without a log line. They should keep degrading quietly but call `Logger.Instance.Warning`.
- **`ContactAliasWriter` cleanup** (`SlingMD.Outlook/Services/Formatting/ContactAliasWriter.cs:77`)
  uses bare `catch` blocks; narrow them to `IOException`/`UnauthorizedAccessException` per the
  style guide.

## Build and tests

- **Tests can only run on Windows with Visual Studio.** Both projects target .NET Framework 4.7.2
  and the add-in imports the VSTO targets, so `dotnet build`/`dotnet test` fail on Linux (MSB3644:
  no 4.7.2 reference assemblies) as well as on Windows without VS (MSB4019). Moving the pure
  helpers in `SlingMD.Outlook/Services/Formatting/` to a multi-targeted library with its own test
  project would let the bulk of the logic be tested on any machine and in CI without Office.
- **Tests added on 2026-11-09 have not been run.** `TemplateServiceTests` (control-character
  escaping, invariant numbers, `{{<list>Yaml}}` escaping) and `FileServiceTests`
  (`BuildObsidianOpenUri`, `EnsureDirectoryExists` on an invalid path) were written on Linux. The
  escaping and URI logic was checked in a throwaway .NET 10 harness, but the tests themselves need a
  run under VSTest.
