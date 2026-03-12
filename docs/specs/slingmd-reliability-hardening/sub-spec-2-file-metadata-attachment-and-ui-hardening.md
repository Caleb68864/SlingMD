---
type: phase-spec
master_spec: "../2026-03-12-slingmd-reliability-hardening.md"
phase: 2
title: "File, Metadata, Attachment, and UI Hardening"
date: 2026-03-12
dependencies:
  - 1
status: ready
---

# Sub-Spec 2: File, Metadata, Attachment, and UI Hardening

## Analysis Summary

- Files: `SlingMD.Outlook/Services/EmailProcessor.cs`, `ThreadService.cs`, `TemplateService.cs`, `AttachmentService.cs`, and `Forms/SettingsForm.cs` all exist and already own the behavior in scope.
- Patterns found: `TemplateService` is already the single place for frontmatter rendering, `AttachmentService` already owns attachment storage-mode decisions, and `ThreadService` already centralizes thread discovery and note re-suffixing.
- Interfaces: this phase defines the metadata/date parsing contract, the frontmatter escaping contract, and the attachment-link contract that the verification phase will test and document.
- Test location: existing service tests live in `SlingMD.Tests/Services/`. `TemplateServiceTests.cs` already exists on disk but is not compiled today. New `ThreadServiceTests.cs` and `AttachmentServiceTests.cs` are the natural additions.

## Current-State Notes

- `EmailProcessor.EnsureEmailCacheIsBuilt(...)` still calls `Directory.GetFiles(...)` without a `Directory.Exists(...)` guard.
- `ThreadService.FindExistingThread(...)` still parses `date: yyyy-MM-dd HH:mm` while `EmailProcessor.BuildEmailMetadata(...)` writes `yyyy-MM-dd HH:mm:ss` as a quoted string.
- `TemplateService.BuildFrontMatter(...)` writes raw double-quoted strings and raw list values, so quotes, backslashes, and newlines can still break YAML.
- `AttachmentService.GenerateWikilink(...)` only receives a filename, which is insufficient for centralized or per-note-subfolder storage.
- `SettingsForm` is already wiring `Browse`, `Add`, `Edit`, `Remove`, and the development toggle in the current branch. Treat UI wiring as verification-only unless a newly discovered control gap appears during implementation.

## Files To Touch

- `SlingMD.Outlook/Services/EmailProcessor.cs`
- `SlingMD.Outlook/Services/ThreadService.cs`
- `SlingMD.Outlook/Services/TemplateService.cs`
- `SlingMD.Outlook/Services/AttachmentService.cs`
- `SlingMD.Outlook/Forms/SettingsForm.cs` only if a new unwired visible control is discovered
- `SlingMD.Tests/Services/EmailProcessorTests.cs`
- `SlingMD.Tests/Services/TemplateServiceTests.cs`
- New `SlingMD.Tests/Services/ThreadServiceTests.cs`
- New `SlingMD.Tests/Services/AttachmentServiceTests.cs`
- `SlingMD.Tests/SlingMD.Tests.csproj` as needed so the targeted tests actually compile

## Patterns To Follow

- Reuse `TemplateService` for any escaping helper instead of adding a new YAML dependency unless absolutely necessary.
- Preserve same-folder attachment output exactly as it works today; only add relative segments when the storage mode requires them.
- Favor backward-compatible parsing in `ThreadService` so existing notes continue to resolve.
- Keep `SettingsForm` changes off the table unless verification proves the current event wiring is incomplete.

## Interface Contracts

**Provides**

- Missing inbox folders are treated as empty state during duplicate-cache population.
- Thread discovery accepts the date shape written by current exports and remains compatible with legacy minute-precision notes.
- `TemplateService.BuildFrontMatter(...)` emits valid YAML-safe scalar and list values for real-world metadata.
- Attachment link generation resolves correctly for same-folder, per-note-subfolder, and centralized storage modes.

**Requires**

- Sub-Spec 1 must already guarantee that failed exports stop before downstream follow-up actions.
- If new service test files are added here, compile them immediately or leave clear notes for Sub-Spec 3 to consolidate the test project entries.

**Shared State**

- Frontmatter fields written by `EmailProcessor.BuildEmailMetadata(...)`
- Inbox-folder cache state in `EmailProcessor`
- Attachment storage settings from `ObsidianSettings`
- Existing notes already written with current second-precision quoted `date` values

**Verification**

- Exporting into a brand-new vault state does not throw during duplicate detection.
- Old and new thread notes both resolve to the same earliest-thread behavior.
- Metadata containing `"`, `\`, and embedded line breaks remains parseable.
- Attachment links inside the note point to the saved file location for every storage mode.

## Implementation Slices

### Slice 1: Treat a missing inbox folder as an empty cache input

1. Add focused coverage in `SlingMD.Tests/Services/EmailProcessorTests.cs`.
   - Prefer a small internal helper or extracted cache-building seam such as `BuildEmailIdCache_WhenInboxFolderMissing_ReturnsEmptySet`.
2. Run the targeted tests.
   - `dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter FullyQualifiedName~EmailProcessorTests`
3. Write the minimal implementation in `SlingMD.Outlook/Services/EmailProcessor.cs`.
   - Guard `EnsureEmailCacheIsBuilt(...)` with `Directory.Exists(inboxPath)`.
   - Preserve the current cache refresh timing and cache-update behavior after a successful save.
4. Re-run the targeted tests.
   - Same command as above.
5. Commit.
   - `fix(cache): treat missing inbox as empty state`

### Slice 2: Align thread date parsing with written metadata

1. Add `SlingMD.Tests/Services/ThreadServiceTests.cs`.
   - Start with `FindExistingThread_ParsesQuotedSecondPrecisionDate`.
   - Add `FindExistingThread_AcceptsLegacyMinutePrecisionDate` so old notes remain supported.
2. If the new file is not compiled yet, add the minimum `Compile Include` entry in `SlingMD.Tests/SlingMD.Tests.csproj`.
3. Run the targeted tests.
   - `dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter FullyQualifiedName~ThreadServiceTests`
4. Write the minimal implementation in `SlingMD.Outlook/Services/ThreadService.cs`.
   - Centralize date parsing in one helper.
   - Accept the currently written quoted `yyyy-MM-dd HH:mm:ss` format first, then legacy minute-precision values for compatibility.
5. Re-run the targeted tests.
   - Same command as above.
6. Commit.
   - `fix(threading): align date parsing with exported metadata`

### Slice 3: Make frontmatter YAML-safe without changing note meaning

1. If `SlingMD.Tests/Services/TemplateServiceTests.cs` is not compiled yet, add the smallest `Compile Include` change needed to run it.
2. Add failing tests in `SlingMD.Tests/Services/TemplateServiceTests.cs`.
   - `BuildFrontMatter_EscapesQuotesBackslashesAndNewlines`
   - `BuildFrontMatter_EscapesListValuesWithoutChangingListShape`
3. Run the targeted tests.
   - `dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter FullyQualifiedName~TemplateServiceTests`
4. Write the minimal implementation in `SlingMD.Outlook/Services/TemplateService.cs`.
   - Add a local escaping helper for double-quoted YAML scalars.
   - Reuse it for both string values and list entries.
   - Preserve current field names, ordering, and overall frontmatter structure.
5. Re-run the targeted tests.
   - Same command as above.
6. Commit.
   - `fix(frontmatter): escape yaml metadata safely`

### Slice 4: Generate attachment links from real saved paths

1. Add `SlingMD.Tests/Services/AttachmentServiceTests.cs`.
   - Start with `GenerateAttachmentLink_SameFolder_UsesBareFilename`.
   - Add `GenerateAttachmentLink_SubfolderPerNote_UsesRelativeSubfolderPath`.
   - Add `GenerateAttachmentLink_CentralizedStorage_UsesRelativeVaultPath`.
2. If the new file is not compiled yet, add the minimum `Compile Include` entry in `SlingMD.Tests/SlingMD.Tests.csproj`.
3. Run the targeted tests.
   - `dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter FullyQualifiedName~AttachmentServiceTests`
4. Write the minimal implementation in `SlingMD.Outlook/Services/AttachmentService.cs` and `SlingMD.Outlook/Services/EmailProcessor.cs`.
   - Change link generation to work from the saved attachment path relative to the note path.
   - Preserve current same-folder output.
   - Keep wikilink vs standard markdown behavior controlled by `UseObsidianWikilinks`.
5. Re-run the targeted tests.
   - Same command as above.
6. Commit.
   - `fix(attachments): generate links for each storage mode`

### Slice 5: Verify the settings UI instead of rewriting it

1. Do a manual UI pass against the current `SettingsForm`.
   - Exercise `Browse`, `Add`, `Edit`, `Remove`, contact-saving toggle behavior, and the development-settings toggle.
2. Only write code if a newly discovered visible control is still inert.
   - If that happens, keep the fix in `SlingMD.Outlook/Forms/SettingsForm.cs` and add a short verification note for Sub-Spec 3.
3. Commit only if code was needed.
   - `fix(settings): wire missing settings control handlers`

## Verification Commands

- Build check: `dotnet build SlingMD.sln --configuration Release`
- Cache/export tests: `dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter FullyQualifiedName~EmailProcessorTests`
- Threading tests: `dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter FullyQualifiedName~ThreadServiceTests`
- Frontmatter tests: `dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter FullyQualifiedName~TemplateServiceTests`
- Attachment tests: `dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter FullyQualifiedName~AttachmentServiceTests`
- Manual acceptance:
  1. Export into a vault where the inbox folder does not exist yet.
  2. Reconstruct a thread from notes using both current second-precision dates and older minute-precision dates.
  3. Export metadata containing quotes, backslashes, and multiline values; inspect the frontmatter.
  4. Save attachments in each storage mode and confirm the inserted links resolve.
  5. Exercise the visible settings controls and confirm their handlers fire.
- Integration handoff to Sub-Spec 3:
  - Confirm the attachment-link and frontmatter contracts are now stable enough to document and lock down with broader regression coverage.
