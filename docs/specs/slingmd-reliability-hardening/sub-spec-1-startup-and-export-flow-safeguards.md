---
type: phase-spec
master_spec: "../2026-03-12-slingmd-reliability-hardening.md"
phase: 1
title: "Startup and Export Flow Safeguards"
date: 2026-03-12
dependencies: []
status: ready
---

# Sub-Spec 1: Startup and Export Flow Safeguards

## Analysis Summary

- Files: `SlingMD.Outlook/ThisAddIn.cs`, `SlingMD.Outlook/Models/ObsidianSettings.cs`, `SlingMD.Outlook/Services/EmailProcessor.cs`, and `SlingMD.Outlook/Services/TaskService.cs` already own the startup/export behavior. `SlingMD.Tests/Models/ObsidianSettingsTests.cs` and `SlingMD.Tests/Services/EmailProcessorTests.cs` are compiled today. `SlingMD.Tests/Services/TaskServiceTests.cs` exists on disk but is not currently included in `SlingMD.Tests/SlingMD.Tests.csproj`.
- Patterns found: service-local fixes are the norm, tests use xUnit `Fact` methods with `Method_State_Expected` naming, and the repo favors explicit typing plus targeted exception handling.
- Interfaces: this phase establishes the per-export task-state contract and the "core export must succeed before contacts or launch" gate that later hardening work assumes.
- Test location: `SlingMD.Tests/Models/` for settings behavior and `SlingMD.Tests/Services/` for task/export seams.

## Current-State Notes

- `ObsidianSettings.Load()` already catches `JsonException` and normalizes defaults. Treat corrupt-JSON handling as regression coverage first, not an automatic code change.
- `TaskService.InitializeTaskSettings(...)` does not restore `_createTasks`, so a canceled task dialog still leaks into later exports.
- `EmailProcessor.ProcessEmail(...)` catches inside the `StatusService` block and then continues into contact creation and Obsidian launch outside that block.
- `ThisAddIn.ShowFirstRunSupportPrompt()` already has a localized warning path for save failures. Do not broaden startup UI behavior unless verification reveals a remaining startup crash path.

## Files To Touch

- `SlingMD.Outlook/Services/TaskService.cs`
- `SlingMD.Outlook/Services/EmailProcessor.cs`
- `SlingMD.Outlook/Models/ObsidianSettings.cs` only if regression coverage exposes a remaining startup failure path
- `SlingMD.Outlook/ThisAddIn.cs` only if unreadable settings still escape `ObsidianSettings.Load()`
- `SlingMD.Tests/Models/ObsidianSettingsTests.cs`
- `SlingMD.Tests/Services/EmailProcessorTests.cs`
- `SlingMD.Tests/Services/TaskServiceTests.cs`
- `SlingMD.Tests/SlingMD.Tests.csproj` if `TaskServiceTests.cs` still needs a `Compile Include`

## Patterns To Follow

- Keep the fix inside the owning service instead of introducing a new abstraction layer.
- Prefer a single success/failure gate in `EmailProcessor` over a broader orchestration rewrite.
- Preserve current success-path behavior; the visible change should be safer failure handling only.
- Use the existing xUnit style and keep any helper seam internal rather than widening public API surface.

## Interface Contracts

**Provides**

- `TaskService.InitializeTaskSettings(...)` establishes per-export defaults and re-enables task creation for that export attempt.
- `EmailProcessor.ProcessEmail(...)` only runs contact creation and Obsidian launch after the core export work completed successfully.
- `ObsidianSettings.Load()` yields a normalized in-memory settings object even when the saved JSON is malformed or partially invalid.

**Requires**

- No prior phase dependency.
- If a targeted regression test lives in a file that is present but not compiled, add the minimum `Compile Include` needed to run it immediately.

**Shared State**

- `%AppData%`-backed `ObsidianSettings.json`
- `_taskService.ShouldCreateTasks`
- `contactNames`, `obsidianLinkPath`, and the success flag that gates post-processing in `EmailProcessor.ProcessEmail(...)`

**Verification**

- A second export in the same Outlook session can create tasks after the first export was canceled.
- A forced core export failure does not show contact follow-up or launch Obsidian.
- A malformed settings file still yields normalized defaults instead of a startup exception.

## Implementation Slices

### Slice 1: Lock in startup fallback behavior

1. Write regression coverage in `SlingMD.Tests/Models/ObsidianSettingsTests.cs`.
   - Add `Load_MalformedJson_KeepsDefaultsAndDoesNotThrow`.
   - Add `Load_TypeMismatchedFields_NormalizesInvalidValuesWithoutClearingValidOnes`.
2. Run the targeted tests.
   - `dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter FullyQualifiedName~ObsidianSettingsTests.Load_`
   - Expected on a compatible VSTO machine: malformed JSON should already stay green; the mixed-validity case will show whether more hardening is still required.
3. Write minimal implementation only if a gap remains.
   - Prefer extending `ObsidianSettings.Load()` before touching `ThisAddIn.cs`.
   - If file-read failures still escape startup, add a narrow fallback and warning path that preserves defaults and avoids crashing the add-in.
4. Re-run the targeted tests.
   - Same command as above.
5. Commit.
   - `test(settings): lock startup fallback behavior`

### Slice 2: Reset task creation per export attempt

1. If `SlingMD.Tests/Services/TaskServiceTests.cs` is not compiled yet, add the smallest `Compile Include` change in `SlingMD.Tests/SlingMD.Tests.csproj` needed to run it.
2. Write the failing tests in `SlingMD.Tests/Services/TaskServiceTests.cs`.
   - Add `InitializeTaskSettings_AfterDisableTaskCreation_ReEnablesTaskGeneration`.
   - Add `ShouldCreateTasks_AfterCancelThenInitialize_ReturnsTrue`.
3. Run the targeted tests.
   - `dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter FullyQualifiedName~TaskServiceTests`
   - Expected before the fix: the new reset coverage fails because `_createTasks` stays `false`.
4. Write the minimal implementation in `SlingMD.Outlook/Services/TaskService.cs`.
   - Re-enable `_createTasks` inside `InitializeTaskSettings(...)` so each export attempt starts from settings-driven defaults.
   - Do not change the cancel behavior for the current export.
5. Re-run the targeted tests.
   - Same command as above.
6. Commit.
   - `fix(tasks): reset task creation for each export`

### Slice 3: Stop post-processing after a fatal export error

1. Add a focused regression seam in `SlingMD.Tests/Services/EmailProcessorTests.cs`.
   - Prefer a small internal helper or result-flag test such as `ShouldRunPostProcessing_WhenCoreExportFails_ReturnsFalse`.
   - If direct unit coverage of `ProcessEmail(...)` remains too interop-heavy, document that limit and keep the seam as small as possible.
2. Run the targeted tests.
   - `dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter FullyQualifiedName~EmailProcessorTests`
3. Write the minimal implementation in `SlingMD.Outlook/Services/EmailProcessor.cs`.
   - Track whether the core export block completed successfully.
   - Return immediately after the `StatusService` block when it did not.
   - Leave contact creation and Obsidian launch unchanged for successful exports.
4. Re-run the targeted tests.
   - Same command as above.
5. Perform a manual failure-path check.
   - Force an I/O failure during note save and confirm there is no contact dialog and no Obsidian launch afterward.
6. Commit.
   - `fix(export): stop follow-up actions after fatal export errors`

## Verification Commands

- Build check: `dotnet build SlingMD.sln --configuration Release`
- Settings tests: `dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter FullyQualifiedName~ObsidianSettingsTests`
- Task tests: `dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter FullyQualifiedName~TaskServiceTests`
- Export-flow tests: `dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter FullyQualifiedName~EmailProcessorTests`
- Manual acceptance:
  1. Replace the saved settings JSON with malformed content and start the add-in.
  2. Cancel task options on one export, then export a second email with task creation enabled.
  3. Force a fatal export failure and confirm there is no contact processing or Obsidian launch.
- Integration handoff to Sub-Spec 2:
  - Confirm `ProcessEmail(...)` now has a reliable "success before follow-up" gate, because later attachment and metadata hardening assumes failed exports stop cleanly.
