---
type: phase-spec
master_spec: "../2026-03-12-slingmd-reliability-hardening.md"
phase: 3
title: "Coverage, Regression Tests, and Verification Guidance"
date: 2026-03-12
dependencies:
  - 2
status: ready
---

# Sub-Spec 3: Coverage, Regression Tests, and Verification Guidance

## Analysis Summary

- Files: `SlingMD.Tests/SlingMD.Tests.csproj`, `SlingMD.Tests/Models/ObsidianSettingsTests.cs`, the service test files under `SlingMD.Tests/Services/`, and `README.md` are the primary surfaces here.
- Patterns found: the test project uses explicit `<Compile Include="...">` entries, so any new or pre-existing test file must be added manually. The repo already uses lightweight temp-directory-based service tests and small fake subclasses in `ContactServiceTests.cs`.
- Interfaces: this phase locks down the contracts introduced in Sub-Specs 1 and 2 and documents the environment needed to run them in a real VSTO-capable machine.
- Test location: keep settings tests in `Models/`, service tests in `Services/`, and use the existing xUnit naming pattern.

## Current-State Notes

- `TaskServiceTests.cs` and `TemplateServiceTests.cs` are present on disk but omitted from `SlingMD.Tests/SlingMD.Tests.csproj` in the current branch.
- `EmailProcessorTests.cs` only covers `BuildEmailMetadata(...)` today, so the export-flow, cache, and failure-path coverage is still very light.
- The original review note about a stale `ContactServiceTests` heading is no longer true on the current branch. The tests and implementation already align on `## Communication History`.
- `README.md` documents commands and features, but it still does not clearly call out the Visual Studio Office/VSTO prerequisite that explains the `Microsoft.VisualStudio.Tools.Office.targets` failure seen in this environment.

## Files To Touch

- `SlingMD.Tests/SlingMD.Tests.csproj`
- `SlingMD.Tests/Models/ObsidianSettingsTests.cs`
- `SlingMD.Tests/Services/EmailProcessorTests.cs`
- `SlingMD.Tests/Services/TaskServiceTests.cs`
- `SlingMD.Tests/Services/TemplateServiceTests.cs`
- New or updated `SlingMD.Tests/Services/ThreadServiceTests.cs`
- New or updated `SlingMD.Tests/Services/AttachmentServiceTests.cs`
- `README.md`

## Patterns To Follow

- Add every test file explicitly to `SlingMD.Tests.csproj`; do not assume SDK-style wildcard inclusion.
- Reuse the existing temp-folder and fake-service test patterns instead of introducing heavy test infrastructure.
- Keep README changes concrete and environment-focused: editing can happen anywhere, but build/test/publish need the right Office tooling installed.
- Do not undo current `Communication History` expectations; treat them as the current correct behavior unless the implementation intentionally changes again.

## Interface Contracts

**Provides**

- The test project compiles every service/regression file needed for this hardening effort.
- Each reliability contract from Sub-Specs 1 and 2 has at least one targeted automated test.
- `README.md` tells contributors why `dotnet test` may fail without Visual Studio Office tooling and what environment is expected.

**Requires**

- Sub-Spec 1 must have landed the export-flow and task-state contracts.
- Sub-Spec 2 must have landed the cache, thread-date, frontmatter, and attachment-link contracts.

**Shared State**

- Explicit `<Compile Include>` items in `SlingMD.Tests/SlingMD.Tests.csproj`
- VSTO/Office targets required by both the add-in and the test project through the referenced Outlook project
- README build/test/publish instructions

**Verification**

- `dotnet test SlingMD.Tests\SlingMD.Tests.csproj` lists and runs the newly added service tests on a compatible machine.
- The documented environment note in `README.md` matches the real missing-targets failure observed here.

## Implementation Slices

### Slice 1: Normalize the test project file list

1. Audit `SlingMD.Tests/SlingMD.Tests.csproj` against the files present under `SlingMD.Tests/Models/` and `SlingMD.Tests/Services/`.
2. Add explicit `Compile Include` entries for every test file that should participate in this hardening pass.
   - At minimum include `Services\TaskServiceTests.cs` and `Services\TemplateServiceTests.cs`.
   - Include any new `ThreadServiceTests.cs` and `AttachmentServiceTests.cs` files created in earlier phases.
3. Run a test discovery check on a compatible machine.
   - `dotnet test SlingMD.Tests\SlingMD.Tests.csproj --list-tests`
4. Confirm the expected classes appear in discovery output.
5. Commit.
   - `test(project): include all hardening test files`

### Slice 2: Finish the regression matrix for the hardening contracts

1. Review Sub-Spec 1 and Sub-Spec 2 acceptance criteria and map each one to at least one automated test.
2. Add or tighten coverage in the existing test files.
   - Settings fallback in `ObsidianSettingsTests.cs`
   - Task reset and cancel behavior in `TaskServiceTests.cs`
   - Export-flow or extracted gate coverage in `EmailProcessorTests.cs`
   - Date compatibility in `ThreadServiceTests.cs`
   - YAML escaping in `TemplateServiceTests.cs`
   - Attachment link resolution in `AttachmentServiceTests.cs`
3. Run the relevant targeted tests while filling gaps.
   - `dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter FullyQualifiedName~ObsidianSettingsTests`
   - `dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter FullyQualifiedName~TaskServiceTests`
   - `dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter FullyQualifiedName~EmailProcessorTests`
   - `dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter FullyQualifiedName~ThreadServiceTests`
   - `dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter FullyQualifiedName~TemplateServiceTests`
   - `dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter FullyQualifiedName~AttachmentServiceTests`
4. Run the full suite on a compatible machine.
   - `dotnet test SlingMD.Tests\SlingMD.Tests.csproj`
5. Commit.
   - `test(reliability): add regression coverage for hardening fixes`

### Slice 3: Document the real build and test prerequisites

1. Update `README.md` near the build/test or contributing guidance.
   - State explicitly that editing can happen in VS Code or another editor.
   - State explicitly that building, testing, and publishing the VSTO add-in require Visual Studio with the Office/VSTO tooling that provides `Microsoft.VisualStudio.Tools.Office.targets`.
2. Add a brief note explaining the failure mode seen without that prerequisite.
   - Mention the missing-targets error directly so contributors can self-diagnose it quickly.
3. Re-read the README sections for consistency with the existing command list.
4. Commit.
   - `docs(readme): clarify VSTO build and test prerequisites`

### Slice 4: Run whole-spec verification and record gaps

1. Run the solution build on a compatible machine.
   - `dotnet build SlingMD.sln --configuration Release`
2. Run the full test suite.
   - `dotnet test SlingMD.Tests\SlingMD.Tests.csproj`
3. Perform the highest-value manual checks from the master spec.
   - Malformed settings startup
   - cancel-then-retry task flow
   - missing inbox export
   - metadata escaping
   - attachment storage-mode link resolution
4. If the local environment is missing Office targets, record that as an environment limitation instead of weakening the acceptance criteria.
5. Commit only if this slice required code or doc updates beyond the earlier slices.
   - `chore(verify): capture hardening verification updates`

## Verification Commands

- Discovery check: `dotnet test SlingMD.Tests\SlingMD.Tests.csproj --list-tests`
- Build check: `dotnet build SlingMD.sln --configuration Release`
- Full suite: `dotnet test SlingMD.Tests\SlingMD.Tests.csproj`
- README/manual acceptance:
  1. Read the build/test documentation and confirm it explains the missing `Microsoft.VisualStudio.Tools.Office.targets` failure clearly.
  2. Re-run the manual reliability checks from the master spec on a compatible Outlook/VSTO machine.
- Final readiness gate:
  - Every hardening behavior from Sub-Specs 1 and 2 is either covered by automation or explicitly documented as an environment-limited manual check.
