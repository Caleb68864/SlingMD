# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview
SlingMD is a .NET Framework 4.7.2 C# Outlook VSTO add-in that exports emails to Obsidian as markdown. The solution consists of three projects, plus one test project kept outside it:
- **SlingMD.Outlook**: Main add-in project, becoming a thin front end (COM reading, ribbon, forms,
  monitors, composition in ThisAddIn). Anything without Office/WinForms/COM goes to SlingMD.Core;
  see "Engine / front end" in `docs/ROADMAP.md` for the per-file plan.
- **SlingMD.Core**: Pure helpers and engines, `ObsidianSettings` (with the models it serialises) and `Logger`,
  with no Office/WinForms dependency (SDK-style,
  `net472;netstandard2.0`). Namespaces stay `SlingMD.Outlook.*`, so moving a file here needs no
  `using` changes. Internals are visible to SlingMD.Outlook, SlingMD.Tests and SlingMD.Core.Tests.
- **SlingMD.Tests**: xUnit test project with Moq for mocking (VSTest, Windows only)
- **SlingMD.Core.Tests** (not in SlingMD.sln): `net10.0` xUnit tests for SlingMD.Core that run
  anywhere with `dotnet test SlingMD.Core.Tests`. It compiles the SlingMD.Tests test files for
  classes in SlingMD.Core as links, so those run in both projects.

## Build and Test Commands
**The `dotnet` CLI does NOT work on this repo** — the VSTO project imports
`Microsoft.VisualStudio.Tools.Office.targets`, which only ships with Visual Studio.
`dotnet build`/`dotnet test` fail with MSB4019. Use VS MSBuild and VSTest instead:
- **Build**: `& "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe" SlingMD.sln -t:Build -p:Configuration=Release`
- **Build single project**: same MSBuild.exe with `SlingMD.Outlook\SlingMD.Outlook.csproj`
- **Run tests**: build Debug, then `& "C:\Program Files\Microsoft Visual Studio\2022\Professional\Common7\IDE\Extensions\TestPlatform\vstest.console.exe" SlingMD.Tests\bin\Debug\SlingMD.Tests.dll` (filter with `/Tests:Name` or `/TestCaseFilter:"FullyQualifiedName~X"`)
- **Publish add-in**: use Visual Studio interactively (Build > Publish SlingMD.Outlook). CLI `msbuild -t:Publish` does not reliably regenerate the ClickOnce payload for VSTO projects.
- **Package for distribution**: `.\package-release.ps1` (creates versioned ZIP in Releases/; refuses to package if the publish payload is stale vs the fresh Release build) — full steps in `RELEASING.md`

- **Run the SlingMD.Core tests on any OS** (Linux, macOS, Windows without VS): `dotnet test SlingMD.Core.Tests`.
  The `dotnet` CLI also builds `SlingMD.Core` itself; it is only the add-in and SlingMD.Tests that need VS.

**Test files must be added to `SlingMD.Tests.csproj` explicitly** — the project uses
explicit `<Compile Include=...>` items, no globbing. A new test .cs file that isn't
referenced there silently never compiles or runs.

## Architecture

### Service-Oriented Design
The add-in uses a service-oriented architecture where `EmailProcessor` orchestrates multiple services:
- **EmailProcessor** (SlingMD.Outlook/Services/EmailProcessor.cs): Main orchestrator that coordinates all services
- **FileService**: Handles file I/O operations, file naming, and path management
- **ThreadService**: Groups related emails into thread folders with summary notes
- **TaskService**: Creates follow-up tasks in Obsidian notes and/or Outlook
- **ContactService**: Creates and manages contact notes with communication history
- **TemplateService**: Processes markdown templates for notes
- **StatusService**: Shows progress UI during email processing

### Core Flow
1. User selects email in Outlook and clicks "Sling" ribbon button
2. `ThisAddIn.ProcessSelectedEmail()` retrieves the `MailItem`
3. `EmailProcessor.ProcessEmail()` orchestrates:
   - Task options dialog (if enabled)
   - Email-to-markdown conversion
   - Thread detection and organization
   - Contact note creation
   - Task creation in Obsidian/Outlook
   - File save and Obsidian launch
4. Settings are persisted to JSON via `ObsidianSettings.Save()`

### Key Models
- **ObsidianSettings** (SlingMD.Core/Models/ObsidianSettings.cs): Centralized configuration with JSON persistence
  - Vault paths, task defaults, threading settings, subject cleanup patterns, template formats
  - Load/Save methods handle serialization to AppData

### Threading System
When `GroupEmailThreads` is enabled:
- Emails are grouped by conversation ID/thread ID
- Thread folders contain individual email notes + a summary note (0-threadname.md)
- Summary notes include timeline view, participants, and Dataview queries
- The system handles duplicate detection and chronological ordering

## Code Style Guidelines
- **Naming**: PascalCase for classes/methods/properties, camelCase for variables/parameters, _camelCase for private fields
- **Types**: Use explicit typing over var
- **Imports**: System namespaces first, then third-party, then project-specific
- **Formatting**: Braces on new lines, 4-space indentation
- **Exception Handling**:
  - Use fully qualified `System.Exception` instead of just `Exception`
  - Catch specific exceptions, not generic Exception
  - Display user-facing errors via `MessageBox.Show()`
  - Use `throw;` to preserve stack trace (never `throw ex;`)
- **Services**: Must have 'Service' suffix
- **Interfaces**: Must start with 'I' prefix

## Project Structure
```
SlingMD.Core/
├── Helpers/        - Logger (file log under AppData; never throws), BoundedHashSet
├── Infrastructure/ - IClock/SystemClock, MapiPropertyTags (constants only)
├── Models/         - ObsidianSettings and the DTOs it serialises (AutoSlingRule, WatchedFolder,
│                     CalendarFolderInfo, AttachmentStorageMode), ContactName, FilenameSubjectRule,
│                     MailItemSnapshot, SlingDecision, TaskDueSettings, TaskDueDates, EmailSlingResult,
│                     ClipboardEmail, AttachmentSignals
├── Services/       - TemplateService, FileService, SubjectFilenameCleaner, ClipboardEmailService,
│                     RuleEngine, SlingDecisionEngine, ReminderDueDateCalculator, BatchSlingReport,
│                     IUserNotifier (what the engine asks the front end to show)
└── Services/Formatting/ - Pure helpers (sanitisers, parsers, cleaners, path budget, thread ids, ...)

SlingMD.Outlook/
├── Forms/          - WinForms dialogs (Settings, TaskOptions, Progress, etc.)
├── Helpers/        - Utility classes (deprecated - use Services instead)
├── Properties/     - Assembly info, resources
├── Ribbon/         - Outlook ribbon UI (SlingRibbon.cs/xml)
├── Services/       - COM readers, processors not yet split (EmailProcessor, ThreadService, ...),
│                     monitors, MessageBoxUserNotifier
├── Templates/      - Default markdown templates
└── ThisAddIn.cs    - VSTO entry point

SlingMD.Tests/
├── Models/         - Tests for model classes
└── Services/       - Tests for service classes
```

## Dependencies
- Newtonsoft.Json for settings serialization
- Microsoft Office Interop assemblies (Outlook, Office.Core)
- xUnit + Moq for testing
- .NET Framework 4.7.2 (VSTO requirement)