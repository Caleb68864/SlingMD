# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview
SlingMD is a .NET Framework 4.7.2 C# Outlook VSTO add-in that exports emails to Obsidian as markdown. The solution consists of two projects:
- **SlingMD.Outlook**: Main add-in project
- **SlingMD.Tests**: xUnit test project with Moq for mocking

## Build and Test Commands
- **Build**: `dotnet build SlingMD.sln --configuration Release`
- **Build single project**: `dotnet build SlingMD.Outlook\SlingMD.Outlook.csproj --configuration Release`
- **Run tests**: `dotnet test SlingMD.Tests\SlingMD.Tests.csproj`
- **Publish add-in**: `dotnet publish SlingMD.Outlook\SlingMD.Outlook.csproj --configuration Release`
- **Package for distribution**: `.\package-release.ps1` (creates versioned ZIP in Releases/)

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
- **ObsidianSettings** (Models/ObsidianSettings.cs): Centralized configuration with JSON persistence
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
SlingMD.Outlook/
├── Forms/          - WinForms dialogs (Settings, TaskOptions, Progress, etc.)
├── Helpers/        - Utility classes (deprecated - use Services instead)
├── Models/         - Data models (ObsidianSettings)
├── Properties/     - Assembly info, resources
├── Ribbon/         - Outlook ribbon UI (SlingRibbon.cs/xml)
├── Services/       - Business logic services (FileService, EmailProcessor, etc.)
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