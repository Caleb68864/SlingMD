---
type: phase-spec
master_spec: "docs/specs/2026-03-13-appointment-processing.md"
sub_spec: 1
title: "ObsidianSettings Extensions + AppointmentProcessor Core"
dependencies: none
date: 2026-03-13
---

# Sub-Spec 1: ObsidianSettings Extensions + AppointmentProcessor Core

## Shared Context

- **Master Spec:** [2026-03-13-appointment-processing.md](../2026-03-13-appointment-processing.md)
- **Trade-off hierarchy:** Match EmailProcessor patterns > Additive-only changes > Working end-to-end
- **Constraints:** .NET Framework 4.7.2, PascalCase, explicit typing, braces on new lines, 4-space indent, fully qualified `System.Exception`
- **Escalation triggers:** Any change to EmailProcessor.cs or existing shared service method signatures

## Codebase Analysis

### Existing Patterns

**ObsidianSettings.cs** (404 lines):
- 35+ properties with defaults declared inline (e.g., `public string InboxFolder { get; set; } = "Inbox";`)
- Path methods: `GetFullVaultPath()` (line 178), `GetInboxPath()` (line 183), `GetContactsPath()` (line 188)
- Save/Load: `JsonConvert.SerializeObject(this, Formatting.Indented)` / `JsonConvert.PopulateObject(json, this, settings)`
- Validation in `Validate()` method (line 211): range checks, regex validation, path validation
- `NormalizeLoadedSettings()` (line 324): fills null/empty values with defaults after Load
- `GetSettingsPath()` returns `AppData/SlingMD.Outlook/ObsidianSettings.json`
- JsonSerializerSettings: MissingMemberHandling.Ignore, ObjectCreationHandling.Replace

**EmailProcessor.cs** (987 lines):
- Constructor (line 58): takes `ObsidianSettings`, creates FileService, TemplateService, ThreadService, TaskService, ContactService, AttachmentService
- ProcessEmail signature (line 77): `public async Task ProcessEmail(MailItem mail, CancellationToken cancellationToken = default)`
- Static `_processedEmailIds` (line 30): `ConcurrentDictionary<string, bool>` for duplicate detection
- Static `_threadFolderLocks` (line 25): `ConcurrentDictionary<string, SemaphoreSlim>` for thread safety
- Compiled regex patterns (lines 36-48) for subject cleaning
- COM cleanup: `Marshal.ReleaseComObject()` in try/finally blocks (lines 119-147)
- StatusService usage: `using (var status = new StatusService())` wrapping main processing block

### Files to Create/Modify

| File | Action | Exists |
|------|--------|--------|
| `SlingMD.Outlook/Models/ObsidianSettings.cs` | Modify (add properties + method) | Yes (404 lines) |
| `SlingMD.Outlook/Services/AppointmentProcessor.cs` | Create new | No |

## Implementation Steps

### Step 1: Add appointment properties to ObsidianSettings

**Test first:**
- File: `SlingMD.Tests/Models/ObsidianSettingsAppointmentTests.cs`
- Test: `DefaultValues_AllAppointmentProperties_HaveCorrectDefaults`
- Asserts: Each of the 10 new properties has its expected default value

**Run test to verify failure:**
```
dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter "FullyQualifiedName~ObsidianSettingsAppointmentTests.DefaultValues"
```
Expected: Compilation failure (properties don't exist yet)

**Implement:**
- File: `SlingMD.Outlook/Models/ObsidianSettings.cs`
- Add after `UseObsidianWikilinks` property (line 143), before `SubjectCleanupPatterns`:

```csharp
// Appointment Settings
public string AppointmentsFolder { get; set; } = "Appointments";
public string AppointmentNoteTitleFormat { get; set; } = "{Date} - {Subject}";
public int AppointmentNoteTitleMaxLength { get; set; } = 50;
public List<string> AppointmentDefaultNoteTags { get; set; } = new List<string> { "Appointment" };
public bool AppointmentSaveAttachments { get; set; } = true;
public bool CreateMeetingNotes { get; set; } = true;
public string MeetingNoteTemplate { get; set; } = string.Empty;
public bool GroupRecurringMeetings { get; set; } = true;
public bool SaveCancelledAppointments { get; set; } = false;
public string AppointmentTaskCreation { get; set; } = "None";
```

**Run test to verify pass:**
```
dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter "FullyQualifiedName~ObsidianSettingsAppointmentTests.DefaultValues"
```

**Commit:** `feat(settings): add 10 appointment-specific properties to ObsidianSettings`

---

### Step 2: Add GetAppointmentsPath() method

**Test first:**
- File: `SlingMD.Tests/Models/ObsidianSettingsAppointmentTests.cs`
- Test: `GetAppointmentsPath_ReturnsCorrectCombinedPath`
- Asserts: `settings.GetAppointmentsPath()` returns `Path.Combine(VaultBasePath, VaultName, AppointmentsFolder)`

**Run test to verify failure:**
```
dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter "FullyQualifiedName~ObsidianSettingsAppointmentTests.GetAppointmentsPath"
```

**Implement:**
- File: `SlingMD.Outlook/Models/ObsidianSettings.cs`
- Add after `GetContactsPath()` (line 191), following the same pattern:

```csharp
public string GetAppointmentsPath()
{
    return Path.Combine(GetFullVaultPath(), AppointmentsFolder);
}
```

**Run test to verify pass:**
```
dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter "FullyQualifiedName~ObsidianSettingsAppointmentTests.GetAppointmentsPath"
```

**Commit:** `feat(settings): add GetAppointmentsPath() method`

---

### Step 3: Add validation for new appointment properties

**Test first:**
- File: `SlingMD.Tests/Models/ObsidianSettingsAppointmentTests.cs`
- Test: `Validate_AppointmentNoteTitleMaxLength_OutOfRange_Throws`
- Test: `Validate_AppointmentTaskCreation_InvalidValue_Throws`
- Asserts: Invalid ranges/values throw ArgumentException

**Implement:**
- File: `SlingMD.Outlook/Models/ObsidianSettings.cs`
- Add to `Validate()` method (after line 256):

```csharp
ValidateFolderName(AppointmentsFolder, nameof(AppointmentsFolder));
if (AppointmentNoteTitleMaxLength < 10 || AppointmentNoteTitleMaxLength > 500)
    throw new System.ArgumentOutOfRangeException(nameof(AppointmentNoteTitleMaxLength), "Must be between 10 and 500");
string[] validTaskCreation = { "None", "Obsidian", "Outlook", "Both" };
if (!validTaskCreation.Contains(AppointmentTaskCreation))
    throw new System.ArgumentException($"Invalid AppointmentTaskCreation value: {AppointmentTaskCreation}");
```

- Add to `NormalizeLoadedSettings()`:

```csharp
if (AppointmentDefaultNoteTags == null) AppointmentDefaultNoteTags = new List<string> { "Appointment" };
if (string.IsNullOrEmpty(AppointmentsFolder)) AppointmentsFolder = "Appointments";
if (string.IsNullOrEmpty(AppointmentNoteTitleFormat)) AppointmentNoteTitleFormat = "{Date} - {Subject}";
if (string.IsNullOrEmpty(AppointmentTaskCreation)) AppointmentTaskCreation = "None";
```

**Commit:** `feat(settings): add validation and normalization for appointment properties`

---

### Step 4: Add serialization round-trip test

**Test first:**
- File: `SlingMD.Tests/Models/ObsidianSettingsAppointmentTests.cs`
- Test: `RoundTrip_AppointmentSettings_PreservedThroughSaveLoad`
- Pattern: Use `ObsidianSettingsTestable` subclass (existing pattern from ObsidianSettingsTests.cs line 335) to override `GetSettingsPath()` to a temp file
- Asserts: Set all 10 properties to non-default values, Save(), create new instance, Load(), verify all values match

**Implement:** Already implemented in steps 1-3 (Save/Load uses JsonConvert which auto-serializes all public properties)

**Run test to verify pass:**
```
dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter "FullyQualifiedName~ObsidianSettingsAppointmentTests.RoundTrip"
```

**Commit:** `test(settings): add round-trip serialization test for appointment properties`

---

### Step 5: Create AppointmentProcessor skeleton

**Test first:**
- File: `SlingMD.Tests/Services/AppointmentProcessorTests.cs`
- Test: `Constructor_WithValidSettings_CreatesInstance`
- Pattern: Follow EmailProcessorTests.cs (try/catch with System.Exception, Assert.NotNull)

**Implement:**
- File: `SlingMD.Outlook/Services/AppointmentProcessor.cs` (NEW)
- Create class skeleton mirroring EmailProcessor constructor pattern:

```csharp
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using SlingMD.Outlook.Models;

namespace SlingMD.Outlook.Services
{
    public class AppointmentProcessor
    {
        private static readonly ConcurrentDictionary<string, byte> _processedAppointmentIds =
            new ConcurrentDictionary<string, byte>();

        private readonly ObsidianSettings _settings;
        private readonly FileService _fileService;
        private readonly TemplateService _templateService;
        private readonly ThreadService _threadService;
        private readonly TaskService _taskService;
        private readonly ContactService _contactService;
        private readonly AttachmentService _attachmentService;

        // Compiled regex patterns (same as EmailProcessor)
        private static readonly Regex WhitespaceRegex = new Regex(@"\s+", RegexOptions.Compiled);

        public AppointmentProcessor(ObsidianSettings settings)
        {
            _settings = settings;
            _fileService = new FileService(settings);
            _templateService = new TemplateService(_fileService);
            _threadService = new ThreadService(_fileService, _templateService, settings);
            _taskService = new TaskService(settings, _templateService);
            _contactService = new ContactService(_fileService, _templateService);
            _attachmentService = new AttachmentService(settings, _fileService);
        }

        public async System.Threading.Tasks.Task ProcessAppointment(
            AppointmentItem appointment,
            bool bulkMode = false,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            // TODO: Implement in subsequent steps
            await System.Threading.Tasks.Task.CompletedTask;
        }
    }
}
```

**Run test to verify pass:**
```
dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter "FullyQualifiedName~AppointmentProcessorTests.Constructor"
```

**Commit:** `feat(processor): create AppointmentProcessor skeleton with service wiring`

---

### Step 6: Implement ProcessAppointment metadata extraction

**Implement:**
- File: `SlingMD.Outlook/Services/AppointmentProcessor.cs`
- Add metadata extraction to `ProcessAppointment()`:
  - Extract: Subject, HTMLBody, Location, Start, End, GetOrganizer(), Recipients, RecurrenceState, GlobalAppointmentID, Attachments
  - Clean subject using same compiled regex patterns as EmailProcessor
  - Apply AppointmentNoteTitleFormat with `{Date}`, `{Subject}`, `{Sender}` placeholders
  - Truncate to AppointmentNoteTitleMaxLength
  - COM object cleanup for Recipients, Organizer in try/finally blocks

Key implementation details:
- Use `appointment.GetOrganizer()` which returns `AddressEntry` (COM object, must release)
- `appointment.Recipients` collection filtered by `OlMeetingRecipientType`
- `appointment.RecurrenceState` enum: olApptNotRecurring, olApptMaster, olApptOccurrence, olApptException
- `appointment.GlobalAppointmentID` as unique identifier for duplicate detection

**Commit:** `feat(processor): implement appointment metadata extraction with COM cleanup`

---

### Step 7: Implement duplicate detection

**Implement:**
- File: `SlingMD.Outlook/Services/AppointmentProcessor.cs`
- Add duplicate detection using `_processedAppointmentIds` ConcurrentDictionary:
  - Key: `GlobalAppointmentID` (unique per appointment instance)
  - Check before processing, add after successful save
  - Method: `IsDuplicateAppointment(string globalAppointmentId)`

Pattern mirrors EmailProcessor's `IsDuplicateEmail()` (line 949) but simpler since we use a single ID field.

**Commit:** `feat(processor): add duplicate detection via GlobalAppointmentID`

---

### Step 8: Implement note content building and file save

**Implement:**
- File: `SlingMD.Outlook/Services/AppointmentProcessor.cs`
- Build frontmatter dictionary with appointment metadata
- Build note content (placeholder - full template rendering comes in sub-spec 2)
- Write note via `_fileService.WriteUtf8File()`
- Handle `.ics` attachment filtering (skip attachments where filename ends with `.ics`)
- Process remaining attachments via `_attachmentService.ProcessAttachments()`

Key details:
- Frontmatter fields: title, type ("Appointment"), organizer, attendees, optionalAttendees, resources, location, startDateTime, endDateTime, recurrence, globalAppointmentId, tags
- File path: `Path.Combine(_settings.GetAppointmentsPath(), sanitizedFileName + ".md")`
- Use `_fileService.CleanFileName()` for sanitization
- Wrap in `using (var status = new StatusService())` for progress UI (skip if bulkMode)

**Commit:** `feat(processor): implement note building, file save, and .ics filtering`

---

### Step 9: Implement cancelled appointment filtering and StatusService

**Implement:**
- File: `SlingMD.Outlook/Services/AppointmentProcessor.cs`
- At start of ProcessAppointment, check:
  - If `!_settings.SaveCancelledAppointments` and `appointment.MeetingStatus == OlMeetingStatus.olMeetingCanceled`: skip
- StatusService integration:
  - If `!bulkMode`: wrap processing in `using (var status = new StatusService())`
  - Update progress at checkpoints: "Processing appointment..." (0%), "Building metadata" (25%), "Writing note" (50%), "Processing attachments" (75%), "Complete" (100%)
  - If bulkMode: skip StatusService entirely

**Commit:** `feat(processor): add cancelled filtering and progress UI`

---

## Interface Contracts

### Provides (to other sub-specs)
- **ObsidianSettings**: 10 new properties + `GetAppointmentsPath()` method (used by sub-specs 2-9)
- **AppointmentProcessor class**: Constructor, `ProcessAppointment()` method signature (used by sub-specs 3-8)
- **`_processedAppointmentIds`**: Static ConcurrentDictionary for duplicate detection (used by sub-spec 5)

### Requires (from other sub-specs)
- None (this is the foundation sub-spec)

### Shared State
- `ObsidianSettings` instance shared across all services
- `_processedAppointmentIds` static dictionary shared across AppointmentProcessor instances

## Verification Commands

### Per-Step
```bash
# After each step:
dotnet build SlingMD.sln --configuration Release
dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter "FullyQualifiedName~AppointmentProcessor"
dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter "FullyQualifiedName~ObsidianSettingsAppointment"
```

### Sub-Spec Acceptance
```bash
# [MECHANICAL] Build succeeds
dotnet build SlingMD.sln --configuration Release

# [STRUCTURAL] AppointmentProcessor.cs exists with ProcessAppointment()
grep -r "ProcessAppointment" SlingMD.Outlook/Services/AppointmentProcessor.cs

# [STRUCTURAL] ObsidianSettings has 10 new properties + GetAppointmentsPath()
grep -c "{ get; set; }" SlingMD.Outlook/Models/ObsidianSettings.cs
grep "GetAppointmentsPath" SlingMD.Outlook/Models/ObsidianSettings.cs

# [BEHAVIORAL] Settings round-trip
dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter "FullyQualifiedName~ObsidianSettingsAppointmentTests.RoundTrip"
```

### Integration Check
- Sub-specs 2-9 depend on this: verify `AppointmentProcessor` constructor compiles and `ObsidianSettings` properties are accessible
