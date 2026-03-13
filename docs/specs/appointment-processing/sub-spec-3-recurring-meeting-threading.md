---
type: phase-spec
master_spec: "docs/specs/2026-03-13-appointment-processing.md"
sub_spec: 3
title: "Recurring Meeting Threading"
dependencies: [1]
date: 2026-03-13
---

# Sub-Spec 3: Recurring Meeting Threading

## Shared Context

- **Master Spec:** [2026-03-13-appointment-processing.md](../2026-03-13-appointment-processing.md)
- **Trade-off hierarchy:** Match EmailProcessor patterns > Additive-only changes > Working end-to-end
- **Key constraint:** Use ThreadService as-is for folder creation and summary notes. AppointmentProcessor handles the routing logic.

## Codebase Analysis

### ThreadService.cs Patterns

- `GetConversationId()` (line 40): Derives 20-char MD5 hash from conversation topic -- for appointments, we'll derive from RecurrencePattern + subject
- `GetThreadFolderName()` (line 91): Returns `{subject}-{sender}-{recipient}` -- for appointments, we'll use `{subject}` only (no sender/recipient in appointment threading)
- `MoveToThreadFolder()` (line 176): Moves file from inbox to thread folder, appends `-eid{ID}` suffix -- can reuse for appointment notes
- `ResuffixThreadNotes()` (line 410): Renames all notes in thread folder with `yyyy-MM-dd_HHmmss` prefix -- reuse for chronological ordering
- `UpdateThreadNote()` (line 138): Creates/refreshes `0-{threadname}.md` summary -- reuse for recurring meeting summaries

### EmailProcessor Threading Pattern (lines 272-400)

1. Get semaphore for thread folder via `_threadFolderLocks.GetOrAdd()`
2. Acquire semaphore: `await semaphore.WaitAsync()`
3. Build temp filename with ID suffix
4. Write file to thread folder
5. Scan for other related files and move them
6. `ResuffixThreadNotes()` for chronological ordering
7. `UpdateThreadNote()` for summary
8. Release semaphore in finally block

### RecurrenceState Values

- `olApptNotRecurring` (0): Single appointment
- `olApptMaster` (1): Series master
- `olApptOccurrence` (2): Instance of recurring series
- `olApptException` (3): Modified instance of series

### Files to Modify

| File | Action | Exists |
|------|--------|--------|
| `SlingMD.Outlook/Services/AppointmentProcessor.cs` | Add recurring detection + ThreadService routing | Yes (from sub-spec 1) |

## Implementation Steps

### Step 1: Add recurring meeting detection

**Test first:**
- File: `SlingMD.Tests/Services/AppointmentProcessorTests.cs`
- Test: `ProcessAppointment_RecurringInstance_WithGroupingEnabled_CreatesThreadFolder`
- Test: `ProcessAppointment_NonRecurring_WritesToFlatFolder`

**Implement:**
- File: `SlingMD.Outlook/Services/AppointmentProcessor.cs`
- Add private method to determine if appointment should be threaded:

```csharp
private bool ShouldGroupAsRecurring(AppointmentItem appointment)
{
    if (!_settings.GroupRecurringMeetings)
    {
        return false;
    }

    try
    {
        return appointment.RecurrenceState == OlRecurrenceState.olApptOccurrence
            || appointment.RecurrenceState == OlRecurrenceState.olApptException;
    }
    catch (System.Runtime.InteropServices.COMException)
    {
        return false;
    }
}
```

**Commit:** `feat(processor): add recurring meeting detection`

---

### Step 2: Generate thread folder name for recurring meetings

**Implement:**
- File: `SlingMD.Outlook/Services/AppointmentProcessor.cs`
- Add method to generate stable thread folder name from recurring series:

```csharp
private string GetRecurringThreadFolderName(AppointmentItem appointment)
{
    string cleanSubject = CleanSubject(appointment.Subject);
    RecurrencePattern pattern = null;
    try
    {
        pattern = appointment.GetRecurrencePattern();
        string patternStart = pattern.PatternStartDate.ToString("yyyy-MM-dd");
        return _fileService.CleanFileName($"{cleanSubject}");
    }
    catch (System.Runtime.InteropServices.COMException)
    {
        return _fileService.CleanFileName(cleanSubject);
    }
    finally
    {
        if (pattern != null)
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(pattern);
        }
    }
}
```

Note: Use cleaned subject as folder name (simpler than email threading which uses subject-sender-recipient). The RecurrencePattern's PatternStartDate can be used for generating a stable conversation ID if needed.

**Commit:** `feat(processor): generate recurring meeting thread folder names`

---

### Step 3: Implement thread folder routing in ProcessAppointment

**Implement:**
- File: `SlingMD.Outlook/Services/AppointmentProcessor.cs`
- In `ProcessAppointment()`, after note content is built:

```csharp
if (ShouldGroupAsRecurring(appointment))
{
    string threadFolderName = GetRecurringThreadFolderName(appointment);
    string threadFolderPath = Path.Combine(_settings.GetAppointmentsPath(), threadFolderName);
    _fileService.EnsureDirectoryExists(threadFolderPath);

    // Date-stamped filename for instance
    string dateStamp = appointment.Start.ToString("yyyy-MM-dd");
    string instanceFileName = $"{dateStamp} - {cleanSubject}.md";

    // Handle same-day collision (append time suffix)
    string instanceFilePath = Path.Combine(threadFolderPath, instanceFileName);
    if (File.Exists(instanceFilePath))
    {
        string timeSuffix = appointment.Start.ToString("_HHmm");
        instanceFileName = $"{dateStamp}{timeSuffix} - {cleanSubject}.md";
        instanceFilePath = Path.Combine(threadFolderPath, instanceFileName);
    }

    _fileService.WriteUtf8File(instanceFilePath, renderedContent);
}
else
{
    // Write flat to AppointmentsFolder
    string filePath = Path.Combine(_settings.GetAppointmentsPath(), fileName + ".md");
    _fileService.WriteUtf8File(filePath, renderedContent);
}
```

Follow EmailProcessor's semaphore pattern for thread safety:
- Use static `ConcurrentDictionary<string, SemaphoreSlim>` for thread folder locks
- Acquire before folder operations, release in finally

**Commit:** `feat(processor): route recurring appointments to thread folders`

---

### Step 4: Update thread summary note for recurring meetings

**Implement:**
- File: `SlingMD.Outlook/Services/AppointmentProcessor.cs`
- After writing instance note to thread folder, update the summary note:

```csharp
// Generate thread summary (0-threadname.md)
string threadNoteName = $"0-{threadFolderName}";
ThreadTemplateContext threadContext = new ThreadTemplateContext
{
    Title = cleanSubject,
    ThreadId = threadConversationId,
    FolderPath = threadFolderPath
};
string threadContent = _templateService.RenderThreadContent(threadContext);
string threadNotePath = Path.Combine(threadFolderPath, threadNoteName + ".md");
_fileService.WriteUtf8File(threadNotePath, threadContent);
```

This reuses the existing ThreadService/TemplateService thread summary pattern. The DataviewJS query in the thread template will automatically pick up all notes in the folder.

**Commit:** `feat(processor): generate thread summary for recurring meetings`

---

### Step 5: Handle series master warning

**Implement:**
- File: `SlingMD.Outlook/Services/AppointmentProcessor.cs`
- At the start of `ProcessAppointment()`, check for series master:

```csharp
if (appointment.RecurrenceState == OlRecurrenceState.olApptMaster)
{
    if (!bulkMode)
    {
        System.Windows.Forms.DialogResult result = System.Windows.Forms.MessageBox.Show(
            "You've selected the recurring series master. Would you like to save the next upcoming instance instead?",
            "Recurring Series",
            System.Windows.Forms.MessageBoxButtons.YesNo,
            System.Windows.Forms.MessageBoxIcon.Question);

        if (result == System.Windows.Forms.DialogResult.No)
        {
            return;
        }

        // Get next upcoming instance
        RecurrencePattern pattern = null;
        AppointmentItem nextInstance = null;
        try
        {
            pattern = appointment.GetRecurrencePattern();
            nextInstance = pattern.GetOccurrence(DateTime.Today);
            await ProcessAppointment(nextInstance, bulkMode, cancellationToken);
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            System.Windows.Forms.MessageBox.Show(
                "Could not find the next upcoming instance of this series.",
                "Error",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Warning);
        }
        finally
        {
            if (nextInstance != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(nextInstance);
            if (pattern != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(pattern);
        }
    }
    return; // In bulk mode, skip series masters entirely
}
```

**Commit:** `feat(processor): handle series master with upcoming instance redirect`

---

### Step 6: Handle deleted instances (bulk mode edge case)

**Implement:**
- File: `SlingMD.Outlook/Services/AppointmentProcessor.cs`
- Wrap RecurrencePattern access in try/catch for COMException:

```csharp
// In bulk processing, deleted instances throw COMException
// Catch and skip gracefully
try
{
    // ... access RecurrencePattern properties
}
catch (System.Runtime.InteropServices.COMException ex)
{
    if (!bulkMode)
    {
        System.Windows.Forms.MessageBox.Show(
            $"This recurring instance may have been deleted from the series: {ex.Message}",
            "Deleted Instance",
            System.Windows.Forms.MessageBoxButtons.OK,
            System.Windows.Forms.MessageBoxIcon.Information);
    }
    return; // Skip deleted instances
}
```

**Commit:** `feat(processor): handle deleted recurring instances gracefully`

---

## Interface Contracts

### Provides (to other sub-specs)
- **Thread folder routing logic**: Sub-spec 4 (meeting notes) needs to know the thread folder path to place companion notes correctly
- **`ShouldGroupAsRecurring()` method**: Internal helper used by sub-spec 4

### Requires (from other sub-specs)
- **Sub-Spec 1**: AppointmentProcessor class, ProcessAppointment method, CleanSubject method, ObsidianSettings.GroupRecurringMeetings property

### Shared State
- Static `ConcurrentDictionary<string, SemaphoreSlim>` for thread folder locks (same pattern as EmailProcessor's `_threadFolderLocks`)

## Verification Commands

### Per-Step
```bash
dotnet build SlingMD.sln --configuration Release
dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter "FullyQualifiedName~AppointmentProcessorTests"
```

### Sub-Spec Acceptance
```bash
# [MECHANICAL] Build succeeds
dotnet build SlingMD.sln --configuration Release

# [STRUCTURAL] AppointmentProcessor checks RecurrenceState
grep "RecurrenceState" SlingMD.Outlook/Services/AppointmentProcessor.cs

# [STRUCTURAL] ThreadService integration
grep "ThreadService\|ThreadTemplateContext\|RenderThreadContent" SlingMD.Outlook/Services/AppointmentProcessor.cs

# [BEHAVIORAL] Recurring instance routing
dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter "FullyQualifiedName~AppointmentProcessorTests.ProcessAppointment_Recurring"
```
