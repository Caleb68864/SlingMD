---
type: phase-spec
master_spec: "docs/specs/2026-03-13-appointment-processing.md"
sub_spec: 5
title: "Bulk Save Today's Appointments"
dependencies: [1]
date: 2026-03-13
---

# Sub-Spec 5: Bulk "Save Today's Appointments"

## Shared Context

- **Master Spec:** [2026-03-13-appointment-processing.md](../2026-03-13-appointment-processing.md)
- **Trade-off hierarchy:** Match EmailProcessor patterns > Additive-only changes > Working end-to-end
- **Key consideration:** Bulk mode must suppress ALL dialogs, collect errors, and show a single summary at the end

## Codebase Analysis

### ThisAddIn.cs Patterns (148 lines)

- `ProcessSelectedEmail()` (line 90): Gets `Explorer.Selection[1]`, validates selection count, casts to MailItem, calls `_emailProcessor.ProcessEmail(mail)`
- Settings lifecycle: Load on startup (line 30), recreate EmailProcessor after settings change (line 128), save on shutdown (line 42)
- Error handling: try/catch with `MessageBox.Show()` (line 112)
- `Application.Session.Accounts` available for iterating accounts

### EmailProcessor Bulk Patterns

EmailProcessor doesn't have a formal bulk mode, but appointment bulk mode needs:
- Suppress `TaskOptionsForm` dialog
- Suppress `StatusService` / `CountdownForm`
- Suppress individual `MessageBox.Show()` errors
- Suppress individual Obsidian launches
- Collect errors in `List<string>` for summary
- Return processing result for counting

### DASL Filtering Pattern

Outlook DASL filtering for calendar items:
```csharp
string filter = $"[Start] >= '{today:g}' AND [End] <= '{tomorrow:g}'";
items.IncludeRecurrences = true;
items.Sort("[Start]");
Items restricted = items.Restrict(filter);
```

### Files to Modify

| File | Action | Exists |
|------|--------|--------|
| `SlingMD.Outlook/ThisAddIn.cs` | Add `SaveTodaysAppointments()` method | Yes (148 lines) |
| `SlingMD.Outlook/Services/AppointmentProcessor.cs` | Ensure bulkMode flag works correctly | Yes (from sub-spec 1) |

## Implementation Steps

### Step 1: Define AppointmentProcessingResult

**Implement:**
- File: `SlingMD.Outlook/Services/AppointmentProcessor.cs`
- Add result enum/class at top of file:

```csharp
public enum AppointmentProcessingResult
{
    Success,
    Skipped,       // Duplicate or cancelled
    Error
}
```

- Modify `ProcessAppointment()` return type to `Task<AppointmentProcessingResult>` (or add an overload)
- In bulk mode, return appropriate result instead of showing MessageBox

**Commit:** `feat(processor): add AppointmentProcessingResult for bulk reporting`

---

### Step 2: Implement bulkMode suppression in AppointmentProcessor

**Test first:**
- File: `SlingMD.Tests/Services/AppointmentProcessorTests.cs`
- Test: `ProcessAppointment_BulkMode_SuppressesDialogs`

**Implement:**
- File: `SlingMD.Outlook/Services/AppointmentProcessor.cs`
- Add error collection field:

```csharp
private List<string> _bulkErrors = new List<string>();

public List<string> GetBulkErrors()
{
    List<string> errors = new List<string>(_bulkErrors);
    _bulkErrors.Clear();
    return errors;
}
```

- In `ProcessAppointment()`, gate dialogs on `!bulkMode`:
  - Skip `TaskOptionsForm` if bulkMode (use default settings)
  - Skip `StatusService` creation if bulkMode
  - Skip `MessageBox.Show()` errors if bulkMode (add to `_bulkErrors` instead)
  - Skip Obsidian launch if bulkMode
  - Skip countdown if bulkMode

**Commit:** `feat(processor): implement bulkMode dialog suppression and error collection`

---

### Step 3: Implement SaveTodaysAppointments in ThisAddIn

**Implement:**
- File: `SlingMD.Outlook/ThisAddIn.cs`
- Add `_appointmentProcessor` field alongside `_emailProcessor`:

```csharp
private AppointmentProcessor _appointmentProcessor;
```

- Initialize in Startup alongside EmailProcessor:

```csharp
_appointmentProcessor = new AppointmentProcessor(_settings);
```

- Add `SaveTodaysAppointments()` method:

```csharp
public async void SaveTodaysAppointments()
{
    int saved = 0;
    int skipped = 0;
    int errors = 0;
    int total = 0;

    try
    {
        Accounts accounts = Application.Session.Accounts;
        try
        {
            foreach (Account account in accounts)
            {
                MAPIFolder calendar = null;
                Items items = null;
                Items restricted = null;
                try
                {
                    calendar = account.Session.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
                    items = calendar.Items;
                    items.IncludeRecurrences = true;
                    items.Sort("[Start]");

                    DateTime today = DateTime.Today;
                    DateTime tomorrow = today.AddDays(1);
                    string filter = string.Format(
                        "[Start] >= '{0}' AND [Start] < '{1}'",
                        today.ToString("g"),
                        tomorrow.ToString("g"));
                    restricted = items.Restrict(filter);

                    foreach (object item in restricted)
                    {
                        AppointmentItem appointment = item as AppointmentItem;
                        if (appointment == null) continue;

                        try
                        {
                            total++;

                            // Filter cancelled if setting is off
                            if (!_settings.SaveCancelledAppointments &&
                                appointment.MeetingStatus == OlMeetingStatus.olMeetingCanceled)
                            {
                                skipped++;
                                continue;
                            }

                            AppointmentProcessingResult result =
                                await _appointmentProcessor.ProcessAppointment(
                                    appointment, bulkMode: true);

                            switch (result)
                            {
                                case AppointmentProcessingResult.Success:
                                    saved++;
                                    break;
                                case AppointmentProcessingResult.Skipped:
                                    skipped++;
                                    break;
                                case AppointmentProcessingResult.Error:
                                    errors++;
                                    break;
                            }
                        }
                        finally
                        {
                            if (appointment != null)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(appointment);
                        }
                    }
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    errors++;
                    // Per-account error -- continue with remaining accounts
                }
                finally
                {
                    if (restricted != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(restricted);
                    if (items != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(items);
                    if (calendar != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(calendar);
                }
            }
        }
        finally
        {
            if (accounts != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(accounts);
        }

        // Show summary
        List<string> bulkErrors = _appointmentProcessor.GetBulkErrors();
        string summary = $"Saved {saved}/{total} appointments.\n" +
                         $"Skipped: {skipped} (duplicates/cancelled)\n" +
                         $"Errors: {errors}";

        if (bulkErrors.Count > 0)
        {
            summary += "\n\nError details:\n" + string.Join("\n", bulkErrors);
        }

        System.Windows.Forms.MessageBox.Show(
            summary,
            "Save Today's Appointments",
            System.Windows.Forms.MessageBoxButtons.OK,
            System.Windows.Forms.MessageBoxIcon.Information);

        // Optional: single Obsidian launch to appointments folder
        if (_settings.LaunchObsidian && saved > 0)
        {
            _fileService.LaunchObsidian(_settings.GetAppointmentsPath());
        }
    }
    catch (System.Exception ex)
    {
        System.Windows.Forms.MessageBox.Show(
            $"Error saving today's appointments: {ex.Message}",
            "Error",
            System.Windows.Forms.MessageBoxButtons.OK,
            System.Windows.Forms.MessageBoxIcon.Error);
    }
}
```

**Commit:** `feat(addin): implement SaveTodaysAppointments with DASL filtering and summary`

---

### Step 4: Add FileService reference to ThisAddIn

**Implement:**
- File: `SlingMD.Outlook/ThisAddIn.cs`
- Need FileService for LaunchObsidian call. Either:
  - Create a local FileService instance for the launch call, or
  - Add a `LaunchObsidianToFolder(string path)` method to AppointmentProcessor that delegates

Prefer the simpler approach: create local FileService.

```csharp
private FileService _fileService;
// In Startup:
_fileService = new FileService(_settings);
```

**Commit:** `feat(addin): add FileService for Obsidian launch in bulk mode`

---

### Step 5: Recreate AppointmentProcessor on settings change

**Implement:**
- File: `SlingMD.Outlook/ThisAddIn.cs`
- In the settings save handler (around line 128), add:

```csharp
_appointmentProcessor = new AppointmentProcessor(_settings);
```

Alongside the existing `_emailProcessor = new EmailProcessor(_settings);`

**Commit:** `feat(addin): recreate AppointmentProcessor on settings change`

---

## Interface Contracts

### Provides (to other sub-specs)
- **`SaveTodaysAppointments()` method in ThisAddIn**: Called by SlingRibbon button (sub-spec 6)
- **`_appointmentProcessor` field in ThisAddIn**: Used by sub-spec 6 for routing
- **`AppointmentProcessingResult` enum**: Return type for bulk processing

### Requires (from other sub-specs)
- **Sub-Spec 1**: AppointmentProcessor class, `ProcessAppointment(appointment, bulkMode: true)`, `SaveCancelledAppointments` setting

## Verification Commands

### Per-Step
```bash
dotnet build SlingMD.sln --configuration Release
```

### Sub-Spec Acceptance
```bash
# [MECHANICAL] Build succeeds
dotnet build SlingMD.sln --configuration Release

# [STRUCTURAL] ThisAddIn has SaveTodaysAppointments
grep "SaveTodaysAppointments" SlingMD.Outlook/ThisAddIn.cs

# [STRUCTURAL] AppointmentProcessor has bulkMode parameter
grep "bulkMode" SlingMD.Outlook/Services/AppointmentProcessor.cs

# [BEHAVIORAL] Bulk mode suppresses dialogs
dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter "FullyQualifiedName~AppointmentProcessorTests.ProcessAppointment_BulkMode"
```
