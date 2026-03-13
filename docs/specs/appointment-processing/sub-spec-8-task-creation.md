---
type: phase-spec
master_spec: "docs/specs/2026-03-13-appointment-processing.md"
sub_spec: 8
title: "Task Creation Integration for Appointments"
dependencies: [1]
date: 2026-03-13
---

# Sub-Spec 8: Task Creation Integration for Appointments

## Shared Context

- **Master Spec:** [2026-03-13-appointment-processing.md](../2026-03-13-appointment-processing.md)
- **Trade-off hierarchy:** Match EmailProcessor patterns > Additive-only changes > Working end-to-end
- **Key constraint:** Mirror the exact pattern from EmailProcessor's task creation flow

## Codebase Analysis

### EmailProcessor Task Flow (lines 92-110, 446-449)

**Step 1: Get task options (line 92-110):**
```csharp
if (_settings.AskForDates)
{
    using (TaskOptionsForm form = new TaskOptionsForm(_settings))
    {
        DialogResult result = form.ShowDialog();
        if (result == DialogResult.OK)
        {
            _taskService.InitializeTaskSettings(
                form.DueDays, form.ReminderDays,
                form.ReminderHour, form.UseRelativeReminder);
        }
        else
        {
            _taskService.DisableTaskCreation();
        }
    }
}
else
{
    _taskService.InitializeTaskSettings(
        _settings.DefaultDueDays, _settings.DefaultReminderDays,
        _settings.DefaultReminderHour, _settings.UseRelativeReminder);
}
```

**Step 2: Generate Obsidian task (within note content building):**
```csharp
string taskBlock = string.Empty;
if (_settings.CreateObsidianTask && _taskService.ShouldCreateTasks)
{
    taskBlock = _taskService.GenerateObsidianTask(fileName, noteTags);
}
```

**Step 3: Create Outlook task (line 446-449):**
```csharp
if (_settings.CreateOutlookTask && _taskService.ShouldCreateTasks)
{
    _taskService.CreateOutlookTask(mail);
}
```

### TaskService.cs Methods

- `InitializeTaskSettings(dueDays, reminderDays, reminderHour, useRelativeReminder)` (line 27)
- `GenerateObsidianTask(fileName, tags)` (line 46): Returns `- [ ] [[{fileName}]] {tags} ... {dates}` line
- `CreateOutlookTask(MailItem mail)` (line 86): Creates TaskItem with subject "Follow up: {subject}"
- `ShouldCreateTasks` (line 36): Property gate
- `DisableTaskCreation()` (line 38)

### AppointmentTaskCreation Setting

- Values: "None", "Obsidian", "Outlook", "Both"
- In AppointmentProcessor, need to map this to the EmailProcessor's dual-flag pattern

### Files to Modify

| File | Action | Exists |
|------|--------|--------|
| `SlingMD.Outlook/Services/AppointmentProcessor.cs` | Add task creation step | Yes (from sub-spec 1) |

## Implementation Steps

### Step 1: Add task initialization to ProcessAppointment

**Test first:**
- File: `SlingMD.Tests/Services/AppointmentProcessorTests.cs`
- Test: `ProcessAppointment_TaskCreationNone_SkipsTaskDialog`
- Test: `ProcessAppointment_TaskCreationObsidian_CreatesTaskLine`

**Implement:**
- File: `SlingMD.Outlook/Services/AppointmentProcessor.cs`
- At the beginning of `ProcessAppointment()`, after cancelled check:

```csharp
// Initialize task creation based on AppointmentTaskCreation setting
bool createObsidianTask = _settings.AppointmentTaskCreation == "Obsidian"
                       || _settings.AppointmentTaskCreation == "Both";
bool createOutlookTask = _settings.AppointmentTaskCreation == "Outlook"
                      || _settings.AppointmentTaskCreation == "Both";

if (createObsidianTask || createOutlookTask)
{
    if (!bulkMode && _settings.AskForDates)
    {
        using (TaskOptionsForm form = new TaskOptionsForm(_settings))
        {
            System.Windows.Forms.DialogResult result = form.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                _taskService.InitializeTaskSettings(
                    form.DueDays, form.ReminderDays,
                    form.ReminderHour, form.UseRelativeReminder);
            }
            else
            {
                _taskService.DisableTaskCreation();
                createObsidianTask = false;
                createOutlookTask = false;
            }
        }
    }
    else
    {
        // Use defaults (always in bulk mode, or when AskForDates is off)
        _taskService.InitializeTaskSettings(
            _settings.DefaultDueDays, _settings.DefaultReminderDays,
            _settings.DefaultReminderHour, _settings.UseRelativeReminder);
    }
}
```

**Commit:** `feat(processor): add task initialization for appointments`

---

### Step 2: Generate Obsidian task block in note content

**Implement:**
- File: `SlingMD.Outlook/Services/AppointmentProcessor.cs`
- When building appointment note content (before template rendering):

```csharp
string taskBlock = string.Empty;
if (createObsidianTask && _taskService.ShouldCreateTasks)
{
    List<string> taskTags = _settings.AppointmentDefaultNoteTags ?? new List<string>();
    taskBlock = _taskService.GenerateObsidianTask(fileName, taskTags);
}
```

- Pass `taskBlock` to `AppointmentTemplateContext.TaskBlock` for rendering

**Commit:** `feat(processor): generate Obsidian task block for appointment notes`

---

### Step 3: Create Outlook task for appointments

**Implement:**
- File: `SlingMD.Outlook/Services/AppointmentProcessor.cs`
- After contact processing, add Outlook task creation:

```csharp
if (createOutlookTask && _taskService.ShouldCreateTasks)
{
    // TaskService.CreateOutlookTask takes MailItem -- need appointment-aware overload
    // For now, create task manually following TaskService pattern
    TaskItem task = null;
    try
    {
        task = Globals.ThisAddIn.Application.CreateItem(OlItemType.olTaskItem) as TaskItem;
        if (task != null)
        {
            task.Subject = $"Follow up: {appointment.Subject}";
            task.Body = $"Follow up on appointment: {appointment.Subject}\n" +
                       $"Date: {appointment.Start:yyyy-MM-dd HH:mm}\n" +
                       $"Location: {appointment.Location}";

            DateTime dueDate = DateTime.Today.AddDays(_settings.DefaultDueDays);
            task.DueDate = dueDate;

            int reminderDays = _settings.DefaultReminderDays;
            int reminderHour = _settings.DefaultReminderHour;
            DateTime reminderTime;
            if (_settings.UseRelativeReminder)
            {
                reminderTime = dueDate.AddDays(-reminderDays).Date.AddHours(reminderHour);
            }
            else
            {
                reminderTime = DateTime.Today.AddDays(reminderDays).Date.AddHours(reminderHour);
            }

            if (reminderTime <= DateTime.Now)
            {
                reminderTime = DateTime.Today.AddDays(1).AddHours(reminderHour);
            }

            task.ReminderSet = true;
            task.ReminderTime = reminderTime;
            task.Save();
        }
    }
    catch (System.Exception ex)
    {
        if (!bulkMode)
        {
            System.Windows.Forms.MessageBox.Show(
                $"Could not create Outlook task: {ex.Message}",
                "Task Creation Error",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Warning);
        }
    }
    finally
    {
        if (task != null)
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(task);
        }
    }
}
```

Note: TaskService.CreateOutlookTask takes `MailItem`. Rather than modifying TaskService (which would violate additive-only constraint), implement appointment task creation directly in AppointmentProcessor following the same pattern.

**Commit:** `feat(processor): create Outlook tasks for appointments`

---

### Step 4: Ensure bulk mode uses defaults without dialog

**Implement:**
- Verify the bulkMode check in step 1 properly skips `TaskOptionsForm`
- In bulk mode with `AppointmentTaskCreation != "None"`, use default due/reminder values
- Task errors in bulk mode go to `_bulkErrors` list (from sub-spec 5)

**Commit:** `feat(processor): ensure bulk mode uses task defaults without dialogs`

---

## Interface Contracts

### Provides (to other sub-specs)
- **Task creation in appointment notes**: Complete task pipeline for appointments

### Requires (from other sub-specs)
- **Sub-Spec 1**: AppointmentProcessor class, `AppointmentTaskCreation` setting, `AppointmentDefaultNoteTags`
- **Sub-Spec 2**: `AppointmentTemplateContext.TaskBlock` field (for rendering task in note)

### Shared State
- `TaskService` instance shared within AppointmentProcessor (initialized in constructor)

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

# [STRUCTURAL] Task creation logic exists
grep "AppointmentTaskCreation\|GenerateObsidianTask\|CreateItem.*olTaskItem" SlingMD.Outlook/Services/AppointmentProcessor.cs

# [BEHAVIORAL] Obsidian task creation
dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter "FullyQualifiedName~AppointmentProcessorTests.ProcessAppointment_TaskCreation"

# [BEHAVIORAL] Bulk mode skips dialog
dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter "FullyQualifiedName~AppointmentProcessorTests.ProcessAppointment_BulkMode"
```
