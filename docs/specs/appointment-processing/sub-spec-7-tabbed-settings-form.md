---
type: phase-spec
master_spec: "docs/specs/2026-03-13-appointment-processing.md"
sub_spec: 7
title: "Tabbed SettingsForm Rewrite"
dependencies: [1]
date: 2026-03-13
---

# Sub-Spec 7: Tabbed SettingsForm Rewrite

## Shared Context

- **Master Spec:** [2026-03-13-appointment-processing.md](../2026-03-13-appointment-processing.md)
- **Trade-off hierarchy:** Match EmailProcessor patterns > Additive-only changes > Working end-to-end
- **Key note:** SettingsForm is entirely programmatic (no Designer.cs file). All UI built in `InitializeComponent()`. Must preserve all existing setting save/load behavior.

## Codebase Analysis

### SettingsForm.cs (598 lines)

**Current structure:** Single scrollable form with 8 GroupBox sections stacked vertically:
1. Vault Settings (lines 107-131)
2. General Settings (lines 133-159)
3. Timing Settings (lines 161-180)
4. Subject Cleanup Patterns (lines 182-202)
5. Note & Tag Customization (lines 204-283)
6. Attachment Settings (lines 285-347)
7. Development Settings (lines 348-360)
8. Footer with Save/Cancel (lines 362-407)

**Key patterns:**
- `InitializeComponent()` (lines 88-422): Builds all UI programmatically
- `LoadSettings()` (lines 424-484): Copies `_settings` properties to controls
- `btnSave_Click()` (lines 500-554): Copies control values back to `_settings`, calls `_settings.Save()`
- Event handlers for conditional enable/disable (lines 268-280, 342-344, 355, 595-598)
- Controls are private fields (lines 15-79)
- Form size: 760x820, FormBorderStyle.Sizable

**Save/Load binding pattern:**
```csharp
// Load:
txtVaultName.Text = _settings.VaultName;
chkLaunchObsidian.Checked = _settings.LaunchObsidian;
numDelay.Value = _settings.ObsidianDelaySeconds;

// Save:
_settings.VaultName = txtVaultName.Text;
_settings.LaunchObsidian = chkLaunchObsidian.Checked;
_settings.ObsidianDelaySeconds = (int)numDelay.Value;
```

### Target Tab Structure

| Tab | Settings (from current GroupBoxes) | New? |
|-----|-----------------------------------|------|
| General | Vault path, vault name, launch Obsidian, countdown, delay, templates folder | Existing fields reorganized |
| Email | Inbox folder, note title format, max length, include date, default note tags, subject cleanup patterns, email filename format, email template file | Existing fields reorganized |
| Appointments | Appointments folder, title format, max length, default tags, save attachments, meeting notes, meeting note template, group recurring, save cancelled, task creation mode | NEW (sub-spec 1 properties) |
| Contacts | Contacts folder, enable contact saving, vault-wide search, contact filename format, contact template file | Existing fields reorganized |
| Tasks | Create Obsidian task, create Outlook task, ask for dates, due days, reminder days, reminder hour, task tags, task template file | Existing fields reorganized |
| Threading | Group email threads, thread debug, thread template file, move date to front | Existing fields reorganized |
| Attachments | Save inline images, save all attachments, use wikilinks, storage mode, attachments folder | Existing fields reorganized |
| Developer | Show development settings, show thread debug | Existing fields reorganized |

### Files to Modify

| File | Action | Exists |
|------|--------|--------|
| `SlingMD.Outlook/Forms/SettingsForm.cs` | Rewrite InitializeComponent with TabControl | Yes (598 lines) |

## Implementation Steps

### Step 1: Refactor InitializeComponent to use TabControl

**Implement:**
- File: `SlingMD.Outlook/Forms/SettingsForm.cs`
- Replace the current `rootLayout` + `mainLayout` + GroupBoxes structure with:

```csharp
// Root layout: TabControl + Footer
TableLayoutPanel rootLayout = new TableLayoutPanel
{
    Dock = DockStyle.Fill,
    RowCount = 2,
    ColumnCount = 1
};
rootLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));  // Tab content
rootLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));        // Footer

TabControl tabControl = new TabControl
{
    Dock = DockStyle.Fill
};

// Create 8 tab pages
TabPage tabGeneral = new TabPage("General");
TabPage tabEmail = new TabPage("Email");
TabPage tabAppointments = new TabPage("Appointments");
TabPage tabContacts = new TabPage("Contacts");
TabPage tabTasks = new TabPage("Tasks");
TabPage tabThreading = new TabPage("Threading");
TabPage tabAttachments = new TabPage("Attachments");
TabPage tabDeveloper = new TabPage("Developer");

tabControl.TabPages.AddRange(new TabPage[] {
    tabGeneral, tabEmail, tabAppointments, tabContacts,
    tabTasks, tabThreading, tabAttachments, tabDeveloper
});
```

Each tab page gets its own `TableLayoutPanel` with the relevant controls. Move existing controls from GroupBoxes into the appropriate tab's layout panel.

**Key: Preserve all existing control field declarations (lines 15-79) and their event handlers. Only change which container they're added to.**

**Commit:** `refactor(settings): replace GroupBox layout with TabControl (8 tabs)`

---

### Step 2: Build General tab

**Implement:**
- File: `SlingMD.Outlook/Forms/SettingsForm.cs`
- General tab contains:
  - Vault Name (TextBox)
  - Vault Base Path (TextBox + Browse button)
  - Launch Obsidian (CheckBox)
  - Show Countdown (CheckBox)
  - Delay seconds (NumericUpDown)
  - Templates Folder (TextBox)
  - Include Daily Note Link (CheckBox)
  - Daily Note Link Format (TextBox)

Move these existing controls from their current GroupBoxes into the General tab's layout panel.

**Commit:** `refactor(settings): populate General tab with vault and launch settings`

---

### Step 3: Build Email tab

**Implement:**
- Email tab contains:
  - Inbox Folder (TextBox)
  - Note Title Format (TextBox) with format tokens help
  - Max Title Length (NumericUpDown)
  - Include Date in Title (CheckBox)
  - Default Note Tags (TextBox, comma-separated)
  - Subject Cleanup Patterns (ListBox + Add/Edit/Remove buttons)
  - Email Filename Format (TextBox)
  - Email Template File (TextBox)

**Commit:** `refactor(settings): populate Email tab with email-specific settings`

---

### Step 4: Build Appointments tab (NEW)

**Implement:**
- Appointments tab contains all new controls bound to sub-spec 1 settings:

```csharp
// Appointments Folder
TextBox txtAppointmentsFolder = new TextBox { ... };

// Note Title Format
TextBox txtAppointmentNoteTitleFormat = new TextBox { ... };

// Max Title Length
NumericUpDown numAppointmentTitleMaxLength = new NumericUpDown { Minimum = 10, Maximum = 500, Value = 50 };

// Default Tags
TextBox txtAppointmentDefaultTags = new TextBox { ... };

// Save Attachments
CheckBox chkAppointmentSaveAttachments = new CheckBox { Text = "Save attachments", Checked = true };

// Create Meeting Notes
CheckBox chkCreateMeetingNotes = new CheckBox { Text = "Create companion meeting notes", Checked = true };

// Meeting Note Template
TextBox txtMeetingNoteTemplate = new TextBox { ... };

// Group Recurring Meetings
CheckBox chkGroupRecurringMeetings = new CheckBox { Text = "Group recurring meeting instances", Checked = true };

// Save Cancelled Appointments
CheckBox chkSaveCancelledAppointments = new CheckBox { Text = "Save cancelled appointments", Checked = false };

// Task Creation Mode
ComboBox cmbAppointmentTaskCreation = new ComboBox
{
    DropDownStyle = ComboBoxStyle.DropDownList,
    Items = { "None", "Obsidian", "Outlook", "Both" }
};
```

Add control field declarations at the class level (alongside existing fields).

**Commit:** `feat(settings): add Appointments tab with all appointment controls`

---

### Step 5: Build Contacts, Tasks, Threading, Attachments, Developer tabs

**Implement:**
- **Contacts tab**: Contacts folder, enable contact saving, vault-wide search, contact filename format, contact template file
- **Tasks tab**: Create Obsidian task, create Outlook task, ask for dates, due days, reminder days, reminder hour, default task tags, task template file
- **Threading tab**: Group email threads, move date to front in thread, thread template file, show thread debug (conditional visibility)
- **Attachments tab**: Storage mode dropdown, attachments folder (conditional), save inline images, save all attachments, use wikilinks
- **Developer tab**: Show development settings, show thread debug

Move existing controls to their new tabs.

**Commit:** `refactor(settings): populate remaining tabs (Contacts, Tasks, Threading, Attachments, Developer)`

---

### Step 6: Update LoadSettings for appointment controls

**Implement:**
- File: `SlingMD.Outlook/Forms/SettingsForm.cs`
- Add to `LoadSettings()`:

```csharp
// Appointment settings
txtAppointmentsFolder.Text = _settings.AppointmentsFolder;
txtAppointmentNoteTitleFormat.Text = _settings.AppointmentNoteTitleFormat;
numAppointmentTitleMaxLength.Value = _settings.AppointmentNoteTitleMaxLength;
txtAppointmentDefaultTags.Text = string.Join(", ", _settings.AppointmentDefaultNoteTags);
chkAppointmentSaveAttachments.Checked = _settings.AppointmentSaveAttachments;
chkCreateMeetingNotes.Checked = _settings.CreateMeetingNotes;
txtMeetingNoteTemplate.Text = _settings.MeetingNoteTemplate ?? string.Empty;
chkGroupRecurringMeetings.Checked = _settings.GroupRecurringMeetings;
chkSaveCancelledAppointments.Checked = _settings.SaveCancelledAppointments;
cmbAppointmentTaskCreation.SelectedItem = _settings.AppointmentTaskCreation;
```

**Commit:** `feat(settings): add LoadSettings for appointment controls`

---

### Step 7: Update btnSave_Click for appointment controls

**Implement:**
- File: `SlingMD.Outlook/Forms/SettingsForm.cs`
- Add to `btnSave_Click()`:

```csharp
// Appointment settings
_settings.AppointmentsFolder = txtAppointmentsFolder.Text;
_settings.AppointmentNoteTitleFormat = txtAppointmentNoteTitleFormat.Text;
_settings.AppointmentNoteTitleMaxLength = (int)numAppointmentTitleMaxLength.Value;
_settings.AppointmentDefaultNoteTags = txtAppointmentDefaultTags.Text
    .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
    .Select(t => t.Trim())
    .Where(t => !string.IsNullOrEmpty(t))
    .ToList();
_settings.AppointmentSaveAttachments = chkAppointmentSaveAttachments.Checked;
_settings.CreateMeetingNotes = chkCreateMeetingNotes.Checked;
_settings.MeetingNoteTemplate = txtMeetingNoteTemplate.Text;
_settings.GroupRecurringMeetings = chkGroupRecurringMeetings.Checked;
_settings.SaveCancelledAppointments = chkSaveCancelledAppointments.Checked;
_settings.AppointmentTaskCreation = cmbAppointmentTaskCreation.SelectedItem?.ToString() ?? "None";
```

**Commit:** `feat(settings): add Save handler for appointment controls`

---

### Step 8: Verify all existing settings still save/load correctly

**Run full test suite:**
```bash
dotnet test SlingMD.Tests\SlingMD.Tests.csproj
```

Manually verify: Open settings, change values across all tabs, save, close, reopen -- all values should persist. Focus on:
- Email settings (most risk of regression from reorganization)
- Subject cleanup patterns (ListBox with Add/Edit/Remove)
- Attachment storage mode conditional visibility
- Development settings conditional visibility

**Commit:** `test(settings): verify all existing settings round-trip after tab rewrite`

---

## Interface Contracts

### Provides (to other sub-specs)
- **Appointments tab UI**: Binds to ObsidianSettings appointment properties from sub-spec 1
- **Tabbed SettingsForm**: All 8 tabs functional

### Requires (from other sub-specs)
- **Sub-Spec 1**: ObsidianSettings appointment properties (AppointmentsFolder, AppointmentNoteTitleFormat, etc.)

## Verification Commands

### Per-Step
```bash
dotnet build SlingMD.sln --configuration Release
```

### Sub-Spec Acceptance
```bash
# [MECHANICAL] Build succeeds
dotnet build SlingMD.sln --configuration Release

# [STRUCTURAL] SettingsForm has 8 tabs
grep -c "TabPage" SlingMD.Outlook/Forms/SettingsForm.cs

# [BEHAVIORAL] All existing settings still save/load
dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter "FullyQualifiedName~ObsidianSettingsTests"

# [BEHAVIORAL] New appointment settings save/load
dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter "FullyQualifiedName~ObsidianSettingsAppointmentTests"

# [HUMAN REVIEW] Visual inspection of tab layout
```
