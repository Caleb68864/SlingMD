---
type: phase-spec
master_spec: "docs/specs/2026-03-13-appointment-processing.md"
sub_spec: 6
title: "Ribbon Extensions"
dependencies: [1, 5]
date: 2026-03-13
---

# Sub-Spec 6: Ribbon Extensions

## Shared Context

- **Master Spec:** [2026-03-13-appointment-processing.md](../2026-03-13-appointment-processing.md)
- **Trade-off hierarchy:** Match EmailProcessor patterns > Additive-only changes > Working end-to-end
- **Key note:** SlingRibbon loads XML from embedded resource. Adding inspector-level customUI requires `GetCustomUI()` routing by ribbonID.

## Codebase Analysis

### SlingRibbon.cs (109 lines)

- Class implements `Office.IRibbonExtensibility` (line 11)
- Constructor takes `ThisAddIn addIn` (line 18), stores as `_addIn` (line 15)
- `GetCustomUI(string ribbonID)` (line 46): Currently returns single XML from `GetResourceText("SlingMD.Outlook.Ribbon.SlingRibbon.xml")` regardless of ribbonID
- `OnSlingButtonClick(Office.IRibbonControl control)` (line 60): Calls `_addIn.ProcessSelectedEmail()`
- `OnSettingsButtonClick(Office.IRibbonControl control)` (line 72): Calls `_addIn.ShowSettings()`
- `GetSlingButtonImage()` (line 84): Returns cached bitmap from embedded resource
- `GetResourceText()` (line 93): Loads XML from manifest resource stream

### SlingRibbon.xml (22 lines)

- Namespace: `http://schemas.microsoft.com/office/2009/07/customui`
- Single `<tab id="SlingTab" label="Sling">`
- Two groups: `EmailGroup` (Sling button) and `SettingsGroup` (Configure button)
- Sling button uses `getImage="GetSlingButtonImage"` callback
- Configure button uses `imageMso="ControlProperties"`

### ThisAddIn.cs (148 lines)

- `ProcessSelectedEmail()` (line 90): Gets `Explorer.Selection[1]`, validates, casts to MailItem
- Needs: `ProcessSelection()` to replace `ProcessSelectedEmail()` (handles both MailItem and AppointmentItem)
- Needs: `ProcessCurrentAppointment()` for inspector-level Sling button

### GetCustomUI Routing

Outlook calls `GetCustomUI()` with different ribbonIDs:
- `"Microsoft.Outlook.Explorer"` -- main Outlook window
- `"Microsoft.Outlook.Appointment"` -- appointment inspector window
- Other inspector types also possible

### Files to Modify

| File | Action | Exists |
|------|--------|--------|
| `SlingMD.Outlook/Ribbon/SlingRibbon.xml` | Add Appointments group, keep as Explorer ribbon | Yes (22 lines) |
| `SlingMD.Outlook/Ribbon/SlingRibbon.cs` | Add handlers, GetCustomUI routing, inspector ribbon | Yes (109 lines) |
| `SlingMD.Outlook/ThisAddIn.cs` | Add ProcessSelection, ProcessCurrentAppointment | Yes (from sub-spec 5) |

## Implementation Steps

### Step 1: Add Appointments group to Explorer ribbon XML

**Implement:**
- File: `SlingMD.Outlook/Ribbon/SlingRibbon.xml`
- Add new group after SettingsGroup:

```xml
<group id="AppointmentsGroup" label="Appointments">
    <button id="SaveTodaysButton"
            label="Save Today"
            size="large"
            imageMso="ShowSchedulingPage"
            onAction="OnSaveTodaysClick"
            supertip="Save all of today's appointments to Obsidian"/>
</group>
```

`imageMso="ShowSchedulingPage"` is a built-in calendar icon. Alternative: `"AppointmentColor"` or `"GoToCalendar"`.

**Commit:** `feat(ribbon): add Appointments group with Save Today button`

---

### Step 2: Create inspector ribbon XML for appointments

**Implement:**
- File: `SlingMD.Outlook/Ribbon/SlingRibbonAppointment.xml` (NEW embedded resource)
- Or inline in `GetCustomUI()` method as string literal

Since the existing pattern uses embedded XML resources, create a new XML file:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"
          onLoad="Ribbon_Load">
    <ribbon>
        <tabs>
            <tab idMso="TabAppointment">
                <group id="SlingAppointmentGroup" label="Sling" insertBeforeMso="GroupActions">
                    <button id="InspectorSlingButton"
                            label="Sling"
                            size="large"
                            getImage="GetSlingButtonImage"
                            onAction="OnInspectorSlingClick"
                            supertip="Save this appointment to Obsidian as a markdown note"/>
                </group>
            </tab>
        </tabs>
    </ribbon>
</customUI>
```

- Mark as embedded resource in the .csproj
- Uses `idMso="TabAppointment"` to add to the existing Appointment tab in the inspector

**Commit:** `feat(ribbon): create inspector ribbon XML for appointment Sling button`

---

### Step 3: Update GetCustomUI to route by ribbonID

**Implement:**
- File: `SlingMD.Outlook/Ribbon/SlingRibbon.cs`
- Modify `GetCustomUI()` (line 46):

```csharp
public string GetCustomUI(string ribbonID)
{
    switch (ribbonID)
    {
        case "Microsoft.Outlook.Appointment":
            return GetResourceText("SlingMD.Outlook.Ribbon.SlingRibbonAppointment.xml");
        default:
            return GetResourceText("SlingMD.Outlook.Ribbon.SlingRibbon.xml");
    }
}
```

**Commit:** `feat(ribbon): route GetCustomUI by ribbonID for appointment inspector`

---

### Step 4: Add button click handlers to SlingRibbon

**Implement:**
- File: `SlingMD.Outlook/Ribbon/SlingRibbon.cs`
- Add after existing handlers:

```csharp
public void OnSaveTodaysClick(Office.IRibbonControl control)
{
    try
    {
        _addIn.SaveTodaysAppointments();
    }
    catch (System.Exception ex)
    {
        System.Windows.Forms.MessageBox.Show(
            $"Error saving today's appointments: {ex.Message}",
            "SlingMD Error",
            System.Windows.Forms.MessageBoxButtons.OK,
            System.Windows.Forms.MessageBoxIcon.Error);
    }
}

public void OnInspectorSlingClick(Office.IRibbonControl control)
{
    try
    {
        _addIn.ProcessCurrentAppointment();
    }
    catch (System.Exception ex)
    {
        System.Windows.Forms.MessageBox.Show(
            $"Error processing appointment: {ex.Message}",
            "SlingMD Error",
            System.Windows.Forms.MessageBoxButtons.OK,
            System.Windows.Forms.MessageBoxIcon.Error);
    }
}
```

**Commit:** `feat(ribbon): add click handlers for Save Today and Inspector Sling`

---

### Step 5: Replace ProcessSelectedEmail with ProcessSelection in ThisAddIn

**Implement:**
- File: `SlingMD.Outlook/ThisAddIn.cs`
- Add `ProcessSelection()` that handles both MailItem and AppointmentItem:

```csharp
public async void ProcessSelection()
{
    try
    {
        Explorer explorer = Application.ActiveExplorer();
        if (explorer == null || explorer.Selection == null || explorer.Selection.Count == 0)
        {
            System.Windows.Forms.MessageBox.Show(
                "Please select an email or appointment first.",
                "SlingMD",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Information);
            return;
        }

        object selectedItem = explorer.Selection[1];

        if (selectedItem is MailItem mail)
        {
            await _emailProcessor.ProcessEmail(mail);
        }
        else if (selectedItem is AppointmentItem appointment)
        {
            await _appointmentProcessor.ProcessAppointment(appointment);
        }
        else
        {
            System.Windows.Forms.MessageBox.Show(
                "Please select an email or appointment.",
                "SlingMD",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Information);
        }
    }
    catch (System.Exception ex)
    {
        System.Windows.Forms.MessageBox.Show(
            $"Error: {ex.Message}",
            "SlingMD Error",
            System.Windows.Forms.MessageBoxButtons.OK,
            System.Windows.Forms.MessageBoxIcon.Error);
    }
}
```

- Keep `ProcessSelectedEmail()` for backwards compatibility (or have it call `ProcessSelection()`)
- Update `OnSlingButtonClick` in SlingRibbon.cs to call `ProcessSelection()` instead of `ProcessSelectedEmail()`

**Commit:** `feat(addin): add ProcessSelection to handle both emails and appointments`

---

### Step 6: Add ProcessCurrentAppointment for inspector

**Implement:**
- File: `SlingMD.Outlook/ThisAddIn.cs`

```csharp
public async void ProcessCurrentAppointment()
{
    try
    {
        Inspector inspector = Application.ActiveInspector();
        if (inspector == null || inspector.CurrentItem == null)
        {
            System.Windows.Forms.MessageBox.Show(
                "No appointment is currently open.",
                "SlingMD",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Information);
            return;
        }

        AppointmentItem appointment = inspector.CurrentItem as AppointmentItem;
        if (appointment == null)
        {
            System.Windows.Forms.MessageBox.Show(
                "The current item is not an appointment.",
                "SlingMD",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Information);
            return;
        }

        // Check for unsaved changes
        if (!appointment.Saved)
        {
            System.Windows.Forms.DialogResult saveResult = System.Windows.Forms.MessageBox.Show(
                "Save appointment changes before slinging?",
                "Unsaved Changes",
                System.Windows.Forms.MessageBoxButtons.YesNoCancel,
                System.Windows.Forms.MessageBoxIcon.Question);

            if (saveResult == System.Windows.Forms.DialogResult.Yes)
            {
                appointment.Save();
            }
            else if (saveResult == System.Windows.Forms.DialogResult.Cancel)
            {
                return;
            }
        }

        await _appointmentProcessor.ProcessAppointment(appointment);
    }
    catch (System.Exception ex)
    {
        System.Windows.Forms.MessageBox.Show(
            $"Error processing appointment: {ex.Message}",
            "SlingMD Error",
            System.Windows.Forms.MessageBoxButtons.OK,
            System.Windows.Forms.MessageBoxIcon.Error);
    }
}
```

**Commit:** `feat(addin): add ProcessCurrentAppointment for inspector-level Sling`

---

### Step 7: Update OnSlingButtonClick to use ProcessSelection

**Implement:**
- File: `SlingMD.Outlook/Ribbon/SlingRibbon.cs`
- Change `OnSlingButtonClick` (line 60) from:

```csharp
_addIn.ProcessSelectedEmail();
```

to:

```csharp
_addIn.ProcessSelection();
```

**Commit:** `refactor(ribbon): route Sling button through ProcessSelection for dual-type support`

---

## Interface Contracts

### Provides (to other sub-specs)
- **Ribbon UI**: Fully functional ribbon buttons for all appointment operations
- **`ProcessSelection()`**: Unified selection handler (replaces `ProcessSelectedEmail()`)
- **`ProcessCurrentAppointment()`**: Inspector-level appointment processing

### Requires (from other sub-specs)
- **Sub-Spec 1**: `AppointmentProcessor` class, `ProcessAppointment()` method
- **Sub-Spec 5**: `SaveTodaysAppointments()` in ThisAddIn, `_appointmentProcessor` field

## Verification Commands

### Per-Step
```bash
dotnet build SlingMD.sln --configuration Release
```

### Sub-Spec Acceptance
```bash
# [MECHANICAL] Build succeeds
dotnet build SlingMD.sln --configuration Release

# [STRUCTURAL] SlingRibbon.xml has Appointments group
grep "AppointmentsGroup\|SaveTodaysButton" SlingMD.Outlook/Ribbon/SlingRibbon.xml

# [STRUCTURAL] GetCustomUI handles appointment inspector
grep "Microsoft.Outlook.Appointment" SlingMD.Outlook/Ribbon/SlingRibbon.cs

# [STRUCTURAL] ThisAddIn has all required methods
grep "ProcessSelection\|ProcessCurrentAppointment\|SaveTodaysAppointments\|_appointmentProcessor" SlingMD.Outlook/ThisAddIn.cs
```
