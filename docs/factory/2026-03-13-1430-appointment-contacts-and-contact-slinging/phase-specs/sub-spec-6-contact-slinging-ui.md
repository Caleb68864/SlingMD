---
type: phase-spec
master_spec: ../spec.md
sub_spec_number: 6
title: Contact Slinging UI (Ribbon + ThisAddIn Dispatch)
date: 2026-03-13
dependencies: [5]
---

# Sub-Spec 6: Contact Slinging UI (Ribbon + ThisAddIn Dispatch)

## Scope

Add ribbon buttons for contact operations (Sling Contact, Sling All Contacts) and wire up ThisAddIn dispatch methods. Update `ProcessSelection()` to detect `ContactItem`.

### Codebase Findings

- **SlingRibbon.xml** at `SlingMD.Outlook/Ribbon/SlingRibbon.xml` — currently has 3 groups: EmailGroup, SettingsGroup, AppointmentsGroup. The new ContactsGroup goes after AppointmentsGroup (before `</tab>`).
- **SlingRibbon.cs** at `SlingMD.Outlook/Ribbon/SlingRibbon.cs` — callback pattern: public void methods taking `Office.IRibbonControl`, wrapping calls to `_addIn.MethodName()` in try/catch with MessageBox.
- **ThisAddIn.cs** at `SlingMD.Outlook/ThisAddIn.cs`:
  - Private fields: `_settings`, `_emailProcessor`, `_appointmentProcessor`, `_fileService`, `_ribbon` (lines 16-20).
  - `ThisAddIn_Startup`: creates processors (lines 33-34).
  - `ProcessSelection()`: lines 96-127 — gets explorer selection, casts to `MailItem` or `AppointmentItem`, dispatches.
  - `SaveTodaysAppointments()`: lines 179-292 — bulk export pattern with counts, COM release, summary MessageBox.
  - `ShowSettings()`: lines 294-314 — recreates processors after settings change (lines 304-306).
- **Explorer selection access:** `Application.ActiveExplorer().Selection[1]` (line 106).
- **Default contacts folder:** `Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderContacts)` — same pattern as calendar access in SaveTodaysAppointments (line 198).
- **ProcessCurrentAppointment inspector pattern:** Lines 129-172 — could serve as reference for inspector-based contact handling if needed.

## Interface Contracts

### Provides
- SlingRibbon.xml ContactsGroup with SlingContactButton and SlingAllContactsButton
- SlingRibbon.cs callbacks: `OnSlingContactClick`, `OnSlingAllContactsClick`
- ThisAddIn: `_contactProcessor` field, `ProcessSelectedContact()`, `SlingAllContacts()` methods
- Updated `ProcessSelection()` detecting ContactItem

### Requires
- `ContactProcessor` class — from Sub-Spec 5
- `ContactProcessor.ProcessContact(ContactItem)` — from Sub-Spec 5
- `ContactProcessor.ProcessAddressBook(MAPIFolder, out int, out int, out int)` — from Sub-Spec 5
- `ContactProcessor.GetBulkErrors()` — from Sub-Spec 5

### Shared State
- `_contactProcessor` field in ThisAddIn, initialized in `ThisAddIn_Startup`
- `_settings` shared with all processors

## Implementation Steps

### Step 1: Add ContactsGroup to SlingRibbon.xml

**File:** `SlingMD.Outlook/Ribbon/SlingRibbon.xml`

Add the following AFTER the `AppointmentsGroup` closing tag (`</group>`) and BEFORE the `</tab>` closing tag:

```xml
<group id="ContactsGroup" label="Contacts">
  <button id="SlingContactButton"
          label="Sling Contact"
          size="large"
          imageMso="ContactPictureMenu"
          onAction="OnSlingContactClick"
          supertip="Export selected contact to Obsidian as a markdown note"/>
  <button id="SlingAllContactsButton"
          label="Sling All Contacts"
          size="large"
          imageMso="DistributionListSelectMembers"
          onAction="OnSlingAllContactsClick"
          supertip="Export all contacts from your address book to Obsidian"/>
</group>
```

### Step 2: Add callbacks to SlingRibbon.cs

**File:** `SlingMD.Outlook/Ribbon/SlingRibbon.cs`

Add in the `#region Ribbon Callbacks` section, after `OnSettingsButtonClick` (after line 121):

```csharp
public void OnSlingContactClick(Office.IRibbonControl control)
{
    try
    {
        _addIn.ProcessSelectedContact();
    }
    catch (Exception ex)
    {
        MessageBox.Show($"Error processing contact: {ex.Message}", "SlingMD Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
    }
}

public void OnSlingAllContactsClick(Office.IRibbonControl control)
{
    try
    {
        _addIn.SlingAllContacts();
    }
    catch (Exception ex)
    {
        MessageBox.Show($"Error exporting contacts: {ex.Message}", "SlingMD Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
    }
}
```

### Step 3: Add _contactProcessor field and initialization

**File:** `SlingMD.Outlook/ThisAddIn.cs`

a) Add field declaration after `_appointmentProcessor` (after line 18):

```csharp
private ContactProcessor _contactProcessor;
```

b) Add initialization in `ThisAddIn_Startup` after `_appointmentProcessor` creation (after line 34):

```csharp
_contactProcessor = new ContactProcessor(_settings);
```

c) Add recreation in `ShowSettings()` after the other processor recreations (after line 305):

```csharp
_contactProcessor = new ContactProcessor(_settings);
```

### Step 4: Update ProcessSelection to detect ContactItem

**File:** `SlingMD.Outlook/ThisAddIn.cs`

In `ProcessSelection()` method, after the `AppointmentItem` check (after line 117), add:

```csharp
ContactItem contact = selected as ContactItem;
```

Then update the dispatch logic. The `else if` for appointment becomes:

```csharp
else if (appointment != null)
{
    await _appointmentProcessor.ProcessAppointment(appointment, bulkMode: false);
}
else if (contact != null)
{
    _contactProcessor.ProcessContact(contact);
}
else
{
    MessageBox.Show("Please select an email, appointment, or contact.", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Information);
}
```

Also update the empty-selection message at line 102 to: `"Please select an email, appointment, or contact first."`

### Step 5: Add ProcessSelectedContact method

**File:** `SlingMD.Outlook/ThisAddIn.cs`

Add after `ProcessSelectedEmail()` (after line 177):

```csharp
public void ProcessSelectedContact()
{
    try
    {
        Explorer explorer = Application.ActiveExplorer();
        if (explorer.Selection.Count == 0)
        {
            MessageBox.Show("Please select a contact first.", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        object selected = explorer.Selection[1];
        ContactItem contact = selected as ContactItem;

        if (contact == null)
        {
            MessageBox.Show("Please select a contact item.", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        _contactProcessor.ProcessContact(contact);
    }
    catch (System.Exception ex)
    {
        MessageBox.Show($"Error processing contact: {ex.Message}", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Error);
    }
}
```

### Step 6: Add SlingAllContacts method

**File:** `SlingMD.Outlook/ThisAddIn.cs`

Add after `ProcessSelectedContact()`:

```csharp
public void SlingAllContacts()
{
    int saved = 0;
    int skipped = 0;
    int errors = 0;

    try
    {
        MAPIFolder contactsFolder = null;
        try
        {
            contactsFolder = Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderContacts);
            _contactProcessor.ProcessAddressBook(contactsFolder, out saved, out skipped, out errors);
        }
        finally
        {
            if (contactsFolder != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(contactsFolder);
            }
        }

        List<string> bulkErrors = _contactProcessor.GetBulkErrors();
        string summary = string.Format(
            "Exported {0} contacts.\nSkipped: {1} (existing/non-contact)\nErrors: {2}",
            saved, skipped, errors);

        if (bulkErrors.Count > 0)
        {
            summary += "\n\nError details:\n" + string.Join("\n", bulkErrors);
        }

        MessageBox.Show(
            summary,
            "Sling All Contacts",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);

        if (_settings.LaunchObsidian && saved > 0)
        {
            _fileService.LaunchObsidian(_settings.VaultName, _settings.GetContactsPath());
        }
    }
    catch (System.Exception ex)
    {
        MessageBox.Show(
            string.Format("Error exporting contacts: {0}", ex.Message),
            "Error",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error);
    }
}
```

### Step 7: Add necessary using statements

**File:** `SlingMD.Outlook/ThisAddIn.cs`

Verify these usings exist (most should already be present):
- `using Microsoft.Office.Interop.Outlook;` — already at line 5
- `using SlingMD.Outlook.Services;` — already at line 8
- `using System.Collections.Generic;` — already at line 3

### Step 8: Verify build

```bash
dotnet build SlingMD.sln --configuration Release
```

### Step 9: Verify EmailProcessor untouched

```bash
git diff main -- SlingMD.Outlook/Services/EmailProcessor.cs
```

### Step 10: Run all tests

```bash
dotnet test SlingMD.Tests\SlingMD.Tests.csproj
```

## Acceptance Criteria

- [STRUCTURAL] SlingRibbon.xml contains a `ContactsGroup` with `SlingContactButton` and `SlingAllContactsButton`.
- [STRUCTURAL] SlingRibbon.cs has `OnSlingContactClick` and `OnSlingAllContactsClick` callback methods.
- [STRUCTURAL] ThisAddIn.cs has `_contactProcessor` field initialized in `ThisAddIn_Startup`.
- [STRUCTURAL] ThisAddIn.cs has `ProcessSelectedContact()` and `SlingAllContacts()` public methods.
- [BEHAVIORAL] `ProcessSelection()` detects `ContactItem` and dispatches to `ContactProcessor`.
- [BEHAVIORAL] `SlingAllContacts()` shows a summary dialog with saved/skipped/error counts after completion.
- [BEHAVIORAL] Error in contact processing shows user-friendly MessageBox (not unhandled exception).
- [MECHANICAL] `EmailProcessor.cs` has zero modifications.
- [HUMAN REVIEW] Ribbon layout looks correct in Outlook — Contacts group appears after Appointments group with appropriate icons.

## Verification Commands

```bash
# Build
dotnet build SlingMD.sln --configuration Release

# Test
dotnet test SlingMD.Tests\SlingMD.Tests.csproj

# Verify ribbon XML
grep -n "ContactsGroup\|SlingContactButton\|SlingAllContactsButton" SlingMD.Outlook/Ribbon/SlingRibbon.xml

# Verify ribbon callbacks
grep -n "OnSlingContactClick\|OnSlingAllContactsClick" SlingMD.Outlook/Ribbon/SlingRibbon.cs

# Verify ThisAddIn dispatch
grep -n "_contactProcessor\|ProcessSelectedContact\|SlingAllContacts\|ContactItem" SlingMD.Outlook/ThisAddIn.cs

# Verify EmailProcessor untouched
git diff main -- SlingMD.Outlook/Services/EmailProcessor.cs
```

## Patterns to Follow

- **Ribbon XML group:** `SlingMD.Outlook/Ribbon/SlingRibbon.xml` lines 20-27 — AppointmentsGroup pattern.
- **Ribbon callback:** `SlingMD.Outlook/Ribbon/SlingRibbon.cs` lines 87-97 — `OnSaveTodaysClick` pattern.
- **ProcessSelection dispatch:** `SlingMD.Outlook/ThisAddIn.cs` lines 96-127 — cast and dispatch pattern.
- **SaveTodaysAppointments bulk:** Lines 179-292 — counts, COM release, summary dialog, LaunchObsidian.
- **ShowSettings processor recreation:** Lines 304-306 — recreate all processors with new settings.

## Files

| File | Action | Notes |
|------|--------|-------|
| `SlingMD.Outlook/Ribbon/SlingRibbon.xml` | Modify | Add ContactsGroup with 2 buttons |
| `SlingMD.Outlook/Ribbon/SlingRibbon.cs` | Modify | Add 2 callback methods |
| `SlingMD.Outlook/ThisAddIn.cs` | Modify | Add field, init, dispatch, ProcessSelectedContact, SlingAllContacts |
| `SlingMD.Outlook/Services/EmailProcessor.cs` | DO NOT TOUCH | Zero modifications |
