---
type: phase-spec
master_spec: ../spec.md
sub_spec_number: 1
title: Appointment Contact Linking
date: 2026-03-13
dependencies: []
---

# Sub-Spec 1: Appointment Contact Linking

## Scope

Add post-export contact resolution to `AppointmentProcessor.ProcessAppointment()`. After a successful appointment export, the processor checks all attendee names against the vault. Existing managed contact notes are refreshed. New contacts trigger a `ContactConfirmationDialog` (single mode) or are silently skipped (bulk mode).

### Codebase Findings

- **Insertion point:** `SlingMD.Outlook/Services/AppointmentProcessor.cs` line 566 — after `coreExportSucceeded = true` and the outer `if (!coreExportSucceeded)` guard at line 563, BEFORE the Outlook task creation block starting at line 568.
- **Attendee variables already populated:** `organizerName` (string, line 176), `requiredAttendees` (List<string>, line 180), `optionalAttendees` (List<string>, line 181), `resourceAttendees` (List<string>, line 182). These are populated via `ContactService.BuildLinkedNames()` which returns `[[Name]]` formatted strings.
- **EmailProcessor pattern to replicate:** Lines 466-516 of `SlingMD.Outlook/Services/EmailProcessor.cs` — DO NOT MODIFY that file. Copy the pattern only.
- **ContactService methods available:** `ManagedContactNoteExists(string)`, `ContactExists(string)`, `CreateContactNote(string)` — all exist, no modifications needed.
- **ContactConfirmationDialog:** Located at `SlingMD.Outlook/Forms/ContactConfirmationDialog.cs`, takes `List<string>`, returns `SelectedContacts`. Reusable as-is.
- **AppointmentProcessor already has:** `using SlingMD.Outlook.Models;` and `_contactService` field (line 61) initialized in constructor (line 80).
- **Missing import:** `SlingMD.Outlook.Forms` is NOT currently imported in AppointmentProcessor — must be added for `ContactConfirmationDialog`.

## Interface Contracts

### Provides
- Contact resolution after successful appointment export
- Attendee name deduplication and bracket stripping
- Managed contact refresh for existing contacts
- New contact dialog (single mode only)

### Requires
- `ContactService.ManagedContactNoteExists(string)` — exists
- `ContactService.ContactExists(string)` — exists
- `ContactService.CreateContactNote(string)` — exists
- `ContactConfirmationDialog` — exists

### Shared State
- Uses existing `_contactService` field in AppointmentProcessor
- Uses existing `_settings.EnableContactSaving` flag
- Reads `organizerName`, `requiredAttendees`, `optionalAttendees` variables from earlier in `ProcessAppointment()`

## Implementation Steps

### Step 1: Write failing test

**File:** `SlingMD.Tests/Services/AppointmentProcessorTests.cs`

Add a new test verifying that the processor constructs successfully with contact-related settings enabled:

```csharp
[Fact]
public void Constructor_WithContactSettings_CreatesInstance()
{
    // Arrange
    ObsidianSettings settings = new ObsidianSettings
    {
        VaultBasePath = _testDir,
        VaultName = "TestVault",
        AppointmentsFolder = "Appointments",
        EnableContactSaving = true,
        ContactsFolder = "Contacts"
    };

    // Act
    AppointmentProcessor processor = new AppointmentProcessor(settings);

    // Assert
    Assert.NotNull(processor);
}
```

Run: `dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter "FullyQualifiedName~AppointmentProcessorTests"`

### Step 2: Add `using` for Forms namespace

**File:** `SlingMD.Outlook/Services/AppointmentProcessor.cs`

Add at the top with other `using` statements (after line 13 `using SlingMD.Outlook.Helpers;`):

```csharp
using SlingMD.Outlook.Forms;
```

### Step 3: Insert contact resolution block

**File:** `SlingMD.Outlook/Services/AppointmentProcessor.cs`

Insert the following block AFTER line 566 (`return AppointmentProcessingResult.Error;` + closing brace) and BEFORE line 568 (`// --- Create Outlook task if enabled ---`):

```csharp
// --- Contact resolution (mirrors EmailProcessor pattern) ---
if (_settings.EnableContactSaving)
{
    try
    {
        // Collect all attendee names, stripping [[ ]] wrappers
        List<string> allAttendeeNames = new List<string>();

        if (!string.IsNullOrWhiteSpace(organizerName))
        {
            string cleanOrganizer = organizerName.Replace("[[", "").Replace("]]", "").Trim();
            if (!string.IsNullOrWhiteSpace(cleanOrganizer))
            {
                allAttendeeNames.Add(cleanOrganizer);
            }
        }

        foreach (string attendee in requiredAttendees)
        {
            string clean = attendee.Replace("[[", "").Replace("]]", "").Trim();
            if (!string.IsNullOrWhiteSpace(clean))
            {
                allAttendeeNames.Add(clean);
            }
        }

        foreach (string attendee in optionalAttendees)
        {
            string clean = attendee.Replace("[[", "").Replace("]]", "").Trim();
            if (!string.IsNullOrWhiteSpace(clean))
            {
                allAttendeeNames.Add(clean);
            }
        }

        // resourceAttendees excluded (conference rooms are not contacts)

        // Deduplicate and sort
        allAttendeeNames = allAttendeeNames.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(n => n).ToList();

        if (allAttendeeNames.Count > 0)
        {
            List<string> newContacts = new List<string>();
            List<string> managedContactsToRefresh = new List<string>();

            foreach (string name in allAttendeeNames)
            {
                if (_contactService.ManagedContactNoteExists(name))
                {
                    managedContactsToRefresh.Add(name);
                }
                else if (!_contactService.ContactExists(name))
                {
                    newContacts.Add(name);
                }
            }

            // Refresh existing managed contacts
            foreach (string name in managedContactsToRefresh)
            {
                _contactService.CreateContactNote(name);
            }

            // Show dialog for new contacts (single mode only)
            if (!bulkMode && newContacts.Count > 0)
            {
                using (ContactConfirmationDialog dialog = new ContactConfirmationDialog(newContacts))
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        foreach (string contactName in dialog.SelectedContacts)
                        {
                            _contactService.CreateContactNote(contactName);
                        }
                    }
                }
            }
        }
    }
    catch (System.Exception ex)
    {
        if (!bulkMode)
        {
            MessageBox.Show(
                $"Error processing contacts: {ex.Message}",
                "SlingMD Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
        else
        {
            _bulkErrors.Add($"Contact resolution error: {ex.Message}");
        }
    }
}
```

### Step 4: Verify build

```
dotnet build SlingMD.sln --configuration Release
```

### Step 5: Run all tests

```
dotnet test SlingMD.Tests\SlingMD.Tests.csproj
```

### Step 6: Verify EmailProcessor untouched

```
git diff main -- SlingMD.Outlook/Services/EmailProcessor.cs
```

Must produce empty output.

## Acceptance Criteria

- [STRUCTURAL] `AppointmentProcessor` has a `using SlingMD.Outlook.Forms;` statement. Verify with grep.
- [MECHANICAL] The contact resolution block appears after `coreExportSucceeded` check and before the Outlook task creation block.
- [BEHAVIORAL] When `EnableContactSaving` is false, no contact resolution logic executes.
- [BEHAVIORAL] In `bulkMode`, no dialog is shown; only existing managed contacts are refreshed.
- [BEHAVIORAL] Resource attendees (conference rooms) are excluded from contact resolution.
- [MECHANICAL] `EmailProcessor.cs` has zero modifications (diff against main branch must be empty for this file).
- [STRUCTURAL] New test in `AppointmentProcessorTests.cs` verifies the processor still constructs successfully with contact-related settings.

## Verification Commands

```bash
# Build
dotnet build SlingMD.sln --configuration Release

# Test
dotnet test SlingMD.Tests\SlingMD.Tests.csproj

# Verify EmailProcessor untouched
git diff main -- SlingMD.Outlook/Services/EmailProcessor.cs

# Verify using statement added
grep -n "using SlingMD.Outlook.Forms" SlingMD.Outlook/Services/AppointmentProcessor.cs
```

## Patterns to Follow

- **EmailProcessor contact resolution:** `SlingMD.Outlook/Services/EmailProcessor.cs` lines 466-516 — replicate the dedup/sort/check/dialog pattern.
- **AppointmentProcessor try/catch style:** Lines 184-203 — individual COM property reads wrapped in try/catch.
- **AppointmentProcessor bulk error pattern:** `_bulkErrors.Add(...)` for errors in bulk mode.
- **Test pattern:** `SlingMD.Tests/Services/AppointmentProcessorTests.cs` — structural tests, no COM mocking, temp directory with cleanup.

## Files

| File | Action | Notes |
|------|--------|-------|
| `SlingMD.Outlook/Services/AppointmentProcessor.cs` | Modify | Add using, insert contact resolution block |
| `SlingMD.Tests/Services/AppointmentProcessorTests.cs` | Modify | Add constructor test with contact settings |
| `SlingMD.Outlook/Services/EmailProcessor.cs` | DO NOT TOUCH | Zero modifications |
