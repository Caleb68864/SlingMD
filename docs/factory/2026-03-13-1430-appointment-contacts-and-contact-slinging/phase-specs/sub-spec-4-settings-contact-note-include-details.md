---
type: phase-spec
master_spec: ../spec.md
sub_spec_number: 4
title: Settings — ContactNoteIncludeDetails
date: 2026-03-13
dependencies: []
---

# Sub-Spec 4: Settings — ContactNoteIncludeDetails

## Scope

Add `ContactNoteIncludeDetails` boolean setting to `ObsidianSettings` with default value `true`. Wire it into `NormalizeLoadedSettings()` for backward compatibility. Add a checkbox to the existing Contacts tab in `SettingsForm`.

### Codebase Findings

- **ObsidianSettings** at `SlingMD.Outlook/Models/ObsidianSettings.cs`. Existing contact-related properties are at lines 13-15: `ContactsFolder`, `EnableContactSaving`, `SearchEntireVaultForContacts`. Also `ContactFilenameFormat` at line 100.
- **NormalizeLoadedSettings()** at lines 404-443. Pattern: `PropertyName = condition ? default : PropertyName`. For booleans, they're typically left as-is since JSON deserialization handles them, but the method ensures string fields have valid defaults.
- **SettingsForm Contacts tab** at lines 349-388 of `SlingMD.Outlook/Forms/SettingsForm.cs`. Uses `TableLayoutPanel` with `cRow` counter. Currently 5 rows (cRow 0-4): Contacts Folder, Enable Contact Saving checkbox, Search Entire Vault checkbox, Contact Filename Format, Contact Template File.
- **Settings load/save in form:** `LoadSettingsToForm()` at line 697+, `SaveFormToSettings()` at line 794+.
- **Test pattern:** `ObsidianSettingsTestable` subclass in `SlingMD.Tests/Models/ObsidianSettingsTests.cs` overrides `GetSettingsPath()`. Tests use temp directory.
- **Checkbox declaration pattern:** Private field like `private CheckBox chkEnableContactSaving;` (line 59).

## Interface Contracts

### Provides
- `ObsidianSettings.ContactNoteIncludeDetails` property (bool, default true)
- Backward-compatible deserialization (old JSON files without this property default to true)
- Settings UI checkbox on Contacts tab

### Requires
- Nothing — this is a standalone sub-spec

### Shared State
- `ContactNoteIncludeDetails` will be read by `ContactProcessor` (Sub-Spec 5) to set `context.IncludeDetails`

## Implementation Steps

### Step 1: Write failing test

**File:** `SlingMD.Tests/Models/ObsidianSettingsTests.cs`

Add test:

```csharp
[Fact]
public void ContactNoteIncludeDetails_DefaultsToTrue()
{
    ObsidianSettings settings = new ObsidianSettings();
    Assert.True(settings.ContactNoteIncludeDetails);
}

[Fact]
public void ContactNoteIncludeDetails_SavedAndLoaded_Correctly()
{
    ObsidianSettingsTestable settings = new ObsidianSettingsTestable
    {
        TestSettingsPath = _testSettingsPath,
        ContactNoteIncludeDetails = false
    };

    settings.Save();

    ObsidianSettingsTestable loadedSettings = new ObsidianSettingsTestable
    {
        TestSettingsPath = _testSettingsPath
    };
    loadedSettings.Load();

    Assert.False(loadedSettings.ContactNoteIncludeDetails);
}

[Fact]
public void NormalizeLoadedSettings_MissingContactNoteIncludeDetails_DefaultsToTrue()
{
    // Simulate loading settings JSON that lacks ContactNoteIncludeDetails
    ObsidianSettingsTestable settings = new ObsidianSettingsTestable
    {
        TestSettingsPath = _testSettingsPath
    };
    // Save with default
    settings.Save();

    // Load — should not throw and should default to true
    ObsidianSettingsTestable loaded = new ObsidianSettingsTestable
    {
        TestSettingsPath = _testSettingsPath
    };
    loaded.Load();

    Assert.True(loaded.ContactNoteIncludeDetails);
}
```

Run: `dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter "FullyQualifiedName~ObsidianSettingsTests"` — should FAIL (property doesn't exist).

### Step 2: Add property to ObsidianSettings

**File:** `SlingMD.Outlook/Models/ObsidianSettings.cs`

Add after `SearchEntireVaultForContacts` (line 15), near the other contact settings:

```csharp
/// <summary>
/// Whether to include detailed contact information (phone, email, company, etc.)
/// when creating contact notes via contact slinging.
/// </summary>
public bool ContactNoteIncludeDetails { get; set; } = true;
```

### Step 3: Update NormalizeLoadedSettings

**File:** `SlingMD.Outlook/Models/ObsidianSettings.cs`

The `ContactNoteIncludeDetails` property is a boolean with a default of `true`. Since `JsonConvert.PopulateObject` will set it to `false` if the JSON field is absent (C# bool default), we must ensure the NormalizeLoadedSettings does NOT override an explicit `false` value from JSON.

However, the safer approach matching existing patterns: Since the JSON deserializer with `ObjectCreationHandling.Replace` and `MissingMemberHandling.Ignore` will leave the property at its C# default (`true`) if missing from JSON, and will set it to `false` if explicitly `false` in JSON, no normalization is needed for this boolean. The C# property default of `true` handles the upgrade case naturally.

No change needed in `NormalizeLoadedSettings()` — the property default handles it. But verify this works by running the test.

### Step 4: Add checkbox to SettingsForm

**File:** `SlingMD.Outlook/Forms/SettingsForm.cs`

a) Add private field declaration (near line 60, after `chkSearchEntireVaultForContacts`):

```csharp
private CheckBox chkContactNoteIncludeDetails;
```

b) Add checkbox to Contacts tab layout (after the Contact Template File row, before `contactsTabLayout.RowCount = cRow + 1;` at line 383):

```csharp
contactsTabLayout.Controls.Add(new Label(), 0, cRow);
this.chkContactNoteIncludeDetails = new CheckBox { Text = "Include contact details (phone, email, company, etc.)", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true };
contactsTabLayout.Controls.Add(this.chkContactNoteIncludeDetails, 1, cRow++);
```

c) Add to `LoadSettingsToForm()` (after line 701):

```csharp
chkContactNoteIncludeDetails.Checked = _settings.ContactNoteIncludeDetails;
```

d) Add to `SaveFormToSettings()` (after line 797):

```csharp
_settings.ContactNoteIncludeDetails = chkContactNoteIncludeDetails.Checked;
```

### Step 5: Verify build and tests

```bash
dotnet build SlingMD.sln --configuration Release
dotnet test SlingMD.Tests\SlingMD.Tests.csproj
```

## Acceptance Criteria

- [STRUCTURAL] `ObsidianSettings` has `ContactNoteIncludeDetails` property with default value `true`.
- [MECHANICAL] `NormalizeLoadedSettings()` does not throw when deserializing settings JSON that lacks `ContactNoteIncludeDetails`.
- [STRUCTURAL] SettingsForm Contacts tab has checkbox for `ContactNoteIncludeDetails`.
- [BEHAVIORAL] Setting persists correctly through Save/Load cycle.
- [STRUCTURAL] New test verifies default value is `true`.
- [MECHANICAL] Existing settings tests still pass.

## Verification Commands

```bash
# Build
dotnet build SlingMD.sln --configuration Release

# Test
dotnet test SlingMD.Tests\SlingMD.Tests.csproj

# Verify property exists
grep -n "ContactNoteIncludeDetails" SlingMD.Outlook/Models/ObsidianSettings.cs

# Verify UI checkbox exists
grep -n "ContactNoteIncludeDetails\|chkContactNoteIncludeDetails" SlingMD.Outlook/Forms/SettingsForm.cs
```

## Patterns to Follow

- **Property declaration:** `SlingMD.Outlook/Models/ObsidianSettings.cs` line 14 — `public bool EnableContactSaving { get; set; } = true;`
- **Checkbox in form:** `SlingMD.Outlook/Forms/SettingsForm.cs` lines 366-369 — `chkEnableContactSaving` pattern.
- **Load/Save in form:** Lines 697-700 and 794-797 — direct property mapping.
- **Test pattern:** `SlingMD.Tests/Models/ObsidianSettingsTests.cs` — `ObsidianSettingsTestable` with `TestSettingsPath`.

## Files

| File | Action | Notes |
|------|--------|-------|
| `SlingMD.Outlook/Models/ObsidianSettings.cs` | Modify | Add `ContactNoteIncludeDetails` property |
| `SlingMD.Outlook/Forms/SettingsForm.cs` | Modify | Add checkbox field, layout row, load/save bindings |
| `SlingMD.Tests/Models/ObsidianSettingsTests.cs` | Modify | Add 3 tests for new setting |
