---
type: phase-spec
master_spec: ../spec.md
sub_spec_number: 3
title: ContactService Extensions
date: 2026-03-13
dependencies: [2]
---

# Sub-Spec 3: ContactService Extensions

## Scope

Add methods to `ContactService` for extracting rich data from Outlook `ContactItem` objects and creating rich contact notes. Two new public methods: `ExtractContactData(ContactItem)` and `CreateRichContactNote(ContactTemplateContext)`.

### Codebase Findings

- **ContactService** is at `SlingMD.Outlook/Services/ContactService.cs` (534 lines). Constructor takes `FileService` and `TemplateService` (line 24). Private fields: `_fileService`, `_templateService`, `_settings`.
- **Existing `CreateContactNote(string)`:** Lines 371-411 — creates `ContactTemplateContext`, calls `_templateService.RenderContactContent(context)`, writes file or merges with existing content.
- **GetManagedContactNotePath(string):** Lines 413-432 — resolves file path using configured format.
- **MergeManagedSections():** Lines 485-515 — merges updated managed section into existing content.
- **COM property read pattern in AppointmentProcessor:** Lines 184-203 — individual try/catch per property read. This is the pattern to follow for `ExtractContactData`.
- **Missing import:** `ContactService` currently has `using Microsoft.Office.Interop.Outlook;` (line 5) and `using System.Runtime.InteropServices;` is NOT imported — needed for `Marshal.ReleaseComObject()`. However, `Marshal.ReleaseComObject` calls in existing code use fully-qualified `System.Runtime.InteropServices.Marshal.ReleaseComObject()`.
- **Test pattern:** `SlingMD.Tests/Services/ContactServiceTests.cs` uses `TestFileService` and `TestTemplateService` subclasses. Tests use temp directories with cleanup.

## Interface Contracts

### Provides
- `ContactService.ExtractContactData(ContactItem contact)` → returns `ContactTemplateContext`
- `ContactService.CreateRichContactNote(ContactTemplateContext context)` → void

### Requires
- `ContactTemplateContext` with rich fields (Phone, Email, Company, etc.) — from Sub-Spec 2
- `TemplateService.RenderRichContactContent(ContactTemplateContext)` — from Sub-Spec 2
- Existing `GetManagedContactNotePath(string)` — already exists
- Existing `MergeManagedSections()` — already exists
- Existing `_fileService.WriteUtf8File()` — already exists
- Existing `_fileService.EnsureDirectoryExists()` — already exists
- Existing `_fileService.CleanFileName()` — already exists

### Shared State
- `_settings.GetContactsPath()` for file output location
- `_settings.EnableContactSaving` is NOT checked in this method (caller is responsible)

## Implementation Steps

### Step 1: Write failing tests

**File:** `SlingMD.Tests/Services/ContactServiceTests.cs`

Add tests for `CreateRichContactNote`:

```csharp
[Fact]
public void CreateRichContactNote_CreatesFileWithExpectedContent()
{
    // Arrange
    ContactTemplateContext context = new ContactTemplateContext
    {
        Metadata = new Dictionary<string, object>
        {
            { "title", "Test Contact" },
            { "type", "contact" },
            { "tags", new List<string> { "contact" } }
        },
        ContactName = "Test Contact",
        ContactShortName = "TestC",
        Created = DateTime.Now.ToString("yyyy-MM-dd HH:mm"),
        Phone = "555-0100",
        Email = "test@example.com",
        Company = "Test Corp",
        JobTitle = "Developer",
        Address = "456 Oak Ave",
        Birthday = "1985-06-15",
        Notes = "Test notes",
        IncludeDetails = true
    };

    // Act
    _contactService.CreateRichContactNote(context);

    // Assert
    string expectedPath = _contactService.GetManagedContactNotePath("Test Contact");
    Assert.True(File.Exists(expectedPath));
    string content = File.ReadAllText(expectedPath);
    Assert.Contains("## Contact Details", content);
    Assert.Contains("555-0100", content);
    Assert.Contains("test@example.com", content);
}

[Fact]
public void CreateRichContactNote_MergesWhenFileExists()
{
    // Arrange - create initial note
    ContactTemplateContext context = new ContactTemplateContext
    {
        Metadata = new Dictionary<string, object>
        {
            { "title", "Test Contact" },
            { "type", "contact" },
            { "tags", new List<string> { "contact" } }
        },
        ContactName = "Test Contact",
        ContactShortName = "TestC",
        Created = DateTime.Now.ToString("yyyy-MM-dd HH:mm"),
        Phone = "555-0100",
        Email = "test@example.com",
        IncludeDetails = true
    };

    _contactService.CreateRichContactNote(context);

    // Act - update with new data
    context.Phone = "555-0200";
    _contactService.CreateRichContactNote(context);

    // Assert
    string expectedPath = _contactService.GetManagedContactNotePath("Test Contact");
    string content = File.ReadAllText(expectedPath);
    Assert.Contains("555-0200", content);
}
```

Run: `dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter "FullyQualifiedName~ContactServiceTests"` — should FAIL.

### Step 2: Add ExtractContactData method

**File:** `SlingMD.Outlook/Services/ContactService.cs`

Add after `GetMeetingResourceData()` method (before `ContactExists()`), approximately after line 284:

```csharp
/// <summary>
/// Extracts rich contact data from an Outlook <see cref="ContactItem"/> into a
/// <see cref="ContactTemplateContext"/>. Each property read is individually try/caught
/// to handle COM failures gracefully.
/// </summary>
public ContactTemplateContext ExtractContactData(ContactItem contact)
{
    ContactTemplateContext context = new ContactTemplateContext();

    // Contact name with fallback chain
    string contactName = string.Empty;
    try { contactName = contact.FullName; } catch { }
    if (string.IsNullOrWhiteSpace(contactName))
    {
        try
        {
            string last = contact.LastName ?? string.Empty;
            string first = contact.FirstName ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(last) || !string.IsNullOrWhiteSpace(first))
            {
                contactName = string.IsNullOrWhiteSpace(last) ? first : $"{last}, {first}";
            }
        }
        catch { }
    }
    if (string.IsNullOrWhiteSpace(contactName))
    {
        try { contactName = contact.FileAs; } catch { }
    }
    if (string.IsNullOrWhiteSpace(contactName))
    {
        contactName = "Unknown Contact";
    }

    context.ContactName = contactName;
    context.ContactShortName = GetShortName(contactName);

    // Phone with fallback chain
    string phone = string.Empty;
    try { phone = contact.BusinessTelephoneNumber; } catch { }
    if (string.IsNullOrWhiteSpace(phone))
    {
        try { phone = contact.MobileTelephoneNumber; } catch { }
    }
    if (string.IsNullOrWhiteSpace(phone))
    {
        try { phone = contact.HomeTelephoneNumber; } catch { }
    }
    context.Phone = phone ?? string.Empty;

    // Email
    try { context.Email = contact.Email1Address ?? string.Empty; } catch { context.Email = string.Empty; }

    // Company
    try { context.Company = contact.CompanyName ?? string.Empty; } catch { context.Company = string.Empty; }

    // Job title
    try { context.JobTitle = contact.JobTitle ?? string.Empty; } catch { context.JobTitle = string.Empty; }

    // Address with fallback
    string address = string.Empty;
    try { address = contact.BusinessAddress; } catch { }
    if (string.IsNullOrWhiteSpace(address))
    {
        try { address = contact.HomeAddress; } catch { }
    }
    context.Address = address ?? string.Empty;

    // Birthday — skip if year 4501 (Outlook's "not set" sentinel)
    try
    {
        DateTime birthday = contact.Birthday;
        if (birthday != DateTime.MinValue && birthday.Year != 4501)
        {
            context.Birthday = birthday.ToString("yyyy-MM-dd");
        }
    }
    catch { context.Birthday = string.Empty; }

    // Notes / Body
    try { context.Notes = contact.Body ?? string.Empty; } catch { context.Notes = string.Empty; }

    // Metadata
    string created = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
    context.Created = created;
    context.Metadata = new Dictionary<string, object>
    {
        { "title", contactName },
        { "type", "contact" },
        { "created", created },
        { "tags", new List<string> { "contact" } }
    };

    string fileNameNoExtension = _fileService.CleanFileName(contactName);
    context.FileName = fileNameNoExtension + ".md";
    context.FileNameWithoutExtension = fileNameNoExtension;

    return context;
}
```

### Step 3: Add CreateRichContactNote method

**File:** `SlingMD.Outlook/Services/ContactService.cs`

Add after `CreateContactNote()` (after line 411):

```csharp
/// <summary>
/// Creates or updates a rich contact note using the extended template context.
/// When the file already exists, managed sections are merged while preserving user content.
/// </summary>
public void CreateRichContactNote(ContactTemplateContext context)
{
    string filePath = GetManagedContactNotePath(context.ContactName);
    _fileService.EnsureDirectoryExists(_settings.GetContactsPath());

    string renderedContent = _templateService.RenderRichContactContent(context);

    if (!File.Exists(filePath))
    {
        _fileService.WriteUtf8File(filePath, renderedContent);
        return;
    }

    string existingContent = File.ReadAllText(filePath);
    string managedSection = ExtractManagedCommunicationHistorySection(renderedContent);
    string updatedContent = MergeManagedSections(existingContent, managedSection);
    _fileService.WriteUtf8File(filePath, updatedContent);
}
```

### Step 4: Verify build and tests

```bash
dotnet build SlingMD.sln --configuration Release
dotnet test SlingMD.Tests\SlingMD.Tests.csproj
```

## Acceptance Criteria

- [STRUCTURAL] `ContactService` has public methods `ExtractContactData` and `CreateRichContactNote`.
- [BEHAVIORAL] `ExtractContactData` handles null/missing properties gracefully — each property read is individually try/caught.
- [BEHAVIORAL] Birthday value of year 4501 (Outlook's "not set" sentinel) is treated as empty string.
- [MECHANICAL] `CreateRichContactNote` writes to the same contacts folder as `CreateContactNote`.
- [STRUCTURAL] New tests verify `CreateRichContactNote` creates a file with expected content.
- [STRUCTURAL] New tests verify `CreateRichContactNote` merges correctly when file already exists.
- [MECHANICAL] No modifications to existing `CreateContactNote()` method signature or behavior.

## Verification Commands

```bash
# Build
dotnet build SlingMD.sln --configuration Release

# Test
dotnet test SlingMD.Tests\SlingMD.Tests.csproj

# Verify new methods exist
grep -n "ExtractContactData\|CreateRichContactNote" SlingMD.Outlook/Services/ContactService.cs

# Verify existing CreateContactNote unchanged
git diff main -- SlingMD.Outlook/Services/ContactService.cs | grep "CreateContactNote(string"
```

## Patterns to Follow

- **COM property reads:** `SlingMD.Outlook/Services/AppointmentProcessor.cs` lines 184-203 — individual try/catch per property.
- **CreateContactNote():** `SlingMD.Outlook/Services/ContactService.cs` lines 371-411 — context creation, render, write/merge.
- **MergeManagedSections():** Lines 485-515 — preserve user content, update managed sections.
- **Test helpers:** `TestFileService` and `TestTemplateService` in `SlingMD.Tests/Services/ContactServiceTests.cs`.

## Files

| File | Action | Notes |
|------|--------|-------|
| `SlingMD.Outlook/Services/ContactService.cs` | Modify | Add `ExtractContactData()` and `CreateRichContactNote()` |
| `SlingMD.Tests/Services/ContactServiceTests.cs` | Modify | Add tests for new methods |
