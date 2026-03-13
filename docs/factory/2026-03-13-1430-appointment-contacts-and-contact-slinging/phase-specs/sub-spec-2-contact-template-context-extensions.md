---
type: phase-spec
master_spec: ../spec.md
sub_spec_number: 2
title: ContactTemplateContext & TemplateService Extensions
date: 2026-03-13
dependencies: []
---

# Sub-Spec 2: ContactTemplateContext & TemplateService Extensions

## Scope

Extend `ContactTemplateContext` with rich contact fields (Phone, Email, Company, JobTitle, Address, Birthday, Notes, IncludeDetails) and add a `RenderRichContactContent()` method to `TemplateService` that renders contact notes with a `## Contact Details` section when `IncludeDetails` is true.

### Codebase Findings

- **ContactTemplateContext** is defined at `SlingMD.Outlook/Services/TemplateService.cs` lines 29-37. Currently has 5 properties: `Metadata`, `ContactName`, `ContactShortName`, `Created`, `FileName`, `FileNameWithoutExtension`.
- **RenderContactContent()** is at `TemplateService.cs` lines 228-244. It loads template, builds replacements from metadata + 5 context fields, calls `ProcessTemplate()`.
- **Template rendering pattern:** `ProcessTemplate()` at lines 135-155 replaces `{{key}}` placeholders. `AddReplacement()` helper at line 617 sets key-value pairs.
- **GetDefaultContactTemplate()** at lines 406-488 returns a string with frontmatter, `# {{contactName}}`, `## Communication History` (Dataview script), and `## Notes`.
- **Existing render methods:** `RenderEmailContent`, `RenderAppointmentContent`, `RenderMeetingNoteContent` all follow the same pattern: load template, build replacements, call ProcessTemplate.
- **Template file loading:** `LoadConfiguredTemplate(configuredFile, defaultFile)` at lines 566-580 tries user template first, falls back to default.
- **Test pattern:** `SlingMD.Tests/Services/TemplateServiceTests.cs` — uses temp directory, `FileService` + `TemplateService` instances, `[Fact]` tests.

## Interface Contracts

### Provides
- Extended `ContactTemplateContext` with 8 new properties: `Phone`, `Email`, `Company`, `JobTitle`, `Address`, `Birthday`, `Notes`, `IncludeDetails`
- `TemplateService.RenderRichContactContent(ContactTemplateContext)` method
- Default rich contact template via `GetDefaultRichContactTemplate()`

### Requires
- Existing `TemplateService.ProcessTemplate()` method
- Existing `TemplateService.BuildMetadataReplacements()` method
- Existing `TemplateService.AddReplacement()` helper
- Existing `TemplateService.BuildFrontMatter()` method

### Shared State
- `ContactTemplateContext` is used by `ContactService.CreateContactNote()` — existing fields must remain unchanged.
- New fields default to `string.Empty` (or `true` for `IncludeDetails`) so existing callers are unaffected.

## Implementation Steps

### Step 1: Write failing tests

**File:** `SlingMD.Tests/Services/TemplateServiceTests.cs`

Add tests at the end of the class:

```csharp
[Fact]
public void RenderRichContactContent_WithAllFields_ContainsContactDetails()
{
    ContactTemplateContext context = new ContactTemplateContext
    {
        Metadata = new Dictionary<string, object>
        {
            { "title", "John Doe" },
            { "type", "contact" },
            { "tags", new List<string> { "contact" } }
        },
        ContactName = "John Doe",
        ContactShortName = "JohnD",
        Created = "2026-03-13 14:30",
        FileName = "John Doe.md",
        FileNameWithoutExtension = "John Doe",
        Phone = "555-1234",
        Email = "john@example.com",
        Company = "Acme Corp",
        JobTitle = "Engineer",
        Address = "123 Main St",
        Birthday = "1990-01-15",
        Notes = "Met at conference",
        IncludeDetails = true
    };

    string result = _templateService.RenderRichContactContent(context);

    Assert.Contains("## Contact Details", result);
    Assert.Contains("555-1234", result);
    Assert.Contains("john@example.com", result);
    Assert.Contains("Acme Corp", result);
    Assert.Contains("Engineer", result);
    Assert.Contains("123 Main St", result);
    Assert.Contains("1990-01-15", result);
}

[Fact]
public void RenderRichContactContent_WithIncludeDetailsFalse_OmitsContactDetails()
{
    ContactTemplateContext context = new ContactTemplateContext
    {
        Metadata = new Dictionary<string, object>
        {
            { "title", "Jane Doe" },
            { "type", "contact" },
            { "tags", new List<string> { "contact" } }
        },
        ContactName = "Jane Doe",
        ContactShortName = "JaneD",
        Created = "2026-03-13 14:30",
        FileName = "Jane Doe.md",
        FileNameWithoutExtension = "Jane Doe",
        IncludeDetails = false
    };

    string result = _templateService.RenderRichContactContent(context);

    Assert.DoesNotContain("## Contact Details", result);
}

[Fact]
public void RenderRichContactContent_WithEmptyFields_DoesNotContainNull()
{
    ContactTemplateContext context = new ContactTemplateContext
    {
        Metadata = new Dictionary<string, object>
        {
            { "title", "Empty Contact" },
            { "type", "contact" },
            { "tags", new List<string> { "contact" } }
        },
        ContactName = "Empty Contact",
        ContactShortName = "EmptyC",
        Created = "2026-03-13 14:30",
        FileName = "Empty Contact.md",
        FileNameWithoutExtension = "Empty Contact",
        IncludeDetails = true
    };

    string result = _templateService.RenderRichContactContent(context);

    Assert.DoesNotContain("null", result, StringComparison.OrdinalIgnoreCase);
    Assert.Contains("## Contact Details", result);
}
```

Run: `dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter "FullyQualifiedName~TemplateServiceTests"` — should FAIL (methods don't exist yet).

### Step 2: Extend ContactTemplateContext

**File:** `SlingMD.Outlook/Services/TemplateService.cs`

Add properties to `ContactTemplateContext` (after line 36, before the closing brace at line 37):

```csharp
public string Phone { get; set; } = string.Empty;
public string Email { get; set; } = string.Empty;
public string Company { get; set; } = string.Empty;
public string JobTitle { get; set; } = string.Empty;
public string Address { get; set; } = string.Empty;
public string Birthday { get; set; } = string.Empty;
public string Notes { get; set; } = string.Empty;
public bool IncludeDetails { get; set; } = true;
```

### Step 3: Add GetDefaultRichContactTemplate method

**File:** `SlingMD.Outlook/Services/TemplateService.cs`

Add after `GetDefaultContactTemplate()` (after line 488):

```csharp
public string GetDefaultRichContactTemplate()
{
    StringBuilder sb = new StringBuilder();
    sb.Append("{{frontmatter}}");
    sb.AppendLine("# {{contactName}}");
    sb.AppendLine();
    sb.AppendLine("## Contact Details");
    sb.AppendLine();
    sb.AppendLine("**Phone:** {{phone}}");
    sb.AppendLine("**Email:** {{email}}");
    sb.AppendLine("**Company:** {{company}}");
    sb.AppendLine("**Title:** {{jobTitle}}");
    sb.AppendLine("**Address:** {{address}}");
    sb.AppendLine("**Birthday:** {{birthday}}");
    sb.AppendLine();
    sb.AppendLine("## Communication History");
    sb.AppendLine();
    // Same Dataview script as default contact template
    sb.Append(GetContactDataviewScript());
    sb.AppendLine();
    sb.AppendLine("## Notes");
    sb.AppendLine();
    sb.AppendLine("{{notes}}");
    return sb.ToString();
}
```

Note: Extract the Dataview script from `GetDefaultContactTemplate()` into a private `GetContactDataviewScript()` helper, or inline the full script. The key point is the `## Contact Details` section must appear BEFORE `## Communication History`.

### Step 4: Add RenderRichContactContent method

**File:** `SlingMD.Outlook/Services/TemplateService.cs`

Add after `RenderContactContent()` (after line 244):

```csharp
public string RenderRichContactContent(ContactTemplateContext context)
{
    if (!context.IncludeDetails)
    {
        return RenderContactContent(context);
    }

    string templateContent = LoadConfiguredTemplate(_settings.ContactTemplateFile, "RichContactTemplate.md");
    if (string.IsNullOrEmpty(templateContent))
    {
        templateContent = GetDefaultRichContactTemplate();
    }

    Dictionary<string, string> replacements = BuildMetadataReplacements(context.Metadata);
    AddReplacement(replacements, "contactName", context.ContactName);
    AddReplacement(replacements, "contactShortName", context.ContactShortName);
    AddReplacement(replacements, "created", context.Created);
    AddReplacement(replacements, "fileName", context.FileName);
    AddReplacement(replacements, "fileNameNoExt", context.FileNameWithoutExtension);
    AddReplacement(replacements, "phone", context.Phone);
    AddReplacement(replacements, "email", context.Email);
    AddReplacement(replacements, "company", context.Company);
    AddReplacement(replacements, "jobTitle", context.JobTitle);
    AddReplacement(replacements, "address", context.Address);
    AddReplacement(replacements, "birthday", context.Birthday);
    AddReplacement(replacements, "notes", context.Notes);

    return ProcessTemplate(templateContent, replacements);
}
```

Note: `BuildMetadataReplacements` and `AddReplacement` are private methods in `TemplateService`. Since `RenderRichContactContent` is being added to the same class, they are accessible.

### Step 5: Verify build and tests

```bash
dotnet build SlingMD.sln --configuration Release
dotnet test SlingMD.Tests\SlingMD.Tests.csproj
```

## Acceptance Criteria

- [STRUCTURAL] `ContactTemplateContext` has properties: Phone, Email, Company, JobTitle, Address, Birthday, Notes, IncludeDetails.
- [MECHANICAL] `RenderRichContactContent()` produces markdown containing `## Contact Details` when `IncludeDetails` is true.
- [MECHANICAL] When `IncludeDetails` is false, the output does NOT contain `## Contact Details`.
- [BEHAVIORAL] Template token replacement works for all new fields (e.g., `{{phone}}`, `{{email}}`, `{{company}}`).
- [STRUCTURAL] New test verifies rich contact template rendering with all fields populated.
- [STRUCTURAL] New test verifies rich contact template rendering with empty fields (graceful — no "null" strings).

## Verification Commands

```bash
# Build
dotnet build SlingMD.sln --configuration Release

# Test
dotnet test SlingMD.Tests\SlingMD.Tests.csproj

# Verify new properties exist
grep -n "Phone\|Email\|Company\|JobTitle\|Address\|Birthday\|IncludeDetails" SlingMD.Outlook/Services/TemplateService.cs

# Verify new method exists
grep -n "RenderRichContactContent" SlingMD.Outlook/Services/TemplateService.cs
```

## Patterns to Follow

- **RenderContactContent():** `SlingMD.Outlook/Services/TemplateService.cs` lines 228-244 — follow same load/replace/process pattern.
- **RenderAppointmentContent():** Lines 285-312 — example of a render method with many fields.
- **AddReplacement helper:** Line 617 — use for all field additions.
- **Context class pattern:** `EmailTemplateContext` (lines 11-27), `AppointmentTemplateContext` (lines 56-75) — all use `string.Empty` defaults.

## Files

| File | Action | Notes |
|------|--------|-------|
| `SlingMD.Outlook/Services/TemplateService.cs` | Modify | Add properties to ContactTemplateContext, add RenderRichContactContent(), add GetDefaultRichContactTemplate() |
| `SlingMD.Tests/Services/TemplateServiceTests.cs` | Modify | Add 3 new tests for rich contact rendering |
