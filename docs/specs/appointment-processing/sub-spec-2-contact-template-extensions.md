---
type: phase-spec
master_spec: "docs/specs/2026-03-13-appointment-processing.md"
sub_spec: 2
title: "ContactService Meeting Extensions + TemplateService Appointment Contexts"
dependencies: [1]
date: 2026-03-13
---

# Sub-Spec 2: ContactService Meeting Extensions + TemplateService Appointment Contexts

## Shared Context

- **Master Spec:** [2026-03-13-appointment-processing.md](../2026-03-13-appointment-processing.md)
- **Trade-off hierarchy:** Match EmailProcessor patterns > Additive-only changes > Working end-to-end
- **Critical constraint:** Additive only -- do NOT modify existing method signatures in ContactService or TemplateService
- **Escalation triggers:** Any modification to existing shared service method signatures

## Codebase Analysis

### ContactService.cs (298 lines)

**Existing methods to reference (not modify):**
- `GetSenderEmail(MailItem mail)` (line 65): Uses MAPI property `0x39FE001E` via PropertyAccessor, falls back to `mail.SenderEmailAddress`
- `BuildLinkedNames(Recipients recipients, OlMailRecipientType type)` (line 81): Iterates Recipients, filters by type cast to int, creates `[[Name]]` wikilinks, COM cleanup via Marshal.ReleaseComObject in try/finally
- `BuildEmailList(Recipients recipients, OlMailRecipientType type)` (line 108): Same pattern but returns plain email addresses, nested try/catch for MAPI resolution with fallback to `recipient.Address`

**COM patterns to follow:**
- Each recipient in try/finally with `Marshal.ReleaseComObject(recipient)` (line 97, 139)
- Type filtering: `recipient.Type == (int)type` (lines 88, 115)
- MAPI property resolution: `const string PrSmtpAddress = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"` (line 69)

### TemplateService.cs (579+ lines)

**Existing context classes (lines 11-54):**
- `EmailTemplateContext`: Metadata dict, NoteTitle, Subject, SenderName, SenderShortName, SenderEmail, Date, Timestamp, Body, TaskBlock, FileName, FileNameWithoutExtension, ThreadNote, ThreadId
- `ContactTemplateContext`: Metadata dict, ContactName, ContactShortName, Created, FileName, FileNameWithoutExtension
- `TaskTemplateContext`: NoteLink, NoteName, Tags, CreatedDate, ReminderDate, DueDate
- `ThreadTemplateContext`: Title, ThreadId, FolderPath

**Render pattern:**
- `RenderEmailContent(EmailTemplateContext context)` (line 170): Loads template via `LoadConfiguredTemplate()`, falls back to `GetDefaultEmailTemplate()`, builds replacements dict, calls `ProcessTemplate()`
- `ProcessTemplate(string templateContent, Dictionary<string, string> replacements)` (line 103): Simple `{{key}}` replacement loop
- `BuildFrontMatter(Dictionary<string, object> metadata)` (line 130): YAML front-matter generation with type-specific serialization

**Template loading:**
- `LoadConfiguredTemplate(configuredFile, defaultFile)` (line 446): Tries configured file, falls back to default
- `BuildTemplateCandidatePaths()` (line 462): 5-path search priority

### Existing Templates

- `EmailTemplate.md`: `{{frontmatter}}{{taskBlock}}{{body}}` (minimal)
- `ContactTemplate.md`: Full template with DataviewJS communication history query
- `TaskTemplate.md`: `- [ ] {{noteLink}} {{tags}} ... {{dueDate}}`
- `ThreadNoteTemplate.md`: YAML front-matter + DataviewJS thread summary

### Files to Create/Modify

| File | Action | Exists |
|------|--------|--------|
| `SlingMD.Outlook/Services/ContactService.cs` | Add 4 new methods (additive) | Yes (298 lines) |
| `SlingMD.Outlook/Services/TemplateService.cs` | Add 2 context classes + 2 render methods + 2 default template methods | Yes (579+ lines) |
| `SlingMD.Outlook/Templates/AppointmentTemplate.md` | Create new | No |
| `SlingMD.Outlook/Templates/MeetingNoteTemplate.md` | Create new | No |

## Implementation Steps

### Step 1: Add AppointmentTemplateContext to TemplateService

**Test first:**
- File: `SlingMD.Tests/Services/TemplateServiceAppointmentTests.cs`
- Test: `AppointmentTemplateContext_AllProperties_Settable`
- Asserts: Create context, set all properties, verify values

**Implement:**
- File: `SlingMD.Outlook/Services/TemplateService.cs`
- Add after `ThreadTemplateContext` (line 54):

```csharp
public class AppointmentTemplateContext
{
    public Dictionary<string, object> Metadata { get; set; }
    public string NoteTitle { get; set; }
    public string Subject { get; set; }
    public string Organizer { get; set; }
    public string OrganizerEmail { get; set; }
    public string Attendees { get; set; }
    public string OptionalAttendees { get; set; }
    public string Resources { get; set; }
    public string Location { get; set; }
    public string StartDateTime { get; set; }
    public string EndDateTime { get; set; }
    public string Recurrence { get; set; }
    public string Date { get; set; }
    public string Body { get; set; }
    public string TaskBlock { get; set; }
    public string FileName { get; set; }
    public string FileNameWithoutExtension { get; set; }
}
```

**Commit:** `feat(template): add AppointmentTemplateContext class`

---

### Step 2: Add MeetingNoteTemplateContext to TemplateService

**Test first:**
- File: `SlingMD.Tests/Services/TemplateServiceAppointmentTests.cs`
- Test: `MeetingNoteTemplateContext_AllProperties_Settable`

**Implement:**
- File: `SlingMD.Outlook/Services/TemplateService.cs`
- Add after `AppointmentTemplateContext`:

```csharp
public class MeetingNoteTemplateContext
{
    public Dictionary<string, object> Metadata { get; set; }
    public string AppointmentTitle { get; set; }
    public string AppointmentLink { get; set; }
    public string Organizer { get; set; }
    public string Attendees { get; set; }
    public string Date { get; set; }
    public string Location { get; set; }
}
```

**Commit:** `feat(template): add MeetingNoteTemplateContext class`

---

### Step 3: Create default AppointmentTemplate.md

**Implement:**
- File: `SlingMD.Outlook/Templates/AppointmentTemplate.md`

```markdown
{{frontmatter}}{{taskBlock}}
## Attendees

**Organizer:** {{organizer}}
**Required:** {{attendees}}
**Optional:** {{optionalAttendees}}
**Resources:** {{resources}}

## Details

**Location:** {{location}}
**Start:** {{startDateTime}}
**End:** {{endDateTime}}
**Recurrence:** {{recurrence}}

## Notes

{{body}}
```

Style matches EmailTemplate.md's minimal approach with `{{frontmatter}}` and `{{body}}`.

**Commit:** `feat(template): add default AppointmentTemplate.md`

---

### Step 4: Create default MeetingNoteTemplate.md

**Implement:**
- File: `SlingMD.Outlook/Templates/MeetingNoteTemplate.md`

```markdown
{{frontmatter}}
## Meeting Notes

**Appointment:** {{appointmentLink}}
**Organizer:** {{organizer}}
**Attendees:** {{attendees}}
**Date:** {{date}}
**Location:** {{location}}

## Agenda

-

## Notes

-

## Action Items

- [ ]
```

**Commit:** `feat(template): add default MeetingNoteTemplate.md`

---

### Step 5: Add RenderAppointmentContent method

**Test first:**
- File: `SlingMD.Tests/Services/TemplateServiceAppointmentTests.cs`
- Test: `RenderAppointmentContent_WithValidContext_ProducesExpectedMarkdown`
- Asserts: Output contains frontmatter, organizer, attendees, location, body content

**Implement:**
- File: `SlingMD.Outlook/Services/TemplateService.cs`
- Add after `RenderThreadContent()` (line 251), following the same pattern:

```csharp
public string RenderAppointmentContent(AppointmentTemplateContext context)
{
    string template = LoadConfiguredTemplate(
        _settings?.AppointmentTemplateFile,
        "AppointmentTemplate.md");

    if (string.IsNullOrEmpty(template))
    {
        template = GetDefaultAppointmentTemplate();
    }

    string frontmatter = BuildFrontMatter(context.Metadata);
    Dictionary<string, string> replacements = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "frontmatter", frontmatter },
        { "noteTitle", context.NoteTitle ?? string.Empty },
        { "subject", context.Subject ?? string.Empty },
        { "organizer", context.Organizer ?? string.Empty },
        { "organizerEmail", context.OrganizerEmail ?? string.Empty },
        { "attendees", context.Attendees ?? string.Empty },
        { "optionalAttendees", context.OptionalAttendees ?? string.Empty },
        { "resources", context.Resources ?? string.Empty },
        { "location", context.Location ?? string.Empty },
        { "startDateTime", context.StartDateTime ?? string.Empty },
        { "endDateTime", context.EndDateTime ?? string.Empty },
        { "recurrence", context.Recurrence ?? string.Empty },
        { "date", context.Date ?? string.Empty },
        { "body", context.Body ?? string.Empty },
        { "taskBlock", context.TaskBlock ?? string.Empty },
        { "fileName", context.FileName ?? string.Empty },
        { "fileNameWithoutExtension", context.FileNameWithoutExtension ?? string.Empty }
    };

    return ProcessTemplate(template, replacements);
}
```

- Add `GetDefaultAppointmentTemplate()` returning the inline version of AppointmentTemplate.md
- Add `AppointmentTemplateFile` property to ObsidianSettings if not present (default: `"AppointmentTemplate.md"`)

**Commit:** `feat(template): add RenderAppointmentContent method`

---

### Step 6: Add RenderMeetingNoteContent method

**Test first:**
- File: `SlingMD.Tests/Services/TemplateServiceAppointmentTests.cs`
- Test: `RenderMeetingNoteContent_WithValidContext_ProducesStubWithBacklink`
- Asserts: Output contains appointment link, organizer, attendees, empty sections

**Implement:**
- File: `SlingMD.Outlook/Services/TemplateService.cs`
- Add `RenderMeetingNoteContent(MeetingNoteTemplateContext context)` following same pattern
- Add `GetDefaultMeetingNoteTemplate()` returning inline version of MeetingNoteTemplate.md
- Add `MeetingNoteTemplateFile` property to ObsidianSettings if not present

**Commit:** `feat(template): add RenderMeetingNoteContent method`

---

### Step 7: Add GetSMTPEmailAddress for Recipient objects

**Test first:**
- File: `SlingMD.Tests/Services/ContactServiceMeetingTests.cs`
- Test: `GetSMTPEmailAddress_ResolvesFromRecipient`
- Note: Will need TestContactService subclass pattern (from existing ContactServiceTests.cs) since COM objects can't be easily mocked

**Implement:**
- File: `SlingMD.Outlook/Services/ContactService.cs`
- Add after `GetSenderEmail()` (line 76):

```csharp
public string GetSMTPEmailAddress(Recipient recipient)
{
    try
    {
        const string PrSmtpAddress = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
        return recipient.PropertyAccessor.GetProperty(PrSmtpAddress) as string ?? recipient.Address;
    }
    catch
    {
        return recipient.Address;
    }
}
```

**Commit:** `feat(contact): add GetSMTPEmailAddress for Recipient objects`

---

### Step 8: Add BuildLinkedNames overload for OlMeetingRecipientType

**Test first:**
- File: `SlingMD.Tests/Services/ContactServiceMeetingTests.cs`
- Test: `BuildLinkedNames_FiltersByMeetingRecipientType`

**Implement:**
- File: `SlingMD.Outlook/Services/ContactService.cs`
- Add overload after existing `BuildLinkedNames()` (line 103):

```csharp
public List<string> BuildLinkedNames(Recipients recipients, params OlMeetingRecipientType[] types)
{
    List<string> linkedNames = new List<string>();
    HashSet<int> typeSet = new HashSet<int>();
    foreach (OlMeetingRecipientType type in types)
    {
        typeSet.Add((int)type);
    }

    foreach (Recipient recipient in recipients)
    {
        try
        {
            if (typeSet.Contains(recipient.Type))
            {
                string name = recipient.Name;
                if (!string.IsNullOrEmpty(name))
                {
                    linkedNames.Add($"[[{name}]]");
                }
            }
        }
        finally
        {
            if (recipient != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(recipient);
            }
        }
    }
    return linkedNames;
}
```

**Commit:** `feat(contact): add BuildLinkedNames overload for meeting recipient types`

---

### Step 9: Add BuildEmailList overload and GetMeetingResourceData

**Implement:**
- File: `SlingMD.Outlook/Services/ContactService.cs`
- Add `BuildEmailList(Recipients recipients, IEnumerable<OlMeetingRecipientType> types)` -- same pattern as step 8 but returns email addresses
- Add `GetMeetingResourceData(Recipients recipients)` -- filters by `OlMeetingRecipientType.olResource`, returns list of resource name/email pairs

Both follow the same COM cleanup pattern with try/finally and Marshal.ReleaseComObject.

**Commit:** `feat(contact): add meeting-role-aware email list and resource data methods`

---

## Interface Contracts

### Provides (to other sub-specs)
- **AppointmentTemplateContext**: Context class for appointment note rendering (used by sub-spec 1 step 8, sub-spec 4)
- **MeetingNoteTemplateContext**: Context class for meeting note rendering (used by sub-spec 4)
- **RenderAppointmentContent()**: Method to render appointment notes (used by sub-spec 1)
- **RenderMeetingNoteContent()**: Method to render meeting note stubs (used by sub-spec 4)
- **GetSMTPEmailAddress(Recipient)**: SMTP resolution for Recipient objects (used by sub-spec 1)
- **BuildLinkedNames(Recipients, params OlMeetingRecipientType[])**: Meeting-role-aware wikilinks (used by sub-spec 1)
- **AppointmentTemplate.md / MeetingNoteTemplate.md**: Default templates (used by rendering methods)

### Requires (from other sub-specs)
- **Sub-Spec 1**: ObsidianSettings appointment properties (AppointmentTemplateFile, MeetingNoteTemplate path)

### Verification
```
# Verify new methods don't break existing ones
dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter "FullyQualifiedName~ContactServiceTests"
dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter "FullyQualifiedName~TemplateServiceTests"
```

## Verification Commands

### Per-Step
```bash
dotnet build SlingMD.sln --configuration Release
dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter "FullyQualifiedName~TemplateServiceAppointmentTests"
dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter "FullyQualifiedName~ContactServiceMeetingTests"
```

### Sub-Spec Acceptance
```bash
# [MECHANICAL] Build succeeds
dotnet build SlingMD.sln --configuration Release

# [STRUCTURAL] ContactService has 4 new methods
grep -c "OlMeetingRecipientType" SlingMD.Outlook/Services/ContactService.cs

# [STRUCTURAL] TemplateService has new context classes and render methods
grep "AppointmentTemplateContext\|MeetingNoteTemplateContext\|RenderAppointmentContent\|RenderMeetingNoteContent" SlingMD.Outlook/Services/TemplateService.cs

# [STRUCTURAL] Template files exist
ls SlingMD.Outlook/Templates/AppointmentTemplate.md SlingMD.Outlook/Templates/MeetingNoteTemplate.md

# [BEHAVIORAL] Existing tests still pass
dotnet test SlingMD.Tests\SlingMD.Tests.csproj
```
