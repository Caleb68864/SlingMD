---
type: phase-spec
master_spec: "docs/specs/2026-03-13-appointment-processing.md"
sub_spec: 4
title: "Companion Meeting Notes"
dependencies: [1, 2]
date: 2026-03-13
---

# Sub-Spec 4: Companion Meeting Notes

## Shared Context

- **Master Spec:** [2026-03-13-appointment-processing.md](../2026-03-13-appointment-processing.md)
- **Critical rule:** NEVER overwrite existing meeting notes content. If file exists, skip creation but ensure backlink.
- **Trade-off hierarchy:** Match EmailProcessor patterns > Additive-only changes > Working end-to-end

## Codebase Analysis

### Relevant Patterns

**FileService.WriteUtf8File():** Writes file to path, creates parent directory if needed. Overwrites existing files -- so we must check existence before calling.

**TemplateService.RenderMeetingNoteContent()** (from sub-spec 2): Renders meeting note stub with backlink to appointment note, attendee list, empty Agenda/Notes/Action Items sections.

**EmailProcessor contact creation pattern (lines 468-516):** Shows how to check for existing files before creating new ones -- `ContactExists()` / `ManagedContactNoteExists()` checks before `CreateContactNote()`.

### Files to Modify

| File | Action | Exists |
|------|--------|--------|
| `SlingMD.Outlook/Services/AppointmentProcessor.cs` | Add meeting notes generation step | Yes (from sub-spec 1) |

## Implementation Steps

### Step 1: Add meeting notes generation method

**Test first:**
- File: `SlingMD.Tests/Services/AppointmentProcessorTests.cs`
- Test: `ProcessAppointment_WithCreateMeetingNotesEnabled_CreatesTwoFiles`
- Test: `ProcessAppointment_WithExistingMeetingNotes_DoesNotOverwrite`

**Implement:**
- File: `SlingMD.Outlook/Services/AppointmentProcessor.cs`
- Add private method:

```csharp
private void CreateCompanionMeetingNote(
    string appointmentNotePath,
    string appointmentNoteTitle,
    string organizer,
    string attendees,
    string date,
    string location,
    string outputFolder)
{
    if (!_settings.CreateMeetingNotes)
    {
        return;
    }

    string meetingNoteFileName = Path.GetFileNameWithoutExtension(appointmentNotePath) + " - Meeting Notes.md";
    string meetingNotePath = Path.Combine(outputFolder, meetingNoteFileName);

    // NEVER overwrite existing meeting notes
    if (File.Exists(meetingNotePath))
    {
        return;
    }

    string appointmentLink = $"[[{Path.GetFileNameWithoutExtension(appointmentNotePath)}]]";

    Dictionary<string, object> metadata = new Dictionary<string, object>
    {
        { "title", Path.GetFileNameWithoutExtension(meetingNoteFileName) },
        { "type", "Meeting Notes" },
        { "appointment", appointmentLink },
        { "date", date },
        { "tags", new List<string> { "MeetingNotes" } }
    };

    MeetingNoteTemplateContext context = new MeetingNoteTemplateContext
    {
        Metadata = metadata,
        AppointmentTitle = appointmentNoteTitle,
        AppointmentLink = appointmentLink,
        Organizer = organizer,
        Attendees = attendees,
        Date = date,
        Location = location
    };

    string content = _templateService.RenderMeetingNoteContent(context);
    _fileService.WriteUtf8File(meetingNotePath, content);
}
```

**Commit:** `feat(processor): add companion meeting notes generation`

---

### Step 2: Wire meeting notes into ProcessAppointment flow

**Implement:**
- File: `SlingMD.Outlook/Services/AppointmentProcessor.cs`
- In `ProcessAppointment()`, after writing the appointment note file and processing attachments:

```csharp
// Create companion meeting notes (after appointment note is saved)
string outputFolder = Path.GetDirectoryName(appointmentFilePath);
CreateCompanionMeetingNote(
    appointmentFilePath,
    noteTitle,
    organizerName,
    string.Join(", ", attendeeNames),
    appointment.Start.ToString("yyyy-MM-dd"),
    appointment.Location ?? string.Empty,
    outputFolder);
```

Key details:
- `outputFolder` is the same folder as the appointment note (flat or thread folder)
- For recurring meetings in thread folders, the meeting note goes in the same thread folder
- This runs after file save but before task creation and Obsidian launch

**Commit:** `feat(processor): wire meeting notes into appointment processing pipeline`

---

### Step 3: Support custom meeting note template path

**Implement:**
- File: `SlingMD.Outlook/Services/AppointmentProcessor.cs`
- In `CreateCompanionMeetingNote()`, check for custom template:

```csharp
// If custom MeetingNoteTemplate path is set, pass it to TemplateService
// The TemplateService.RenderMeetingNoteContent() already handles
// LoadConfiguredTemplate() with fallback to default
```

This is already handled by `RenderMeetingNoteContent()` in sub-spec 2, which uses `LoadConfiguredTemplate(_settings.MeetingNoteTemplate, "MeetingNoteTemplate.md")`. No additional code needed in AppointmentProcessor -- just verify the setting flows through.

**Commit:** `feat(processor): support custom meeting note template path`

---

### Step 4: Add appointment frontmatter backlink to meeting notes

**Implement:**
- File: `SlingMD.Outlook/Services/AppointmentProcessor.cs`
- When building appointment note frontmatter metadata, include link to meeting notes if they exist or will be created:

```csharp
if (_settings.CreateMeetingNotes)
{
    string meetingNoteLink = $"[[{Path.GetFileNameWithoutExtension(fileName)} - Meeting Notes]]";
    metadata.Add("meetingNotes", meetingNoteLink);
}
```

This ensures bidirectional linking: appointment note links to meeting notes, and meeting notes link back to appointment note.

**Commit:** `feat(processor): add bidirectional backlinks between appointment and meeting notes`

---

## Interface Contracts

### Provides (to other sub-specs)
- **`CreateCompanionMeetingNote()` method**: Internal helper, not directly used by other sub-specs but its output (meeting note files) may be referenced

### Requires (from other sub-specs)
- **Sub-Spec 1**: AppointmentProcessor class, `ProcessAppointment()` method, `CreateMeetingNotes` setting
- **Sub-Spec 2**: `MeetingNoteTemplateContext` class, `RenderMeetingNoteContent()` method, `MeetingNoteTemplate.md`

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

# [STRUCTURAL] Meeting notes generation logic exists
grep "CreateCompanionMeetingNote\|MeetingNoteTemplateContext\|Meeting Notes" SlingMD.Outlook/Services/AppointmentProcessor.cs

# [BEHAVIORAL] Creates two files when enabled
dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter "FullyQualifiedName~AppointmentProcessorTests.ProcessAppointment_WithCreateMeetingNotesEnabled"

# [BEHAVIORAL] Does not overwrite existing
dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter "FullyQualifiedName~AppointmentProcessorTests.ProcessAppointment_WithExistingMeetingNotes"
```
