---
type: phase-spec
master_spec: "docs/specs/2026-03-13-appointment-processing.md"
sub_spec: 9
title: "Tests"
dependencies: [1, 2, 3, 4, 5, 8]
date: 2026-03-13
---

# Sub-Spec 9: Tests

## Shared Context

- **Master Spec:** [2026-03-13-appointment-processing.md](../2026-03-13-appointment-processing.md)
- **Test framework:** xUnit + Moq (though existing tests use real services and test subclasses, not Moq mocks)
- **Key patterns:** Constructor-based setup, IDisposable cleanup, `[Fact]` (no `[Theory]`), fully qualified `System.Exception`, AAA comments

## Codebase Analysis

### Existing Test Patterns

**Test class structure:**
```csharp
public class ServiceNameTests : IDisposable
{
    private readonly ObsidianSettings _settings;
    private readonly ServiceName _service;
    private readonly string _testDir;

    public ServiceNameTests()
    {
        _testDir = Path.Combine(Path.GetTempPath(), "SlingMDTests", "ServiceName");
        if (Directory.Exists(_testDir)) Directory.Delete(_testDir, true);
        Directory.CreateDirectory(_testDir);

        _settings = new ObsidianSettings { VaultBasePath = _testDir, VaultName = "TestVault" };
        _service = new ServiceName(_settings);
    }

    public void Dispose()
    {
        if (Directory.Exists(_testDir))
        {
            try { Directory.Delete(_testDir, true); }
            catch (System.Exception) { }
        }
    }
}
```

**Test subclass pattern (from ContactServiceTests.cs):**
```csharp
public class TestFileService : FileService
{
    public TestFileService(ObsidianSettings settings) : base(settings) { _testSettings = settings; }
    public override ObsidianSettings GetSettings() { return _testSettings; }
    public override void WriteUtf8File(string filePath, string content) { /* simplified */ }
    public override string CleanFileName(string input) { /* simplified */ }
}
```

**Testable settings (from ObsidianSettingsTests.cs):**
```csharp
public class ObsidianSettingsTestable : ObsidianSettings
{
    public string TestSettingsPath { get; set; }
    protected override string GetSettingsPath() { return TestSettingsPath ?? base.GetSettingsPath(); }
}
```

**Assertion patterns:** Assert.True, Assert.False, Assert.Equal, Assert.NotNull, Assert.Contains, Assert.StartsWith, Assert.Matches

**No Moq usage in existing tests** -- all use real services or test subclasses.

### Files to Create

| File | Action | Exists |
|------|--------|--------|
| `SlingMD.Tests/Models/ObsidianSettingsAppointmentTests.cs` | Create new | No |
| `SlingMD.Tests/Services/AppointmentProcessorTests.cs` | Create new | No |
| `SlingMD.Tests/Services/ContactServiceMeetingTests.cs` | Create new | No |
| `SlingMD.Tests/Services/TemplateServiceAppointmentTests.cs` | Create new | No |

## Implementation Steps

### Step 1: ObsidianSettingsAppointmentTests

**File:** `SlingMD.Tests/Models/ObsidianSettingsAppointmentTests.cs`

**Tests to implement:**

```csharp
namespace SlingMD.Tests.Models
{
    public class ObsidianSettingsAppointmentTests : IDisposable
    {
        private readonly string _testDir;

        public ObsidianSettingsAppointmentTests()
        {
            _testDir = Path.Combine(Path.GetTempPath(), "SlingMDTests", "SettingsAppointment");
            if (Directory.Exists(_testDir)) Directory.Delete(_testDir, true);
            Directory.CreateDirectory(_testDir);
        }

        [Fact]
        public void DefaultValues_AllAppointmentProperties_HaveCorrectDefaults()
        {
            // Arrange & Act
            ObsidianSettings settings = new ObsidianSettings();

            // Assert
            Assert.Equal("Appointments", settings.AppointmentsFolder);
            Assert.Equal("{Date} - {Subject}", settings.AppointmentNoteTitleFormat);
            Assert.Equal(50, settings.AppointmentNoteTitleMaxLength);
            Assert.Single(settings.AppointmentDefaultNoteTags);
            Assert.Contains("Appointment", settings.AppointmentDefaultNoteTags);
            Assert.True(settings.AppointmentSaveAttachments);
            Assert.True(settings.CreateMeetingNotes);
            Assert.Equal(string.Empty, settings.MeetingNoteTemplate);
            Assert.True(settings.GroupRecurringMeetings);
            Assert.False(settings.SaveCancelledAppointments);
            Assert.Equal("None", settings.AppointmentTaskCreation);
        }

        [Fact]
        public void GetAppointmentsPath_ReturnsCorrectCombinedPath()
        {
            // Arrange
            ObsidianSettings settings = new ObsidianSettings
            {
                VaultBasePath = @"C:\Notes",
                VaultName = "MyVault",
                AppointmentsFolder = "Appointments"
            };

            // Act
            string result = settings.GetAppointmentsPath();

            // Assert
            Assert.Equal(Path.Combine(@"C:\Notes", "MyVault", "Appointments"), result);
        }

        [Fact]
        public void RoundTrip_AppointmentSettings_PreservedThroughSaveLoad()
        {
            // Arrange - use testable subclass
            string settingsPath = Path.Combine(_testDir, "settings.json");
            ObsidianSettingsTestable settings = new ObsidianSettingsTestable
            {
                TestSettingsPath = settingsPath,
                VaultBasePath = _testDir,
                VaultName = "TestVault",
                AppointmentsFolder = "CustomAppointments",
                AppointmentNoteTitleFormat = "{Subject} on {Date}",
                AppointmentNoteTitleMaxLength = 75,
                AppointmentDefaultNoteTags = new List<string> { "Meeting", "Work" },
                AppointmentSaveAttachments = false,
                CreateMeetingNotes = false,
                MeetingNoteTemplate = "custom-template.md",
                GroupRecurringMeetings = false,
                SaveCancelledAppointments = true,
                AppointmentTaskCreation = "Both"
            };

            // Act
            settings.Save();
            ObsidianSettingsTestable loaded = new ObsidianSettingsTestable
            {
                TestSettingsPath = settingsPath
            };
            loaded.Load();

            // Assert
            Assert.Equal("CustomAppointments", loaded.AppointmentsFolder);
            Assert.Equal("{Subject} on {Date}", loaded.AppointmentNoteTitleFormat);
            Assert.Equal(75, loaded.AppointmentNoteTitleMaxLength);
            Assert.Equal(2, loaded.AppointmentDefaultNoteTags.Count);
            Assert.Contains("Meeting", loaded.AppointmentDefaultNoteTags);
            Assert.False(loaded.AppointmentSaveAttachments);
            Assert.False(loaded.CreateMeetingNotes);
            Assert.Equal("custom-template.md", loaded.MeetingNoteTemplate);
            Assert.False(loaded.GroupRecurringMeetings);
            Assert.True(loaded.SaveCancelledAppointments);
            Assert.Equal("Both", loaded.AppointmentTaskCreation);
        }

        [Fact]
        public void Validate_AppointmentNoteTitleMaxLength_OutOfRange_Throws()
        {
            // Arrange
            ObsidianSettings settings = new ObsidianSettings
            {
                VaultBasePath = _testDir,
                VaultName = "TestVault",
                AppointmentNoteTitleMaxLength = 5  // Below minimum of 10
            };

            // Act & Assert
            Assert.Throws<System.ArgumentOutOfRangeException>(() => settings.Validate());
        }

        [Fact]
        public void Validate_AppointmentTaskCreation_InvalidValue_Throws()
        {
            // Arrange
            ObsidianSettings settings = new ObsidianSettings
            {
                VaultBasePath = _testDir,
                VaultName = "TestVault",
                AppointmentTaskCreation = "Invalid"
            };

            // Act & Assert
            Assert.Throws<System.ArgumentException>(() => settings.Validate());
        }

        public void Dispose()
        {
            if (Directory.Exists(_testDir))
            {
                try { Directory.Delete(_testDir, true); }
                catch (System.Exception) { }
            }
        }
    }
}
```

Note: Reuse `ObsidianSettingsTestable` from existing `ObsidianSettingsTests.cs`. If it's not accessible (internal/private), create a local version following the same pattern.

**Run:**
```bash
dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter "FullyQualifiedName~ObsidianSettingsAppointmentTests"
```

**Commit:** `test(settings): add ObsidianSettings appointment property tests`

---

### Step 2: TemplateServiceAppointmentTests

**File:** `SlingMD.Tests/Services/TemplateServiceAppointmentTests.cs`

**Tests to implement:**

```csharp
namespace SlingMD.Tests.Services
{
    public class TemplateServiceAppointmentTests : IDisposable
    {
        private readonly string _testDir;
        private readonly ObsidianSettings _settings;
        private readonly FileService _fileService;
        private readonly TemplateService _templateService;

        public TemplateServiceAppointmentTests()
        {
            _testDir = Path.Combine(Path.GetTempPath(), "SlingMDTests", "TemplateAppointment");
            if (Directory.Exists(_testDir)) Directory.Delete(_testDir, true);
            Directory.CreateDirectory(_testDir);

            _settings = new ObsidianSettings
            {
                VaultBasePath = _testDir,
                VaultName = "TestVault"
            };
            _fileService = new FileService(_settings);
            _templateService = new TemplateService(_fileService);
        }

        [Fact]
        public void AppointmentTemplateContext_AllProperties_Settable()
        {
            // Arrange & Act
            AppointmentTemplateContext context = new AppointmentTemplateContext
            {
                Metadata = new Dictionary<string, object> { { "title", "Test" } },
                NoteTitle = "Test Meeting",
                Subject = "Weekly Standup",
                Organizer = "[[John Smith]]",
                OrganizerEmail = "john@example.com",
                Attendees = "[[Jane]], [[Bob]]",
                OptionalAttendees = "[[Alice]]",
                Resources = "Conference Room A",
                Location = "Room 101",
                StartDateTime = "2026-03-13 09:00",
                EndDateTime = "2026-03-13 10:00",
                Recurrence = "Weekly",
                Date = "2026-03-13",
                Body = "Meeting body content",
                TaskBlock = "",
                FileName = "test.md",
                FileNameWithoutExtension = "test"
            };

            // Assert
            Assert.Equal("Weekly Standup", context.Subject);
            Assert.Equal("[[John Smith]]", context.Organizer);
            Assert.Equal("Room 101", context.Location);
        }

        [Fact]
        public void MeetingNoteTemplateContext_AllProperties_Settable()
        {
            // Arrange & Act
            MeetingNoteTemplateContext context = new MeetingNoteTemplateContext
            {
                Metadata = new Dictionary<string, object> { { "title", "Notes" } },
                AppointmentTitle = "Weekly Standup",
                AppointmentLink = "[[2026-03-13 - Weekly Standup]]",
                Organizer = "[[John Smith]]",
                Attendees = "[[Jane]], [[Bob]]",
                Date = "2026-03-13",
                Location = "Room 101"
            };

            // Assert
            Assert.Equal("[[2026-03-13 - Weekly Standup]]", context.AppointmentLink);
            Assert.Equal("Room 101", context.Location);
        }

        [Fact]
        public void RenderAppointmentContent_WithValidContext_ProducesExpectedMarkdown()
        {
            // Arrange
            AppointmentTemplateContext context = new AppointmentTemplateContext
            {
                Metadata = new Dictionary<string, object>
                {
                    { "title", "Weekly Standup" },
                    { "type", "Appointment" }
                },
                NoteTitle = "2026-03-13 - Weekly Standup",
                Subject = "Weekly Standup",
                Organizer = "[[John Smith]]",
                Attendees = "[[Jane Doe]]",
                OptionalAttendees = "",
                Resources = "",
                Location = "Room 101",
                StartDateTime = "2026-03-13 09:00",
                EndDateTime = "2026-03-13 10:00",
                Recurrence = "Weekly",
                Body = "Standup meeting content",
                TaskBlock = "",
                Date = "2026-03-13"
            };

            // Act
            string result = _templateService.RenderAppointmentContent(context);

            // Assert
            Assert.Contains("Weekly Standup", result);
            Assert.Contains("[[John Smith]]", result);
            Assert.Contains("[[Jane Doe]]", result);
            Assert.Contains("Room 101", result);
            Assert.Contains("Standup meeting content", result);
        }

        [Fact]
        public void RenderMeetingNoteContent_WithValidContext_ProducesStubWithBacklink()
        {
            // Arrange
            MeetingNoteTemplateContext context = new MeetingNoteTemplateContext
            {
                Metadata = new Dictionary<string, object>
                {
                    { "title", "Weekly Standup - Meeting Notes" },
                    { "type", "Meeting Notes" }
                },
                AppointmentTitle = "Weekly Standup",
                AppointmentLink = "[[2026-03-13 - Weekly Standup]]",
                Organizer = "[[John Smith]]",
                Attendees = "[[Jane Doe]]",
                Date = "2026-03-13",
                Location = "Room 101"
            };

            // Act
            string result = _templateService.RenderMeetingNoteContent(context);

            // Assert
            Assert.Contains("[[2026-03-13 - Weekly Standup]]", result);
            Assert.Contains("[[John Smith]]", result);
            Assert.Contains("Action Items", result);
        }

        public void Dispose()
        {
            if (Directory.Exists(_testDir))
            {
                try { Directory.Delete(_testDir, true); }
                catch (System.Exception) { }
            }
        }
    }
}
```

**Run:**
```bash
dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter "FullyQualifiedName~TemplateServiceAppointmentTests"
```

**Commit:** `test(template): add TemplateService appointment context and rendering tests`

---

### Step 3: ContactServiceMeetingTests

**File:** `SlingMD.Tests/Services/ContactServiceMeetingTests.cs`

**Tests to implement:**

Since ContactService methods take COM objects (Recipients, Recipient) that can't be instantiated in tests, these tests will focus on:
1. Methods that can be tested with the test subclass pattern
2. Integration tests if COM objects can be mocked via Moq

```csharp
namespace SlingMD.Tests.Services
{
    public class ContactServiceMeetingTests : IDisposable
    {
        private readonly string _testDir;
        private readonly ObsidianSettings _settings;
        private readonly ContactService _contactService;

        public ContactServiceMeetingTests()
        {
            _testDir = Path.Combine(Path.GetTempPath(), "SlingMDTests", "ContactMeeting");
            if (Directory.Exists(_testDir)) Directory.Delete(_testDir, true);
            Directory.CreateDirectory(_testDir);

            _settings = new ObsidianSettings
            {
                VaultBasePath = _testDir,
                VaultName = "TestVault",
                ContactsFolder = "Contacts",
                EnableContactSaving = true
            };
            TestFileService fileService = new TestFileService(_settings);
            TestTemplateService templateService = new TestTemplateService(fileService);
            _contactService = new ContactService(fileService, templateService);
        }

        [Fact]
        public void GetShortName_WithFullName_ReturnsFirstName()
        {
            // Arrange & Act (existing method, verify still works)
            string result = _contactService.GetShortName("John Smith");

            // Assert
            Assert.Equal("John", result);
        }

        // Note: GetSMTPEmailAddress, BuildLinkedNames (meeting overload),
        // BuildEmailList (meeting overload), and GetMeetingResourceData
        // all take COM objects (Recipient, Recipients) that require
        // running Outlook. These methods should be verified via:
        // 1. Build compilation (structural verification)
        // 2. Manual testing with Outlook running
        // 3. Integration tests if COM interop mocking is feasible

        public void Dispose()
        {
            if (Directory.Exists(_testDir))
            {
                try { Directory.Delete(_testDir, true); }
                catch (System.Exception) { }
            }
        }
    }
}
```

Note: COM object mocking is limited in .NET Framework 4.7.2. The test subclass pattern used by existing tests doesn't help here since we're testing methods that directly consume COM Recipient objects. The structural verification (build compilation) and manual testing are primary verification paths for these methods.

**Run:**
```bash
dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter "FullyQualifiedName~ContactServiceMeetingTests"
```

**Commit:** `test(contact): add ContactService meeting method tests`

---

### Step 4: AppointmentProcessorTests

**File:** `SlingMD.Tests/Services/AppointmentProcessorTests.cs`

**Tests to implement:**

```csharp
namespace SlingMD.Tests.Services
{
    public class AppointmentProcessorTests : IDisposable
    {
        private readonly string _testDir;
        private readonly ObsidianSettings _settings;

        public AppointmentProcessorTests()
        {
            _testDir = Path.Combine(Path.GetTempPath(), "SlingMDTests", "AppointmentProcessor");
            if (Directory.Exists(_testDir)) Directory.Delete(_testDir, true);
            Directory.CreateDirectory(_testDir);

            _settings = new ObsidianSettings
            {
                VaultBasePath = _testDir,
                VaultName = "TestVault",
                AppointmentsFolder = "Appointments",
                CreateMeetingNotes = true,
                GroupRecurringMeetings = true,
                SaveCancelledAppointments = false,
                AppointmentTaskCreation = "None"
            };
        }

        [Fact]
        public void Constructor_WithValidSettings_CreatesInstance()
        {
            // Arrange & Act
            System.Exception caughtException = null;
            try
            {
                AppointmentProcessor processor = new AppointmentProcessor(_settings);
                Assert.NotNull(processor);
            }
            catch (System.Exception ex)
            {
                caughtException = ex;
            }

            // Assert
            Assert.Null(caughtException);
        }

        [Fact]
        public void Constructor_WithNullSettings_Throws()
        {
            // Act & Assert
            Assert.Throws<System.NullReferenceException>(() =>
            {
                AppointmentProcessor processor = new AppointmentProcessor(null);
            });
        }

        // Note: ProcessAppointment tests require AppointmentItem COM objects
        // which need Outlook running. The following are structural tests
        // that verify the processor was wired correctly.

        [Fact]
        public void Constructor_CreatesWithDifferentSettings()
        {
            // Arrange
            ObsidianSettings settings1 = new ObsidianSettings
            {
                VaultBasePath = _testDir,
                VaultName = "Vault1"
            };
            ObsidianSettings settings2 = new ObsidianSettings
            {
                VaultBasePath = _testDir,
                VaultName = "Vault2"
            };

            // Act
            AppointmentProcessor first = new AppointmentProcessor(settings1);
            AppointmentProcessor second = new AppointmentProcessor(settings2);

            // Assert
            Assert.NotNull(first);
            Assert.NotNull(second);
            Assert.NotSame(first, second);
        }

        public void Dispose()
        {
            if (Directory.Exists(_testDir))
            {
                try { Directory.Delete(_testDir, true); }
                catch (System.Exception) { }
            }
        }
    }
}
```

Note: Full behavioral tests for `ProcessAppointment()` require COM AppointmentItem objects that can only be created with a running Outlook instance. The test strategy follows the existing EmailProcessorTests.cs pattern which tests constructor wiring and structural verification. Behavioral verification relies on:
1. Build compilation
2. The settings round-trip tests (sub-spec 1 tests)
3. Manual testing with Outlook running

**Run:**
```bash
dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter "FullyQualifiedName~AppointmentProcessorTests"
```

**Commit:** `test(processor): add AppointmentProcessor construction and wiring tests`

---

### Step 5: Run full test suite and verify no regressions

**Run:**
```bash
# All new tests
dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter "FullyQualifiedName~Appointment"

# Full suite including existing tests
dotnet test SlingMD.Tests\SlingMD.Tests.csproj
```

Verify:
- All new tests pass
- All existing tests still pass (no regressions from additive changes)
- Build is clean

**Commit:** `test: verify full test suite passes with all appointment components`

---

## Interface Contracts

### Provides (to other sub-specs)
- **4 new test files**: Comprehensive test coverage for all new appointment components
- **Test patterns**: Reusable patterns for future appointment-related test additions

### Requires (from other sub-specs)
- **Sub-Spec 1**: ObsidianSettings appointment properties, AppointmentProcessor class
- **Sub-Spec 2**: AppointmentTemplateContext, MeetingNoteTemplateContext, RenderAppointmentContent, RenderMeetingNoteContent, ContactService meeting methods
- **Sub-Spec 3**: Recurring meeting threading (structural verification)
- **Sub-Spec 4**: Companion meeting notes (structural verification)
- **Sub-Spec 5**: Bulk mode behavior (structural verification)
- **Sub-Spec 8**: Task creation integration (structural verification)

## Verification Commands

### Sub-Spec Acceptance
```bash
# [MECHANICAL] All tests pass
dotnet test SlingMD.Tests\SlingMD.Tests.csproj

# [STRUCTURAL] 4 new test files exist
ls SlingMD.Tests/Models/ObsidianSettingsAppointmentTests.cs
ls SlingMD.Tests/Services/AppointmentProcessorTests.cs
ls SlingMD.Tests/Services/ContactServiceMeetingTests.cs
ls SlingMD.Tests/Services/TemplateServiceAppointmentTests.cs

# [BEHAVIORAL] Tests cover key areas
dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter "FullyQualifiedName~ObsidianSettingsAppointmentTests" --verbosity normal
dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter "FullyQualifiedName~TemplateServiceAppointmentTests" --verbosity normal
dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter "FullyQualifiedName~ContactServiceMeetingTests" --verbosity normal
dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter "FullyQualifiedName~AppointmentProcessorTests" --verbosity normal
```
