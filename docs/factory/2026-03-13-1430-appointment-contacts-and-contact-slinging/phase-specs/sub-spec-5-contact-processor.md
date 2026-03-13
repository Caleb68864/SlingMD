---
type: phase-spec
master_spec: ../spec.md
sub_spec_number: 5
title: ContactProcessor Orchestrator
date: 2026-03-13
dependencies: [2, 3]
---

# Sub-Spec 5: ContactProcessor Orchestrator

## Scope

Create a new `ContactProcessor` class that orchestrates single-contact and address-book export flows. Follows the structural pattern of `AppointmentProcessor` exactly: constructor creates service instances, public Process methods orchestrate, `_bulkErrors` list with `GetBulkErrors()`.

### Codebase Findings

- **AppointmentProcessor pattern:** `SlingMD.Outlook/Services/AppointmentProcessor.cs`
  - Constructor (lines 73-82): takes `ObsidianSettings`, creates `FileService`, `TemplateService`, `ThreadService`, `TaskService`, `ContactService`, `AttachmentService`.
  - `_bulkErrors` list (line 64) with `GetBulkErrors()` method (lines 66-71).
  - `AppointmentProcessingResult` enum (lines 19-24): `Success`, `Skipped`, `Error`.
  - Async method signature: `public async Task<AppointmentProcessingResult> ProcessAppointment(...)`.
- **StatusService:** Exists at `SlingMD.Outlook/Services/StatusService.cs`. Used in AppointmentProcessor for progress feedback.
- **ContactService methods needed:** `ExtractContactData(ContactItem)` and `CreateRichContactNote(ContactTemplateContext)` from Sub-Spec 3. Also `ContactExists(string)` (existing).
- **COM release pattern:** `System.Runtime.InteropServices.Marshal.ReleaseComObject()` in finally blocks — seen in `ThisAddIn.SaveTodaysAppointments()` lines 239-255.
- **MAPIFolder iteration:** `ThisAddIn.SaveTodaysAppointments()` iterates folder items, casts to target type, skips nulls — same pattern for address book.
- **No existing ContactProcessor.cs** — new file creation required.
- **No existing ContactProcessorTests.cs** — new file creation required.

## Interface Contracts

### Provides
- `ContactProcessor` class in namespace `SlingMD.Outlook.Services`
- `ContactProcessingResult` enum: `Success`, `Skipped`, `Error`
- `ContactProcessor(ObsidianSettings settings)` constructor
- `ProcessContact(ContactItem contact)` → `ContactProcessingResult`
- `ProcessAddressBook(MAPIFolder contactsFolder)` → returns `(int saved, int skipped, int errors)` tuple or individual out params
- `GetBulkErrors()` → `List<string>`

### Requires
- `ContactService.ExtractContactData(ContactItem)` — from Sub-Spec 3
- `ContactService.CreateRichContactNote(ContactTemplateContext)` — from Sub-Spec 3
- `ContactService.ContactExists(string)` — existing
- `ContactTemplateContext.IncludeDetails` — from Sub-Spec 2
- `ObsidianSettings.ContactNoteIncludeDetails` — from Sub-Spec 4 (optional; can also use `true` default)

### Shared State
- Reads `_settings.ContactNoteIncludeDetails` to set `context.IncludeDetails`
- `_bulkErrors` accumulated during `ProcessAddressBook()`

## Implementation Steps

### Step 1: Create test file

**File:** `SlingMD.Tests/Services/ContactProcessorTests.cs` (NEW)

```csharp
using System;
using System.IO;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;
using Xunit;

namespace SlingMD.Tests.Services
{
    public class ContactProcessorTests : IDisposable
    {
        private readonly string _testDir;
        private readonly ObsidianSettings _settings;

        public ContactProcessorTests()
        {
            _testDir = Path.Combine(Path.GetTempPath(), "SlingMDTests", "ContactProcessor");
            if (Directory.Exists(_testDir))
            {
                Directory.Delete(_testDir, true);
            }
            Directory.CreateDirectory(_testDir);

            _settings = new ObsidianSettings
            {
                VaultBasePath = _testDir,
                VaultName = "TestVault",
                ContactsFolder = "Contacts",
                EnableContactSaving = true,
                ContactNoteIncludeDetails = true
            };
        }

        [Fact]
        public void Constructor_WithValidSettings_CreatesInstance()
        {
            System.Exception caughtException = null;
            ContactProcessor processor = null;
            try
            {
                processor = new ContactProcessor(_settings);
            }
            catch (System.Exception ex)
            {
                caughtException = ex;
            }

            Assert.Null(caughtException);
            Assert.NotNull(processor);
        }

        [Fact]
        public void Constructor_WithNullSettings_DoesNotThrow()
        {
            System.Exception caughtException = null;
            ContactProcessor processor = null;
            try
            {
                processor = new ContactProcessor(null);
            }
            catch (System.Exception ex)
            {
                caughtException = ex;
            }

            Assert.Null(caughtException);
            Assert.NotNull(processor);
        }

        [Fact]
        public void GetBulkErrors_InitiallyEmpty()
        {
            ContactProcessor processor = new ContactProcessor(_settings);
            var errors = processor.GetBulkErrors();
            Assert.Empty(errors);
        }

        public void Dispose()
        {
            try
            {
                if (Directory.Exists(_testDir))
                {
                    Directory.Delete(_testDir, true);
                }
            }
            catch { }
        }
    }
}
```

Run: `dotnet test SlingMD.Tests\SlingMD.Tests.csproj --filter "FullyQualifiedName~ContactProcessorTests"` — should FAIL (class doesn't exist).

### Step 2: Create ContactProcessor class

**File:** `SlingMD.Outlook/Services/ContactProcessor.cs` (NEW)

```csharp
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using SlingMD.Outlook.Models;

namespace SlingMD.Outlook.Services
{
    public enum ContactProcessingResult
    {
        Success,
        Skipped,
        Error
    }

    /// <summary>
    /// Orchestrates the export of Outlook <see cref="ContactItem"/> objects to rich contact notes
    /// in the Obsidian vault. Follows the same structural pattern as <see cref="AppointmentProcessor"/>.
    /// </summary>
    public class ContactProcessor
    {
        private readonly ObsidianSettings _settings;
        private readonly FileService _fileService;
        private readonly TemplateService _templateService;
        private readonly ContactService _contactService;
        private readonly StatusService _statusService;

        private List<string> _bulkErrors = new List<string>();

        public List<string> GetBulkErrors()
        {
            List<string> errors = new List<string>(_bulkErrors);
            _bulkErrors.Clear();
            return errors;
        }

        public ContactProcessor(ObsidianSettings settings)
        {
            _settings = settings;
            _fileService = new FileService(settings);
            _templateService = new TemplateService(_fileService);
            _contactService = new ContactService(_fileService, _templateService);
            _statusService = new StatusService();
        }

        /// <summary>
        /// Exports a single <see cref="ContactItem"/> to a rich contact note in the vault.
        /// </summary>
        public ContactProcessingResult ProcessContact(ContactItem contact)
        {
            try
            {
                ContactTemplateContext context = _contactService.ExtractContactData(contact);
                context.IncludeDetails = _settings?.ContactNoteIncludeDetails ?? true;

                // Check for existing note
                if (_contactService.ContactExists(context.ContactName))
                {
                    DialogResult choice = MessageBox.Show(
                        $"A contact note for \"{context.ContactName}\" already exists. Update it?",
                        "SlingMD — Contact Exists",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                    if (choice == DialogResult.No)
                    {
                        return ContactProcessingResult.Skipped;
                    }
                }

                _contactService.CreateRichContactNote(context);
                return ContactProcessingResult.Success;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(
                    $"Error processing contact: {ex.Message}",
                    "SlingMD Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return ContactProcessingResult.Error;
            }
        }

        /// <summary>
        /// Exports all contacts from the specified <paramref name="contactsFolder"/>.
        /// Non-ContactItem items (e.g. DistListItem) are silently skipped.
        /// </summary>
        public void ProcessAddressBook(MAPIFolder contactsFolder, out int saved, out int skipped, out int errors)
        {
            saved = 0;
            skipped = 0;
            errors = 0;

            Items items = null;
            try
            {
                items = contactsFolder.Items;
                int total = items.Count;
                int current = 0;

                using (StatusService status = new StatusService())
                {
                    status.Show($"Exporting contacts (0/{total})...");

                    foreach (object item in items)
                    {
                        ContactItem contact = item as ContactItem;
                        if (contact == null)
                        {
                            // Skip DistListItem and other non-contact items
                            if (item != null)
                            {
                                Marshal.ReleaseComObject(item);
                            }
                            skipped++;
                            continue;
                        }

                        try
                        {
                            current++;
                            string contactName = string.Empty;
                            try { contactName = contact.FullName ?? "Unknown"; } catch { contactName = "Unknown"; }
                            status.UpdateProgress($"Processing: {contactName} ({current}/{total})", (current * 100) / total);

                            ContactTemplateContext context = _contactService.ExtractContactData(contact);
                            context.IncludeDetails = _settings?.ContactNoteIncludeDetails ?? true;

                            // Skip if already exists (no prompt in bulk mode)
                            if (_contactService.ContactExists(context.ContactName))
                            {
                                skipped++;
                                continue;
                            }

                            _contactService.CreateRichContactNote(context);
                            saved++;
                        }
                        catch (System.Exception ex)
                        {
                            errors++;
                            _bulkErrors.Add($"Error processing contact: {ex.Message}");
                        }
                        finally
                        {
                            if (contact != null)
                            {
                                Marshal.ReleaseComObject(contact);
                            }
                        }
                    }
                }
            }
            finally
            {
                if (items != null)
                {
                    Marshal.ReleaseComObject(items);
                }
            }
        }
    }
}
```

### Step 3: Verify build and tests

```bash
dotnet build SlingMD.sln --configuration Release
dotnet test SlingMD.Tests\SlingMD.Tests.csproj
```

## Acceptance Criteria

- [STRUCTURAL] `ContactProcessor` class exists in `SlingMD.Outlook/Services/` namespace `SlingMD.Outlook.Services`.
- [STRUCTURAL] Constructor signature: `ContactProcessor(ObsidianSettings settings)`.
- [STRUCTURAL] Public methods: `ProcessContact(ContactItem)`, `ProcessAddressBook(MAPIFolder, out int, out int, out int)`, `GetBulkErrors()`.
- [BEHAVIORAL] `ProcessContact` checks for duplicates before creating.
- [BEHAVIORAL] `ProcessAddressBook` skips non-ContactItem items (DistListItem, etc.).
- [BEHAVIORAL] `ProcessAddressBook` collects errors and continues processing remaining contacts.
- [MECHANICAL] COM objects released in finally blocks.
- [STRUCTURAL] `ContactProcessorTests.cs` has constructor test verifying instantiation with valid settings.
- [STRUCTURAL] `ContactProcessorTests.cs` has constructor test verifying instantiation does not throw with null settings.

## Verification Commands

```bash
# Build
dotnet build SlingMD.sln --configuration Release

# Test
dotnet test SlingMD.Tests\SlingMD.Tests.csproj

# Verify class exists
grep -rn "class ContactProcessor" SlingMD.Outlook/Services/

# Verify test file exists
grep -rn "class ContactProcessorTests" SlingMD.Tests/Services/

# Verify enum exists
grep -n "ContactProcessingResult" SlingMD.Outlook/Services/ContactProcessor.cs
```

## Patterns to Follow

- **AppointmentProcessor constructor:** `SlingMD.Outlook/Services/AppointmentProcessor.cs` lines 73-82.
- **AppointmentProcessingResult enum:** Lines 19-24.
- **GetBulkErrors():** Lines 66-71.
- **SaveTodaysAppointments iteration:** `SlingMD.Outlook/ThisAddIn.cs` lines 179-292 — folder iteration, COM release, summary counts.
- **StatusService usage:** `AppointmentProcessor.cs` — `using (StatusService status = new StatusService())` with `Show()` and `UpdateProgress()`.
- **Test pattern:** `SlingMD.Tests/Services/AppointmentProcessorTests.cs` — constructor tests, IDisposable cleanup.

## Files

| File | Action | Notes |
|------|--------|-------|
| `SlingMD.Outlook/Services/ContactProcessor.cs` | Create | New orchestrator class |
| `SlingMD.Tests/Services/ContactProcessorTests.cs` | Create | New test file |
