# Spec: Appointment Contact Linking & Contact Slinging

## Meta

- **Run ID:** 2026-03-13-1430-appointment-contacts-and-contact-slinging
- **Stage:** 2 (Forge)
- **Source:** design.md from Stage 1
- **Features:** 3 (Appointment Contact Linking, Contact Slinging Infrastructure, Contact Slinging UI)
- **Sub-specs:** 6

### Quality Scores

| Dimension | Score | Notes |
|-----------|-------|-------|
| outcome_clarity | 5 | Each sub-spec has a concrete "done when" with testable criteria |
| scope_boundaries | 5 | Explicit out-of-scope list, EmailProcessor untouched constraint repeated per sub-spec |
| decision_guidance | 4 | Trade-off hierarchy provided, open questions resolved; some COM edge cases left to implementer judgment |
| edge_coverage | 4 | COM failures, duplicates, empty attendee lists, bulk mode all covered; edge cases for malformed ContactItem data deferred to implementer |
| acceptance_criteria | 5 | All criteria typed [MECHANICAL]/[STRUCTURAL]/[BEHAVIORAL]/[HUMAN REVIEW] |
| decomposition | 5 | 6 sub-specs, each 1-3 files, independently executable with clear dependency chain |
| **Total** | **28** | Walk-away quality |

---

## Outcome

When this spec is fully implemented:

1. **Appointment contact linking:** After an appointment is exported to Obsidian, the processor checks all attendee names against the vault. Existing managed contact notes are refreshed. New contacts trigger a `ContactConfirmationDialog` (single mode) or are silently skipped (bulk mode). The attendee `[[Name]]` wiki-links in the exported note now have a meaningful target.

2. **Contact slinging (single):** A user can select a `ContactItem` in Outlook and click "Sling" (or "Sling Contact" ribbon button) to export it as a rich contact note with phone, email, company, title, address, birthday, and notes fields.

3. **Contact slinging (address book):** A user can click "Sling All Contacts" to bulk-export their entire default Contacts folder, with progress feedback and a summary dialog.

**Done-when:** All 6 sub-spec acceptance criteria pass, the solution builds without errors, and all new tests pass.

---

## Intent (Trade-off Hierarchy)

When making implementation decisions, prioritize in this order:

1. **Do not break email slinging.** EmailProcessor.cs must have ZERO modifications. Not one line.
2. **Follow existing patterns.** Copy the proven EmailProcessor contact resolution flow (lines 466-516) rather than inventing new approaches.
3. **Minimize settings bloat.** One new setting (`ContactNoteIncludeDetails`). Reuse existing contact settings.
4. **Prefer duplication over abstraction risk.** Duplicating 30 lines of contact resolution is safer than extracting a shared method that EmailProcessor would also call.
5. **Sequential over parallel.** COM interop is single-threaded. Address book export processes contacts one at a time.

---

## Context

### Codebase References

| Component | Path | Relevance |
|-----------|------|-----------|
| AppointmentProcessor | `SlingMD.Outlook/Services/AppointmentProcessor.cs` | Insert contact linking after line 566 (after `coreExportSucceeded` check, before Outlook task creation) |
| EmailProcessor contact flow | `SlingMD.Outlook/Services/EmailProcessor.cs:466-516` | Pattern to replicate — do NOT modify this file |
| ContactService | `SlingMD.Outlook/Services/ContactService.cs` | Existing methods: `ContactExists()`, `ManagedContactNoteExists()`, `CreateContactNote()`, `GetShortName()`. Add `CreateRichContactNote()` and `ExtractContactData()` |
| ContactConfirmationDialog | `SlingMD.Outlook/Forms/ContactConfirmationDialog.cs` | Reuse as-is for appointment contact linking and contact slinging from-email |
| ContactTemplateContext | `SlingMD.Outlook/Services/TemplateService.cs:29-37` | Extend with rich fields (Phone, Email, Company, Title, Address, Birthday, Notes) |
| TemplateService | `SlingMD.Outlook/Services/TemplateService.cs` | Add `RenderRichContactContent()` or extend `RenderContactContent()` |
| ObsidianSettings | `SlingMD.Outlook/Models/ObsidianSettings.cs` | Add `ContactNoteIncludeDetails` property + update `NormalizeLoadedSettings()` |
| ThisAddIn | `SlingMD.Outlook/ThisAddIn.cs` | Add `ProcessSelectedContact()`, `SlingAllContacts()`, update `ProcessSelection()` to detect `ContactItem` |
| SlingRibbon.xml | `SlingMD.Outlook/Ribbon/SlingRibbon.xml` | Add Contacts group with two buttons |
| SlingRibbon.cs | `SlingMD.Outlook/Ribbon/SlingRibbon.cs` | Add callback methods for new ribbon buttons |
| SettingsForm | `SlingMD.Outlook/Forms/SettingsForm.cs` | Add checkbox on existing Contacts tab |

### Test Conventions

- Tests live in `SlingMD.Tests/Services/` and `SlingMD.Tests/Models/`
- xUnit with `[Fact]` attributes
- No COM mocking — structural tests only (constructor, method existence)
- Filesystem tests use `Path.GetTempPath()` temp directories with cleanup in `Dispose()`
- `TestFileService` subclass overrides `GetSettings()`, `EnsureDirectoryExists()`, `WriteUtf8File()`, `CleanFileName()`
- Test class naming: `{ClassName}Tests.cs`

### Attendee Name Extraction (AppointmentProcessor)

The AppointmentProcessor already extracts attendee names into three lists at lines 180-182:
- `requiredAttendees` (List<string>)
- `optionalAttendees` (List<string>)
- `resourceAttendees` (List<string>)
- `organizerName` (string)

These are populated via `ContactService.BuildLinkedNames()` which returns `[[Name]]` formatted strings. The contact resolution step must strip the `[[` and `]]` wrappers to get plain names.

---

## Requirements

### R1: Appointment Contact Linking
After a successful appointment export, resolve all attendee names against the vault and offer to create missing contact notes.

### R2: ContactProcessor Infrastructure
Create a new `ContactProcessor` orchestrator class that exports `ContactItem` objects to rich contact notes in the vault.

### R3: ContactService Extensions
Extend `ContactService` with methods to extract rich data from `ContactItem` objects and create detailed contact notes.

### R4: ContactTemplateContext Extensions
Extend the template context and rendering to support rich contact fields.

### R5: Contact Slinging UI
Add ribbon buttons and ThisAddIn dispatch methods for single-contact and address-book export.

### R6: Settings
Add `ContactNoteIncludeDetails` setting with UI toggle.

---

## Sub-Specs

### Sub-Spec 1: Appointment Contact Linking

**Scope:** Add post-export contact resolution to `AppointmentProcessor.ProcessAppointment()`.

**Files likely touched:**
- `SlingMD.Outlook/Services/AppointmentProcessor.cs` (primary)
- `SlingMD.Tests/Services/AppointmentProcessorTests.cs` (tests)

**What to implement:**

After the `coreExportSucceeded` check (line 566) and BEFORE the Outlook task creation block (line 568), insert a contact resolution block that:

1. Checks `_settings.EnableContactSaving` — skip if false.
2. Collects all attendee names from `organizerName`, `requiredAttendees`, and `optionalAttendees`. These variables are already populated earlier in the method. Note: `requiredAttendees` and `optionalAttendees` contain `[[Name]]` formatted strings — strip the brackets to get plain names. `resourceAttendees` should be excluded (conference rooms are not contacts).
3. Deduplicates and sorts the name list (use `.Distinct().OrderBy(n => n).ToList()`).
4. For each name:
   - If `_contactService.ManagedContactNoteExists(name)` is true, add to `managedContactsToRefresh` list.
   - Else if `_contactService.ContactExists(name)` is false, add to `newContacts` list.
   - Otherwise (user-managed note exists elsewhere), skip.
5. Refresh all managed contacts: call `_contactService.CreateContactNote(name)` for each.
6. If `!bulkMode` and `newContacts.Count > 0`, show `ContactConfirmationDialog` and create selected contacts.
7. Wrap entire block in try/catch with `System.Exception`, showing `MessageBox` on error (matching EmailProcessor pattern).

**Acceptance Criteria:**
- [STRUCTURAL] `AppointmentProcessor` has a using/reference for `SlingMD.Outlook.Forms` (for `ContactConfirmationDialog`). Verify with grep.
- [MECHANICAL] The contact resolution block appears after `coreExportSucceeded` check and before the Outlook task creation block.
- [BEHAVIORAL] When `EnableContactSaving` is false, no contact resolution logic executes.
- [BEHAVIORAL] In `bulkMode`, no dialog is shown; only existing managed contacts are refreshed.
- [BEHAVIORAL] Resource attendees (conference rooms) are excluded from contact resolution.
- [MECHANICAL] `EmailProcessor.cs` has zero modifications (diff against main branch must be empty for this file).
- [STRUCTURAL] New test in `AppointmentProcessorTests.cs` verifies the processor still constructs successfully with contact-related settings.

**Dependencies:** None (can be implemented first).

**Estimated effort:** 30-45 minutes.

---

### Sub-Spec 2: ContactTemplateContext & TemplateService Extensions

**Scope:** Extend `ContactTemplateContext` with rich fields and add rendering support.

**Files likely touched:**
- `SlingMD.Outlook/Services/TemplateService.cs` (primary)
- `SlingMD.Outlook/Templates/RichContactTemplate.md` (new default template)
- `SlingMD.Tests/Services/TemplateServiceTests.cs` (tests)

**What to implement:**

1. Add the following properties to `ContactTemplateContext` (at `TemplateService.cs:29-37`):
   - `Phone` (string, default empty)
   - `Email` (string, default empty)
   - `Company` (string, default empty)
   - `JobTitle` (string, default empty)
   - `Address` (string, default empty)
   - `Birthday` (string, default empty)
   - `Notes` (string, default empty)
   - `IncludeDetails` (bool, default true)

2. Add a `RenderRichContactContent(ContactTemplateContext context)` method or extend `RenderContactContent()` to handle the new fields. When `IncludeDetails` is true, render a "## Contact Details" section above the Communication History section with the rich fields. When false, render the standard contact template only.

3. Create a default rich contact template (embedded resource or inline string, following the pattern of `GetDefaultContactTemplate()`). Include frontmatter fields for phone, email, company, title. Include a `## Contact Details` section with the rich fields in a readable format.

**Acceptance Criteria:**
- [STRUCTURAL] `ContactTemplateContext` has properties: Phone, Email, Company, JobTitle, Address, Birthday, Notes, IncludeDetails.
- [MECHANICAL] `RenderRichContactContent()` (or extended `RenderContactContent()`) produces markdown containing `## Contact Details` when `IncludeDetails` is true.
- [MECHANICAL] When `IncludeDetails` is false, the output does NOT contain `## Contact Details`.
- [BEHAVIORAL] Template token replacement works for all new fields (e.g., `{Phone}`, `{Email}`, `{Company}`).
- [STRUCTURAL] New test verifies rich contact template rendering with all fields populated.
- [STRUCTURAL] New test verifies rich contact template rendering with empty fields (graceful — no "null" strings).

**Dependencies:** None (can be implemented in parallel with Sub-Spec 1).

**Estimated effort:** 30-45 minutes.

---

### Sub-Spec 3: ContactService Extensions

**Scope:** Add methods to `ContactService` for extracting rich data from Outlook `ContactItem` objects and creating rich contact notes.

**Files likely touched:**
- `SlingMD.Outlook/Services/ContactService.cs` (primary)
- `SlingMD.Tests/Services/ContactServiceTests.cs` (tests)

**What to implement:**

1. Add `ExtractContactData(ContactItem contact)` method that returns a `ContactTemplateContext` populated with:
   - `ContactName` from `contact.FullName` (fallback to `contact.LastName + ", " + contact.FirstName`, then `contact.FileAs`, then "Unknown Contact")
   - `Phone` from `contact.BusinessTelephoneNumber` (fallback to `contact.MobileTelephoneNumber`, then `contact.HomeTelephoneNumber`)
   - `Email` from `contact.Email1Address`
   - `Company` from `contact.CompanyName`
   - `JobTitle` from `contact.JobTitle`
   - `Address` from `contact.BusinessAddress` (fallback to `contact.HomeAddress`)
   - `Birthday` from `contact.Birthday` (format as `yyyy-MM-dd`, skip if `DateTime.MinValue` or 4501 year which Outlook uses for "not set")
   - `Notes` from `contact.Body`
   - Each property read wrapped in individual try/catch with fallback to empty string (matching AppointmentProcessor pattern at lines 184-203).

2. Add `CreateRichContactNote(ContactTemplateContext context)` method that:
   - Uses `context.ContactName` to determine file path via existing `GetManagedContactNotePath()`
   - Calls `_templateService.RenderRichContactContent(context)` to get content
   - If file does not exist, writes new file
   - If file exists, merges using existing `MergeManagedSections()` pattern but also updates the Contact Details section

3. Ensure COM objects are properly released with `Marshal.ReleaseComObject()` where applicable.

**Acceptance Criteria:**
- [STRUCTURAL] `ContactService` has public methods `ExtractContactData` and `CreateRichContactNote`.
- [BEHAVIORAL] `ExtractContactData` handles null/missing properties gracefully — each property read is individually try/caught.
- [BEHAVIORAL] Birthday value of year 4501 (Outlook's "not set" sentinel) is treated as empty string.
- [MECHANICAL] `CreateRichContactNote` writes to the same contacts folder as `CreateContactNote`.
- [STRUCTURAL] New tests verify `CreateRichContactNote` creates a file with expected content.
- [STRUCTURAL] New tests verify `CreateRichContactNote` merges correctly when file already exists.
- [MECHANICAL] No modifications to existing `CreateContactNote()` method signature or behavior.

**Dependencies:** Sub-Spec 2 (needs extended `ContactTemplateContext` and `RenderRichContactContent()`).

**Estimated effort:** 45-60 minutes.

---

### Sub-Spec 4: Settings — ContactNoteIncludeDetails

**Scope:** Add the `ContactNoteIncludeDetails` setting and wire it into the Settings UI.

**Files likely touched:**
- `SlingMD.Outlook/Models/ObsidianSettings.cs` (primary)
- `SlingMD.Outlook/Forms/SettingsForm.cs` (UI)
- `SlingMD.Tests/Models/ObsidianSettingsTests.cs` (tests)

**What to implement:**

1. Add to `ObsidianSettings`:
   ```
   public bool ContactNoteIncludeDetails { get; set; } = true;
   ```
   Place it near the existing contact settings (`ContactsFolder`, `EnableContactSaving`, `SearchEntireVaultForContacts`, `ContactFilenameFormat`).

2. Update `NormalizeLoadedSettings()` to handle this new property (ensure backward compatibility — old settings files without this property default to `true`).

3. Add a checkbox to the existing Contacts tab in `SettingsForm.cs`:
   - Label: "Include contact details (phone, email, company, etc.)"
   - Bound to `ContactNoteIncludeDetails`
   - Placed after existing contact settings on the Contacts tab

**Acceptance Criteria:**
- [STRUCTURAL] `ObsidianSettings` has `ContactNoteIncludeDetails` property with default value `true`.
- [MECHANICAL] `NormalizeLoadedSettings()` does not throw when deserializing settings JSON that lacks `ContactNoteIncludeDetails`.
- [STRUCTURAL] SettingsForm Contacts tab has checkbox for `ContactNoteIncludeDetails`.
- [BEHAVIORAL] Setting persists correctly through Save/Load cycle.
- [STRUCTURAL] New test verifies default value is `true`.
- [MECHANICAL] Existing settings tests still pass.

**Dependencies:** None (can be implemented in parallel with Sub-Specs 1-3).

**Estimated effort:** 20-30 minutes.

---

### Sub-Spec 5: ContactProcessor Orchestrator

**Scope:** Create the `ContactProcessor` class that orchestrates single-contact and address-book export flows.

**Files likely touched:**
- `SlingMD.Outlook/Services/ContactProcessor.cs` (new file)
- `SlingMD.Tests/Services/ContactProcessorTests.cs` (new file)

**What to implement:**

1. Create `ContactProcessor` class following the structural pattern of `AppointmentProcessor`:
   - Constructor takes `ObsidianSettings`, creates `FileService`, `TemplateService`, `ContactService` instances (same pattern as AppointmentProcessor constructor at lines 73-82).
   - Private fields: `_settings`, `_fileService`, `_templateService`, `_contactService`.
   - Private `_bulkErrors` list with `GetBulkErrors()` method (matching AppointmentProcessor pattern).

2. `ProcessContact(ContactItem contact)` method:
   - Call `_contactService.ExtractContactData(contact)` to get `ContactTemplateContext`.
   - Set `context.IncludeDetails` from `_settings.ContactNoteIncludeDetails`.
   - Check `_contactService.ContactExists(context.ContactName)` — if exists, warn and offer to update (using `MessageBox` with Yes/No).
   - Call `_contactService.CreateRichContactNote(context)`.
   - Return success/skip/error result (define a `ContactProcessingResult` enum matching `AppointmentProcessingResult`).

3. `ProcessAddressBook(MAPIFolder contactsFolder)` method:
   - Iterate all items in the folder, cast each to `ContactItem` (skip non-contacts like `DistListItem`).
   - Show progress via `StatusService` (instantiate in constructor).
   - For each contact: check exists (skip if so, unless refresh needed), extract data, create note.
   - Collect errors in `_bulkErrors`.
   - Return counts: saved, skipped, errors.
   - Release COM objects in finally blocks.

4. Add `ContactProcessor.cs` to the `.csproj` if needed (the project likely uses wildcard includes).

**Acceptance Criteria:**
- [STRUCTURAL] `ContactProcessor` class exists in `SlingMD.Outlook/Services/` namespace `SlingMD.Outlook.Services`.
- [STRUCTURAL] Constructor signature: `ContactProcessor(ObsidianSettings settings)`.
- [STRUCTURAL] Public methods: `ProcessContact(ContactItem)`, `ProcessAddressBook(MAPIFolder)`, `GetBulkErrors()`.
- [BEHAVIORAL] `ProcessContact` checks for duplicates before creating.
- [BEHAVIORAL] `ProcessAddressBook` skips non-ContactItem items (DistListItem, etc.).
- [BEHAVIORAL] `ProcessAddressBook` collects errors and continues processing remaining contacts.
- [MECHANICAL] COM objects released in finally blocks.
- [STRUCTURAL] `ContactProcessorTests.cs` has constructor test verifying instantiation with valid settings.
- [STRUCTURAL] `ContactProcessorTests.cs` has constructor test verifying instantiation does not throw with null settings.

**Dependencies:** Sub-Spec 2, Sub-Spec 3 (needs ContactService extensions and template extensions).

**Estimated effort:** 60-90 minutes.

---

### Sub-Spec 6: Contact Slinging UI (Ribbon + ThisAddIn Dispatch)

**Scope:** Add ribbon buttons for contact operations and wire up ThisAddIn dispatch.

**Files likely touched:**
- `SlingMD.Outlook/Ribbon/SlingRibbon.xml` (ribbon layout)
- `SlingMD.Outlook/Ribbon/SlingRibbon.cs` (callbacks)
- `SlingMD.Outlook/ThisAddIn.cs` (dispatch methods)

**What to implement:**

1. **SlingRibbon.xml** — Add a "Contacts" group after the Appointments group:
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

2. **SlingRibbon.cs** — Add callback methods following the existing pattern (try/catch with MessageBox):
   - `OnSlingContactClick` → calls `_addIn.ProcessSelectedContact()`
   - `OnSlingAllContactsClick` → calls `_addIn.SlingAllContacts()`

3. **ThisAddIn.cs** — Add:
   - Private field `_contactProcessor` (initialized in `ThisAddIn_Startup` alongside `_emailProcessor` and `_appointmentProcessor`).
   - Update `ProcessSelection()` to detect `ContactItem`: after the `AppointmentItem` check, add `ContactItem contact = selected as ContactItem;` and dispatch to `_contactProcessor.ProcessContact(contact)`.
   - `ProcessSelectedContact()` method: get selected item from explorer, validate it's a `ContactItem`, call `_contactProcessor.ProcessContact(contact)`.
   - `SlingAllContacts()` method: get default Contacts folder via `Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderContacts)`, call `_contactProcessor.ProcessAddressBook(folder)`, show summary dialog with counts (matching `SaveTodaysAppointments()` pattern at lines 179+).

4. Update the "Please select an email or appointment" message in `ProcessSelection()` to include "contact": "Please select an email, appointment, or contact."

**Acceptance Criteria:**
- [STRUCTURAL] SlingRibbon.xml contains a `ContactsGroup` with `SlingContactButton` and `SlingAllContactsButton`.
- [STRUCTURAL] SlingRibbon.cs has `OnSlingContactClick` and `OnSlingAllContactsClick` callback methods.
- [STRUCTURAL] ThisAddIn.cs has `_contactProcessor` field initialized in `ThisAddIn_Startup`.
- [STRUCTURAL] ThisAddIn.cs has `ProcessSelectedContact()` and `SlingAllContacts()` public methods.
- [BEHAVIORAL] `ProcessSelection()` detects `ContactItem` and dispatches to `ContactProcessor`.
- [BEHAVIORAL] `SlingAllContacts()` shows a summary dialog with saved/skipped/error counts after completion.
- [BEHAVIORAL] Error in contact processing shows user-friendly MessageBox (not unhandled exception).
- [MECHANICAL] `EmailProcessor.cs` has zero modifications.
- [HUMAN REVIEW] Ribbon layout looks correct in Outlook — Contacts group appears after Appointments group with appropriate icons.

**Dependencies:** Sub-Spec 5 (needs ContactProcessor).

**Estimated effort:** 45-60 minutes.

---

## Edge Cases

| Edge Case | Handling |
|-----------|----------|
| Appointment with no attendees (personal event) | Contact resolution block skips when attendee list is empty. No dialog shown. |
| Attendee name is a conference room or resource | Resource attendees (`resourceAttendees` list) are excluded from contact resolution. |
| `ContactItem` with all null/empty properties | `ExtractContactData` returns context with empty strings for all fields. Note is still created with name "Unknown Contact". |
| Duplicate contact sling (re-export same contact) | `ProcessContact` detects existing note, prompts user to update or skip. |
| Address book with 0 contacts | `ProcessAddressBook` shows summary with 0 saved, 0 skipped, 0 errors. No error. |
| Address book contains `DistListItem` (distribution lists) | `ProcessAddressBook` casts to `ContactItem` — null result is skipped. |
| COM interop failure mid-address-book | Error collected in `_bulkErrors`, processing continues to next contact. Summary shows error count. |
| `ContactItem.Birthday` returns year 4501 | Treated as "not set" — Birthday field rendered as empty string. |
| Settings file lacks `ContactNoteIncludeDetails` (upgrade scenario) | `NormalizeLoadedSettings()` ensures default value `true` is applied. |
| `EnableContactSaving` is false | Appointment contact linking skips entirely. Contact slinging still works (it creates notes regardless — the setting only gates the post-export resolution in email/appointment flows). |
| Attendee names wrapped in `[[` and `]]` | Strip brackets before passing to ContactService lookup methods. |

---

## Out of Scope

- **Contact photo export** — Extracting and embedding contact photos adds AttachmentService complexity. Deferred to future run.
- **Multi-account address book** — "Sling All Contacts" exports from default account only. Future run can add account selection.
- **SubjectCleanupService extraction** — Both processors keep their own regex copies. Future run can consolidate.
- **EmailProcessor modifications** — Zero changes to EmailProcessor.cs. Period.
- **Interface extraction** — No `IContactService`, `IFileService`, etc. Consistent with existing codebase.
- **Per-field toggle settings** — Users cannot individually choose which contact fields to include. The single `ContactNoteIncludeDetails` bool controls all-or-nothing.
- **Contact photo as note icon** — Not part of this run.
- **Contact groups/distribution lists** — Not exported; silently skipped.
- **From-email contact export with ContactItem resolution** — The design mentions attempting `Recipient.AddressEntry.GetContact()` to resolve rich data. This is deferred; the existing email contact flow (name-only notes) remains unchanged.

---

## Constraints

1. **.NET Framework 4.7.2** — No C# 8+ features (no nullable reference types, no default interface methods, no async streams).
2. **COM interop** — All Outlook object property reads must be individually try/caught. COM objects must be released in finally blocks.
3. **Single-threaded COM** — Address book processing must be sequential. No `Task.WhenAll` or `Parallel.ForEach`.
4. **Naming conventions** — PascalCase for public members, _camelCase for private fields, explicit typing (no `var`).
5. **Exception handling** — Use `System.Exception` fully qualified. User-facing errors via `MessageBox.Show()`. Use `throw;` not `throw ex;`.
6. **EmailProcessor.cs** — Must have exactly zero modifications. Verified by git diff.

---

## Verification

### Build Gate
```
dotnet build SlingMD.sln --configuration Release
```
Must complete with 0 errors.

### Test Gate
```
dotnet test SlingMD.Tests\SlingMD.Tests.csproj
```
All existing tests must pass. All new tests must pass.

### Diff Gate
```
git diff main -- SlingMD.Outlook/Services/EmailProcessor.cs
```
Must produce empty output (no changes to EmailProcessor).

### Manual Verification (HUMAN REVIEW)
1. Open Outlook with the add-in loaded.
2. Verify "Contacts" group appears on the Sling ribbon tab with two buttons.
3. Select an appointment with attendees → click Sling → verify contact confirmation dialog appears for new contacts.
4. Select a ContactItem → click Sling → verify rich contact note is created in vault.
5. Click "Sling All Contacts" → verify progress and summary dialog.
6. Open Settings → Contacts tab → verify "Include contact details" checkbox.
