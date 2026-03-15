# Design: Appointment Contact Linking, Modular Architecture, Contact Slinging

## Summary

Three related features that expand SlingMD's contact capabilities:

1. **Appointment Contact Linking** — The `AppointmentProcessor` currently writes `[[Name]]` wiki-links for attendees in frontmatter and note body but never checks whether those contacts actually exist in the vault, and never offers to create them. This feature adds vault lookup and the same creation-offer flow that `EmailProcessor` already uses.

2. **Modular Architecture** — Extract shared logic (contact resolution, subject cleanup, Obsidian launch, task creation orchestration) into reusable services so that `EmailProcessor`, `AppointmentProcessor`, and the new `ContactProcessor` share code without modifying `EmailProcessor`'s battle-tested flow.

3. **Contact Slinging** — A new top-level capability to export Outlook `ContactItem` objects to Obsidian as contact notes, supporting single-contact, all-contacts-from-an-email, and full-address-book export modes.

---

## Architecture

### Current State

```
ThisAddIn
  ├─ EmailProcessor    (orchestrates email → markdown)
  │    ├─ FileService, TemplateService, ThreadService
  │    ├─ TaskService, ContactService, AttachmentService
  │    └─ StatusService (progress UI)
  └─ AppointmentProcessor  (orchestrates appointment → markdown)
       ├─ FileService, TemplateService, ThreadService
       ├─ TaskService, ContactService, AttachmentService
       └─ StatusService
```

Both processors instantiate their own copies of every service in their constructors. `ContactService` provides name utilities, vault lookup (`ContactExists`, `ManagedContactNoteExists`), and note creation (`CreateContactNote`), but the appointment path never calls the lookup/create flow.

### Proposed State

```
ThisAddIn
  ├─ EmailProcessor          (UNCHANGED — no modifications)
  ├─ AppointmentProcessor    (+ contact linking post-export)
  ├─ ContactProcessor  [NEW] (orchestrates contact → markdown)
  │
  ├─ Shared services (same instances, no code changes):
  │    FileService, TemplateService, ThreadService,
  │    TaskService, AttachmentService, StatusService
  │
  └─ ContactService          (+ new methods for ContactItem export)
```

Key constraint: **EmailProcessor is not touched.** All shared extraction happens in new or extended services.

---

## Approach

### Feature 1: Appointment Contact Linking

**What changes:**
- `AppointmentProcessor.ProcessAppointment()` gains a post-export contact resolution step identical to the one in `EmailProcessor` (lines 466-516 of EmailProcessor.cs).
- After `coreExportSucceeded = true`, collect all attendee names (organizer + required + optional), deduplicate, check `ContactService.ManagedContactNoteExists()` / `ContactService.ContactExists()`, refresh managed notes, and show `ContactConfirmationDialog` for truly new contacts.
- In `bulkMode`, skip the dialog and silently refresh existing managed contacts only (matching the email processor's approach of not blocking bulk operations with UI).

**What does NOT change:**
- The frontmatter already contains `[[Name]]` wiki-links — that stays as-is.
- `ContactService` methods are already sufficient; no changes needed there for this feature.

### Feature 2: Modular Architecture

**Philosophy:** Extract duplicated code into shared utilities without restructuring the working orchestrators. This is a "pull shared code down" refactor, not a "push orchestration up" refactor. EmailProcessor stays completely untouched.

**Candidates for extraction:**

| Duplicated Code | Current Location | Proposed Location |
|---|---|---|
| `CleanSubject()` + 14 compiled regex fields | Both processors (identical) | `SubjectCleanupService` (new) |
| Contact resolution flow (check exists, show dialog, create) | `EmailProcessor` lines 466-516 | `ContactResolutionService` (new) |
| Obsidian launch flow (countdown + delay + launch) | Both processors (near-identical) | `LaunchService` (new) |

**Decision: Defer extraction to a follow-up refactor.** The brain dump says "if it makes sense" and "don't touch the bullet proof and tested code." The safest approach for this run is:

1. **Do NOT extract from EmailProcessor.** Leave its code entirely in place.
2. **In AppointmentProcessor,** call the same service methods (`ContactService`, `ContactConfirmationDialog`) directly — copy the pattern, not the code. This is a small amount of duplication but avoids any risk of regressing email slinging.
3. **For ContactProcessor,** reuse existing services directly. It is a new orchestrator on par with EmailProcessor and AppointmentProcessor.
4. **Create a shared `SubjectCleanupService`** that both AppointmentProcessor and ContactProcessor can use. EmailProcessor keeps its own copy. When confidence is high (after tests pass), a future run can migrate EmailProcessor to use it too.

This honors the constraint: "don't touch the bullet proof and tested code for email slinging if you don't need to."

### Feature 3: Contact Slinging

**New class: `ContactProcessor`** — orchestrates exporting Outlook `ContactItem` objects to Obsidian markdown contact notes.

**Three entry points (all from ribbon/ThisAddIn):**

| Mode | Trigger | Behavior |
|---|---|---|
| Single contact | Select a contact in Outlook, click Sling | Export one `ContactItem` to a contact note |
| From email | New ribbon button or menu | Collect all names from the selected email, show `ContactConfirmationDialog`, create selected |
| Full address book | New ribbon button | Iterate all contacts in the default Contacts folder, show progress, export each |

**Contact note content for ContactItem-sourced notes:**
- Richer than email-derived contact notes: includes phone numbers, email addresses, company, job title, physical address, birthday, notes field.
- Uses an extended `ContactTemplateContext` with additional fields.
- Reuses `ContactService.CreateContactNote()` for the managed section (Communication History), but prepends a richer "Contact Details" section.

---

## Key Decisions

1. **Do not modify EmailProcessor.** All three features can be delivered without changing a single line in `EmailProcessor.cs`. The contact resolution pattern is duplicated into `AppointmentProcessor` rather than extracted into a shared method that EmailProcessor also calls. This is deliberate — the risk/reward ratio of touching proven code is unfavorable.

2. **ContactProcessor is a new top-level orchestrator** on par with EmailProcessor and AppointmentProcessor. It gets its own `ContactProcessor.cs` in `Services/` and follows the same structural pattern (constructor creates service instances, public `Process*` methods orchestrate).

3. **No new interfaces at this time.** The codebase has no existing interfaces for services. Adding `IContactService` etc. would be a broader architectural change that should be its own run. Services continue to be concrete classes.

4. **Bulk address book export uses StatusService** for progress feedback and processes contacts sequentially to avoid overwhelming COM interop.

5. **AppointmentProcessor contact linking in bulk mode** silently refreshes existing managed contacts but does NOT prompt for new contact creation (matches EmailProcessor's implicit behavior where bulk mode skips UI).

6. **ContactService gets new methods** for extracting rich contact data from `ContactItem` objects, but existing methods remain unchanged.

7. **New settings for contact slinging** are minimal — the existing `ContactsFolder`, `EnableContactSaving`, `SearchEntireVaultForContacts`, and `ContactFilenameFormat` settings are sufficient. One new setting is needed: `ContactNoteIncludeDetails` (bool, default true) to control whether the rich details section is included.

8. **Ribbon additions** — Add a "Contacts" group to the Sling tab with buttons for "Sling Contact" (single selected contact) and "Sling All Contacts" (full address book). The existing "Sling" button already handles contacts-from-email via the `ProcessSelection()` flow detection.

---

## Data Flow

### Appointment Contact Linking

```
AppointmentProcessor.ProcessAppointment()
  │
  ├─ [existing] Extract attendee names (organizer, required, optional)
  ├─ [existing] Build frontmatter with [[Name]] wiki-links
  ├─ [existing] Write note, attachments, tasks
  │
  └─ [NEW] Post-export contact resolution
       ├─ Collect: organizerName + requiredAttendees + optionalAttendees
       ├─ Deduplicate, remove resources/rooms
       ├─ For each name:
       │    ├─ ManagedContactNoteExists? → refresh via CreateContactNote()
       │    ├─ ContactExists? → skip (user-managed note exists elsewhere)
       │    └─ Neither → add to newContacts list
       ├─ If !bulkMode && newContacts.Count > 0:
       │    └─ Show ContactConfirmationDialog → create selected
       └─ Done
```

### Contact Slinging (Single)

```
User selects ContactItem in Outlook → clicks "Sling" or "Sling Contact"
  │
  └─ ThisAddIn.ProcessSelection() or ProcessSelectedContact()
       │
       └─ ContactProcessor.ProcessContact(ContactItem)
            ├─ Extract: FullName, Email, Phone, Company, Title, etc.
            ├─ Check ContactService.ContactExists() → warn if duplicate
            ├─ Build extended ContactTemplateContext
            ├─ Render via TemplateService
            ├─ Write via FileService
            └─ Launch Obsidian (if enabled)
```

### Contact Slinging (Address Book)

```
User clicks "Sling All Contacts" ribbon button
  │
  └─ ThisAddIn.SlingAllContacts()
       │
       └─ ContactProcessor.ProcessAddressBook()
            ├─ Get default Contacts folder from Outlook
            ├─ Show StatusService progress
            ├─ For each ContactItem:
            │    ├─ Check exists → skip or refresh
            │    └─ Create contact note with rich details
            ├─ Show summary (saved/skipped/errors)
            └─ Launch Obsidian (if enabled, to contacts folder)
```

---

## Components and Responsibilities

### New Components

| Component | File | Responsibility |
|---|---|---|
| `ContactProcessor` | `Services/ContactProcessor.cs` | Orchestrates single-contact, from-email, and address-book contact export flows |

### Modified Components

| Component | File | Changes |
|---|---|---|
| `AppointmentProcessor` | `Services/AppointmentProcessor.cs` | Add post-export contact resolution (attendee lookup + creation offer) |
| `ContactService` | `Services/ContactService.cs` | Add `CreateRichContactNote(ContactItem)` method for ContactItem-sourced notes with full details; add `ExtractContactData(ContactItem)` helper |
| `ObsidianSettings` | `Models/ObsidianSettings.cs` | Add `ContactNoteIncludeDetails` (bool) setting |
| `TemplateService` | `Services/TemplateService.cs` | Extend `ContactTemplateContext` with rich fields (Phone, Email, Company, Title, Address, Birthday, Notes); add `RenderRichContactContent()` or extend existing render method |
| `ThisAddIn` | `ThisAddIn.cs` | Add `ProcessSelectedContact()`, `SlingAllContacts()` methods; update `ProcessSelection()` to detect `ContactItem` selection |
| `SlingRibbon` | `Ribbon/SlingRibbon.cs` + `.xml` | Add "Contacts" group with "Sling Contact" and "Sling All Contacts" buttons |
| `SettingsForm` | `Forms/SettingsForm.cs` | Add `ContactNoteIncludeDetails` checkbox to the Contacts tab |

### Unchanged Components

| Component | Reason |
|---|---|
| `EmailProcessor` | Explicitly protected — no modifications |
| `FileService` | Already has all needed methods |
| `ThreadService` | Not relevant to contact slinging |
| `TaskService` | Not relevant to contact slinging |
| `AttachmentService` | Not relevant to contact slinging |
| `StatusService` | Already has all needed methods |

---

## Error Handling

- **COM interop failures:** All `ContactItem` property reads wrapped in individual try/catch with fallback to empty string, matching the pattern in `AppointmentProcessor` (lines 184-203).
- **Contact creation failures:** Caught and logged per-contact; processing continues to next contact. Non-fatal for the overall operation.
- **Address book access failures:** If the default Contacts folder is inaccessible, show a single error message and abort the bulk operation.
- **Duplicate detection:** Uses existing `ContactService.ContactExists()` and `ManagedContactNoteExists()`. No new duplicate detection logic needed.
- **Bulk mode (address book):** Errors are collected and displayed in a summary dialog at the end, matching the pattern in `ThisAddIn.SaveTodaysAppointments()`.

---

## Trade-offs

| Decision | Pro | Con |
|---|---|---|
| Duplicate contact resolution logic in AppointmentProcessor instead of extracting shared method | Zero risk to EmailProcessor | ~30 lines of duplicated code |
| No interfaces for services | Consistent with existing codebase; simpler | Harder to mock in tests |
| Sequential address book processing | Simpler COM interop handling; no threading issues | Slower for large address books (acceptable — COM is inherently single-threaded) |
| Single new setting (`ContactNoteIncludeDetails`) | Minimal settings bloat | Users cannot individually toggle which contact fields to include (can be added later) |
| Extending `ContactTemplateContext` rather than creating new context class | Reuses existing template infrastructure | Context class grows larger |
| Not extracting `SubjectCleanupService` in this run | Reduced scope and risk | Duplicated regex patterns remain in both processors |

---

## Open Questions

1. **Contact photo export:** Should contact photos from Outlook be saved as attachments and embedded in the contact note? Recommendation: defer to a future run — adds complexity with `AttachmentService` integration and image file management.

2. **Address book scope:** Should "Sling All Contacts" export from all Outlook accounts or just the default account? Recommendation: default account only for v1, with a future option to select accounts.

3. **Contact update vs. create:** When a contact note already exists and the user re-slings the same ContactItem, should it update/merge the rich details section? Recommendation: yes, using the same managed-section merge pattern that `ContactService.CreateContactNote()` already implements for Communication History.

4. **From-email contact export:** Should the "export contacts from email" flow use `ContactItem` lookup (resolving the recipient to an Outlook contact for rich data) or just create basic contact notes from the name/email? Recommendation: attempt `ContactItem` resolution via `Recipient.AddressEntry.GetContact()`, fall back to basic name/email note if no Outlook contact exists.

---

## Implementation Phases

### Phase 1: Appointment Contact Linking
- Modify `AppointmentProcessor.ProcessAppointment()` to add post-export contact resolution
- Add tests for the new flow
- Estimated: ~100 lines of new code in AppointmentProcessor

### Phase 2: Contact Slinging Infrastructure
- Create `ContactProcessor` with single-contact export
- Extend `ContactService` with rich contact data extraction
- Extend `ContactTemplateContext` and `TemplateService`
- Add `ContactNoteIncludeDetails` setting
- Update `ThisAddIn.ProcessSelection()` to handle `ContactItem`
- Add tests

### Phase 3: Contact Slinging UI + Address Book
- Add ribbon buttons for contact operations
- Implement `ProcessSelectedContact()` and `SlingAllContacts()` in ThisAddIn
- Implement address book bulk export in `ContactProcessor`
- Update SettingsForm with new setting
- Add tests
