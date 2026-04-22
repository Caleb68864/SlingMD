# SlingMD

Tools for Use with ObsidianMD - Seamlessly integrate your Outlook emails with Obsidian notes.

![SlingMD Logo](SlingMD_pixel.png)

## Overview

SlingMD is a powerful Outlook add-in that bridges the gap between your Outlook emails, calendar appointments, and contacts with Obsidian notes. It allows you to easily export and manage your emails, appointments, and contact information within your Obsidian knowledge base, helping you maintain a comprehensive personal knowledge management system.

## Features

- Export Outlook emails directly to Obsidian markdown format
- **Export Outlook calendar appointments to Obsidian markdown** with full attendee, location, and recurrence metadata
- Preserve email metadata and formatting
- Create follow-up tasks in Obsidian notes and/or Outlook
- **Appointment task creation** (None / Obsidian / Outlook / Both)
- Seamless integration with Outlook's interface
- Easy-to-use ribbon interface with **appointment inspector support**
- Support for attachments and email threading
- Automatic email thread organization
- **Recurring meeting threading** — group recurring meeting instances into thread folders with summary notes
- Thread summary pages with timeline views
- **Companion meeting notes** — optional blank note for real-time meeting capture, linked to the appointment note
- **Bulk "Save Today's Appointments"** — one-click export of all today's calendar items
- Automatic contact note creation with communication-history Dataview tables
- Customisable note title formatting (placeholders for {Subject}, {Sender}, {Date}) with max-length trimming
- Advanced subject clean-up engine using user-defined regex patterns
- Configurable default tags for notes and tasks
- Relative vs. absolute reminder modes with optional per-task prompt
- Development/debug mode to surface internal thread-matching diagnostics
- Duplicate-email protection and safe file-naming, including chronological prefixes for threads
- **Contact slinging** — export single contacts or your entire address book to Obsidian with rich detail notes (phone, email, company, etc.)
- First-class markdown templates for email notes, contact notes, inline task lines, thread notes, **appointment notes, and meeting notes**
- **Tabbed settings dialog** organized into 8 focused tabs (General, Email, Appointments, Contacts, Tasks, Threading, Attachments, Developer)

## Installation

1. Go to the [Releases](./Releases) folder in this repository
2. Download the newest versioned ZIP from the [Releases](./Releases) folder (for example, `SlingMD.Outlook_1_1_0_7.zip` at the time of writing)
3. **Important Security Step - Unblock the ZIP File**:
   - Right-click the downloaded ZIP file
   - Click "Properties"
   - At the bottom of the General tab, check the "Unblock" box
   - Click "Apply"
   - If you've already extracted the ZIP, delete the extracted folder first
   - Extract the ZIP file again after unblocking
   
   This step ensures all extracted files are trusted and in the same security zone, preventing potential issues.

4. Run the setup executable to install the Outlook add-in
5. Restart Outlook after installation
6. Enable the Sling Ribbon:
   - In Outlook, click "File" > "Options" > "Customize Ribbon"
   - In the left-hand dropdown menu, select "All Tabs"
   - Find "Sling Tab" in the left column
   - Click "Add >>" to add it to your ribbon
   - Click "OK" to save changes
7. The SlingMD ribbon will now appear in your Outlook interface

## System Requirements

- Microsoft Outlook (Office 365 or 2019+)
- Windows 10 or later
- Obsidian installed on your system

## Development

### Prerequisites for building and testing

Building, running tests, and publishing the VSTO add-in require **Visual Studio** (2019 or later) with the **Office/VSTO development workload** installed. This workload provides `Microsoft.VisualStudio.Tools.Office.targets`, which both `SlingMD.Outlook.csproj` and `SlingMD.Tests.csproj` import at build time.

**If you run `dotnet build` or `dotnet test` from the command line without Visual Studio tooling installed, the build will fail with:**
```
error MSB4019: The imported project "...\Microsoft.VisualStudio.Tools.Office.targets" was not found.
```
This is expected. To resolve it, open the solution in Visual Studio with the Office developer tools workload enabled, or install that workload via the Visual Studio Installer under *Workloads > Office/SharePoint development*.

Editing the codebase (reading, searching, and modifying files) can happen in any editor such as VS Code, but a full build or test run requires the Visual Studio tooling described above.

## Usage

1. Open Microsoft Outlook
2. Select any email you want to save to Obsidian
3. In the Outlook ribbon menu, locate and click the "Sling" button from the Sling Ribbon
4. The email will be converted to Markdown format and saved to your configured Obsidian vault
5. If enabled, follow-up tasks will be created in Obsidian and/or Outlook

## Configuration

Before using SlingMD, you'll need to configure your Obsidian vault settings:

1. Click the "Settings" button in the Sling Ribbon
2. Configure the following options:
   - **Vault Name**: Enter the name of your Obsidian vault
   - **Vault Base Path**: Set the path to your Obsidian vault folder (e.g., C:\Users\YourName\Documents\Notes)
   - **Inbox Folder**: Specify the folder within your vault where emails should be saved (default: "Inbox")
   - **Launch Obsidian**: Toggle whether Obsidian should automatically open after saving an email
   - **Delay (seconds)**: Set how long to wait before launching Obsidian (default: 1 second)
   - **Show countdown**: Toggle whether to show a countdown before launching Obsidian
   - **Create task in Obsidian note**: Toggle whether to create a follow-up task in the Obsidian note
   - **Create task in Outlook**: Toggle whether to create a follow-up task in Outlook
   - **Due in Days**: Set the default number of days until tasks are due (0 = today, 1 = tomorrow, etc.)
   - **Use Relative Dates**: Choose between relative or absolute date mode:
     - When enabled (relative): Reminder days are calculated backwards from the due date
     - When disabled (absolute): Reminder date is calculated forward from today
   - **Reminder Days**: Set the default reminder timing:
     - In relative mode: How many days before the due date
     - In absolute mode: How many days from today
   - **Reminder Hour**: Set the default hour for task reminders (24-hour format)
   - **Ask for dates**: Toggle whether to prompt for dates each time (shows the Task Options form)
   - **Group Email Threads**: Toggle whether to automatically organize related emails into thread folders
   - **Subject Cleanup Patterns**: Configure patterns for cleaning up email subjects (e.g., removing "Re:", "[EXTERNAL]", etc.)
   - **Contacts Folder**: Where new contact notes will be stored (default: "Contacts")
   - **Enable Contact Saving**: Toggle automatic creation of contact notes
   - **Search Entire Vault For Contacts**: When enabled, SlingMD will look outside the contacts folder before creating a new contact note
   - **Note Title Format / Max Length / Include Date**: Fine-tune how note titles are constructed
   - **Move Date To Front In Thread**: When grouping emails, place the date at the beginning of the filename
   - **Default Note Tags / Task Tags**: Tags automatically assigned to new notes or tasks
   - **Templates Folder / Template Files**: Point SlingMD at Dataview-friendly templates for email notes, contact notes, inline task lines, and thread summaries (task templates customize the inline task line only)
   - **Email / Contact Filename Format**: Optional filename formats using tokens like `{Subject}`, `{Sender}`, `{Timestamp}`, `{ContactName}`, and `{ContactShortName}`
   - **Show Development Settings**: Reveals additional debug options in the settings dialog
   - **Show Thread Debug**: Pops up a diagnostic window listing every file that matches a conversationId

3. Click "Save" to apply your settings

Note: Make sure your Vault Base Path points to an existing Obsidian vault directory. If you haven't created a vault yet, please set one up in Obsidian first.

### Customization

SlingMD exposes several format strings so you can tailor the output to your vault or mention-plugin setup.

#### `EmailDateFormat`, `ContactDateFormat`, `AppointmentDateFormat`

Controls how `{{date}}` (and related date placeholders) render in exported notes. Uses the standard .NET `DateTime` format syntax.

- Default: `"yyyy-MM-dd HH:mm:ss"` — e.g., `2026-04-21 09:30:00`

Each domain has its own setting so you can keep email timestamps precise while rendering appointments as date-only.

**Examples (non-default):**

```
EmailDateFormat = "yyyy-MM-dd"
```
Produces: `2026-04-21`

```
ContactDateFormat = "MMMM d, yyyy"
```
Produces: `April 21, 2026`

```
AppointmentDateFormat = "dddd, MMMM d, yyyy h:mm tt"
```
Produces: `Tuesday, April 21, 2026 9:30 AM`

#### `ContactLinkFormat`

Controls how `{{to}}`, `{{from}}`, and `{{cc}}` render recipients in email and appointment notes. Supports tokens: `{FullName}`, `{FirstName}`, `{LastName}`, `{MiddleName}`, `{Suffix}`, `{DisplayName}`, `{ShortName}`, `{Email}`, `{FirstInitial}`, `{LastInitial}`.

- Default: `"[[{FullName}]]"` — produces `[[John Smith]]`

**Examples (non-default):**

```
ContactLinkFormat = "@{FirstName}{LastName}"
```
Produces: `@JohnSmith` — compatible with the At People plugin for inline mentions.

```
ContactLinkFormat = "[[{LastName}]]"
```
Produces: `[[Smith]]`

```
ContactLinkFormat = "@{FirstInitial}{LastInitial}"
```
Produces: `@JS`

#### Subject cleanup

The default subject-cleanup patterns now correctly preserve words that contain `re-` (like `pre-release`). If you upgraded from a previous version, SlingMD silently migrates the broken default pattern on first load — your custom patterns are never touched.

## Task Creation

When task creation is enabled, SlingMD can create follow-up tasks in two locations:

### Obsidian Tasks
- Created at the top of the note
- Uses Obsidian's task format with date metadata
- Links back to the email note
- Tagged with #FollowUp for easy tracking

### Outlook Tasks
- Creates a task in your Outlook task list
- Includes the email subject and content
- Configurable due date (0-30 days from creation)
- Configurable reminder time (0-23 hour)
- Option to prompt for due date and reminder time for each task

## Task Options

When "Ask for dates" is enabled, the Task Options form will appear when creating tasks. This form allows you to:

1. Set the due date using either:
   - Relative mode: Specify number of days from today
   - Absolute mode: Pick a specific date from a calendar

2. Set the reminder using either:
   - Relative mode: Specify number of days before the due date
   - Absolute mode: Pick a specific date from a calendar

3. Set the reminder hour (in 24-hour format)

The "Use Relative Dates" toggle switches between:
- Relative mode: Reminder is set relative to the due date (e.g., "remind me 2 days before it's due")
- Absolute mode: Reminder is set to a specific date (e.g., "remind me next Tuesday")

## Email Threading

When email threading is enabled (via the "Group Email Threads" setting), SlingMD will:

1. Automatically detect related emails using conversation topics and thread IDs
2. Create a dedicated folder for each email thread
3. Generate a thread summary note (0-threadname.md) containing:
   - Thread start and end dates
   - Number of messages
   - List of participants
   - Timeline view of all emails in the thread
4. Move all related emails into the thread folder
5. Update thread summary when new emails are added
6. Link emails to their thread summary for easy navigation

This organization helps keep related emails together and provides a clear overview of email conversations in your vault.

## Contact Slinging

SlingMD can export your Outlook contacts directly to Obsidian as rich markdown notes:

### Single Contact
1. Select a contact in Outlook
2. Click "Sling Contact" in the Contacts group on the Sling ribbon
3. The contact is exported with all available details (phone, email, company, address, etc.)

### All Contacts
1. Click "Sling All Contacts" in the Contacts group on the Sling ribbon
2. SlingMD exports every contact from your default address book
3. Existing contact notes are skipped (no duplicates)
4. A summary dialog shows how many contacts were saved, skipped, or had errors

### Appointment Contact Linking
When you sling an appointment, SlingMD automatically checks the attendees against your vault. Existing contact notes are refreshed, and you're offered the option to create notes for new contacts (same behavior as email slinging).

### Contact Note Include Details
In Settings > Contacts, the "Include contact details" checkbox controls whether slung contact notes get a `## Contact Details` section with phone, email, company, title, address, and birthday. When disabled, contacts get the simpler template (name + communication history only).

## Custom Templates

SlingMD uses `{{placeholder}}` templates to generate markdown notes. You can customize the layout and content of every note type by creating your own template files.

### How It Works

1. Create a `.md` file in your vault's Templates folder (configured in Settings > General)
2. Use `{{placeholders}}` for dynamic content — SlingMD replaces them at export time
3. Point the matching template setting (e.g., "Contact Template File") to your file name
4. Your template is used for all notes of that type going forward

If no custom template is configured, SlingMD uses sensible built-in defaults.

> **Note:** If you drop a `ContactTemplate.md` into your vault's Templates folder, SlingMD will load and use it automatically. (Earlier versions only honored the file when the template filename was changed to a non-default name — this is now fixed.)

### Template Types and Settings

| Note Type | Setting Name | Default File | Description |
|-----------|-------------|--------------|-------------|
| Email | Email Template File | `EmailTemplate.md` | Exported email notes |
| Contact | Contact Template File | `ContactTemplate.md` | Contact notes (both email-created and slung) |
| Task | Task Template File | `TaskTemplate.md` | Inline task lines in notes |
| Thread | Thread Template File | `ThreadNoteTemplate.md` | Thread summary notes |
| Appointment | Appointment Template File | `AppointmentTemplate.md` | Exported appointment notes |
| Meeting Note | Meeting Note Template File | `MeetingNoteTemplate.md` | Companion meeting notes |

### Available Placeholders

Every template type supports `{{frontmatter}}` (generates the YAML frontmatter block) plus type-specific fields:

**Email Templates:**
`{{subject}}`, `{{from}}`, `{{to}}`, `{{cc}}`, `{{date}}`, `{{body}}`, `{{attachments}}`, `{{fileName}}`, `{{fileNameNoExt}}`, `{{threadNote}}`, `{{threadId}}`

**Contact Templates:**
`{{contactName}}`, `{{contactShortName}}`, `{{created}}`, `{{fileName}}`, `{{fileNameNoExt}}`, `{{phone}}`, `{{email}}`, `{{company}}`, `{{jobTitle}}`, `{{address}}`, `{{birthday}}`, `{{notes}}`

All 12 contact placeholders are available regardless of how the contact was created. When a contact comes from email (name only), the detail fields (`{{phone}}`, `{{email}}`, etc.) render as empty strings. When a contact is slung from Outlook, all available fields are populated.

**Appointment Templates:**
`{{subject}}`, `{{organizer}}`, `{{start}}`, `{{end}}`, `{{location}}`, `{{body}}`, `{{requiredAttendees}}`, `{{optionalAttendees}}`, `{{resourceAttendees}}`, `{{categories}}`, `{{recurrenceInfo}}`, `{{fileName}}`, `{{fileNameNoExt}}`

**Task Templates:**
`{{taskText}}`, `{{dueDate}}`, `{{scheduledDate}}`, `{{tags}}`

### Example Custom Contact Template

Create a file called `MyContactTemplate.md` in your vault's Templates folder:

```markdown
{{frontmatter}}
# {{contactName}}

> **{{jobTitle}}** at **{{company}}**

## Details
| Field | Value |
|-------|-------|
| Phone | {{phone}} |
| Email | {{email}} |
| Address | {{address}} |
| Birthday | {{birthday}} |

## Communication History

(Add your Dataview query here)

## Notes

{{notes}}
```

Then set "Contact Template File" to `MyContactTemplate.md` in Settings > Contacts.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the terms included in the [LICENSE](LICENSE) file.

## Support

If you encounter any issues or have questions, please open an issue in the GitHub repository.

## Changelog

### Version 1.1.0.7
- **Contact Slinging** — export single contacts or your entire address book to Obsidian with rich detail notes (phone, email, company, address, birthday)
- **Appointment Contact Linking** — after exporting an appointment, attendees are checked against vault contacts; existing managed notes are refreshed, new contacts offered via dialog
- **Unified Sling Button** — main Sling button now handles emails, appointments, and contacts; dispatches based on selected item type
- **Obsidian Launch on Contact Export** — single-contact export now launches Obsidian for immediate feedback, matching email behavior
- **Unified Contact Template System** — one `RenderContactContent()` path for all contacts; custom templates get all 13 placeholders regardless of source (email or slung)
- **ContactNoteIncludeDetails Setting** — toggle `## Contact Details` section on/off for slung contacts in Settings > Contacts
- **"Sling All Contacts" Ribbon Button** — bulk-export entire address book with progress tracking and summary dialog
- **Contact Confirmation Dialog Clarity** — Cancel button renamed to "Skip Creating Contacts" for clearer intent
- **Custom Templates Documentation** — README now documents all template types, available placeholders, and includes example custom contact template
- New `ContactProcessor` orchestrator, `ExtractContactData()`, `CreateContactNote(ContactTemplateContext)` overload
- 12 new tests (87/87 total passing)

### Version 1.1.0.1
- **Appointment Processing** — full pipeline for exporting Outlook calendar appointments to Obsidian markdown, with attendee lists, location, recurrence metadata, and configurable templates
- **Appointment Ribbon Integration** — "Save Today's Appointments" bulk-export button and appointment inspector "Sling" button
- **Recurring Meeting Threading** — recurring instances grouped into thread folders with summary notes
- **Companion Meeting Notes** — optional blank meeting note linked bidirectionally to the appointment note
- **Appointment Task Creation** — configurable task creation for appointments (None / Obsidian / Outlook / Both)
- **Tabbed Settings Form** — settings reorganized into 8 focused tabs (General, Email, Appointments, Contacts, Tasks, Threading, Attachments, Developer)
- **Settings Form UI Fix** — last row on all tabs no longer stretches; form now displays the SlingMD icon
- 10 new appointment-related settings and dedicated appointment/meeting note templates
- 13 new tests (75/75 total passing)

### Version 1.0.0.124
- **Reliability Hardening** — corrupt or malformed settings files no longer crash the add-in on startup; safe defaults are loaded instead ([#6](https://github.com/Caleb68864/SlingMD/issues/6))
- **Settings Persistence Fix** — settings now persist correctly across Outlook restarts ([#6](https://github.com/Caleb68864/SlingMD/issues/6))
- **Export Flow Safeguards** — a fatal export error no longer continues into contact creation or Obsidian launch
- **Task State Reset** — canceling the task-options dialog no longer disables task creation for the rest of the Outlook session
- **Thread Summary Fix** — the DataviewJS query in thread summary notes (0-file) now correctly scopes to the thread folder instead of producing a JavaScript syntax error
- **Thread Date Alignment** — thread discovery now parses both second-precision and legacy minute-precision date formats for backward compatibility
- **YAML-Safe Frontmatter** — email metadata containing double quotes, backslashes, or newlines is now properly escaped in frontmatter
- **Attachment Link Hardening** — attachment links now resolve correctly for same-folder, per-note-subfolder, and centralized storage modes
- **Missing Vault Handling** — exporting into a fresh vault with no inbox folder no longer throws during duplicate detection
- **Customizable Templates** — email, contact, task, and thread templates are fully configurable via the Templates settings ([#8](https://github.com/Caleb68864/SlingMD/issues/8))
- **Contact Note Dataview Fix** — contact notes now use `## Communication History` with a working DataviewJS query for email history display ([#4](https://github.com/Caleb68864/SlingMD/issues/4))
- **VSTO Build Documentation** — README now documents the Visual Studio Office/VSTO prerequisite for building and testing
- Added regression test coverage for all hardened behaviors

### Version 1.0.0.44
- Added automatic email thread detection and organization
- Added thread summary pages with timeline views
- Added configurable subject cleanup patterns
- Added thread folder creation for related emails
- Added participant tracking in thread summaries
- Added dataview integration for thread visualization
- Improved email relationship detection
- Enhanced thread navigation with bidirectional links
- Fixed various bugs and improved stability

### Version 1.0.0.14
- Added ability to create follow-up tasks in Obsidian notes
- Added ability to create follow-up tasks in Outlook
- Added configurable due dates for tasks
- Added configurable reminder times for tasks
- Added option to prompt for due date and reminder time
- Added task options dialog for custom timing
- Updated settings interface with task configuration options

### Version 1.0.0.8
- Initial release
- Basic email to Obsidian note conversion
- Email metadata preservation
- Obsidian vault configuration
- Launch delay settings

## Disclaimer

THIS SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

The author is not responsible for any data loss, corruption, or other issues that may occur while using this software. Always ensure you have proper backups of your data before using any software that modifies your files.

---

☕ Like what I'm building? Help fuel my next project (or my next coffee)!  
Support me on [Buy Me a Coffee](https://buymeacoffee.com/plainsprepper) 💻🧵🔥


