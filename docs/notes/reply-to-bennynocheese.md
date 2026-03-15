# Reply to bennynocheese

**Context:** Reddit post requesting calendar item export to Obsidian with attendees, categories, etc.

**Reply:**

Good news -- this is fully built and shipping now in v1.1.0.7!

When you sling a calendar item, it exports everything: subject, organizer, required/optional/resource attendees, location, categories, recurrence info, and the full body -- all as markdown with YAML frontmatter. Attendees show up as `[[wiki-links]]` so they connect to your contact notes automatically.

Here's what you get out of the box:

- **Single appointment export** -- select any appointment and hit the Sling button in the ribbon
- **Bulk "Save Today" button** -- one click exports all of today's calendar items
- **Inspector support** -- Sling button right inside an open appointment window
- **Recurring meeting threading** -- recurring instances get grouped into thread folders with a summary note
- **Companion meeting notes** -- optionally creates a blank linked note for real-time meeting capture
- **Appointment task creation** -- create follow-up tasks in Obsidian and/or Outlook
- **Contact linking** -- after export, attendees are checked against your vault and you're offered the option to create contact notes for anyone new

Templates are fully customizable too -- you can create your own markdown template with `{{placeholders}}` for any field (subject, organizer, start, end, location, attendees, categories, etc.).

Grab the latest release from the [Releases folder](https://github.com/Caleb68864/SlingMD/tree/main/Releases) and let me know how it works for you!
