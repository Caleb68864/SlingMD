# Reply to bennynocheese

**Context:** Reddit post requesting calendar item export to Obsidian with attendees, categories, etc.

**Reply:**

Hey! Good news -- just shipped this in v1.1.0.7!

You can sling calendar items now and it pulls everything -- attendees, location, categories, recurrence info, the whole thing. Attendees come through as `[[wiki-links]]` so they hook right into your contact notes.

A few things you might like:

- Hit the Sling button on any appointment and it just exports
- There's a "Save Today" button that bulk exports all of today's calendar items in one click
- Works from inside an open appointment too
- Recurring meetings get grouped into thread folders with a summary note
- You can optionally get a blank companion note linked to the appointment for live meeting notes
- After export it checks the attendees against your vault and offers to create contact notes for anyone new

Templates are customizable if you want to tweak the layout -- just drop a markdown file with `{{placeholders}}` in your templates folder.

Grab it from the [Releases folder](https://github.com/Caleb68864/SlingMD/tree/main/Releases) and let me know how it goes!
