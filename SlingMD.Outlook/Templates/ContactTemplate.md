{{frontmatter}}
# {{contactName}}

## Communication History

```dataviewjs
// Find all emails where this contact appears in from, to, or cc fields
// Use title from frontmatter (original name) rather than file.name (cleaned name)
const contact = dv.current().title || dv.current().file.name;

// Helper to check if a field contains this contact
// Handles both Dataview Link objects and plain strings
function containsContact(field, contactName) {
    if (!field) return false;
    // Handle Dataview Link objects (have .path property)
    if (field.path) return field.path === contactName;
    // Handle string format - check for [[Name]] or exact match
    const str = String(field);
    return str.includes(`[[${contactName}]]`) || str === contactName;
}

// Helper to check arrays (to/cc fields can be arrays)
function checkArray(arr, contactName) {
    if (!arr) return false;
    if (!Array.isArray(arr)) return containsContact(arr, contactName);
    return arr.some(item => containsContact(item, contactName));
}

// Query all pages, then filter to only emails (pages with fromEmail field)
// and where this contact is mentioned in from, to, or cc
const emails = dv.pages()
    .where(p => {
        // Only include pages that are emails (have fromEmail field)
        if (!p.fromEmail) return false;
        // Check if this contact is mentioned in from, to, or cc
        return containsContact(p.from, contact) ||
               checkArray(p.to, contact) ||
               checkArray(p.cc, contact);
    })
    .sort(p => p.date, 'desc');

dv.table(["Date", "Subject", "Type"],
    emails.map(p => {
        // Determine message type
        const isFrom = containsContact(p.from, contact);
        const isTo = checkArray(p.to, contact);
        const type = isFrom ? "From" : isTo ? "To" : "CC";
        return [p.date, p.file.link, type];
    })
);
```

## Notes
