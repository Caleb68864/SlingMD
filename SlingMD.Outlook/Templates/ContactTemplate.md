{{frontmatter}}
# {{contactName}}

## Contact Details

**Phone:** {{phone}}
**Email:** {{email}}
**Company:** {{company}}
**Title:** {{jobTitle}}
**Address:** {{address}}
**Birthday:** {{birthday}}

## Communication History

```dataviewjs
const current = dv.current();
const contactSources = [current.title, current.file?.name, current.file?.path];

function normalizeSingle(value) {
    if (!value && value !== 0) return [];

    if (typeof value === "object") {
        const candidates = [];
        if (value.path) candidates.push(value.path);
        if (value.display) candidates.push(value.display);
        if (value.file?.path) candidates.push(value.file.path);
        return candidates.flatMap(normalizeSingle);
    }

    const text = String(value).trim();
    if (!text) return [];

    const unwrapped = text
        .replace(/^\[\[/, "")
        .replace(/\]\]$/, "")
        .split("|")[0]
        .trim();

    if (!unwrapped) return [];

    const withoutExtension = unwrapped.replace(/\.md$/i, "");
    const pathParts = withoutExtension.split("/");
    const fileName = pathParts[pathParts.length - 1];

    return [...new Set([unwrapped, withoutExtension, fileName]
        .map(item => item.trim().toLowerCase())
        .filter(Boolean))];
}

function normalizeValue(value) {
    if (!value && value !== 0) return [];
    if (Array.isArray(value)) return value.flatMap(normalizeSingle);
    return normalizeSingle(value);
}

const contactKeys = new Set(normalizeValue(contactSources));

function containsContact(field) {
    return normalizeValue(field).some(value => contactKeys.has(value));
}

function isEmailPage(page) {
    const types = normalizeValue(page.type);
    return types.includes('email') || !!page.fromEmail || !!page.internetMessageId || !!page.entryId;
}

const emails = dv.pages()
    .where(page => isEmailPage(page) && (
        containsContact(page.from) ||
        containsContact(page.to) ||
        containsContact(page.cc)
    ))
    .sort(page => page.date, 'desc');

dv.table(["Date", "Subject", "Type"],
    emails.map(page => {
        const role = containsContact(page.from)
            ? "From"
            : containsContact(page.to)
                ? "To"
                : "CC";

        return [page.date, page.file.link, role];
    })
);
```

## Notes

{{notes}}