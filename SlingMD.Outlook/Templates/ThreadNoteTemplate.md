---
title: "{{title}}"
type: email-thread
threadId: "{{threadId}}"
tags: [email-thread]
---

# {{title}}

```dataviewjs
// Get all emails with matching threadId from current folder
const threadId = "{{threadId}}";
const emails = dv.pages("")
    .where(p => p.threadId === threadId && p.file.name !== dv.current().file.name)
    .sort(p => p.date, 'desc');

// Display thread summary
if (emails.length > 0) {
    const startDate = emails[emails.length-1].date;
    const latestDate = emails[0].date;
    const participants = new Set();
    emails.forEach(e => {
        // Handle from field
        if (e.from) {
            const fromName = String(e.from).match(/\[\[(.*?)\]\]/)?.[1];
            if (fromName) participants.add(fromName);
        }

        // Handle to field
        if (e.to) {
            const toList = Array.isArray(e.to) ? e.to : [e.to];
            toList.forEach(to => {
                const name = String(to).match(/\[\[(.*?)\]\]/)?.[1];
                if (name) participants.add(name);
            });
        }

        // Handle cc field
        if (e.cc) {
            const ccList = Array.isArray(e.cc) ? e.cc : [e.cc];
            ccList.forEach(cc => {
                const name = String(cc).match(/\[\[(.*?)\]\]/)?.[1];
                if (name) participants.add(name);
            });
        }
    });

    dv.header(2, 'Thread Summary');
    dv.list([
        `Started: ${startDate}`,
        `Latest: ${latestDate}`,
        `Messages: ${emails.length}`,
        `Participants: ${Array.from(participants).map(p => `[[${p}]]`).join(', ')}`
    ]);
}

// Display email timeline
dv.header(2, 'Email Timeline');
for (const email of emails) {
    dv.header(3, `${email.file.name} - ${email.date}`);
    dv.paragraph(`![[${email.file.name}]]`);
}
```