---
"@stll/docx-core": patch
---

Concurrent typing in one paragraph is sequenced by position. An `insertContent` of a closed inline slice moves later offsets like text, at one gap the later-sequenced insertion goes after the earlier one, and ranges that start or end at the insertion point keep out of the inserted content; a format or direct deletion around it is split. `splitInline` and `joinInline` keep offsets. Sequenced `insertContent` records whether it cut an identified record (`cutIdentified`). A comment anchor moves over typing in its paragraph. Two replies to one parent now commute: a reply's markers take their place after the thread's earlier comments, siblings by id.
