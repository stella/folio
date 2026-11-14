---
"@stll/docx-core": patch
---

A direct paragraph join beside a paragraph mark pending deletion now gives the joined paragraph the properties the editing view shows for that join, whichever paragraph it keeps. Joining a paragraph whose mark is already deleted or moved away with the paragraph after it, as a direct deleteBlock of a container's last paragraph does, now applies and leaves that paragraph's words and properties unchanged.
