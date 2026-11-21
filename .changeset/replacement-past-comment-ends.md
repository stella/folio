---
"@stll/docx-core": patch
---

Text replacing every word a comment covers goes past the comment's end, and past other ranges' ends standing before it, while staying in the link, content control or wrapper the words were in. A direct replacement takes the comment, and a link left holding only its markers goes with them; a tracked one keeps the comment on the struck words, and accepting it gives the direct edit.
