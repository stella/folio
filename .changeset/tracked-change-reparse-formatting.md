---
"@stll/folio-core": patch
---

Text next to a tracked change no longer picks up underline, strikethrough or a text color when a native edit makes the editor read its own DOM back. Those were how the review mark is drawn, not formatting of the text, and they stayed behind once the change was rejected.
