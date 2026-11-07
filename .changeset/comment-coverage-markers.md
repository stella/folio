---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

A comment whose covered text is all deleted is now removed on save also when a bookmark or another range boundary sat inside its range. Such a boundary marks a position, not content, so it no longer keeps an empty comment alive. A source comment whose range holds nothing but such a boundary still opens as a point comment and is kept.
