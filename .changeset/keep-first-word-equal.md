---
"@stll/folio-core": patch
---

Word diffs match words on their text alone, so text inserted or deleted before a
paragraph's first word no longer strikes that word through; a whitespace change
around an unchanged word is marked as a whitespace change only.
