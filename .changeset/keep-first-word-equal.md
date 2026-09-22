---
"@stll/folio-core": patch
---

Word diffs match words on their text alone and treat punctuation at a word's
edges as its own unit. Text inserted before a paragraph's first word no longer
strikes that word through, `jmění.` becoming `jmění,` marks only the mark, and a
whitespace change around an unchanged word marks only the whitespace.
Punctuation inside a word (`d.o.o`, `1.1.2026`, `3.5`) stays part of it.
