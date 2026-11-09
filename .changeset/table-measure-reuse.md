---
"@stll/folio-core": patch
---

Typing no longer measures every table again on each key: a table the edit did not reach keeps its measure, whether the next layout pass hands it out again or moves it. The measure is taken again when the width, the styles or the fonts it was measured with change.
