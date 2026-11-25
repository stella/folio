---
"@stll/docx-core": patch
---

A replacement removes an empty hyperlink, content control or wrapper exactly when the part of its range between two range boundaries or comment anchors removes something else, in both editing and suggesting. Content already deleted or moved away no longer splits that part in a direct replacement, so it removes the same containers as accepting the suggestion does.
