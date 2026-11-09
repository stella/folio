---
"@stll/docx-core": patch
---

Replacing a range directly whose deletion keeps a bookmark or another range boundary inside a wrapper or inline content control now puts the new text in that wrapper, after the boundary, instead of splitting the wrapper in two. The direct edit now gives what accepting the same edit made with tracked changes gives.
