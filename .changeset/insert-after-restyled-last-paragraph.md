---
"@stll/docx-core": patch
---

Let a tracked insertion follow a last paragraph whose properties carry a pending change: the change goes with the mark it was recorded on to the new last paragraph, so Reject All still restores the original properties.
