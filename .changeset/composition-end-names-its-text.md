---
"@stll/folio-core": patch
---

In the canonical session, a composition beside a tracked change commits the text it ended with, once, at the caret. A cancelled one leaves nothing behind. The browser could restart such a composition part way through, which left its earlier text in the document.
