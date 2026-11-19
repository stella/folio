---
"@stll/folio-core": patch
---

Apply a tracked replacement over comment ranges instead of refusing it: the deleted words keep their comments, so rejecting gives every comment back, and accepting drops and reports a comment whose words all went.
