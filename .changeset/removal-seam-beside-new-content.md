---
"@stll/docx-core": patch
---

Accepting a deletion inside a link and then an insertion of new content beside it, with the rest of the link after that insertion, leaves the same records as accepting both at once: the inserted content stays its own record. Rejecting the insertion still joins the link's pieces around the deleted text.

Rejecting a link while text inserted inside it is pending, and then accepting that text, places the link's kept range boundary inside the piece of the link the text keeps, as accepting the text first does.
