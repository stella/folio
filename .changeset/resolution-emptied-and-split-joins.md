---
"@stll/docx-core": patch
---

An accepted insertion whose content the same resolution removes cuts nothing between its neighbours, as resolving that content first leaves them. Accepting the deletion of an inserted paragraph break rejoins the records its split cut, as rejecting the insertion does. A paragraph with no content of its own, removed before a table or block while the paragraph before it joins in the same resolution, leaves all it held in that paragraph, as resolving the marks one at a time from the first does. A kept range marker whose named piece of its cut an accepted deletion takes goes into the piece across from it, as rejecting the revision it was kept from after that deletion places it.
