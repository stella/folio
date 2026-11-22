---
"@stll/docx-core": patch
---

A pending insertion's cut is read between the records its own pieces stand between, past insertions made after it, so resolving those records first or with it leaves the same cut. Two pieces of a pending insertion with only kept range markers between them join once accepted, whatever stands beyond them. A range marker kept from outside the record a later insertion was made in stays beside that insertion once accepted.
