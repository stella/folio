---
"@stll/folio-core": patch
---

An unchanged paragraph now writes its `w:pPr` properties back as they arrived, preserving validated unmodeled non-revision attributes and children and avoiding direct overrides synthesized from style-sourced numbering. The source is used only while a canonical formatting snapshot still matches and its structure passes validation; current section properties and tracked revisions are composed around it without accepting duplicates from the source. Unmodeled run-property revisions are excluded from captured full-repack replay until their accept/reject lifecycle is represented. Tracked-change UTC metadata now survives parser, editor, and serializer round trips under a canonical namespace binding.
