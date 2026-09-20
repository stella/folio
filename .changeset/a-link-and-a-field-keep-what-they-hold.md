---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Keep every child a `w:hyperlink` or a `w:fldSimple` holds and folio does not model, where it stood, through the editor as well as through a save.

`CT_Hyperlink` and `CT_SimpleField` are both `EG_PContent`: either may hold a permission range, a proofing error, a transparent wrapper or one of the eight custom-XML revision ranges between its runs. The link parser modelled the run and the two bookmark boundaries and returned `null` for everything else; the field parser read `w:r` and `w:hyperlink` and skipped the rest. Both walks now go through the shared child dispatcher, over generated declared-child sets the compiler makes their handler maps total over, with the verbatim sink as the default.

`Hyperlink["children"]` and `SimpleField["content"]` gain `PreservedInline`, so the capture is a member of the container's own content union and stands between the same two children in the model, in the ProseMirror document and in the saved part. The editor carries it as the opaque atom the paragraph level already uses, inside the link mark, so it moves, is accepted and is rejected with the link. A simple field holding one keeps its children rather than collapsing to its display text, which would have dropped the markup on the way out of the editor.

The link's handler map is exported and read twice: by the link parser, and by the paragraph parser's revision-segmenting walk, which overrides only the four `CT_RunTrackChange` wrappers because OOXML nests a revision inside a link and the model nests the link inside the revision. A child one of them starts recognising is recognised by both.
