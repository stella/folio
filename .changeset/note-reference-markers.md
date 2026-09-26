---
"@stll/folio-core": minor
"@stll/folio-agents": minor
---

Text change: footnote and endnote references read as markers numbered in reading order, never as their package ids. `getContent()`, the AI snapshot, `getContentAsText()`, tracked-change and comment text, and the agent tools (`read_document`, `find_text`) show `Term A[^1] and Term B[^2]` where they showed `Term A10 and Term B30`; an endnote reads `[^e1]`. The Markdown export uses the same markers: footnotes and endnotes are numbered separately (it used one shared sequence), and a note gets one definition however often it is referenced. Blocks carry the markers as `structuralBoundaries` entries of type `noteReference`.

Behaviour change: a text operation can no longer delete or rewrite a reference. A `replaceInBlock`, `replaceRange` or `replaceBlock` whose match cuts into a marker, whose replacement drops or reorders one, or which writes a marker-shaped string is skipped with the new reason `protectedReference` (recovery `narrowMatch`), in direct, tracked and suggested mode alike. Replacements that keep each marker they cover edit the prose around it and leave the reference in place; `find_text` no longer matches inside a marker. Deleting a whole block still deletes its references.
