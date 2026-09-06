---
"@stll/folio-core": minor
---

Blank paragraphs are blocks. The AI-facing snapshot skipped every paragraph
with no words, so a blank line had no id, nothing could address it, and a
comparison could neither add nor remove one. It carries them now, under an
additive `blank-NNNN` id shape that leaves the published `seq-NNNN` numbering
over the paragraphs that DO carry words exactly where it was, and the reading
surfaces filter them through one shared helper.

`FolioAIEditSnapshot.emptyDocumentAnchorId` is gone, which is why this is a
minor rather than a patch. It named the anchor an empty document had no block
for; an empty document is one blank paragraph, and that paragraph is now a
block, so its id is `blocks[0].id` like any other. Read that instead.

Three operations follow from that:

- `insertAfterBlock` / `insertBeforeBlock` with `text: ""` insert a blank
  paragraph rather than being refused as empty.
- `deleteBlock` on a blank removes it, with its paragraph mark tracked, rather
  than doing nothing.
- `insertTableRow` sizes the new row against the table's own column count, so a
  row a table can hold is no longer refused because a cell in it spans columns.

Five defects the blanks made visible, each of which was already losing content:

- An attribute-only edit — a paragraph mark, a list level, a style on a blank
  paragraph — never reached the saved file. The change tracker reads positions
  off each step, and `AttrStep` carries an empty step map, so the selective save
  wrote the paragraph's original XML and the edit was gone from the document
  while the editor still showed it.
- Deleting a block left its images behind. The deletion range is built from the
  block's text, so an image outside the words kept no mark, and accepting the
  deletion left a paragraph standing around an orphan picture.
- Zero-width anchors counted as content. A bookmark boundary, a text-box
  anchor, or Word's cached pagination boundary (`w:lastRenderedPageBreak`)
  survived a deletion that took every word, and the emptied paragraph stayed as
  a blank line nobody asked for. They are not content, they never carry a
  revision, and a revision landing on one produced a document the serializer
  refuses to write.
- Resolving a paragraph's mark could take a section with it. A section's
  properties live on a paragraph mark, so the paragraph is where the section
  ends: removing it merged two sections into one and dropped the removed
  section's page size, margins, and header and footer references. The
  properties now travel to the paragraph the resolution leaves behind, and a
  paragraph with nothing after it to carry them keeps its mark.
- Appending a paragraph at the end of a container handed its paragraph mark to
  the anchor. That is equivalent only while the two stay adjacent, and a table
  inserted between them by a later operation in the same batch separated them:
  the mark then joined the anchor to the table, which is nothing, and the
  appended paragraph survived a reject that should have closed it. Every
  inserted paragraph carries its own mark, and resolving a mark with nothing
  after it removes the emptied paragraph instead of joining.
