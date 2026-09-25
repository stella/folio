---
"@stll/docx-core": minor
"@stll/folio-core": patch
---

Lay out and paint the text boxes inside a DrawingML group (`wpg:wgp`, nested `wpg:grpSp`) as text boxes: each child's frame is mapped through the group's `a:chOff`/`a:chExt` child coordinate space, keeps its `wps:bodyPr` insets, anchoring and autofit, and carries its own rotation and flips. Editing that text writes it back into the group on save, which keeps the group intact. The group preview also draws zero-extent lines, scales line widths into the child space and renders nested groups. Adds `Shape.groupChild` / `TextBox.groupChild` (`DrawingGroupChild`) to the model.
