---
"@stll/folio-core": minor
---

Carry a stored `outlineLevel` node attr forward to the union shape.

`FOLIO_YJS_ATTR_SCHEMA_VERSION` moves to 4, and the version-3 step rewrites every paragraph's `outlineLevel` from the `w:outlineLvl w:val` number it stored into `OutlineLevel`: 0..8 become the heading arm, 9 becomes the body-text arm, and a value the format never defined is dropped, matching the parse boundary.

The step cannot be skipped and read lazily. ProseMirror copies a stored attr into the node without validating it, so an unmigrated snapshot would reach the strict validator as a bare number; the version gate fires first, on both the editor and the materialization load paths, which is what turns a silent misread into a rewrite.
