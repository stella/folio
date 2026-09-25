---
"@stll/folio-core": patch
---

Paint a bullet level's Latin-range `w:lvlText` character as itself when the level's `w:rFonts` names an ordinary text face, such as the `o` bullet in a monospace font. Symbol, Wingdings and Webdings levels, and levels that name no face, keep their symbol-glyph mapping.
