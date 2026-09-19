---
"@stll/folio-core": patch
---

Read each OOXML reserved value through a single owner, so no two call sites can disagree about it. `w:u w:val="none"` now cancels an inherited underline everywhere instead of painting a solid one when text is typed; `wp:anchor behindDoc="true"` puts a shape or text box behind the text, as `behindDoc="1"` already did; every `ST_OnOff` attribute accepts `on` and `off`, so `w:beforeAutospacing="on"` is no longer read as its opposite and saved back inverted; `w:shd w:fill="auto"` survives a save on paragraphs, styles and table cells, not only on runs; and `w:tblW`/`w:tcW` with `w:type="auto"` autofit instead of being pinned to the meaningless width Word leaves in `w:w`.
