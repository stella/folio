---
"@stll/folio-core": patch
---

Stop giving a shape outline a dash pattern it never stated. An `a:ln` with no
`a:prstDash` came back through the editor as `style: "solid"`, which is the
shape's own decision rather than the absence the source had, so a later change
to what an unstated outline renders as could no longer reach it. The shape
node's `outlineStyle` now defaults to absent, and the renderer already draws an
unstated outline solid.
