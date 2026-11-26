---
"@stll/folio-core": minor
---

A template plan's output is written without text effects that show nothing.

A run's properties can carry `w14` text effects whose values make them invisible: a glow of radius zero, a shadow or reflection of no size, an outline without fill, a fill in the run's own colour, a 3-D scene with the default camera and nothing 3-D beside it. Such an effect still changes how some consumers treat the text. A run with a 3-D scene, even the default one, is drawn as a picture: a list number that carries one cannot be found, copied or read aloud in an exported document.

- `applyTemplatePlan` leaves these effects out of the parts the output takes from the template: its document, numbering, styles, headers and footers. One rule per effect names the values that show nothing; an effect that meets no rule is not touched, and a part without one keeps its bytes.
- `TemplatePlan.textEffects: "keep"` writes the parts as they are.
- `report.textEffects` counts, per part and effect, what the output no longer carries and what it still does.
- `inspectDocxTemplate` lists a template's `invisibleEffects` with the list level, style or text each belongs to, so the template itself can be mended.
