---
"@stll/folio-core": minor
---

A template plan keeps tables apart and can keep justified lines in front of a manual line break unstretched.

- The writer puts an empty paragraph between two tables a plan places next to each other, and after a table that would end the body. Two tables with nothing between them read as one table, and a body has to end in a paragraph. `report.counts.tableSeparators` counts them; verification expects them in place.
- `TemplatePlan.compat.doNotExpandShiftReturn` switches `w:compat/w:doNotExpandShiftReturn` on in the output's settings, at its place in the schema's order and with every other byte of the part kept. A consumer that stretches a justified line ending in a manual line break stops doing so. Verification holds the output to it (`compat-mismatch`).
- A block manifest reports `compat.doNotExpandShiftReturn` of its document and, per paragraph, the `alignment` in effect (its own `w:jc`, else its style's).
