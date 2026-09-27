---
"@stll/folio-react": minor
"@stll/folio-core": patch
---

The document outline and the comments no longer cover the page or each other. The editor lays the page out between the outline (start side) and the comments (end side) by the width it has, measured once: in a wide editor each gets a column; in a medium one the outline shrinks to a rail of heading ticks (hover or focus names a heading, its button opens the full outline as a drawer); in a narrow one a panel that no longer fits opens as a drawer over the page from a toolbar toggle, and closes on Escape or a press outside it. The thresholds derive from the page width at the current zoom and the panel widths.

The outline is a proper panel: an "Outline" header, headings nested by level on one line each, the active heading marked with an accent bar and kept in view, and one tab stop with arrow-key navigation. `OutlineRail` receives a new optional `presentation` prop (`"panel"` or `"rail"`) and is mounted inside the outline's surface rather than positioning itself. The outline now also appears once the body view exists, not only after the first edit.

The comments sidebar opens on load only when it has a thread to show; a document with tracked changes but no comments no longer opens an empty panel. A comments toggle in the toolbar shows the number of open threads. `showOutline` is documented with its actual default, `true`.

The horizontal ruler has an opaque surface (a new `--doc-canvas-surface` token: the canvas tint over the page colour), so page text no longer shows through it when a translucent `--muted` is in use; a comments drawer starts at the top of the editor rather than below the ruler.
