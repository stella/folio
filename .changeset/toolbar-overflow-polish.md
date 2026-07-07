---
"@stll/folio-react": patch
---

Fix `FormattingBar` overflow collapse: the secondary control group (font, color, alignment, lists) now collapses into the "More" popover based on live measurement of the toolbar's actual content width (via `ResizeObserver`), instead of a fixed bar-width breakpoint that ignored host `priorityExtra`/`inlineExtra` width and could leave controls scrolled out of view with no visible affordance; the "More" trigger is now rendered outside the scrollable region so it can never scroll away. Also fixes the zoom control and font-size picker truncating their labels, normalizes the alignment/list active-state affordance and icon sizes to match bold/italic/underline.
