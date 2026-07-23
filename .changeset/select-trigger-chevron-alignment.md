---
"@stll/folio-react": patch
---

Align the default select trigger's trailing chevron to the button's right edge. The `.folio-default-select-trigger` rule set `align-items: center` but no `justify-content`, so on fixed-width triggers (the style, markup, and zoom pickers) the chevron sat mid-button next to the label instead of at the right edge. Adding `justify-content: space-between` pins the label left and the chevron right; content-width triggers have no free space to distribute, so they are unaffected.
