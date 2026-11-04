---
"@stll/folio-react": minor
---

The injected `ColorPicker` takes its popover trigger as `trigger`, a single element that becomes the trigger itself, instead of wrapping `children` in a second element with button semantics. `children` is removed from `FolioColorPickerProps`. `ToolbarButton` forwards extra button attributes and a ref, so the formatting bar's font color button is the picker's trigger.
