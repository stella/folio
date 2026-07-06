---
"@stll/folio-vue": minor
---

Close the Vue `components` UI-injection gap so a host can override all ten of
folio's chrome primitives, matching React: `FolioUIComponents` (and
`DEFAULT_COMPONENTS`) grow from `Button`/`ColorPicker`/`Popover`/`Menu` to the
full set, adding `Dialog`, `Select`, `Input`, `Checkbox`, `DatePickerPopover`,
and `OutlineRail` with a real Vue default per primitive (each collapsed into
one monolithic, data-driven component rather than React's base-ui part-object
shape — the pattern the existing four already used). Chrome consumers now
resolve their primitive through `useFolioUI()` instead of a static import:
`MenuBar` (Menu), 8 previously-static `Popover` consumers
(`EditingModeDropdown`, `ReviewControls`, `TableGridPicker`,
`TableBorderWidthPicker`, `IconGridDropdown`, `TableBorderPicker`,
`TableMoreDropdown`, `AlignmentButtons`), `ZoomControl` / `StylePicker`
(Select), `FindReplaceDialog` (Button/Input/Checkbox), the five modal dialogs
— `PageSetupDialog`, `ImagePropertiesDialog`, `FootnotePropertiesDialog`,
`ImagePositionDialog`, `TablePropertiesDialog` (Dialog), and
`DocumentOutline`'s item list (OutlineRail). `components` moves from
`deferredInVue` to `paired` in the parity contract. New default components and
the `Folio*Props` contract types are exported from the `@stll/folio-vue/ui`
subpath. Minor fix bundled in: Escape now closes all five modal dialogs
(`ImagePropertiesDialog` previously swallowed the keydown before any handler
saw it).
