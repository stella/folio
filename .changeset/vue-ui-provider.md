---
"@stll/folio-vue": minor
---

Add the Vue UI-injection layer (`FolioUIProvider` / `provideFolioUI` / `useFolioUI`
/ `resolveFolioComponents` / `DEFAULT_COMPONENTS`), mirroring the React adapter's
`FolioUIComponents` contract. `DocxEditor` now provides the resolved primitive map
from its `components` prop, and the toolbar / formatting-bar chrome resolve their
`ColorPicker` through the provider, so a host override takes effect. Exports the
`FolioUIComponents`, `ColorPreset`, `FolioButtonProps`, and `OutlineItem` types plus
the `FolioUIProvider` component, closing the corresponding React↔Vue export-parity
gaps.
