---
"@stll/folio-core": patch
"@stll/folio-react": patch
---

Move the initial-layout font-readiness logic (`collectInitialLayoutFontFaces`, `collectInitialLayoutFontFamilies`, `documentFontsAreLoaded`, `getDocumentFontSet`, `waitForInitialLayoutFonts`) out of React's `PagedEditor.tsx` into a framework-neutral `@stll/folio-core/controller/fontReadiness` module. First slice of extracting orchestration from the React God component into the core controller; behavior is unchanged (the existing font-collection test moves to core alongside the code). No public API change.
