---
"@stll/folio-core": minor
"@stll/folio-react": minor
"@stll/folio-vue": minor
---

The inline autocomplete suggestion works in the canonical session as in the default one, through `DocxEditorRef.autocomplete`: `start({ at?, requestId? })` (at the caret by default), `append(requestId, delta)`, `finish(requestId)`, `clear()`, `accept("all" | "word")`, `get()` and `canTrigger()`, with the `onAutocompleteChange` prop reporting each change. The suggestion is view state anchored at a text position: it never changes the document, its undo history or its pending changes, and a peer's edit in a shared room moves it as it moves the caret. Typing its next characters consumes them; any other edit dismisses it. Accepting inserts the text as typed text at the anchor, one undo step, tracked when suggesting. Tab accepts, Mod-ArrowRight accepts the next word and Escape dismisses, the same in React and Vue.

The transaction builders `startAutocompleteSuggestion`, `appendAutocompleteToken`, `finishAutocompleteSuggestion`, `clearAutocompleteSuggestion`, the commands `acceptAutocompleteSuggestion` and `acceptAutocompleteWord`, `getAutocompleteSuggestion`, `autocompleteSuggestionKey`, `autocompleteSuggestionPlugin`, `shouldTriggerAutocomplete` and their types are removed: use `DocxEditorRef.autocomplete` and the `FolioAutocompleteApi` / `FolioAutocompleteState` types.
