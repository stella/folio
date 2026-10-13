# Document search

React and Vue expose document search through `DocxEditorRef`:

```tsx
const result = editorRef.current?.find("clause", { matchCase: false });
editorRef.current?.findNext();
editorRef.current?.findPrevious();
editorRef.current?.getFindResult();
editorRef.current?.clearFind();
```

Use `onFindResultChange` to receive query and cursor updates. `totalCount` is the
number of matches; `currentIndex` is zero-based. Display `currentIndex + 1` when
`totalCount > 0`, otherwise display `0 / 0`. Clearing search reports `null`.
Vue also emits `find-result-change` with the same result.

Every match is highlighted. The current match uses `--doc-find-current`; other
matches use `--doc-find-highlight`. Next and previous wrap around and reveal the
current match. Search works in read-only and editing modes. Highlights are view
state and do not change selections, document content, saved DOCX, or undo history.

Navigation order is body paragraphs (table rows and cells in document order,
shape text after its anchor), headers, footers, footnotes, endnotes, then comments.
Header and footer parts follow their first section reference, with default,
first-page, and even-page variants ordered within that section. Only variants
selected by an actual page are searched; a shared part is counted once and
navigation reveals its first rendered occurrence. Note separators are excluded.
Comments and replies are searched where the adapter exposes the comment sidebar;
navigation opens the thread. Cached field results are searched as displayed text.
Page-dependent fields (`PAGE`, `NUMPAGES`, `SECTIONPAGES`, and `PAGEREF`) are
excluded: their values depend on pagination, while shared parts are counted once.

`matchCase`, `matchWholeWord`, and `useRegex` control matching. Diacritics remain
significant; search does not fold accents. Search respects the editor's display
mode: final views omit deleted and moved-away text, and original view omits
inserted and moved-to text. Query input and document refreshes are debounced.

Replacement remains a body editing operation. A secondary-story, shape, or comment
result cannot be replaced through the existing body replacement commands. Replace
All refuses a mixed result set rather than applying story-local offsets to the
body. Read-only mode rejects replacements.

Core consumers can call `findInDocument` from `@stll/folio-core/utils/findReplace`.
Pass the active page-selected stories as `FindSearchContext.renderedStories` to
limit header/footer eligibility; `getRenderedSearchStories` from
`@stll/folio-core/utils/findMatchView` derives them through the same selector as the
painter. Omitting that context searches all package parts. `includeComments` is
opt-in for hosts with an existing comment surface.
