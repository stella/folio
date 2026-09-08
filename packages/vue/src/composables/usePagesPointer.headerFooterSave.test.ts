import { expect, mock, test } from "bun:test";
import { shallowRef } from "vue";

import { createEmptyDoc } from "@stll/folio-core/prosemirror/conversion/toProseDoc";
import type { Document, HeaderFooter } from "@stll/folio-core/types/document";
import { saveAndCloseHeaderFooterEdit, type HfEditState } from "./usePagesPointer";

const edit = (): HfEditState => ({
  isFirstPage: false,
  pageNumber: 1,
  position: "header",
  rId: "rId-stale",
  targetRect: null,
});

const document = (source: HeaderFooter | null): Document => ({
  package: {
    document: { content: [] },
    headers: new Map(source ? [["rId-stale", source]] : []),
  },
});

test.each([
  { name: "view", source: { type: "header", content: [] } satisfies HeaderFooter, view: null },
  { name: "source", source: null, view: { state: { doc: createEmptyDoc() } } },
])("a stale $name closes the header editor without publishing", ({ source, view }) => {
  const editState = shallowRef<HfEditState | null>(edit());
  const setDocument = mock((_document: Document) => {});
  const syncHfPMs = mock(() => {});
  const reLayout = mock(() => {});
  const onDocumentChange = mock((_document: Document) => {});

  expect(() =>
    saveAndCloseHeaderFooterEdit({
      editState,
      getDocument: () => document(source),
      getHfPmView: () => view,
      onDocumentChange,
      reLayout,
      setDocument,
      syncHfPMs,
    }),
  ).not.toThrow();

  expect(editState.value).toBeNull();
  expect(setDocument).not.toHaveBeenCalled();
  expect(syncHfPMs).not.toHaveBeenCalled();
  expect(reLayout).not.toHaveBeenCalled();
  expect(onDocumentChange).not.toHaveBeenCalled();
});
