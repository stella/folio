import { expect, mock, test } from "bun:test";

import { createEmptyDoc } from "@stll/folio-core/prosemirror/conversion/toProseDoc";
import type { Document } from "@stll/folio-core/types/document";
import { saveAndCloseHeaderFooterEdit } from "./useHeaderFooterEditor";

test("a stale header view closes when its source disappeared", () => {
  const document = {
    package: {
      document: { content: [] },
      headers: new Map(),
    },
  } satisfies Document;
  const pushDocument = mock((next: Document) => next);
  const setEditPosition = mock((_position: "header" | "footer" | null) => {});

  expect(() =>
    saveAndCloseHeaderFooterEdit({
      activeRId: "rId-stale",
      document,
      editPosition: "header",
      isFirstPage: false,
      pushDocument,
      setEditPosition,
      view: { state: { doc: createEmptyDoc() } },
    }),
  ).not.toThrow();

  expect(setEditPosition).toHaveBeenCalledTimes(1);
  expect(setEditPosition).toHaveBeenCalledWith(null);
  expect(pushDocument).not.toHaveBeenCalled();
});
