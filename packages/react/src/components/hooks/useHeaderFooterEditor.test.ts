import { expect, mock, test } from "bun:test";

import { deepFreeze } from "@stll/folio-core/__tests__/deepFreeze";
import {
  createEmptyDoc,
  toProseDoc,
} from "@stll/folio-core/prosemirror/conversion/toProseDoc";
import type { Document, HeaderFooter } from "@stll/folio-core/types/document";
import type { ReadonlyDocument } from "@stll/folio-core/types/documentModel";

import { saveAndCloseHeaderFooterEdit } from "./useHeaderFooterEditor";

test("a stale header view closes when its source disappeared", () => {
  const document = {
    package: {
      document: { content: [] },
      headers: new Map(),
    },
  } satisfies Document;
  const pushDocument = mock((next: ReadonlyDocument) => next);
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

test("saving a header writes a copy, never the frozen model", () => {
  const header: HeaderFooter = {
    type: "header",
    hdrFtrType: "default",
    content: [
      {
        type: "paragraph",
        content: [{ type: "run", content: [{ type: "text", text: "Head" }] }],
      },
    ],
  };
  const model = deepFreeze({
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            content: [
              { type: "run", content: [{ type: "text", text: "Body" }] },
            ],
          },
        ],
      },
      headers: new Map([["rId1", header]]),
    },
  } satisfies Document);
  const pushed: ReadonlyDocument[] = [];
  const pushDocument = (next: ReadonlyDocument) => {
    pushed.push(next);
    return next;
  };
  const view = {
    state: {
      doc: toProseDoc({
        package: { document: { content: structuredClone(header.content) } },
      }),
    },
  };

  saveAndCloseHeaderFooterEdit({
    activeRId: "rId1",
    document: model,
    editPosition: "header",
    isFirstPage: false,
    pushDocument,
    setEditPosition: () => {},
    view,
  });

  expect(pushed).toHaveLength(1);
  const saved = pushed[0];
  expect(saved).not.toBe(model);
  expect(saved?.package.document).not.toBe(model.package.document);
  expect(saved?.package.headers?.get("rId1")).not.toBe(header);
  expect(model.package.headers.get("rId1")).toBe(header);
});
