import { act, createRef } from "react";
import { createRoot } from "react-dom/client";

import { GlobalRegistrator } from "@happy-dom/global-registrator";
import { panic } from "better-result";
import { describe, expect, test } from "bun:test";

import type { Document, HeaderFooter } from "@stll/folio-core/types/document";

import {
  enumerateHfSlots,
  HiddenHeaderFooterPMs,
} from "./HiddenHeaderFooterPMs";
import type { HiddenHeaderFooterPMsRef } from "./HiddenHeaderFooterPMs";

const persistentHeaderDocument = {
  package: {
    document: { content: [] },
    headers: new Map([
      [
        "rIdH",
        {
          type: "header",
          hdrFtrType: "default",
          content: [
            {
              type: "paragraph",
              content: [
                { type: "run", content: [{ type: "text", text: "Header" }] },
              ],
            },
          ],
        },
      ],
    ]),
  },
} satisfies Document;

test("persistent header/footer views refuse edits through read-only switches", async () => {
  const ownsDom = typeof document === "undefined";
  if (ownsDom) GlobalRegistrator.register();
  const actEnvironment = Reflect.get(globalThis, "IS_REACT_ACT_ENVIRONMENT");
  Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", true);
  const host = document.createElement("div");
  document.body.append(host);
  const root = createRoot(host);
  const ref = createRef<HiddenHeaderFooterPMsRef>();
  try {
    await act(async () =>
      root.render(
        <HiddenHeaderFooterPMs ref={ref} document={persistentHeaderDocument} />,
      ),
    );
    const view = ref.current?.getView("rIdH");
    if (!view) panic("Expected a persistent header view");
    const authored = view.state.doc;
    for (const readOnly of [true, false, true]) {
      await act(async () =>
        root.render(
          <HiddenHeaderFooterPMs
            ref={ref}
            document={persistentHeaderDocument}
            readOnly={readOnly}
          />,
        ),
      );
      expect(ref.current?.getView("rIdH")).toBe(view);
      expect(view.editable).toBe(!readOnly);
      if (readOnly) {
        await act(async () =>
          view.dispatch(view.state.tr.insertText("Refused", 1)),
        );
        expect(view.state.doc.eq(authored)).toBe(true);
      }
    }
  } finally {
    await act(async () => root.unmount());
    host.remove();
    Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", actEnvironment);
    if (ownsDom) GlobalRegistrator.unregister();
  }
});

function hf(kind: "header" | "footer" = "header"): HeaderFooter {
  return {
    type: kind,
    hdrFtrType: "default",
    content: [],
  };
}

function makeDocument(
  headers: Map<string, HeaderFooter> | undefined,
  footers: Map<string, HeaderFooter> | undefined,
): Document {
  return {
    package: {
      document: { content: [], sections: [] },
      ...(headers ? { headers } : {}),
      ...(footers ? { footers } : {}),
    },
  } as unknown as Document;
}

describe("enumerateHfSlots", () => {
  test("returns [] for null document", () => {
    expect(enumerateHfSlots(null)).toEqual([]);
  });

  test("returns [] for document with no headers or footers", () => {
    const doc = makeDocument(undefined, undefined);
    expect(enumerateHfSlots(doc)).toEqual([]);
  });

  test("enumerates each header rId once", () => {
    const headers = new Map<string, HeaderFooter>([
      ["rId1", hf("header")],
      ["rId2", hf("header")],
    ]);
    const slots = enumerateHfSlots(makeDocument(headers, undefined));
    expect(slots).toEqual([
      { rId: "rId1", kind: "header" },
      { rId: "rId2", kind: "header" },
    ]);
  });

  test("enumerates headers and footers together", () => {
    const headers = new Map([["rIdH", hf("header")]]);
    const footers = new Map([["rIdF", hf("footer")]]);
    const slots = enumerateHfSlots(makeDocument(headers, footers));
    expect(slots).toEqual([
      { rId: "rIdH", kind: "header" },
      { rId: "rIdF", kind: "footer" },
    ]);
  });

  test("two sections sharing one header rId share one slot", () => {
    // The package-level Map is the storage. The "shared by rId" pattern
    // (ECMA-376 §17.10.1) means two sections both reference rIdH; the bag
    // still has one entry. Slot enumeration must reflect that.
    const headers = new Map([["rIdH", hf("header")]]);
    const slots = enumerateHfSlots(makeDocument(headers, undefined));
    expect(slots.length).toBe(1);
    expect(slots[0]).toEqual({ rId: "rIdH", kind: "header" });
  });

  test("defensively dedupes an rId that appears in both bags", () => {
    // The OOXML schema keeps header/footer rIds disjoint, but if both bags
    // accidentally registered the same rId we should emit one slot, not two.
    const headers = new Map([["rIdShared", hf("header")]]);
    const footers = new Map([["rIdShared", hf("footer")]]);
    const slots = enumerateHfSlots(makeDocument(headers, footers));
    expect(slots).toEqual([{ rId: "rIdShared", kind: "header" }]);
  });
});
