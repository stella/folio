import { describe, expect, test } from "bun:test";

import { schema } from "../../schema";

class FakeHTMLElement {
  constructor(private readonly attrs: Readonly<Record<string, string>>) {}

  getAttribute(name: string): string | null {
    return this.attrs[name] ?? null;
  }
}

const parseBookmarkBoundary = (attrs: Readonly<Record<string, string>>) => {
  const getAttrs = schema.nodes.bookmarkBoundary.spec.parseDOM?.at(0)?.getAttrs;
  if (!getAttrs) {
    throw new Error("BookmarkBoundaryExtension must define parseDOM[0].getAttrs");
  }
  return getAttrs(new FakeHTMLElement(attrs) as unknown as HTMLElement);
};

describe("BookmarkBoundaryExtension DOM round-trip", () => {
  test("preserves table-column bookmark bounds", () => {
    const attrs = {
      type: "start",
      id: 12,
      name: "clause",
      colFirst: 2,
      colLast: 4,
    } as const;
    const node = schema.node("bookmarkBoundary", attrs);
    const toDOM = node.type.spec.toDOM;
    if (!toDOM) {
      throw new Error("BookmarkBoundaryExtension must define toDOM");
    }

    expect(toDOM(node)).toEqual([
      "span",
      expect.objectContaining({
        "data-docx-bookmark-boundary": "start",
        "data-docx-bookmark-id": "12",
        "data-docx-bookmark-name": "clause",
        "data-docx-bookmark-col-first": "2",
        "data-docx-bookmark-col-last": "4",
        "aria-hidden": "true",
        contenteditable: "false",
        style: "display: none;",
        "data-docx-internal-clipboard": expect.any(String),
      }),
    ]);
    expect(
      parseBookmarkBoundary({
        "data-docx-bookmark-boundary": "start",
        "data-docx-bookmark-id": "12",
        "data-docx-bookmark-name": "clause",
        "data-docx-bookmark-col-first": "2",
        "data-docx-bookmark-col-last": "4",
      }),
    ).toEqual(attrs);
  });

  test("rejects malformed table-column bookmark bounds", () => {
    expect(
      parseBookmarkBoundary({
        "data-docx-bookmark-boundary": "start",
        "data-docx-bookmark-id": "12",
        "data-docx-bookmark-name": "clause",
        "data-docx-bookmark-col-first": "-1",
      }),
    ).toBe(false);
    expect(
      parseBookmarkBoundary({
        "data-docx-bookmark-boundary": "start",
        "data-docx-bookmark-id": "12junk",
        "data-docx-bookmark-name": "clause",
      }),
    ).toBe(false);
  });

  test("shares one internal clipboard capability with text-box anchors", () => {
    const boundary = schema.node("bookmarkBoundary", { type: "end", id: 12 });
    const anchor = schema.node("textBoxAnchor", { anchorId: "0:0" });
    const boundaryToDOM = boundary.type.spec.toDOM;
    const anchorToDOM = anchor.type.spec.toDOM;
    if (!boundaryToDOM || !anchorToDOM) {
      throw new Error("Reconstruction atoms must define toDOM");
    }

    const boundaryDom = boundaryToDOM(boundary) as [string, Record<string, string>];
    const anchorDom = anchorToDOM(anchor) as [string, Record<string, string>];

    expect(boundaryDom[1]["data-docx-internal-clipboard"]).toBeTruthy();
    expect(anchorDom[1]["data-docx-internal-clipboard"]).toBe(
      boundaryDom[1]["data-docx-internal-clipboard"],
    );
  });
});
