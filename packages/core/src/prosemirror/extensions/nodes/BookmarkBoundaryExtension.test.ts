import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { EditorState, type Plugin } from "prosemirror-state";

import { schema } from "../../schema";
import { validateProseMirrorDocument } from "../../validation";
import { BookmarkBoundaryExtension } from "./BookmarkBoundaryExtension";

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

const getBoundaryIntegrityPlugin = (): Plugin => {
  const plugin = BookmarkBoundaryExtension().onSchemaReady({ schema }).plugins?.at(0);
  if (!plugin) {
    throw new Error("BookmarkBoundaryExtension must enforce boundary integrity");
  }
  return plugin;
};

const createBoundaryDocument = () =>
  schema.node("doc", null, [
    schema.node("paragraph", null, [
      schema.node("bookmarkBoundary", { type: "start", id: 101, name: "range-a" }),
      schema.text("alpha"),
      schema.node("bookmarkBoundary", { type: "start", id: 202, name: "range-b" }),
      schema.text("beta"),
      schema.node("bookmarkBoundary", { type: "end", id: 101 }),
      schema.text("gamma"),
      schema.node("bookmarkBoundary", { type: "end", id: 202 }),
    ]),
  ]);

const expectValidBoundaryStructure = (state: EditorState): void => {
  const result = validateProseMirrorDocument(state.doc);
  expect(result.issues.filter(({ message }) => message.includes("Bookmark"))).toEqual([]);
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

describe("BookmarkBoundaryExtension editing integrity", () => {
  test("removes the surviving endpoint when an edit deletes its pair", () => {
    const state = EditorState.create({
      doc: createBoundaryDocument(),
      plugins: [getBoundaryIntegrityPlugin()],
    });

    const applied = state.applyTransaction(state.tr.delete(1, 2));

    expect(applied.transactions).toHaveLength(2);
    expectValidBoundaryStructure(applied.state);
    expect(applied.state.doc.textContent).toBe("alphabetagamma");
  });

  test("preserves complete crossing pairs after ordinary text edits", () => {
    const state = EditorState.create({
      doc: createBoundaryDocument(),
      plugins: [getBoundaryIntegrityPlugin()],
    });

    const applied = state.applyTransaction(state.tr.insertText("x", 3));

    expect(applied.transactions).toHaveLength(1);
    expectValidBoundaryStructure(applied.state);
    let boundaryCount = 0;
    applied.state.doc.descendants((node) => {
      if (node.type.name === "bookmarkBoundary") {
        boundaryCount += 1;
      }
      return true;
    });
    expect(boundaryCount).toBe(4);
  });

  test("removes ambiguous node pairs that collide with paragraph bookmark ids", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", { bookmarks: [{ id: 303, name: "paragraph-range" }] }, [
        schema.node("bookmarkBoundary", { type: "start", id: 303, name: "node-range" }),
        schema.text("synthetic"),
        schema.node("bookmarkBoundary", { type: "end", id: 303 }),
      ]),
    ]);
    const state = EditorState.create({ doc, plugins: [getBoundaryIntegrityPlugin()] });

    const applied = state.applyTransaction(state.tr.insertText("x", 3));

    expect(applied.transactions).toHaveLength(2);
    expectValidBoundaryStructure(applied.state);
    let boundaryCount = 0;
    applied.state.doc.descendants((node) => {
      if (node.type.name === "bookmarkBoundary") {
        boundaryCount += 1;
      }
      return true;
    });
    expect(boundaryCount).toBe(0);
  });

  test("preserves paired-boundary validity under arbitrary deletion sequences", () => {
    fc.assert(
      fc.property(
        fc.array(fc.tuple(fc.nat(), fc.nat()), { minLength: 1, maxLength: 20 }),
        (deletions) => {
          let state = EditorState.create({
            doc: createBoundaryDocument(),
            plugins: [getBoundaryIntegrityPlugin()],
          });

          for (const [first, second] of deletions) {
            const paragraph = state.doc.firstChild;
            if (!paragraph) {
              throw new Error("Synthetic document must retain its paragraph");
            }
            const boundary = paragraph.content.size + 1;
            const from = 1 + (Math.min(first, second) % boundary);
            const to = 1 + (Math.max(first, second) % boundary);
            const applied = state.applyTransaction(state.tr.delete(from, to));
            state = applied.state;
            expectValidBoundaryStructure(state);
          }
        },
      ),
    );
  });
});
