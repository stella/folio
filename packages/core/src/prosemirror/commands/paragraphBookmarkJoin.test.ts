import { expect, test } from "bun:test";
import { EditorState, type Command, type Transaction } from "prosemirror-state";

import { schema } from "../schema";
import { acceptChange } from "./comments";

const boundary = (type: "start" | "end", id: number) =>
  schema.node("blockBookmarkBoundary", {
    type,
    id,
    ...(type === "start" ? { name: `bookmark${id}` } : {}),
  });

const resolve = (state: EditorState, command: Command) => {
  let transaction: Transaction | null = null;
  expect(command(state, (dispatched) => (transaction = dispatched))).toBe(true);
  if (!transaction) throw new Error("Expected a paragraph-mark resolution transaction");
  return { doc: state.apply(transaction).doc, mapping: transaction.mapping };
};

const bookmarkedParagraphBreak = () => {
  const first = schema.node(
    "paragraph",
    { pPrMark: { kind: "del", info: { id: 1, author: "Reviewer", date: "2026-09-09" } } },
    [schema.text("before")],
  );
  const second = schema.node("paragraph", null, [schema.text("after")]);
  const source = schema.node("doc", null, [
    boundary("start", 1),
    first,
    boundary("end", 1),
    boundary("start", 2),
    second,
    boundary("end", 2),
  ]);
  return { first, source, state: EditorState.create({ schema, doc: source }) };
};

test("a deleted paragraph break carries interstitial bookmarks into the joined paragraph", () => {
  const { first, source, state } = bookmarkedParagraphBreak();
  const single = resolve(state, acceptChange(0, source.content.size));

  expect(() => single.doc.check()).not.toThrow();
  expect(single.doc.childCount).toBe(3);
  const joined = single.doc.child(1);
  expect(joined.type.name).toBe("paragraph");
  expect(joined.textContent).toBe("beforeafter");
  expect(
    joined.content.content.map((node) => [node.type.name, node.attrs["type"], node.attrs["id"]]),
  ).toEqual([
    ["text", undefined, undefined],
    ["bookmarkBoundary", "end", 1],
    ["bookmarkBoundary", "start", 2],
    ["text", undefined, undefined],
  ]);

  const finalSecondText = single.doc.child(0).nodeSize + 1 + first.content.size + 2;
  expect(single.doc.nodeAt(finalSecondText)?.text).toBe("after");
});
