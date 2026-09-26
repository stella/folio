import { expect, test } from "bun:test";
import { EditorState, type Command, type Transaction } from "prosemirror-state";
import { Step } from "prosemirror-transform";

import { schema } from "../schema";
import { acceptAllChanges, acceptChange } from "./comments";

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

test("whole-story resolution preserves positions through a bookmark-spanning join", () => {
  const { first, source, state } = bookmarkedParagraphBreak();
  const bulk = resolve(state, acceptAllChanges());
  const single = resolve(state, acceptChange(0, source.content.size));
  expect(bulk.doc.eq(single.doc)).toBe(true);

  const sourceBoundary = source.child(0).nodeSize + first.nodeSize;
  const finalBoundary = bulk.doc.child(0).nodeSize + 1 + first.content.size;
  const sourceSecondText = sourceBoundary + 2 + 1;
  expect(bulk.mapping.map(sourceBoundary)).toBe(finalBoundary);
  expect(bulk.mapping.map(sourceSecondText)).toBe(finalBoundary + 2);

  let transaction: Transaction | null = null;
  acceptAllChanges()(state, (dispatched) => {
    transaction = dispatched;
  });
  if (!transaction) throw new Error("Expected a whole-story resolution transaction");
  const step = transaction.steps.at(0);
  if (!step) throw new Error("Expected a whole-story resolution step");
  const replayed = Step.fromJSON(schema, step.toJSON()).apply(source);
  expect(replayed.failed).toBeNull();
  expect(replayed.doc?.eq(bulk.doc)).toBe(true);
});

test("whole-story resolution joins a chain across two bookmark pairs", () => {
  const marked = (text: string, id: number) =>
    schema.node(
      "paragraph",
      { pPrMark: { kind: "del", info: { id, author: "Reviewer", date: "2026-09-09" } } },
      schema.text(text),
    );
  const source = schema.node("doc", null, [
    boundary("start", 1),
    marked("A", 1),
    boundary("end", 1),
    boundary("start", 2),
    marked("B", 2),
    boundary("end", 2),
    boundary("start", 3),
    schema.node("paragraph", null, schema.text("C")),
    boundary("end", 3),
  ]);
  const { doc, mapping } = resolve(EditorState.create({ schema, doc: source }), acceptAllChanges());
  expect(() => doc.check()).not.toThrow();
  expect(doc.childCount).toBe(3);
  expect(doc.child(1).textContent).toBe("ABC");
  let sourceThirdText = 1;
  for (let index = 0; index < 7; index++) sourceThirdText += source.child(index).nodeSize;
  expect(mapping.map(sourceThirdText)).toBe(doc.child(0).nodeSize + doc.child(1).nodeSize - 2);
  expect(
    doc
      .child(1)
      .content.content.filter((node) => node.type.name === "bookmarkBoundary")
      .map((node) => [node.attrs["type"], node.attrs["id"]]),
  ).toEqual([
    ["end", 1],
    ["start", 2],
    ["end", 2],
    ["start", 3],
  ]);
});
