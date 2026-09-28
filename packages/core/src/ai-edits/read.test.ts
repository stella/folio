import { expect, test } from "bun:test";

import { schema } from "../prosemirror/schema";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { getCommentAnchorsFromDoc, getTrackedChangesFromDoc } from "./read";

test("orders changes by document block and stable revision id", () => {
  const insertion = schema.mark("insertion", {
    revisionId: 2,
    author: "Reviewer",
    date: "2026-06-20T00:00:00Z",
  });
  const deletion = schema.mark("deletion", {
    revisionId: 1,
    author: "Reviewer",
    date: "2026-06-20T00:00:00Z",
  });
  const live = schema.node("doc", null, [
    schema.node("paragraph", { paraId: "00000001" }, [
      schema.text("later revision", [insertion]),
      schema.text("earlier revision", [deletion]),
    ]),
    schema.node("paragraph", { paraId: "00000002" }, [schema.text("first block", [deletion])]),
  ]);
  const changes = (doc: typeof live) =>
    getTrackedChangesFromDoc(doc).map(({ id, type, text }) => ({ id, type, text }));

  expect(changes(live)).toEqual([
    { id: 1, type: "deletion", text: "earlier revision" },
    { id: 2, type: "insertion", text: "later revision" },
    { id: 1, type: "deletion", text: "first block" },
  ]);
  expect(changes(toProseDoc(fromProseDoc(live)))).toEqual(changes(live));
});

test("comment quotes span unmarked runs and paragraphs after a save", () => {
  const comment = schema.mark("comment", { commentId: 17 });
  const insertion = schema.mark("insertion", {
    revisionId: 1,
    author: "Reviewer",
    date: "2026-06-20T00:00:00Z",
  });
  const live = schema.node("doc", null, [
    schema.node("paragraph", null, [
      schema.text("before"),
      schema.text("hello", [comment]),
      schema.text("XXX", [insertion]),
      schema.text("world", [comment]),
    ]),
    schema.node("paragraph", null, [schema.text("outside")]),
    schema.node("paragraph", null, [schema.text("again", [comment]), schema.text("after")]),
  ]);
  const reopened = toProseDoc(fromProseDoc(live));
  const quotes = (doc: typeof live) => getCommentAnchorsFromDoc(doc).map(({ quote }) => quote);

  expect(quotes(live)).toEqual(["helloXXXworldoutsideagain"]);
  expect(quotes(reopened)).toEqual(quotes(live));
});
