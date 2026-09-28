import { expect, test } from "bun:test";

import { schema } from "../prosemirror/schema";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { getCommentAnchorsFromDoc } from "./read";

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
