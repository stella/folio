import { expect, test } from "bun:test";

import { schema } from "../prosemirror/schema";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import {
  getCommentAnchorsFromDoc,
  getTrackedChangeStatsFromDoc,
  getTrackedChangesFromDoc,
} from "./read";

test("keeps change carriers in document order when ids go backward", () => {
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
    { id: 2, type: "insertion", text: "later revision" },
    { id: 1, type: "deletion", text: "earlier revision" },
    { id: 1, type: "deletion", text: "first block" },
  ]);
  expect(changes(toProseDoc(fromProseDoc(live)))).toEqual(changes(live));
});

test("folds a revision split only at a hyperlink boundary", () => {
  const insertion = (revisionId: number) =>
    schema.mark("insertion", { revisionId, author: "Reviewer", date: "2026-06-20T00:00:00Z" });
  const link = schema.mark("hyperlink", { href: "https://example.test/terms" });
  const doc = schema.node("doc", null, [
    schema.node("paragraph", { paraId: "00000001" }, [
      schema.text("Stages ", [insertion(7)]),
      schema.text("are", [insertion(0), link]),
      schema.text(" invoiced separately.", [insertion(1)]),
      schema.text(" Separate edit.", [insertion(2)]),
    ]),
  ]);

  expect(getTrackedChangesFromDoc(doc).map(({ id, text }) => ({ id, text }))).toEqual([
    { id: 7, text: "Stages are invoiced separately." },
    { id: 2, text: " Separate edit." },
  ]);
});

test("folds a revision split by a comment range and reference", () => {
  const insertion = (revisionId: number) =>
    schema.mark("insertion", { revisionId, author: "Reviewer", date: "2026-06-20T00:00:00Z" });
  const comment = schema.mark("comment", { commentId: 17 });
  const doc = schema.node("doc", null, [
    schema.node("paragraph", { paraId: "00000001" }, [
      schema.text("Stages ", [insertion(7)]),
      schema.text("are", [insertion(80), comment]),
      schema.node("commentReference", { commentId: 17 }),
      schema.text(" invoiced separately.", [insertion(100)]),
      schema.text(" Separate edit.", [insertion(2)]),
    ]),
  ]);

  expect(getTrackedChangesFromDoc(doc).map(({ id, text }) => ({ id, text }))).toEqual([
    { id: 7, text: "Stages are invoiced separately." },
    { id: 2, text: " Separate edit." },
  ]);
  expect(getTrackedChangeStatsFromDoc(doc).highestId).toBe(100);
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
