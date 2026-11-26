import { expect, spyOn, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import * as comments from "../__tests__/reference-editor/comments";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { createDocx } from "../docx/rezip";
import { paragraph, run } from "../docx/server/build";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { schema } from "../prosemirror/schema";
import { FolioDocxReviewer } from "./headless";
import * as read from "./read";

const date = "2026-01-02T03:04:05Z";

test("bulk membership preserves folded ids, duplicate listed ids and first suggestion ownership", () => {
  const mark = (revisionId: number, suggestionId?: string) =>
    schema.mark("insertion", {
      revisionId,
      author: "Writer",
      date,
      provenance: suggestionId ? "suggested" : "user",
      suggestionId,
    });
  const doc = schema.node("doc", null, [
    schema.node("paragraph", null, [
      schema.text("one", [mark(7, "first")]),
      schema.text("two", [
        mark(4, "first"),
        schema.mark("hyperlink", { href: "https://example.test/" }),
      ]),
    ]),
    schema.node("paragraph", null, [schema.text("three", [mark(7, "second")])]),
    schema.node("paragraph", null, [schema.text("ordinary", [mark(11)])]),
  ]);
  const state = EditorState.create({ doc });
  const owners = comments.suggestionOwnersByRevision(state);
  expect(owners.get(7)).toBe("first");
  for (const id of [7, 4, 11])
    expect(owners.get(id) ?? null).toBe(
      comments.suggestionIdOfRevision(state, id),
    );
  const listed = read.getTrackedChangesFromDoc(doc).map(({ id }) => ({
    id,
    groupIds: read.getTrackedChangeGroupIdsFromDoc(doc, id),
  }));
  expect(listed.map(({ id }) => id)).toEqual([7, 7, 11]);
  expect(read.listedRevisionGroups(doc)).toEqual(listed);
  expect(read.getTrackedChangeGroupsForIds(doc, [4, 7, 999, 11])).toEqual([
    [7, 4],
    [999],
    [11],
  ]);
});

for (const decision of ["accept", "reject"] as const)
  test(`canonical ${decision} classifies revision groups and suggestions once per changed story`, async () => {
    for (const count of [8, 32]) {
      const document = fromMarkdown("Body.");
      document.package.document.content = Array.from(
        { length: count },
        (_, id) =>
          paragraph([
            run("before "),
            {
              type: "insertion",
              info: { id, author: "Writer", date },
              content: [run("added")],
            },
          ]),
      );
      const prepared = await prepareCanonicalDocxInput(
        await createDocx(document),
      );
      if (prepared.isErr()) throw prepared.error;
      const reviewer = await FolioDocxReviewer.fromBuffer(
        prepared.value.slice().buffer,
        { session: "canonical" },
      );
      using groups = spyOn(read, "listedRevisionGroups");
      using owners = spyOn(comments, "suggestionOwnersByRevision");
      using repeatedGroups = spyOn(read, "getTrackedChangeGroupIdsFromDoc");
      using repeatedOwners = spyOn(comments, "suggestionIdOfRevision");
      const result =
        decision === "accept" ? reviewer.acceptAll() : reviewer.rejectAll();
      expect(result.resolved).toBe(count);
      expect(groups).toHaveBeenCalledTimes(1);
      expect(owners).toHaveBeenCalledTimes(1);
      expect(repeatedGroups).not.toHaveBeenCalled();
      expect(repeatedOwners).not.toHaveBeenCalled();
    }
  });
