import { Result } from "better-result";
import { expect, spyOn, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import * as ops from "@stll/docx-core/ops";

import * as comments from "../__tests__/reference-editor/comments";
import * as canonicalRoom from "../canonical/canonicalRoomBase";
import { createCanonicalSession } from "../canonical/canonicalSession";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { createDocx } from "../docx/rezip";
import { hyperlink, paragraph, run } from "../docx/server/build";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Insertion } from "../types/document";
import { FolioDocxReviewer } from "./headless";
import * as read from "./read";

const date = "2026-01-02T03:04:05Z";

test("bulk membership preserves folded ids, duplicate listed ids and first suggestion ownership", () => {
  const insertion = ({
    id,
    content,
    suggestionId,
  }: {
    id: number;
    content: Insertion["content"];
    suggestionId?: string;
  }): Insertion => ({
    type: "insertion",
    info: {
      id,
      author: "Reviewer",
      date,
      ...(suggestionId === undefined ? {} : { suggestionId }),
    },
    content,
  });
  const model = fromMarkdown("Body.");
  model.package.document.content = [
    paragraph([
      insertion({ id: 7, content: [run("one")], suggestionId: "first" }),
      insertion({
        id: 4,
        content: [hyperlink({ text: "two", href: "https://example.test/" })],
        suggestionId: "first",
      }),
    ]),
    paragraph([
      insertion({ id: 7, content: [run("three")], suggestionId: "second" }),
    ]),
    paragraph([insertion({ id: 11, content: [run("ordinary")] })]),
  ];
  const doc = toProseDoc(model);
  const state = EditorState.create({ doc });
  const owners = comments.suggestionOwnersByRevision(state);
  expect(ops.suggestionOwnersInStory(model, ops.OP_STORIES.MAIN)).toEqual(
    owners,
  );
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

test("headless review adopts a proposal carried by one of several run property changes", async () => {
  const document = fromMarkdown("Body.");
  const suggestionId = "proposed-emphasis";
  document.package.document.content = [
    {
      ...paragraph([
        {
          ...run("Body", { italic: true, bold: true }),
          propertyChanges: [
            {
              type: "runPropertyChange",
              info: { id: 11, author: "Reviewer", date, suggestionId },
              previousFormatting: { bold: true },
            },
            {
              type: "runPropertyChange",
              info: { id: 12, author: "Author", date },
              previousFormatting: {},
            },
          ],
        },
      ]),
      paraId: "10000001",
    },
  ];
  const session = createCanonicalSession(document).unwrap();
  using open = spyOn(
    canonicalRoom,
    "openCanonicalDocxSession",
  ).mockResolvedValue(
    Result.ok({ document, session, mintedParaIds: new Set() }),
  );
  const reviewer = await FolioDocxReviewer.fromBuffer(new ArrayBuffer(0));

  expect(open).toHaveBeenCalledTimes(1);
  expect(ops.pendingSuggestionIdsIn(document)).toEqual(new Set([suggestionId]));
  expect(reviewer.acceptSuggestion(suggestionId)).toBe(true);
  expect(reviewer.acceptSuggestion(suggestionId)).toBe(false);
  expect(reviewer.getChanges().map(({ id }) => id)).toEqual([11, 12]);
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
      using owners = spyOn(ops, "suggestionOwnersInStory");
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
