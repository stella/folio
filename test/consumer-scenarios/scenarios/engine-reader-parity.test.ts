/**
 * A reviewer opened on the canonical session reads every synthetic document
 * exactly as a reviewer on the legacy engine does, through the packed public
 * entry points.
 *
 * The fixtures are packed through `ensureParaIds`, so both engines see the
 * same paragraph ids. A note story's anchor positions are left out: the
 * canonical model owns a normal note's automatic mark, an atom before the
 * note's text, so those positions may move by one while its text and block
 * ids stay (the in-repo reader parity property states that difference
 * exactly). The body's anchors are compared in full.
 */

import assert from "node:assert/strict";
import { describe, test } from "node:test";
import { isDeepStrictEqual } from "node:util";

import {
  FolioDocxReviewer,
  type FolioDocxReviewerOptions,
} from "@stll/folio-core/server";

import {
  AUTHOR,
  COLLISION_FIXTURES,
  FIXTURES,
  STORY_FIXTURES,
  toArrayBuffer,
} from "../support/documents.ts";

const ALL_FIXTURES = {
  ...FIXTURES,
  ...COLLISION_FIXTURES,
  ...STORY_FIXTURES,
} as const satisfies Record<string, () => Promise<Uint8Array>>;

type Reviewer = Awaited<ReturnType<typeof FolioDocxReviewer.fromBuffer>>;
type Engine = NonNullable<FolioDocxReviewerOptions["session"]>;

const openOn = (bytes: Uint8Array, session: Engine): Promise<Reviewer> =>
  FolioDocxReviewer.fromBuffer(toArrayBuffer(bytes), {
    author: AUTHOR,
    session,
  });

/** Every public read of a reviewer, as plain data. */
const reads = (reviewer: Reviewer) => {
  const { blocks, anchors } = reviewer.snapshot();
  const stories = reviewer.listStories();
  return {
    snapshot: { blocks, anchors },
    content: reviewer.getContent(),
    text: reviewer.getContentAsText(),
    annotated: reviewer.getContentAsText({ annotated: true }),
    notes: reviewer.getNotesAsText(),
    stories,
    readStory: stories.map(({ handle }) => reviewer.readStory(handle)),
    storyBlocks: stories.map(
      ({ handle }) => reviewer.snapshotStory(handle)?.blocks ?? null,
    ),
    reviewedStories: stories.map(({ handle }) => {
      const read = reviewer.readReviewedStory({
        story: handle,
        view: "current-markup",
      });
      return read && { text: read.text, changes: read.changes };
    }),
    changes: reviewer.getChanges(),
    comments: reviewer.getComments(),
    numbering: reviewer.readNumberingDefinitions(),
    properties: reviewer.getDocumentProperties(),
    pendingSuggestions: reviewer.exportPendingSuggestions(),
  };
};

type Reads = ReturnType<typeof reads>;

describe("reader parity across reviewer engines", () => {
  for (const [name, build] of Object.entries(ALL_FIXTURES)) {
    test(`${name}: canonical reads equal legacy reads`, async () => {
      const bytes = await build();
      const legacy = reads(await openOn(bytes, "legacy"));
      const canonical = reads(await openOn(bytes, "canonical"));
      const differing = (Object.keys(legacy) as (keyof Reads)[]).filter(
        (read) => !isDeepStrictEqual(canonical[read], legacy[read]),
      );
      const first = differing.at(0);
      assert.deepEqual(
        { differing, first: first && canonical[first] },
        { differing: [], first: first && legacy[first] },
        `${name}: reads that differ on the canonical session`,
      );
    });
  }

  test("the canonical reviewer reads the documents' content", async () => {
    // Guards the parity above against two engines that both read nothing.
    const stories = reads(
      await openOn(await STORY_FIXTURES.stories(), "canonical"),
    );
    assert.ok(stories.content.length > 0);
    assert.ok(stories.stories.some(({ handle }) => handle.type === "header"));
    assert.ok(stories.stories.some(({ handle }) => handle.type === "footnote"));
    const tracked = reads(
      await openOn(await FIXTURES.trackedChanges(), "canonical"),
    );
    assert.ok(tracked.changes.length > 0);
    const comments = reads(
      await openOn(await FIXTURES.comments(), "canonical"),
    );
    assert.ok(comments.comments.length > 0);
    const lists = reads(await openOn(await FIXTURES.lists(), "canonical"));
    assert.deepEqual(
      lists.content
        .filter((block) => block.kind === "listItem")
        .map((block) => block.displayLabel),
      ["•", "•", "1.", "2.", "3."],
    );
  });
});
