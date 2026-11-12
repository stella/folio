/**
 * The notes a document references, read reusing what earlier versions read,
 * equal the notes a read of every paragraph finds. An edit that removes a
 * note's last reference removes the note, so a stale read would keep a note
 * nothing references or drop one still referenced; every version of a random
 * run of edits, direct and tracked, in the body and in notes, is read both
 * ways.
 */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import { applyDocumentOp } from "../apply";
import { referencedNoteKeys, referencedNoteKeysInFull } from "../noteMarks";
import { documentStories } from "../stories";
import { OP_STORIES } from "../types";
import {
  documentArbitrary,
  opForStory,
  opSeedArbitrary,
  reviewDocumentArbitrary,
  trackedOpFor,
} from "./documentArbitraries";

setDefaultTimeout(propertyTestTimeout(240_000));

describe("the notes a shared document references", () => {
  test("equal a read of every paragraph over a run of edits", async () => {
    await assertProperty(
      fc.property(
        fc.oneof(documentArbitrary, reviewDocumentArbitrary),
        fc.array(
          fc.record({
            seed: opSeedArbitrary,
            tracked: fc.boolean(),
            story: fc.nat(),
          }),
          { minLength: 1, maxLength: 12 },
        ),
        (generated, edits) => {
          let document = generated;
          expect([...referencedNoteKeys(document)]).toEqual([
            ...referencedNoteKeysInFull(document),
          ]);
          edits.forEach(({ seed, tracked, story: storyIndex }, index) => {
            const stories = documentStories(document);
            const story =
              stories[storyIndex % stories.length] ?? OP_STORIES.MAIN;
            const applied = applyDocumentOp(
              document,
              tracked && story === OP_STORIES.MAIN
                ? trackedOpFor(document, seed, index)
                : opForStory({ document, seed, story }),
            );
            if (applied.isErr()) return;
            document = applied.value.document;
            expect([...referencedNoteKeys(document)]).toEqual([
              ...referencedNoteKeysInFull(document),
            ]);
          });
        },
      ),
      {
        id: "note reference reads share records soundly",
        numRuns: 200,
      },
    );
  });
});
