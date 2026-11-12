/**
 * A canonical session never writes into a version of its document it has
 * handed out.
 *
 * A host may hold any version the session handed out as its read-only model
 * (in its history, or as the document layout reads) while the session goes
 * on editing. Random runs of edits through a canonical session, over
 * documents with tables, lists, comments, tracked changes, notes, fields and
 * text boxes, typing, deleting, splitting, joining, lists, accepting and
 * rejecting revisions, undo and redo, deep-freeze every version as it is
 * handed out: a later edit that wrote into one would throw.
 */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";

import { checkCoverageSample } from "../../../../test/property-coverage";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { deepFreeze } from "../__tests__/deepFreeze";
import { createCanonicalSession } from "../canonical/canonicalSession";
import { schema } from "../prosemirror/schema";
import {
  applyEdit,
  editArbitrary,
  parsedSources,
} from "./__tests__/canonicalEditRuns";

setDefaultTimeout(propertyTestTimeout(320_000));

describe("canonical document versions", () => {
  test("a write into a frozen version throws", async () => {
    const [{ document }] = await parsedSources();
    const session = createCanonicalSession(structuredClone(document)).unwrap();
    const version = session.document;
    deepFreeze(version);
    expect(() => {
      (version.package.document.content as unknown[]).push(undefined);
    }).toThrow(TypeError);
  });

  test("edits never write into a version handed out before them", async () => {
    const sources = await parsedSources();
    const law = (counts: { versions: number }) =>
      fc.property(
        fc.constantFrom(...sources),
        fc.array(editArbitrary, { minLength: 1, maxLength: 16 }),
        ({ document }, edits) => {
          const session = createCanonicalSession(
            structuredClone(document),
          ).unwrap();
          let state = EditorState.create({
            schema,
            doc: session.projection.doc,
          });
          deepFreeze(session.document);
          for (const edit of edits) {
            const handedOut = session.document;
            state = applyEdit(session, state, edit);
            if (session.document !== handedOut) counts.versions += 1;
            // The host now holds this version too.
            deepFreeze(session.document);
          }
        },
      );
    await assertProperty(law({ versions: 0 }), {
      id: "canonical edits leave handed-out versions unwritten",
      numRuns: 150,
    });
    // A fixed sample hands out more than one new version a case. Each case
    // replays up to 16 edits, so a third of the runs, at the same ratio.
    const RUNS = 50;
    const counts = { versions: 0 };
    checkCoverageSample(law(counts), RUNS);
    expect(counts.versions).toBeGreaterThan(RUNS);
  });
});
