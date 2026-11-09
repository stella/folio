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

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { schema } from "../prosemirror/schema";
import {
  applyEdit,
  editArbitrary,
  parsedSources,
} from "./__tests__/canonicalEditRuns";
import { createCanonicalSession } from "./canonicalSession";

setDefaultTimeout(propertyTestTimeout(240_000));

const refuseWrite = () => {
  throw new TypeError("A handed-out document version was written to.");
};

/**
 * Freeze everything reachable from `value`. Maps and sets cannot be frozen,
 * so their writing methods throw instead.
 */
const deepFreeze = (value: unknown, seen = new Set<unknown>()): void => {
  if (
    typeof value !== "object" ||
    value === null ||
    seen.has(value) ||
    // Versions share what an edit left: a part frozen with an earlier one.
    Object.isFrozen(value)
  )
    return;
  seen.add(value);
  if (ArrayBuffer.isView(value) || value instanceof ArrayBuffer) return;
  if (value instanceof Map) {
    for (const [key, entry] of value) {
      deepFreeze(key, seen);
      deepFreeze(entry, seen);
    }
    Object.assign(value, {
      set: refuseWrite,
      delete: refuseWrite,
      clear: refuseWrite,
    });
  } else if (value instanceof Set) {
    for (const entry of value) deepFreeze(entry, seen);
    Object.assign(value, {
      add: refuseWrite,
      delete: refuseWrite,
      clear: refuseWrite,
    });
  } else {
    for (const key of Reflect.ownKeys(value))
      deepFreeze((value as Record<PropertyKey, unknown>)[key], seen);
  }
  Object.freeze(value);
};

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
    let versions = 0;
    await assertProperty(
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
            if (session.document !== handedOut) versions += 1;
            // The host now holds this version too.
            deepFreeze(session.document);
          }
        },
      ),
      {
        id: "canonical edits leave handed-out versions unwritten",
        numRuns: 150,
      },
    );
    expect(versions).toBeGreaterThan(150);
  });
});
