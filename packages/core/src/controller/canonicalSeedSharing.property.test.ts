/**
 * A canonical session opens the host's document as it is: it takes no copy
 * of it and never writes into it.
 *
 * The host's model is read-only, shared with its history and its layout, and
 * every step of a seed and of an edit makes a new version sharing what it
 * left as it was, so the session needs no copy of its own. Opening takes no
 * whole-document copy and keeps the blocks it left as the host's own objects;
 * random edits over a deep-frozen seed never throw (a write into the host's
 * document would) and end where the same edits over an unfrozen seed end.
 */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { deepFreeze } from "../__tests__/deepFreeze";
import { cloneDocumentWithParagraphPropertySources } from "../docx/paragraphPropertySource";
import type { LayoutInstrumentation } from "../layout-engine/layoutInstrumentation";
import { schema } from "../prosemirror/schema";
import type { Document } from "../types/document";
import {
  applyEdit,
  editArbitrary,
  parsedSources,
} from "./__tests__/canonicalEditRuns";
import {
  type CanonicalSession,
  createCanonicalSession,
} from "./canonicalSession";

setDefaultTimeout(propertyTestTimeout(240_000));

/** Whole-document copies taken while `run` runs. */
const copiesDuring = <Value>(
  run: () => Value,
): { value: Value; copies: number } => {
  const previous: LayoutInstrumentation | undefined =
    globalThis.__folioLayoutInstrumentation;
  let copies = 0;
  globalThis.__folioLayoutInstrumentation = {
    ...previous,
    onDocumentBuilt: (event) => {
      previous?.onDocumentBuilt?.(event);
      if (event.kind === "copy") copies += 1;
    },
  };
  try {
    return { value: run(), copies };
  } finally {
    globalThis.__folioLayoutInstrumentation = previous;
  }
};

/** A document as text, maps and sets as their entries, for comparing two runs. */
const serialized = (document: unknown): string =>
  JSON.stringify(document, (_key, value: unknown) =>
    value instanceof Map || value instanceof Set ? [...value] : value,
  );

const seeded = (document: Document): CanonicalSession =>
  createCanonicalSession(document).unwrap();

describe("canonical seed sharing", () => {
  test("opening a session takes no copy of the document and keeps its blocks", async () => {
    const sources = await parsedSources();
    expect(sources.length).toBeGreaterThan(0);
    for (const { id, document } of sources) {
      const opened = cloneDocumentWithParagraphPropertySources(document);
      const { value: session, copies } = copiesDuring(() => seeded(opened));
      expect({ id, copies }).toEqual({ id, copies: 0 });
      const blocks = new Set<object>(opened.package.document.content);
      const kept = session.document.package.document.content.filter((block) =>
        blocks.has(block),
      );
      expect({ id, kept: kept.length > 0 }).toEqual({ id, kept: true });
    }
  });

  test("edits over a frozen seed never write into it and end as over an unfrozen one", async () => {
    const sources = await parsedSources();
    let edited = 0;
    await assertProperty(
      fc.property(
        fc.constantFrom(...sources),
        fc.array(editArbitrary, { minLength: 1, maxLength: 12 }),
        ({ document }, edits) => {
          const frozen = seeded(
            deepFreeze(cloneDocumentWithParagraphPropertySources(document)),
          );
          const plain = seeded(
            cloneDocumentWithParagraphPropertySources(document),
          );
          let frozenState = EditorState.create({
            schema,
            doc: frozen.projection.doc,
          });
          let plainState = EditorState.create({
            schema,
            doc: plain.projection.doc,
          });
          for (const edit of edits) {
            frozenState = applyEdit(frozen, frozenState, edit);
            plainState = applyEdit(plain, plainState, edit);
          }
          if (frozen.version > 0) edited += 1;
          expect(frozenState.doc.eq(plainState.doc)).toBe(true);
          expect(frozen.version).toBe(plain.version);
          expect(serialized(frozen.document)).toBe(serialized(plain.document));
        },
      ),
      {
        id: "canonical edits leave the seeded host document unwritten",
        numRuns: 80,
      },
    );
    expect(edited).toBeGreaterThan(40);
  });
});
