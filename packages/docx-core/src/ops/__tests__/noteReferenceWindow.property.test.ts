/**
 * An edit whose changed blocks carry the same note references as the blocks
 * they replace leaves the document referencing exactly the notes it did.
 * Runs of generated operations over generated documents with notes compare
 * that answer, and the note stories an edit removes or touches, with every
 * paragraph's references read again.
 */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type { Document } from "../../model/document";
import { applyDocumentOp } from "../apply";
import {
  noteStoriesTouchedByReferences,
  referencedNoteKeysInFull,
  removedNoteStories,
  sameNoteReferences,
} from "../noteMarks";
import {
  documentArbitrary,
  opFor,
  opSeedArbitrary,
} from "./documentArbitraries";

setDefaultTimeout(propertyTestTimeout(240_000));

const keyOf = ({ kind, id }: { kind: string; id: number }) => `${kind}:${id}`;

describe("an edit's note references read from its changed blocks", () => {
  test("equal the references read from every paragraph", () => {
    let same = 0;
    let changed = 0;
    let referencing = 0;
    assertProperty(
      fc.property(
        documentArbitrary,
        fc.array(opSeedArbitrary, { minLength: 1, maxLength: 12 }),
        (initial, seeds) => {
          let document = initial;
          for (const seed of seeds) {
            const applied = applyDocumentOp(document, opFor(document, seed));
            if (applied.isErr()) continue;
            const after = applied.value.document;
            const before = [...referencedNoteKeysInFull(document)].sort();
            const now = [...referencedNoteKeysInFull(after)].sort();
            if (before.length > 0) referencing += 1;
            if (sameNoteReferences(document, after)) {
              same += 1;
              expect(now).toEqual(before);
            } else changed += 1;
            expect(
              removedNoteStories(document, after).map(keyOf).sort(),
            ).toEqual(before.filter((key) => !now.includes(key)));
            expect(
              noteStoriesTouchedByReferences(document, after).map(keyOf).sort(),
            ).toEqual(
              [...new Set([...before, ...now])]
                .filter((key) => before.includes(key) !== now.includes(key))
                .sort(),
            );
            document = after;
          }
        },
      ),
      {
        id: "note references read from changed blocks equal every paragraph",
        numRuns: 300,
      },
    );
    // Both answers, over documents that reference notes.
    expect(same).toBeGreaterThan(300);
    expect(changed).toBeGreaterThan(20);
    expect(referencing).toBeGreaterThan(100);
  });

  test("tell a renumbered reference from the one it replaces", () => {
    let renumbered = 0;
    assertProperty(
      fc.property(documentArbitrary, fc.nat(), (document, choice) => {
        const content = document.package.document.content;
        const at = content.findIndex(
          (block, index) =>
            index >= choice % Math.max(1, content.length) &&
            block.type === "paragraph" &&
            /"(?:footnoteRef|endnoteRef)"/u.test(JSON.stringify(block)),
        );
        const block = content[at];
        if (block === undefined) return;
        // The same paragraph referencing another note in place of its own.
        const replaced: unknown = JSON.parse(
          JSON.stringify(block),
          (key, value) =>
            key === "id" && typeof value === "number" ? value + 1000 : value,
        );
        const after: Document = {
          ...document,
          package: {
            ...document.package,
            document: {
              ...document.package.document,
              content: content.map((existing, index) =>
                index === at ? (replaced as typeof existing) : existing,
              ),
            },
          },
        };
        const before = [...referencedNoteKeysInFull(document)].sort();
        const now = [...referencedNoteKeysInFull(after)].sort();
        if (now.join() === before.join()) return;
        renumbered += 1;
        expect(sameNoteReferences(document, after)).toBe(false);
        expect(removedNoteStories(document, after).map(keyOf).sort()).toEqual(
          before.filter((key) => !now.includes(key)),
        );
      }),
      {
        id: "note references read from changed blocks tell renumbered references",
        numRuns: 300,
      },
    );
    expect(renumbered).toBeGreaterThan(30);
  });
});
