/**
 * The indexed lookups do exactly what the ProseMirror methods they replace
 * do: `nodesBetween` visits the same nodes in the same order with the same
 * arguments (and honours a `false` return the same way), `nodeAt` returns the
 * same node, and the enclosing paragraph is the innermost paragraph
 * `doc.resolve` passes through.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import type { Node as PMNode } from "prosemirror-model";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { randomDocument } from "./__tests__/randomTransactions";
import { enclosingParagraphIndexed, nodeAtIndexed, nodesBetweenIndexed } from "./indexedNodeLookup";

setDefaultTimeout(propertyTestTimeout(30_000));

type Visit = [node: PMNode, pos: number, parent: PMNode | null, index: number];

const resolvedParagraph = (doc: PMNode, pos: number): { node: PMNode; pos: number } | null => {
  const $pos = doc.resolve(pos);
  for (let depth = $pos.depth; depth > 0; depth--) {
    const node = $pos.node(depth);
    if (node.type.name === "paragraph") {
      return { node, pos: $pos.before(depth) };
    }
  }
  return null;
};

describe("indexed node lookups", () => {
  test("nodesBetween visits what Node.nodesBetween visits", () => {
    fc.assert(
      fc.property(
        randomDocument,
        fc.nat({ max: 10_000 }),
        fc.nat({ max: 10_000 }),
        fc.constantFrom("paragraph", "blockquote", "table", null),
        (doc, left, right, prune) => {
          const size = doc.content.size;
          const from = Math.min(left % (size + 1), right % (size + 1));
          const to = Math.max(left % (size + 1), right % (size + 1));
          const collect = (walk: (visit: (...visit: Visit) => boolean) => void): Visit[] => {
            const visits: Visit[] = [];
            walk((node, pos, parent, index) => {
              visits.push([node, pos, parent, index]);
              return node.type.name !== prune;
            });
            return visits;
          };
          const expected = collect((visit) => doc.nodesBetween(from, to, visit));
          const actual = collect((visit) => nodesBetweenIndexed(doc, from, to, visit));
          expect(actual.length).toBe(expected.length);
          for (const [index, [node, pos, parent, childIndex]] of expected.entries()) {
            const [actualNode, actualPos, actualParent, actualIndex] = actual[index] ?? [];
            expect(actualNode).toBe(node);
            expect(actualPos).toBe(pos);
            expect(actualParent).toBe(parent);
            expect(actualIndex).toBe(childIndex);
          }
        },
      ),
      propertyConfig({ numRuns: 300 }),
    );
  });

  test("nodeAt and the enclosing paragraph agree with the tree at every position", () => {
    fc.assert(
      fc.property(randomDocument, (doc) => {
        for (let pos = 0; pos <= doc.content.size; pos++) {
          expect(nodeAtIndexed(doc, pos)).toBe(doc.nodeAt(pos));
          const expected = resolvedParagraph(doc, pos);
          const actual = enclosingParagraphIndexed(doc, pos);
          expect(actual?.node).toBe(expected?.node);
          expect(actual?.pos).toBe(expected?.pos);
        }
      }),
      propertyConfig({ numRuns: 150 }),
    );
  });
});
