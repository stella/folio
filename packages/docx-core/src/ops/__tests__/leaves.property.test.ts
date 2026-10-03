import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { assertProperty, propertyTestTimeout } from "../../../../../test/property-testing";
import type { ParagraphContent } from "../../model/document";
import { mergeLists, type InlineNode } from "../leaves";

setDefaultTimeout(propertyTestTimeout(30_000));

// Resolving an inner wrapper can shorten the surviving branch below its
// recorded cut depth. The depth bounds an alike merge; exact cuts stay exact.
test("alike merges consume surviving text within the maximum source depth", () => {
  assertProperty(
    fc.property(fc.integer({ min: 0, max: 3 }), fc.integer({ min: 1, max: 5 }), (levels, extra) => {
      const branch = (text: string): InlineNode => {
        let node: ParagraphContent = { type: "run", content: [{ type: "text", text }] };
        for (let level = 0; level < levels; level += 1)
          node = { type: "inlineWrapper", kind: "bidi", control: "embedding", content: [node] };
        return node;
      };
      const depth = levels + 2;
      expect(
        mergeLists([branch("a")], [branch("b")], depth + extra, { mode: "asFarAsAlike" }),
      ).toStrictEqual([branch("ab")]);
      expect(mergeLists([branch("a")], [branch("b")], depth + extra)).toBeUndefined();
      expect(mergeLists([branch("a")], [branch("b")], depth)).toStrictEqual([branch("ab")]);
      expect(mergeLists([branch("a")], [branch("b")], 0)).toStrictEqual([branch("a"), branch("b")]);
    }),
    { numRuns: 50 },
  );
});

test("open continuations preserve the source wrapper's cut facts", () => {
  assertProperty(
    fc.property(
      fc.integer({ min: 0, max: 5 }),
      fc.constantFrom("first", "second"),
      (depth, identity) => {
        const joins = { before: depth, after: depth, remove: depth };
        const source = {
          type: "insertion",
          info: { id: 7, author: "Editor" },
          resolutionJoins: joins,
          content: [{ type: "run", content: [{ type: "text", text: "a" }] }],
        } satisfies InlineNode;
        const incoming = {
          type: "insertion",
          info: { id: 8, author: "Editor" },
          content: [{ type: "run", content: [{ type: "text", text: "b" }] }],
        } satisfies InlineNode;
        const left = identity === "first" ? source : incoming;
        const right = identity === "first" ? incoming : source;
        expect(
          mergeLists([left], [right], 3, { mode: "exact", fields: "source", identity }),
        ).toStrictEqual([
          {
            ...source,
            content: [
              {
                type: "run",
                content: [{ type: "text", text: identity === "first" ? "ab" : "ba" }],
              },
            ],
          },
        ]);
        expect(mergeLists([left], [right], 3)).toBeUndefined();
        expect(
          mergeLists([source], [{ ...incoming, info: { id: 8, author: "Other" } }], 3, {
            mode: "exact",
            fields: "source",
          }),
        ).toBeUndefined();
      },
    ),
    { numRuns: 50 },
  );
});
