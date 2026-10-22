import { expect, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { RUN_IDENTITY_MARK_NAME } from "./runIdentity";
import {
  continueSourceCarrierIdentitiesAcrossRevisions,
  SOURCE_CARRIER_MARK_NAMES,
} from "./runIdentityAcrossRevisions";
import { schema } from "./schema";

const CARRIERS = {
  [RUN_IDENTITY_MARK_NAME]: {
    id: "id",
    payload: { emptyFormatting: true },
    changed: {
      emptyFormatting: true,
      preservedAttributes: [{ name: "rsidR", value: "00ABCDEF" }],
    },
  },
  hyperlink: {
    id: "_docxHyperlinkIndex",
    payload: { href: "https://example.com/link", rId: "rId1" },
    changed: { href: "https://example.com/other", rId: "rId2" },
  },
} as const;

test("source-carrier continuation exercises every declared identity family", () => {
  expect(Object.keys(CARRIERS).toSorted()).toEqual(
    [...SOURCE_CARRIER_MARK_NAMES].toSorted(),
  );
});

test(
  "generated source-carrier identities continue only across matching revision cuts",
  () => {
    assertProperty(
      fc.property(
        fc.integer({ min: 0, max: 64 }),
        fc.constantFrom("boundary", "bridge"),
        (id, transition) => {
          const revision = schema.mark("deletion", {
            revisionId: 7,
            author: "Reviewer",
            provenance: "user",
          });
          for (const [name, carrier] of Object.entries(CARRIERS)) {
            const identity = (nextId: number, changed = false) =>
              schema.mark(name, {
                ...(changed ? carrier.changed : carrier.payload),
                [carrier.id]: nextId,
              });
            const left = schema.text("left", [identity(id)]);
            const right = schema.text("right", [identity(id + 1)]);
            const adjacent = [left, right];
            expect(
              continueSourceCarrierIdentitiesAcrossRevisions(adjacent),
            ).toBe(adjacent);
            const marked =
              transition === "boundary"
                ? [left, right.mark(revision.addToSet(right.marks))]
                : [left, schema.text("changed", [revision]), right];
            const continued =
              continueSourceCarrierIdentitiesAcrossRevisions(marked);
            expect(
              continued.at(-1)?.marks.find(({ type }) => type.name === name)
                ?.attrs[carrier.id],
            ).toBe(id);
            const different = schema.text("different", [
              identity(id + 1, true),
            ]);
            const split =
              transition === "boundary"
                ? [left, different.mark(revision.addToSet(different.marks))]
                : [left, schema.text("changed", [revision]), different];
            expect(continueSourceCarrierIdentitiesAcrossRevisions(split)).toBe(
              split,
            );
          }
        },
      ),
      {
        id: "generated source-carrier identities continue only across matching revision cuts",
        numRuns: 25,
        seed: 1683011,
      },
    );
  },
  propertyTestTimeout(10_000),
);
