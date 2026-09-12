import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig } from "../../../../test/property-testing";
import { alignFolioContentBlocks, type FolioContentAlignedBlockEvent } from "./content-alignment";
import type { FolioContentBlock, FolioContentParagraphKind } from "./content-types";

const tableLocation = {
  outerTableIndex: 0,
  tableIndex: 0,
  rowIndex: 0,
  cellIndex: 0,
  gridColumnIndex: 0,
  columnSpan: 1,
  rowSpan: 1,
  paragraphIndex: 0,
} as const;

const reconstruct = (
  events: readonly FolioContentAlignedBlockEvent[],
  side: "base" | "revised",
): FolioContentBlock[] =>
  events.flatMap((event) => {
    switch (event.type) {
      case "pair":
        return [side === "base" ? event.baseBlock : event.revisedBlock];
      case "baseOnly":
        return side === "base" ? [event.block] : [];
      case "revisedOnly":
        return side === "revised" ? [event.block] : [];
      default: {
        const unreachable: never = event;
        return unreachable;
      }
    }
  });

const PAIRING_PASSES = ["stable", "exact", "continuity", "positional"] as const;
type PairingPass = (typeof PAIRING_PASSES)[number];

const STRUCTURAL_BARRIERS = ["container", "body-table", "table-cell"] as const;
type StructuralBarrier = (typeof STRUCTURAL_BARRIERS)[number];

const PARAGRAPH_KINDS = [
  "paragraph",
  "heading",
  "listItem",
] as const satisfies readonly FolioContentParagraphKind[];
const DISTINCT_STRUCTURAL_KINDS = ["codeBlock", "figure"] as const;

const candidateForPass = (
  pass: PairingPass,
  text: string,
): { base: FolioContentBlock; revised: FolioContentBlock } => {
  switch (pass) {
    case "stable":
      return {
        base: { id: "shared", kind: "paragraph", text: `${text} base` },
        revised: { id: "shared", kind: "paragraph", text: `${text} revised` },
      };
    case "exact":
      return {
        base: { id: "base", idStability: "positional", kind: "paragraph", text },
        revised: { id: "revised", idStability: "positional", kind: "paragraph", text },
      };
    case "continuity":
      return {
        base: { id: "shared", idStability: "positional", kind: "paragraph", text: `${text} base` },
        revised: { id: "shared", kind: "paragraph", text: `${text} revised` },
      };
    case "positional":
      return {
        base: { id: "base", idStability: "positional", kind: "paragraph", text: `${text} base` },
        revised: {
          id: "revised",
          idStability: "positional",
          kind: "paragraph",
          text: `${text} revised`,
        },
      };
    default: {
      const unreachable: never = pass;
      return unreachable;
    }
  }
};

const applyBarrier = (
  barrier: StructuralBarrier,
  base: FolioContentBlock,
  revised: FolioContentBlock,
): void => {
  switch (barrier) {
    case "container":
      base.containerPath = [{ kind: "section", id: "base" }];
      revised.containerPath = [{ kind: "section", id: "revised" }];
      return;
    case "body-table":
      revised.table = tableLocation;
      return;
    case "table-cell":
      base.table = tableLocation;
      revised.table = { ...tableLocation, cellIndex: 1, gridColumnIndex: 1 };
      return;
    default: {
      const unreachable: never = barrier;
      return unreachable;
    }
  }
};

describe("block alignment ownership", () => {
  test("pairs every presentation-kind transition through every pairing pass", () => {
    for (const pass of PAIRING_PASSES) {
      for (const baseKind of PARAGRAPH_KINDS) {
        for (const revisedKind of PARAGRAPH_KINDS) {
          fc.assert(
            fc.property(fc.string({ minLength: 1, maxLength: 40 }), (text) => {
              const { base, revised } = candidateForPass(pass, text);
              base.kind = baseKind;
              revised.kind = revisedKind;

              const events = alignFolioContentBlocks([base], [revised], {
                stableIdMismatch: "pair",
              });
              expect(events.map(({ type }) => type)).toEqual(["pair"]);
              expect(reconstruct(events, "base")).toEqual([base]);
              expect(reconstruct(events, "revised")).toEqual([revised]);
            }),
            propertyConfig({ numRuns: 10 }),
          );
        }
      }
    }
  });

  test("keeps distinct structural kinds separate through every pairing pass", () => {
    for (const pass of PAIRING_PASSES) {
      for (const paragraphKind of PARAGRAPH_KINDS) {
        for (const structuralKind of DISTINCT_STRUCTURAL_KINDS) {
          for (const direction of ["fromParagraph", "toParagraph"] as const) {
            fc.assert(
              fc.property(fc.string({ minLength: 1, maxLength: 40 }), (text) => {
                const { base, revised } = candidateForPass(pass, text);
                base.kind = direction === "fromParagraph" ? paragraphKind : structuralKind;
                revised.kind = direction === "fromParagraph" ? structuralKind : paragraphKind;

                const events = alignFolioContentBlocks([base], [revised], {
                  stableIdMismatch: "pair",
                });
                expect(events.map(({ type }) => type)).toEqual(["baseOnly", "revisedOnly"]);
                expect(reconstruct(events, "base")).toEqual([base]);
                expect(reconstruct(events, "revised")).toEqual([revised]);
              }),
              propertyConfig({ numRuns: 10 }),
            );
          }
        }
      }
    }
  });

  test("applies every structural barrier to every pairing pass", () => {
    for (const pass of PAIRING_PASSES) {
      for (const barrier of STRUCTURAL_BARRIERS) {
        fc.assert(
          fc.property(fc.string({ minLength: 1, maxLength: 40 }), (text) => {
            const unblocked = candidateForPass(pass, text);
            expect(
              alignFolioContentBlocks([unblocked.base], [unblocked.revised], {
                stableIdMismatch: "pair",
              }).map(({ type }) => type),
            ).toEqual(["pair"]);

            const blocked = candidateForPass(pass, text);
            applyBarrier(barrier, blocked.base, blocked.revised);
            const events = alignFolioContentBlocks([blocked.base], [blocked.revised], {
              stableIdMismatch: "pair",
            });
            expect(events.map(({ type }) => type)).toEqual(["baseOnly", "revisedOnly"]);
            expect(reconstruct(events, "base")).toEqual([blocked.base]);
            expect(reconstruct(events, "revised")).toEqual([blocked.revised]);
          }),
          propertyConfig({ numRuns: 20 }),
        );
      }
    }
  });

  test("does not promote a stable ID duplicated on either side", () => {
    const base: FolioContentBlock[] = [
      { id: "duplicate", kind: "paragraph", text: "Alpha" },
      { id: "duplicate", kind: "paragraph", text: "Beta" },
    ];
    const revised: FolioContentBlock[] = [
      { id: "duplicate", kind: "paragraph", text: "Inserted" },
      { id: "duplicate", kind: "paragraph", text: "Beta revised" },
      { id: "duplicate", kind: "paragraph", text: "Alpha revised" },
    ];

    const events = alignFolioContentBlocks(base, revised, { stableIdMismatch: "pair" });
    expect(events.map(({ type }) => type)).toEqual(["pair", "pair", "revisedOnly"]);
    expect(reconstruct(events, "base")).toEqual(base);
    expect(reconstruct(events, "revised")).toEqual(revised);
  });

  test("keeps authored-ID policy independent from structural ownership", () => {
    fc.assert(
      fc.property(
        fc.string({ minLength: 1, maxLength: 50 }),
        fc.uuid(),
        fc.uuid(),
        (text, baseId, revisedId) => {
          fc.pre(baseId !== revisedId);
          const base: FolioContentBlock = { id: baseId, kind: "paragraph", text };
          const revised: FolioContentBlock = { id: revisedId, kind: "paragraph", text };

          expect(alignFolioContentBlocks([base], [revised]).map(({ type }) => type)).toEqual([
            "baseOnly",
            "revisedOnly",
          ]);
          expect(
            alignFolioContentBlocks([base], [revised], { stableIdMismatch: "pair" }).map(
              ({ type }) => type,
            ),
          ).toEqual(["pair"]);

          revised.kind = "figure";
          expect(
            alignFolioContentBlocks([base], [revised], { stableIdMismatch: "pair" }).map(
              ({ type }) => type,
            ),
          ).toEqual(["baseOnly", "revisedOnly"]);
        },
      ),
      propertyConfig({ numRuns: 160 }),
    );
  });
});
