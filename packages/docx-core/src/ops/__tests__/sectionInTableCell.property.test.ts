/** No generated document, and no operation on one, ends a section in a table cell. */

import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { assertExactModel } from "../../../../../test/exactModel";
import { checkCoverageSample } from "../../../../../test/property-coverage";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type { Document } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { storyParagraphs } from "../blocks";
import { hasSectionInTableCell } from "../cellSections";
import { IDENTITY_SPACES, identityKeysIn, packageIdentityKeys } from "../ids";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  REVISION_DECISIONS,
} from "../types";
import {
  documentArbitrary,
  opFor,
  opSeedArbitrary,
  reviewDocumentArbitrary,
  trackedOpFor,
} from "./documentArbitraries";

setDefaultTimeout(propertyTestTimeout(60_000));

const NUM_RUNS = 1_000;

test("no operation on a generated document ends a section in a table cell", () => {
  assertProperty(
    fc.property(
      fc.oneof(documentArbitrary, reviewDocumentArbitrary),
      opSeedArbitrary,
      (document, seed) => {
        expect(hasSectionInTableCell(document)).toBe(false);
        for (const op of [
          opFor(document, seed),
          trackedOpFor(document, seed),
        ]) {
          const applied = applyDocumentOp(document, op);
          if (applied.isOk())
            expect(hasSectionInTableCell(applied.value.document)).toBe(false);
        }
      },
    ),
    {
      id: "no operation on a generated document ends a section in a table cell",
      numRuns: NUM_RUNS,
    },
  );
});

/** The paragraphs in a table cell, at any depth. */
const cellParagraphs = (document: Document) =>
  storyParagraphs(document.package.document).filter(({ list }) =>
    list.some(({ kind }) => kind === "tableCell"),
  );

const sectionsInCells = (document: Document): (string | undefined)[] =>
  cellParagraphs(document)
    .filter(({ paragraph }) => paragraph.sectionProperties !== undefined)
    .map(({ paragraph }) => paragraph.paraId);

const revisionIdsIn = (document: Document): number[] => {
  const prefix = `${IDENTITY_SPACES.REVISION}:`;
  return packageIdentityKeys(document.package)
    .filter((key) => key.startsWith(prefix))
    .map((key) => Number(key.slice(prefix.length)));
};

type Outcome = "kept" | "removed" | "refused";

/** Operations on a generated document one of whose cell paragraphs arrived ending a section. */
const arrivedSections = (tally: Map<Outcome, number>) =>
  fc.property(
    fc
      .oneof(documentArbitrary, reviewDocumentArbitrary)
      .filter((document) => cellParagraphs(document).length > 0),
    fc.nat(),
    fc.nat(),
    opSeedArbitrary,
    (generated, at, kind, seed) => {
      const document = structuredClone(generated);
      const cells = cellParagraphs(document);
      const arrived = cells[at % cells.length]?.paragraph;
      const count = (outcome: Outcome) =>
        tally.set(outcome, (tally.get(outcome) ?? 0) + 1);
      if (arrived === undefined) return;
      arrived.sectionProperties = { pageWidth: 11_906 };
      // The section view may hold its own copy of the body's blocks.
      Reflect.deleteProperty(document.package.document, "sections");
      const ids = identityKeysIn(arrived).length;
      const ops: DocumentOp[] = [
        opFor(document, seed),
        trackedOpFor(document, seed),
        {
          type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
          story: OP_STORIES.MAIN,
          revisionIds: revisionIdsIn(document),
          decision:
            (kind + ids) % 2 === 0
              ? REVISION_DECISIONS.ACCEPT
              : REVISION_DECISIONS.REJECT,
        },
      ];
      const op = ops[kind % ops.length];
      if (op === undefined) return;
      const before = structuredClone(document);
      const applied = applyDocumentOp(document, op);
      assertExactModel(document, before);
      // The same content as a fresh value has the same answer.
      const again = applyDocumentOp(structuredClone(before), op);
      expect(again.isOk()).toBe(applied.isOk());
      if (applied.isErr()) {
        if (again.isErr())
          expect(again.error.reason).toBe(applied.error.reason);
        count("refused");
        return;
      }
      const after = applied.value.document;
      const stood = sectionsInCells(before);
      const stands = sectionsInCells(after);
      for (const id of stands) expect(stood).toContain(id);
      const inverse: DocumentOp[] = JSON.parse(
        JSON.stringify(applied.value.inverse),
      );
      const undone = applyDocumentOps(after, inverse);
      if (stands.length === stood.length) {
        assertExactModel(undone.unwrap().document, before);
        count("kept");
        return;
      }
      expect(undone.isErr() ? undone.error.reason : "applied").toBe(
        DOCUMENT_OP_REFUSAL_REASONS.SECTION_BOUNDARY,
      );
      count("removed");
    },
  );

test("a section a cell arrived with only stays or goes", () => {
  assertProperty(arrivedSections(new Map()), {
    id: "a section a cell arrived with only stays or goes",
    numRuns: NUM_RUNS,
  });
});

test("the cases keep, remove and refuse", () => {
  const tally = new Map<Outcome, number>();
  checkCoverageSample(arrivedSections(tally), 400);
  expect({
    kept: (tally.get("kept") ?? 0) >= 100,
    removed: (tally.get("removed") ?? 0) >= 10,
    refused: (tally.get("refused") ?? 0) >= 20,
    tally: Object.fromEntries(tally),
  }).toMatchObject({ kept: true, removed: true, refused: true });
});
