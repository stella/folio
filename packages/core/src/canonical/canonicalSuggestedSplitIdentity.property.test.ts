/**
 * Block identity across suggested insertions and splits on the canonical
 * session (#3521): every block keeps one unique id while its suggestions are
 * pending, rejecting them all in any order gives every original paragraph
 * back under its own id, and accepting them all keeps the ids the pending
 * blocks showed, which are the ids the same edits make directly.
 */
import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { FolioDocxReviewer } from "../ai-edits/headless";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperationMode,
} from "../document-operations";
import { createDocx } from "../docx/rezip";
import type { Paragraph } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";

const PROPERTY_BUDGET = 90_000;

setDefaultTimeout(propertyTestTimeout(PROPERTY_BUDGET));

const ORIGINAL = [
  ["3CEFFC83", "This agreement is made between the parties."],
  ["31617F9A", "The Supplier delivers the goods."],
  ["5E88C024", "Signed."],
] as const;

const paragraph = ([paraId, text]: readonly [string, string]): Paragraph => ({
  type: "paragraph",
  paraId,
  textId: paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const open = async () => {
  const document = createEmptyDocument();
  document.package.document.content = ORIGINAL.map(paragraph);
  return FolioDocxReviewer.fromBuffer(await createDocx(document), {
    author: "Editor",
    session: "canonical",
  });
};

type Step =
  | { type: "insertAfterBlock" | "insertBeforeBlock"; at: number; text: string }
  | { type: "splitBlock"; at: number; fraction: number };

const stepArbitrary: fc.Arbitrary<Step> = fc.oneof(
  fc.record({
    type: fc.constantFrom(
      "insertAfterBlock" as const,
      "insertBeforeBlock" as const,
    ),
    at: fc.nat({ max: 7 }),
    text: fc.constantFrom("Added clause.", "Second added clause."),
  }),
  fc.record({
    type: fc.constant("splitBlock" as const),
    at: fc.nat({ max: 7 }),
    fraction: fc.double({ min: 0, max: 1, noNaN: true }),
  }),
);

type Row = { id: string; text: string };

const rows = (reviewer: FolioDocxReviewer): Row[] =>
  reviewer.getContent().map(({ id, text }) => ({ id, text }));

/** The saved model's paragraphs: what the package holds, suggestions rejected. */
const savedRows = (reviewer: FolioDocxReviewer): Row[] =>
  reviewer.toDocument().package.document.content.flatMap((block) =>
    block.type === "paragraph"
      ? [
          {
            id: block.paraId ?? "",
            text: block.content
              .flatMap((item) =>
                item.type === "run"
                  ? item.content.flatMap((part) =>
                      part.type === "text" ? [part.text] : [],
                    )
                  : [],
              )
              .join(""),
          },
        ]
      : [],
  );

const expectUniqueIds = (listed: readonly Row[]) => {
  const ids = listed.map(({ id }) => id);
  expect(ids.every((id) => id !== "")).toBe(true);
  expect(new Set(ids).size).toBe(ids.length);
};

/** Apply `steps` one batch each; the suggestion ids of the applied ones. */
const applySteps = (
  reviewer: FolioDocxReviewer,
  steps: readonly Step[],
  mode: FolioDocumentOperationMode,
): string[] => {
  const suggestions: string[] = [];
  for (const [index, step] of steps.entries()) {
    const listed = rows(reviewer);
    const target = listed[step.at % listed.length];
    if (target === undefined) throw new TypeError("A document has blocks.");
    const operation =
      step.type === "splitBlock"
        ? {
            id: `s${index}`,
            type: step.type,
            blockId: target.id,
            offset: Math.round(step.fraction * target.text.length),
          }
        : {
            id: `s${index}`,
            type: step.type,
            blockId: target.id,
            text: step.text,
          };
    const result = reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode,
        operations: [operation],
      },
      { undefinedReferences: "refuse" },
    );
    for (const { suggestionId } of result.applied)
      if (suggestionId !== undefined) suggestions.push(suggestionId);
    expectUniqueIds(rows(reviewer));
    expectUniqueIds(savedRows(reviewer));
  }
  return suggestions;
};

const original = ORIGINAL.map(([id, text]) => ({ id, text }));

test("suggested insertions and splits keep every block's id unique and stable", async () => {
  await assertProperty(
    fc.asyncProperty(
      fc.array(stepArbitrary, { minLength: 1, maxLength: 5 }),
      fc.nat(),
      async (steps, order) => {
        const suggested = await open();
        const suggestions = applySteps(suggested, steps, "suggested");
        // Pending suggestions stay out of the package.
        expect(savedRows(suggested)).toEqual(original);
        const pending = rows(suggested);

        const rejected = await open();
        applySteps(rejected, steps, "suggested");
        // One at a time, in a generated order, as Reject All does at once.
        const remaining = [...suggestions];
        let seed = order;
        while (remaining.length > 0) {
          const [id] = remaining.splice(seed % remaining.length, 1);
          seed = Math.floor(seed / 2) + 1;
          if (id !== undefined) rejected.rejectSuggestion(id);
          expectUniqueIds(rows(rejected));
        }
        expect(rows(rejected)).toEqual(original);
        expect(savedRows(rejected)).toEqual(original);

        // Adopted as tracked changes, then accepted.
        for (const id of suggestions) suggested.acceptSuggestion(id);
        suggested.acceptAll();
        expect(rows(suggested)).toEqual(pending);
        expect(savedRows(suggested)).toEqual(pending);

        const direct = await open();
        applySteps(direct, steps, "direct");
        expect(rows(direct)).toEqual(pending);
      },
    ),
    {
      id: "suggested insertions and splits keep every block's id unique and stable",
      numRuns: 30,
      examples: [
        // #3521: a suggested split of a suggested inserted paragraph.
        [
          [
            { type: "insertBeforeBlock", at: 0, text: "Second added clause." },
            { type: "splitBlock", at: 0, fraction: 0.3 },
          ],
          0,
        ],
      ],
    },
  );
});
