import { expect, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { paragraph } from "../docx/server/build";
import {
  JOINED_PARAGRAPH_ATTR_NAMES,
  joinedParagraphAttrs,
} from "../prosemirror/paragraphMarkJoin";
import { schema } from "../prosemirror/schema";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";

const SOURCE = [
  ["10000001", "Opening clause."],
  ["10000002", "Signed in two copies."],
  ["10000003", "Closing clause."],
] as const;

const source = () => {
  const document = createEmptyDocument();
  document.package.document.content = SOURCE.map(([paraId, text]) => ({
    ...paragraph(text),
    paraId,
  }));
  return createDocx(document);
};

/** Each body paragraph's words, as a saved package reads. */
const textsOf = async (bytes: ArrayBuffer) =>
  (
    await parseDocx(bytes.slice(0), {
      preloadFonts: false,
      detectVariables: false,
    })
  ).package.document.content.map((block) =>
    block.type === "paragraph"
      ? block.content
          .flatMap((item) => (item.type === "run" ? item.content : []))
          .map((leaf) => (leaf.type === "text" ? leaf.text : ""))
          .join("")
      : block.type,
  );

const SOURCE_TEXTS = SOURCE.map(([, text]) => text);

const apply = (
  reviewer: FolioDocxReviewer,
  mode: "tracked-changes" | "suggested",
  operations: FolioDocumentOperation[],
) =>
  reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode,
      operations,
    },
    { undefinedReferences: "refuse" },
  );

// A suggested paragraph inserted after the first half of a tracked split.
// Rejecting the split joins that half into the suggested paragraph, which must
// not make the source words part of the suggestion.
test("rejecting a split beside a suggested paragraph keeps the split's words", async () => {
  const reviewer = await FolioDocxReviewer.fromBuffer(await source());
  expect(
    apply(reviewer, "tracked-changes", [
      { id: "split", type: "splitBlock", blockId: "10000002", offset: 14 },
    ]).skipped,
  ).toEqual([]);
  expect(
    apply(reviewer, "suggested", [
      {
        id: "proposal",
        type: "insertAfterBlock",
        blockId: "10000002",
        text: "Payment is due in thirty days.",
      },
    ]).skipped,
  ).toEqual([]);
  const split = reviewer
    .getChanges()
    .find(({ type }) => type === "paragraphMarkInserted");
  expect(split).toBeDefined();
  expect(split !== undefined && reviewer.rejectChange(split)).toBe(true);
  // The proposal is held back from the save; the source words are not.
  expect(await textsOf(await reviewer.toBuffer())).toEqual(SOURCE_TEXTS);
  expect(reviewer.rejectSuggestion("proposal")).toBe(true);
  expect(await textsOf(await reviewer.toBuffer())).toEqual(SOURCE_TEXTS);
});

// A suggested paragraph whose own mark another revision holds (the break of a
// paragraph inserted after it, or a deleted break it took over), or that ends
// the body, which carries no added break. Its break is then the free mark
// before it, so rejecting the suggestion, before or after the other revision,
// must take that break out and leave the other one.
type Batch = {
  mode: "tracked-changes" | "suggested";
  operation: FolioDocumentOperation;
};
const SLOT_TAKEN: readonly (readonly Batch[])[] = [
  [
    {
      mode: "suggested",
      operation: {
        id: "op-0",
        type: "insertAfterBlock",
        blockId: "10000002",
        text: "New term",
      },
    },
    {
      mode: "tracked-changes",
      operation: { id: "op-1", type: "deleteBlock", blockId: "10000003" },
    },
    {
      mode: "tracked-changes",
      operation: {
        id: "op-2",
        type: "insertAfterBlock",
        blockId: "10000002",
        text: "Payment due.",
      },
    },
  ],
  [
    {
      mode: "tracked-changes",
      operation: { id: "op-0", type: "deleteBlock", blockId: "10000003" },
    },
    {
      mode: "suggested",
      operation: {
        id: "op-1",
        type: "insertAfterBlock",
        blockId: "10000002",
        text: "New term",
      },
    },
    {
      mode: "tracked-changes",
      operation: {
        id: "op-2",
        type: "insertBeforeBlock",
        blockId: "10000003",
        text: "Payment due.",
      },
    },
  ],
  [
    {
      mode: "suggested",
      operation: {
        id: "op-0",
        type: "insertAfterBlock",
        blockId: "10000003",
        text: "New term",
      },
    },
    {
      mode: "tracked-changes",
      operation: {
        id: "op-1",
        type: "insertAfterBlock",
        blockId: "10000003",
        text: "Payment due.",
      },
    },
  ],
];

const applyBatches = async (bytes: ArrayBuffer, batches: readonly Batch[]) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(bytes.slice(0));
  for (const { mode, operation } of batches)
    expect(apply(reviewer, mode, [operation]).skipped).toEqual([]);
  return reviewer;
};

test("a suggested paragraph whose mark another revision holds leaves with its suggestion, whichever is rejected first", async () => {
  const bytes = await source();
  for (const batches of SLOT_TAKEN) {
    const suggestions = batches
      .filter(({ mode }) => mode === "suggested")
      .map(({ operation }) => operation.id);
    const changes = (await applyBatches(bytes, batches))
      .getChanges()
      .map(({ id }) => id);
    for (const first of [null, ...changes]) {
      const reviewer = await applyBatches(bytes, batches);
      let rejected = true;
      if (first === null) {
        for (const id of suggestions)
          rejected = reviewer.rejectSuggestion(id) && rejected;
      } else {
        rejected = reviewer.rejectChange(first);
        for (const id of suggestions) reviewer.rejectSuggestion(id);
      }
      for (const change of reviewer.getChanges().toReversed())
        reviewer.rejectChange(change);
      reviewer.rejectAll();
      expect({
        first,
        rejected,
        texts: await textsOf(await reviewer.toBuffer()),
      }).toEqual({ first, rejected: true, texts: SOURCE_TEXTS });
    }
  }
});

// Mostly one paragraph, so the batches meet at its marks.
const blockId = fc.oneof(
  { weight: 3, arbitrary: fc.constant("10000002") },
  { weight: 1, arbitrary: fc.constantFrom("10000001", "10000003") },
);
const words = fc.constantFrom("Payment due.", "New term", "Amended text.");
const operation: fc.Arbitrary<FolioDocumentOperation> = fc.oneof(
  {
    weight: 3,
    arbitrary: fc.record({
      type: fc.constant("splitBlock" as const),
      blockId,
      offset: fc.integer({ min: 1, max: 12 }),
    }),
  },
  {
    weight: 3,
    arbitrary: fc.record({
      type: fc.constantFrom(
        "insertAfterBlock" as const,
        "insertBeforeBlock" as const,
      ),
      blockId,
      text: words,
    }),
  },
  {
    weight: 1,
    arbitrary: fc.record({
      type: fc.constant("replaceBlock" as const),
      blockId,
      text: words,
    }),
  },
  {
    weight: 1,
    arbitrary: fc.record({
      type: fc.constant("deleteBlock" as const),
      blockId,
    }),
  },
) as fc.Arbitrary<FolioDocumentOperation>;

test(
  "rejecting any mix of tracked and suggested changes, in any order, gives the source back",
  async () => {
    const bytes = await source();
    await assertProperty(
      fc.asyncProperty(
        fc.array(
          fc.record({
            mode: fc.constantFrom(
              "tracked-changes" as const,
              "suggested" as const,
            ),
            operation,
          }),
          { minLength: 1, maxLength: 5 },
        ),
        fc.array(fc.nat(), { minLength: 12, maxLength: 12 }),
        async (batches, picks) => {
          const reviewer = await FolioDocxReviewer.fromBuffer(bytes.slice(0));
          const suggestions: string[] = [];
          for (const [index, { mode, operation: op }] of batches.entries()) {
            const id = `op-${String(index)}`;
            const result = apply(reviewer, mode, [{ ...op, id }]);
            if (mode === "suggested" && result.applied.length > 0)
              suggestions.push(id);
          }
          // One change or one suggestion at a time, in the order drawn.
          for (const pick of picks) {
            const changes = reviewer.getChanges();
            if (changes.length === 0 && suggestions.length === 0) break;
            if (pick % 2 === 0 && suggestions.length > 0) {
              const [id] = suggestions.splice(pick % suggestions.length, 1);
              if (id !== undefined) reviewer.rejectSuggestion(id);
              continue;
            }
            const change = changes[pick % Math.max(changes.length, 1)];
            if (change !== undefined) reviewer.rejectChange(change);
          }
          for (const id of suggestions) reviewer.rejectSuggestion(id);
          reviewer.rejectAll();
          expect(await textsOf(await reviewer.toBuffer())).toEqual(
            SOURCE_TEXTS,
          );
        },
      ),
      { id: "rejected-mark-beside-suggestion-mixed-order", numRuns: 300 },
    );
  },
  propertyTestTimeout(120_000),
);

// Every paragraph attribute has a stated join rule, so a new one cannot reach
// a join without deciding which side it comes from.
test("a join decides every paragraph attribute", () => {
  expect(
    Object.keys(schema.nodes["paragraph"]?.spec.attrs ?? {}).toSorted(),
  ).toEqual([...JOINED_PARAGRAPH_ATTR_NAMES].toSorted());
});

test("a join keeps both paragraphs' boundary records and turns a suggested paragraph into its mark", () => {
  const paragraphType = schema.nodes["paragraph"];
  if (!paragraphType) throw new Error("No paragraph node");
  const first = paragraphType.create(
    {
      paraId: "10000001",
      bookmarks: [{ id: 1, name: "First" }],
      _emptyHyperlinks: [{ offset: 0, href: "https://example.com/a" }],
    },
    schema.text("ab"),
  );
  const suggested = {
    suggestionId: "proposal",
    revisionId: 7,
    author: "Reviewer",
    date: "2026-01-01T00:00:00Z",
  };
  const next = paragraphType.create(
    {
      paraId: "10000002",
      bookmarks: [{ id: 2, name: "Second" }],
      _emptyHyperlinks: [{ offset: 1, href: "https://example.com/b" }],
      _suggestedInsert: suggested,
    },
    schema.text("cd"),
  );
  const joined = joinedParagraphAttrs(first, next, 2);
  expect(joined["paraId"]).toBe("10000002");
  expect(joined["bookmarks"]).toEqual([
    { id: 1, name: "First" },
    { id: 2, name: "Second" },
  ]);
  expect(joined["_emptyHyperlinks"]).toEqual([
    { offset: 0, href: "https://example.com/a" },
    { offset: 5, href: "https://example.com/b" },
  ]);
  expect(joined["_suggestedInsert"]).toBeNull();
  expect(joined["pPrMark"]).toEqual({
    kind: "ins",
    info: {
      id: 7,
      author: "Reviewer",
      date: "2026-01-01T00:00:00Z",
      provenance: "suggested",
      suggestionId: "proposal",
    },
  });
  // Both paragraphs one suggestion's: the joined one still is.
  const both = joinedParagraphAttrs(
    paragraphType.create({ _suggestedInsert: suggested }, schema.text("ab")),
    next,
  );
  expect(both["_suggestedInsert"]).toEqual(suggested);
  expect(both["pPrMark"]).toBeNull();
});
