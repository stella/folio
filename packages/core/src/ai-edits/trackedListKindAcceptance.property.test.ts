/**
 * A suggested batch of list edits on a numbered document, accepted, gives
 * every block the kind, style and list level the same batch gives it applied
 * directly to the accepted document: the block a paragraph's text ends up in
 * reads as the edit asked, whatever review history its marks carry.
 *
 * Other authors' earlier suggestions leave inserted paragraph marks behind
 * (a split, an insertion), and a later merge or deletion deletes one: the
 * mark is then added and deleted at once, and accepting removes it as it
 * removes any deleted mark. Read as a mark that stays, a paragraph put before
 * one such mark copied that paragraph instead of the joined one the merge
 * leaves, a merge chain through it gave the joined paragraph the last part's
 * properties, and a paragraph put before the join's last part landed inside
 * the join.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import {
  paragraphNumberingFromSlots,
  paragraphNumberingReference,
} from "@stll/docx-core/model";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown";
import { FolioDocxReviewer } from "./headless";
import type {
  FolioAIBlockParagraphProperties,
  FolioAIEditOperation,
  FolioAIListNumbering,
  FolioAIParagraphBlock,
} from "./types";

type NumberingLevel = NonNullable<
  ReturnType<typeof fromMarkdown>["package"]["numbering"]
>["abstractNums"][number]["levels"][number];

/** The instance `Heading 2` numbers through its style. */
const HEADING_LIST = 5;

const level = (ilvl: number, lvlText: string): NumberingLevel => ({
  ilvl,
  start: 1,
  numFmt: "decimal",
  lvlText,
  suffix: "space",
  pPr: { indentLeft: 0, indentFirstLine: 0 },
});

/**
 * Numbered headings (through `Heading 2`'s style), a directly numbered list
 * and plain paragraphs between them.
 */
const numberedDocx = async (): Promise<ArrayBuffer> => {
  const document = fromMarkdown(
    [
      "# Agreement",
      "## Scope",
      "The Supplier delivers the goods.",
      "1. Deposit on signature\n2. Balance on delivery\n3. Retention after inspection",
      "## Payment",
      "Payment is due in ten days.",
      "Closing paragraph.",
    ].join("\n\n"),
  );
  const numbering = document.package.numbering;
  if (numbering === undefined) throw new Error("The list has no numbering");
  document.package.numbering = {
    ...numbering,
    abstractNums: [
      ...numbering.abstractNums,
      {
        abstractNumId: HEADING_LIST,
        multiLevelType: "multilevel",
        levels: [level(0, "%1."), level(1, "%1.%2.")],
      },
    ],
    nums: [
      ...numbering.nums,
      { numId: HEADING_LIST, abstractNumId: HEADING_LIST },
    ],
  };
  const heading = document.package.styles?.styles.find(
    (style) => style.styleId === "Heading2",
  );
  if (heading === undefined) throw new Error("Heading2 is missing");
  heading.pPr = {
    ...heading.pPr,
    numPr: paragraphNumberingFromSlots({ numId: HEADING_LIST, ilvl: 0 }),
  };
  return createDocx(document);
};

type Reviewer = Awaited<ReturnType<typeof FolioDocxReviewer.fromBuffer>>;

const open = (bytes: ArrayBuffer, author: string): Promise<Reviewer> =>
  FolioDocxReviewer.fromBuffer(bytes.slice(0), {
    author,
    session: "canonical",
  });

/** The numbering an edit states, named against the document's lists. */
type NumberingChoice = "none" | "item" | "nested" | "newList" | "heading";

const numberingOf = (
  choice: NumberingChoice,
  listId: number,
): FolioAIListNumbering => {
  switch (choice) {
    case "none":
      return { kind: "none" };
    case "item":
      return paragraphNumberingReference({ numId: listId, ilvl: 0 });
    case "nested":
      return paragraphNumberingReference({ numId: listId, ilvl: 1 });
    case "newList":
      return { kind: "newList", format: "numbered" };
    case "heading":
      return paragraphNumberingReference({ numId: HEADING_LIST, ilvl: 0 });
    default: {
      const unreachable: never = choice;
      return unreachable;
    }
  }
};

/** Paragraph properties an edit asks for: a style, numbering, or both. */
type PropertiesChoice = {
  styleId: "Heading2" | "Normal" | undefined;
  numbering: NumberingChoice | undefined;
};

const propertiesChoice: fc.Arbitrary<PropertiesChoice> = fc.record({
  styleId: fc.constantFrom("Heading2" as const, "Normal" as const, undefined),
  numbering: fc.constantFrom<NumberingChoice | undefined>(
    "none",
    "item",
    "nested",
    "newList",
    "heading",
    undefined,
  ),
});

/** One list edit, its block named by position among the blocks it meets. */
type Edit =
  | { type: "split"; at: number; offset: number }
  | { type: "merge"; at: number; properties: PropertiesChoice | undefined }
  | {
      type: "insertBefore" | "insertAfter";
      at: number;
      numbering: NumberingChoice | undefined;
      styleId: "Heading2" | undefined;
    }
  | { type: "properties"; at: number; properties: PropertiesChoice }
  | { type: "delete"; at: number };

const at = fc.nat({ max: 11 });

const edit: fc.Arbitrary<Edit> = fc.oneof(
  fc.record({
    type: fc.constant("split" as const),
    at,
    offset: fc.integer({ min: 1, max: 12 }),
  }),
  fc.record({
    type: fc.constant("merge" as const),
    at,
    properties: fc.option(propertiesChoice, { nil: undefined }),
  }),
  fc.record({
    type: fc.constantFrom("insertBefore" as const, "insertAfter" as const),
    at,
    numbering: fc.constantFrom<NumberingChoice | undefined>(
      "none",
      "item",
      "newList",
      "heading",
      undefined,
    ),
    styleId: fc.constantFrom("Heading2" as const, undefined),
  }),
  fc.record({
    type: fc.constant("properties" as const),
    at,
    properties: propertiesChoice,
  }),
  fc.record({ type: fc.constant("delete" as const), at }),
);

/** The paragraphs a reader lists with words, in order. */
const textBlocks = (reviewer: Reviewer): FolioAIParagraphBlock[] =>
  reviewer
    .getContent()
    .flatMap((block) =>
      block.kind !== "diagnostic" && block.text.length > 0 ? [block] : [],
    );

const requested = (
  { styleId, numbering }: PropertiesChoice,
  listId: number,
): FolioAIBlockParagraphProperties => ({
  ...(styleId === undefined ? {} : { styleId }),
  ...(numbering === undefined
    ? {}
    : { numbering: numberingOf(numbering, listId) }),
});

/** An operation and the block it names. */
type Named = { operation: FolioAIEditOperation; blockId: string };

/** `edit` as an operation on `blocks`, or null when it names nothing there. */
const operationOf = (
  change: Edit,
  blocks: readonly FolioAIParagraphBlock[],
  listId: number,
  id: string,
): Named | null => {
  const block = blocks[change.at % Math.max(blocks.length, 1)];
  if (block === undefined) return null;
  const blockId = block.id;
  const named = (operation: FolioAIEditOperation): Named => ({
    operation,
    blockId,
  });
  switch (change.type) {
    case "split":
      return change.offset < block.text.length
        ? named({ id, type: "splitBlock", blockId, offset: change.offset })
        : null;
    case "merge": {
      const merged =
        change.properties === undefined
          ? {}
          : requested(change.properties, listId);
      return named({
        id,
        type: "mergeBlockWithNext",
        blockId,
        separator: " ",
        ...(Object.keys(merged).length === 0
          ? {}
          : { mergedParagraphProperties: merged }),
      });
    }
    case "insertBefore":
    case "insertAfter": {
      const inserted = {
        id,
        blockId,
        text: `New clause ${id}.`,
        ...(change.styleId === undefined ? {} : { styleId: change.styleId }),
        ...(change.numbering === undefined
          ? {}
          : { numbering: numberingOf(change.numbering, listId) }),
      };
      return named(
        change.type === "insertBefore"
          ? { ...inserted, type: "insertBeforeBlock" }
          : { ...inserted, type: "insertAfterBlock" },
      );
    }
    case "properties": {
      const properties = requested(change.properties, listId);
      return Object.keys(properties).length === 0
        ? null
        : named({
            id,
            type: "setBlockParagraphProperties",
            blockId,
            properties,
          });
    }
    case "delete":
      return named({ id, type: "deleteBlock", blockId });
    default: {
      const unreachable: never = change;
      return unreachable;
    }
  }
};

/** What a reader tells of a block's kind: the parts the edits decide. */
const kinds = (reviewer: Reviewer) =>
  textBlocks(reviewer).map((block) => ({
    text: block.text,
    kind: block.kind,
    styleId: block.styleId,
    headingLevel: block.headingLevel,
    listLevel: block.listReference?.level,
    numbered: block.listReference !== undefined,
  }));

/**
 * Earlier suggestions, one round per author, each round saved; then the
 * batch, by an author of its own.
 */
type Case = { earlier: readonly (readonly Edit[])[]; batch: readonly Edit[] };

/**
 * The kinds the batch leaves accepted after other authors' earlier
 * suggestions, and those it leaves applied directly to their accepted
 * document; null when an edit does not apply in either.
 */
const outcomes = async (base: ArrayBuffer, { earlier, batch }: Case) => {
  let pending = base;
  let listId = HEADING_LIST;
  for (const [round, changes] of earlier.entries()) {
    const reviewer = await open(pending, `Reviewer ${round}`);
    listId =
      textBlocks(reviewer).find((block) => block.kind === "listItem")
        ?.listReference?.numId ?? listId;
    for (const [index, change] of changes.entries()) {
      const named = operationOf(
        change,
        textBlocks(reviewer),
        listId,
        `earlier-${round}-${index}`,
      );
      if (named === null) continue;
      reviewer.applyOperations([named.operation], {
        mode: "tracked-changes",
        undefinedReferences: "refuse",
      });
    }
    pending = await reviewer.toBuffer();
  }
  const tracked = await open(pending, "Batch reviewer");
  const blocks = textBlocks(tracked);
  const named = batch.flatMap((change, index) => {
    const one = operationOf(change, blocks, listId, `batch-${index}`);
    return one === null ? [] : [one];
  });
  const operations = named.map(({ operation }) => operation);
  if (operations.length === 0) return null;
  const suggested = tracked.applyOperations(operations, {
    mode: "tracked-changes",
    undefinedReferences: "refuse",
  });
  if (suggested.skipped.length > 0) return null;
  const accepted = await open(await tracked.toBuffer(), "Reader");
  accepted.acceptAll();
  const direct = await open(pending, "Batch reviewer");
  direct.acceptAll();
  // The batch names blocks as its reviewer reads them, pending changes and
  // all; directly it runs only on blocks the accepted document holds, and a
  // split only where the block reads the same there.
  const texts = new Map(
    textBlocks(direct).map((block) => [block.id, block.text]),
  );
  const live = new Map(blocks.map((block) => [block.id, block.text]));
  if (
    named.some(({ operation, blockId }) => {
      const text = texts.get(blockId);
      return (
        text === undefined ||
        (operation.type === "splitBlock" && text !== live.get(blockId))
      );
    })
  )
    return null;
  const applied = direct.applyOperations(operations, {
    mode: "direct",
    undefinedReferences: "refuse",
  });
  if (applied.skipped.length > 0) return null;
  return { accepted: kinds(accepted), direct: kinds(direct) };
};

describe("a suggested list edit beside another author's inserted break", () => {
  test("a paragraph put before one a merge joins on copies the joined paragraph", async () => {
    const base = await numberedDocx();
    // "Closing " keeps an inserted break the first author split it with;
    // the second puts a paragraph before it and merges it on as a heading.
    const result = await outcomes(base, {
      earlier: [[{ type: "split", at: 8, offset: 8 }]],
      batch: [
        {
          type: "insertBefore",
          at: 8,
          numbering: undefined,
          styleId: undefined,
        },
        {
          type: "merge",
          at: 8,
          properties: { styleId: "Heading2", numbering: undefined },
        },
      ],
    });
    expect(result).not.toBeNull();
    expect(result?.accepted).toEqual(result?.direct ?? []);
    expect(result?.accepted.at(-2)).toMatchObject({
      text: "New clause batch-0.",
      kind: "heading",
      styleId: "Heading2",
      listLevel: 0,
    });
  });

  test("merges run on through a removed inserted break with the first paragraph's properties", async () => {
    const base = await numberedDocx();
    // "S" keeps the break the first author split "Scope" with; the second
    // merges it on, then merges the title into it.
    const result = await outcomes(base, {
      earlier: [[{ type: "split", at: 1, offset: 1 }]],
      batch: [
        { type: "merge", at: 1, properties: undefined },
        { type: "merge", at: 0, properties: undefined },
      ],
    });
    expect(result).not.toBeNull();
    expect(result?.accepted).toEqual(result?.direct ?? []);
    expect(result?.accepted.at(0)).toMatchObject({
      text: "Agreement S cope",
      kind: "heading",
      styleId: "Heading1",
      numbered: false,
    });
  });

  test("a list item put before a paragraph an earlier merge removed an inserted break into goes before the join", async () => {
    const base = await numberedDocx();
    // One reviewer splits "Closing paragraph.", another merges the halves
    // back; the batch numbers a paragraph put before the second half.
    const result = await outcomes(base, {
      earlier: [
        [{ type: "split", at: 8, offset: 8 }],
        [{ type: "merge", at: 8, properties: undefined }],
      ],
      batch: [
        { type: "insertBefore", at: 9, numbering: "item", styleId: undefined },
      ],
    });
    expect(result).not.toBeNull();
    expect(result?.accepted).toEqual(result?.direct ?? []);
    expect(result?.accepted.slice(-2)).toMatchObject([
      { text: "New clause batch-0.", kind: "listItem", listLevel: 0 },
      { text: "Closing  paragraph.", kind: "paragraph", numbered: false },
    ]);
  });

  test(
    "accepting a suggested batch of list edits reads as the direct batch",
    async () => {
      const base = await numberedDocx();
      await assertProperty(
        fc.asyncProperty(
          fc.record({
            earlier: fc.array(fc.array(edit, { minLength: 1, maxLength: 2 }), {
              minLength: 1,
              maxLength: 2,
            }),
            batch: fc.array(edit, { minLength: 1, maxLength: 3 }),
          }),
          async (generated) => {
            const result = await outcomes(base, generated);
            if (result === null) return;
            // Where the words land is the accepted-versus-direct text rule
            // (#1350); this property holds each block's kind to its text.
            const texts = (rows: readonly { text: string }[]) =>
              rows.map(({ text }) => text);
            if (
              JSON.stringify(texts(result.accepted)) !==
              JSON.stringify(texts(result.direct))
            )
              return;
            expect(result.accepted).toEqual(result.direct);
          },
        ),
        {
          numRuns: 40,
          id: "suggested list edits accept to the direct batch's kinds",
        },
      );
    },
    propertyTestTimeout(120_000),
  );
});
