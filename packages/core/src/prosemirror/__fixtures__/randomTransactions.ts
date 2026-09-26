/**
 * Random documents and random multi-step transactions for the equivalence
 * properties of the editor plugins that map positions through a transaction.
 *
 * The schema nests paragraphs in quotes and table cells, so positions cross
 * container boundaries, and the operations cover every step shape the plugins
 * read: plain replacements (typing, deleting, splitting, joining, pasting),
 * replace-around steps (wrapping, lifting, re-marking a paragraph), attribute
 * steps and mark steps. An operation that does not fit the document it lands
 * in is skipped, so every generated transaction is valid.
 */

import fc from "fast-check";
import { Schema, type Node as PMNode } from "prosemirror-model";
import { EditorState, type Transaction } from "prosemirror-state";
import { canJoin, canSplit, findWrapping, liftTarget } from "prosemirror-transform";

import {
  PROSE_PARAGRAPH_SOURCE_CONTRACT_ATTR,
  PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR,
} from "../../docx/paragraphPropertySource";
import { markChangedParagraphRanges } from "../extensions/features/ParagraphChangeTrackerExtension";
import { RUN_IDENTITY_MARK_NAME } from "../runIdentity";

export const randomSchema = new Schema({
  nodes: {
    doc: {
      content: "block+",
      attrs: { [PROSE_PARAGRAPH_SOURCE_CONTRACT_ATTR]: { default: null } },
    },
    paragraph: {
      group: "block",
      content: "inline*",
      attrs: {
        paraId: { default: null },
        direction: { default: null },
        [PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR]: { default: null },
      },
    },
    blockquote: { group: "block", content: "block+" },
    table: { group: "block", content: "tableRow+" },
    tableRow: { content: "tableCell+" },
    tableCell: { content: "block+" },
    text: { group: "inline" },
  },
  marks: {
    bold: {},
    deletion: {},
    [RUN_IDENTITY_MARK_NAME]: { attrs: { id: {} }, inclusive: false },
  },
});

/** Few enough ids that duplicates and collisions are common. */
const PARA_IDS = ["00000001", "00000002", "00000003", "0000000A", "0000000B", null] as const;
const TOKENS = ["token-a", "token-b", "token-c", null] as const;
const WORDS = ["alpha", "beta", "gamma delta", "مرحبا بالعالم", "שלום", " "] as const;

type ParagraphSpec = {
  paraId: (typeof PARA_IDS)[number];
  token: (typeof TOKENS)[number];
  word: (typeof WORDS)[number];
  identity: number | null;
};

type BlockSpec =
  | { kind: "paragraph"; paragraph: ParagraphSpec }
  | { kind: "quote"; paragraphs: ParagraphSpec[] }
  | { kind: "table"; paragraphs: ParagraphSpec[] };

const paragraphSpec: fc.Arbitrary<ParagraphSpec> = fc.record({
  paraId: fc.constantFrom(...PARA_IDS),
  token: fc.constantFrom(...TOKENS),
  word: fc.constantFrom(...WORDS),
  identity: fc.option(fc.nat({ max: 3 })),
});

const blockSpec: fc.Arbitrary<BlockSpec> = fc.oneof(
  {
    weight: 6,
    arbitrary: paragraphSpec.map((paragraph): BlockSpec => ({ kind: "paragraph", paragraph })),
  },
  {
    weight: 1,
    arbitrary: fc
      .array(paragraphSpec, { minLength: 1, maxLength: 3 })
      .map((paragraphs): BlockSpec => ({ kind: "quote", paragraphs })),
  },
  {
    weight: 1,
    arbitrary: fc
      .array(paragraphSpec, { minLength: 1, maxLength: 3 })
      .map((paragraphs): BlockSpec => ({ kind: "table", paragraphs })),
  },
);

const buildParagraph = ({ paraId, token, word, identity }: ParagraphSpec): PMNode => {
  const marks =
    identity === null ? [] : [randomSchema.mark(RUN_IDENTITY_MARK_NAME, { id: identity })];
  return randomSchema.node(
    "paragraph",
    { paraId, [PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR]: token },
    word.trim().length === 0 ? [] : [randomSchema.text(word, marks)],
  );
};

const buildBlock = (block: BlockSpec): PMNode => {
  if (block.kind === "paragraph") {
    return buildParagraph(block.paragraph);
  }
  if (block.kind === "quote") {
    return randomSchema.node("blockquote", null, block.paragraphs.map(buildParagraph));
  }
  return randomSchema.node("table", null, [
    randomSchema.node(
      "tableRow",
      null,
      block.paragraphs.map((paragraph) =>
        randomSchema.node("tableCell", null, [buildParagraph(paragraph)]),
      ),
    ),
  ]);
};

/**
 * Up to forty top-level blocks: past the size where the indexed lookups stop
 * scanning linearly and start binary-searching.
 */
export const randomDocument: fc.Arbitrary<PMNode> = fc
  .array(blockSpec, { minLength: 1, maxLength: 40, size: "large" })
  .map((blocks) => randomSchema.node("doc", null, blocks.map(buildBlock)));

const OPERATION_KINDS = [
  "insertText",
  "insertText",
  "insertText",
  "delete",
  "deleteLarge",
  "split",
  "join",
  "insertParagraph",
  "addMark",
  "removeMark",
  "setMarkup",
  "setAttribute",
  "paste",
  "wrap",
  "lift",
  "recordRange",
] as const;

export type RandomOperation = {
  kind: (typeof OPERATION_KINDS)[number];
  a: number;
  b: number;
  c: number;
};

export const randomOperation: fc.Arbitrary<RandomOperation> = fc.record({
  kind: fc.constantFrom(...OPERATION_KINDS),
  a: fc.nat({ max: 10_000 }),
  b: fc.nat({ max: 10_000 }),
  c: fc.nat({ max: 10_000 }),
});

export const randomOperations = (maxLength: number): fc.Arbitrary<RandomOperation[]> =>
  fc.array(randomOperation, { minLength: 1, maxLength, size: "large" });

const paragraphPositions = (doc: PMNode): number[] => {
  const positions: number[] = [];
  doc.descendants((node, pos) => {
    if (node.type.name === "paragraph") {
      positions.push(pos);
      return false;
    }
    return true;
  });
  return positions;
};

const pick = <T>(values: readonly T[], index: number): T | undefined =>
  values.length === 0 ? undefined : values[index % values.length];

const MARKS = ["bold", "deletion", RUN_IDENTITY_MARK_NAME] as const;

const markFor = (index: number) => {
  const name = MARKS[index % MARKS.length] ?? "bold";
  return name === RUN_IDENTITY_MARK_NAME
    ? randomSchema.mark(name, { id: index % 3 })
    : randomSchema.mark(name);
};

/** Apply one operation to `tr`, or nothing when it does not fit. */
export const applyRandomOperation = (tr: Transaction, operation: RandomOperation): void => {
  const { kind, a, b, c } = operation;
  const size = tr.doc.content.size;
  const at = a % (size + 1);
  const span = kind === "deleteLarge" ? b % (size + 1) : b % 12;
  const to = Math.min(size, at + span);
  try {
    switch (kind) {
      case "insertText":
        tr.insertText(pick(WORDS, c) ?? "x", at);
        return;
      case "delete":
      case "deleteLarge":
        tr.delete(at, to);
        return;
      case "split":
        if (canSplit(tr.doc, at)) {
          tr.split(at);
        }
        return;
      case "join":
        if (canJoin(tr.doc, at)) {
          tr.join(at);
        }
        return;
      case "insertParagraph":
        tr.replaceWith(
          at,
          at,
          randomSchema.node("paragraph", { paraId: pick(PARA_IDS, c) ?? null }, [
            randomSchema.text(pick(WORDS, b) ?? "x"),
          ]),
        );
        return;
      case "addMark":
        tr.addMark(at, to, markFor(c));
        return;
      case "removeMark":
        tr.removeMark(at, to, markFor(c).type);
        return;
      case "setMarkup":
      case "setAttribute": {
        const pos = pick(paragraphPositions(tr.doc), c);
        const paragraph = pos === undefined ? null : tr.doc.nodeAt(pos);
        if (pos === undefined || !paragraph) {
          return;
        }
        if (kind === "setMarkup") {
          tr.setNodeMarkup(pos, undefined, { ...paragraph.attrs, paraId: pick(PARA_IDS, b) });
        } else {
          tr.setNodeAttribute(pos, "direction", b % 2 === 0 ? null : { source: "manual" });
        }
        return;
      }
      case "paste": {
        // A bounded slice: repeated pastes of the whole document would double
        // it every time.
        const source = c % (size + 1);
        tr.replace(at, at, tr.doc.slice(source, Math.min(size, source + (b % 40))));
        return;
      }
      case "wrap":
      case "lift": {
        const range = tr.doc.resolve(at).blockRange(tr.doc.resolve(to));
        if (!range) {
          return;
        }
        const blockquote = randomSchema.nodes["blockquote"];
        if (kind === "wrap" && blockquote) {
          const wrapping = findWrapping(range, blockquote);
          if (wrapping) {
            tr.wrap(range, wrapping);
          }
          return;
        }
        const target = liftTarget(range);
        if (target !== null) {
          tr.lift(range, target);
        }
        return;
      }
      case "recordRange":
        markChangedParagraphRanges(tr, {
          ranges: [{ from: at, to }],
          mappingFrom: tr.steps.length,
        });
        return;
    }
  } catch {
    // The operation does not fit this document; the transaction goes on without it.
  }
};

/** One transaction built from `operations` on `state`. */
export const buildRandomTransaction = (
  state: EditorState,
  operations: readonly RandomOperation[],
): Transaction => {
  const tr = state.tr;
  for (const operation of operations) {
    applyRandomOperation(tr, operation);
  }
  return tr;
};

/**
 * `batches` transactions applied one after another to a plugin-free state, as
 * a batch `appendTransaction` would see them.
 */
export const buildRandomBatch = (
  doc: PMNode,
  batches: readonly (readonly RandomOperation[])[],
): Transaction[] => {
  let state = EditorState.create({ doc });
  const transactions: Transaction[] = [];
  for (const operations of batches) {
    const tr = buildRandomTransaction(state, operations);
    transactions.push(tr);
    state = state.apply(tr);
  }
  return transactions;
};
