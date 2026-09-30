/** Cross-paragraph plans preserve review, inverse and survivor laws at partial endpoints. */
import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { assertProperty, propertyTestTimeout } from "../../../../../test/property-testing";
import {
  paragraphNumberingReference,
  type BlockContent,
  type Document,
  type Paragraph,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps, type AppliedDocumentOp } from "../apply";
import { endsItsContainer, storyParagraphs } from "../blocks";
import { contractViolation } from "../contract";
import { sameRunFormatting } from "../inline";
import { paragraphLogicalText } from "../offsets";
import { planTrackedDeletion } from "../plan";
import { planTrackedReplace } from "../rangeReplacement";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
  SPLIT_HALVES,
  type DocumentOp,
  type InlineSlice,
  type RevisionDecision,
} from "../types";

setDefaultTimeout(propertyTestTimeout(120_000));

const apply = (document: Document, op: DocumentOp) => {
  const result = applyDocumentOp(document, op);
  if (result.isErr()) throw result.error;
  return result.value;
};

const applyAll = (document: Document, ops: readonly DocumentOp[]) => {
  const result = applyDocumentOps(document, ops);
  if (result.isErr()) throw result.error;
  return result.value;
};

const resolve = (applied: AppliedDocumentOp, decision: RevisionDecision) =>
  apply(applied.document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: applied.revisions,
    decision,
  });

const paragraphsOf = (document: Document) =>
  storyParagraphs(document.package.document).map(({ paragraph }) => paragraph);

/** Comparing formatted text units allows the two valid seam merge shapes. */
const projection = (document: Document) =>
  paragraphsOf(document).map(({ content, ...fields }) =>
    Object.assign({}, fields, {
      units: content.flatMap((item) => {
        expect(item.type).toBe("run");
        if (item.type !== "run")
          throw new Error("Resolved generated content must contain only runs.");
        return item.content.flatMap((leaf) => {
          expect(leaf.type).toBe("text");
          if (leaf.type !== "text") throw new Error("Generated runs contain only text.");
          return leaf.text.split("").map((text) => ({ text, formatting: item.formatting }));
        });
      }),
    }),
  );

const expectInverse = (applied: AppliedDocumentOp, original: Document) => {
  const restored = applyAll(applied.document, applied.inverse);
  expect(restored.document).toStrictEqual(original);
  expect(applyAll(restored.document, restored.inverse).document).toStrictEqual(applied.document);
};

const runArbitrary = fc
  .record({
    text: fc.string({ unit: fc.constantFrom("a", "b", " ", "ž"), minLength: 1, maxLength: 6 }),
    bold: fc.boolean(),
    italic: fc.boolean(),
  })
  .map(({ text, bold, italic }) => ({
    type: "run" as const,
    formatting: { bold, italic },
    content: [{ type: "text" as const, text }],
  }));

type GeneratedRun = ReturnType<typeof runArbitrary.generate>["value"];

const canonicalRuns = (runs: readonly GeneratedRun[]) => {
  const content: GeneratedRun[] = [];
  for (const run of runs) {
    const previous = content.at(-1);
    const previousText = previous?.content.at(0);
    const text = run.content.at(0);
    if (
      previous !== undefined &&
      previousText !== undefined &&
      text !== undefined &&
      sameRunFormatting(previous.formatting, run.formatting)
    ) {
      previous.content = [{ type: "text", text: previousText.text + text.text }];
      continue;
    }
    content.push({ ...run, content: [...run.content] });
  }
  return content;
};

const paragraphArbitrary = fc
  .record({
    content: fc.array(runArbitrary, { maxLength: 3 }),
    alignment: fc.constantFrom("start", "end", "center"),
    listed: fc.boolean(),
    markItalic: fc.boolean(),
    runInWithNext: fc.boolean(),
  })
  .map(
    ({ content, alignment, listed, markItalic, runInWithNext }) =>
      ({
        type: "paragraph",
        content: canonicalRuns(content),
        formatting: {
          alignment,
          runProperties: { italic: markItalic },
          runInWithNext,
          ...(listed ? { numPr: paragraphNumberingReference({ numId: 1, ilvl: 0 }) } : {}),
        },
      }) satisfies Paragraph,
  );

const caseArbitrary = fc.record({
  source: fc.array(paragraphArbitrary, { minLength: 2, maxLength: 4 }),
  inserted: fc.array(paragraphArbitrary, { maxLength: 3 }),
  tail: paragraphArbitrary,
  fromPick: fc.nat(),
  toPick: fc.nat(),
  container: fc.constantFrom("body", "cell"),
  following: fc.boolean(),
});

type PlannedCase = ReturnType<typeof caseArbitrary.generate>["value"];

const makeCase = ({
  source,
  inserted,
  tail,
  fromPick,
  toPick,
  container,
  following,
}: PlannedCase) => {
  const paragraphs = source.map((paragraph, index) => ({
    ...paragraph,
    paraId: (16 + index).toString(16).padStart(8, "0"),
  }));
  const first = paragraphs.at(0);
  const last = paragraphs.at(-1);
  if (first === undefined || last === undefined)
    throw new Error("The generated source has endpoints.");
  const outside: Paragraph = {
    type: "paragraph",
    paraId: "00000001",
    content: [{ type: "run", content: [{ type: "text", text: "Outside" }] }],
  };
  const after: Paragraph = { type: "paragraph", paraId: "00000002", content: [] };
  const cellEnd: Paragraph = { type: "paragraph", paraId: "00000003", content: [] };
  const target = following ? [...paragraphs, after] : paragraphs;
  const blocks: BlockContent[] =
    container === "body"
      ? [outside, ...target]
      : [
          outside,
          {
            type: "table",
            rows: [{ type: "tableRow", cells: [{ type: "tableCell", content: target }] }],
          },
          cellEnd,
        ];
  const document: Document = { package: { document: { content: blocks } } };
  const from = {
    story: OP_STORIES.MAIN,
    blockId: first.paraId,
    offset: fromPick % (paragraphLogicalText(first).length + 1),
  };
  const to = {
    story: OP_STORIES.MAIN,
    blockId: last.paraId,
    offset: toPick % (paragraphLogicalText(last).length + 1),
  };
  return {
    document,
    paragraphs,
    from,
    to,
    replacement: {
      paragraphs: inserted.map((paragraph, index) => ({
        ...paragraph,
        paraId: (256 + index).toString(16).padStart(8, "0"),
      })),
      tail: { content: tail.content, openStart: 0, openEnd: 0 } satisfies InlineSlice,
    },
  };
};

type IdentifiedCase = ReturnType<typeof makeCase>;

type DirectDeletionOptions = Pick<IdentifiedCase, "document" | "from" | "to"> & {
  paragraphs: readonly (Paragraph & { paraId: string })[];
};

/** Direct range edits followed by reverse joins retain the original final identity. */
const directDeletion = ({ document, paragraphs, from, to }: DirectDeletionOptions) => {
  let current = document;
  for (const [index, paragraph] of paragraphs.entries()) {
    current = apply(current, {
      type: DOCUMENT_OP_TYPES.DELETE_RANGE,
      from: { ...from, blockId: paragraph.paraId, offset: index === 0 ? from.offset : 0 },
      to: {
        ...to,
        blockId: paragraph.paraId,
        offset:
          index === paragraphs.length - 1 ? to.offset : paragraphLogicalText(paragraph).length,
      },
    }).document;
  }
  const last = paragraphs.at(-1);
  if (last === undefined) throw new Error("The generated source has a final paragraph.");
  for (const paragraph of paragraphs.slice(0, -1).toReversed()) {
    current = apply(current, {
      type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
      story: OP_STORIES.MAIN,
      blockId: paragraph.paraId,
      nextBlockId: last.paraId,
    }).document;
  }
  return current;
};

/** Delete first, then place replacement paragraphs at the retained prefix boundary. */
const directReplacement = (identified: IdentifiedCase) => {
  const { replacement, from } = identified;
  const last = identified.paragraphs.at(-1);
  if (last === undefined) throw new Error("The generated source has a final paragraph.");
  let current = directDeletion(identified);
  let at = { ...from, blockId: last.paraId };
  for (const paragraph of replacement.paragraphs) {
    const width = paragraphLogicalText(paragraph).length;
    if (width > 0) {
      current = apply(current, {
        type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
        at,
        slice: { content: paragraph.content, openStart: 0, openEnd: 0 },
      }).document;
    }
    current = apply(current, {
      type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
      at: { ...at, offset: at.offset + width },
      newBlockId: paragraph.paraId,
      newHalf: SPLIT_HALVES.FIRST,
      newParagraph: { formatting: paragraph.formatting },
    }).document;
    at = { story: from.story, blockId: last.paraId, offset: 0 };
  }
  const tailWidth = paragraphLogicalText({
    type: "paragraph",
    content: [...replacement.tail.content],
  }).length;
  if (tailWidth > 0) {
    current = apply(current, {
      type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
      at,
      slice: replacement.tail,
    }).document;
  }
  return current;
};

const expectLaws = ({
  document,
  ops,
  direct,
  sourceIds,
}: {
  document: Document;
  ops: readonly DocumentOp[];
  direct: Document;
  sourceIds: ReadonlySet<string>;
}) => {
  const original = structuredClone(document);
  const tracked = applyAll(document, ops);
  const accepted = resolve(tracked, REVISION_DECISIONS.ACCEPT);
  const rejected = resolve(tracked, REVISION_DECISIONS.REJECT);
  expect(projection(accepted.document)).toStrictEqual(projection(direct)); // L1
  expect(rejected.document).toStrictEqual(document); // L2
  expectInverse(tracked, document); // L4
  expectInverse(accepted, tracked.document);
  expectInverse(rejected, tracked.document);
  expect(applyAll(structuredClone(document), structuredClone(ops))).toStrictEqual(tracked); // L5
  const trackedById = new Map(
    paragraphsOf(tracked.document).map((paragraph) => [paragraph.paraId, paragraph]),
  );
  for (const paragraph of paragraphsOf(document)) {
    if (!sourceIds.has(paragraph.paraId ?? "")) {
      expect(trackedById.get(paragraph.paraId)).toBe(paragraph); // L6
    }
  }
  for (const reviewed of [tracked, accepted, rejected]) {
    const body = reviewed.document.package.document;
    expect(contractViolation(reviewed.document)).toBeUndefined();
    expect(
      storyParagraphs(body).filter(
        (at) => endsItsContainer(body, at) && at.paragraph.pPrMark !== undefined,
      ),
    ).toEqual([]); // L7
  }
  for (const [decision, reviewed] of [
    [REVISION_DECISIONS.ACCEPT, accepted],
    [REVISION_DECISIONS.REJECT, rejected],
  ] as const) {
    const twice = resolve({ ...tracked, document: reviewed.document }, decision);
    expect(twice.document).toBe(reviewed.document);
    expect(twice.inverse).toEqual([]);
  }
  let sequential = document;
  const edits: AppliedDocumentOp[] = [];
  for (const op of ops) {
    const edit = apply(sequential, op);
    edits.push(edit);
    sequential = edit.document;
  }
  expect(sequential).toStrictEqual(tracked.document); // L3
  expect(
    applyAll(
      sequential,
      edits.toReversed().flatMap(({ inverse }) => inverse),
    ).document,
  ).toStrictEqual(document);
  expect(document).toStrictEqual(original);
};

test("cross-paragraph deletion and replacement plans satisfy generated review laws", () => {
  assertProperty(
    fc.property(caseArbitrary, (seed) => {
      const identified = makeCase(seed);
      const { document, from, to, paragraphs, replacement } = identified;
      const options = {
        from,
        to,
        revision: { id: 1000, author: "Reviewer", date: "2026-02-03T04:05:06Z" },
        newIds: { revision: Array.from({ length: 128 }, (_, index) => 1001 + index) },
      };
      const deletion = planTrackedDeletion(document, options);
      if (deletion.isErr()) throw deletion.error;
      expect(
        planTrackedDeletion(structuredClone(document), structuredClone(options)),
      ).toStrictEqual(deletion);
      const sourceIds = new Set(paragraphs.map(({ paraId }) => paraId));
      expectLaws({ document, ops: deletion.value, direct: directDeletion(identified), sourceIds });
      const replaced = planTrackedReplace(document, { ...options, replacement });
      if (replaced.isErr()) throw replaced.error;
      expect(
        planTrackedReplace(structuredClone(document), structuredClone({ ...options, replacement })),
      ).toStrictEqual(replaced);
      expectLaws({
        document,
        ops: replaced.value,
        direct: directReplacement(identified),
        sourceIds,
      });
      const finalId = paragraphs.at(-1)?.paraId;
      expect(
        paragraphsOf(
          resolve(applyAll(document, replaced.value), REVISION_DECISIONS.ACCEPT).document,
        )
          .filter(({ paraId }) => sourceIds.has(paraId ?? ""))
          .map(({ paraId }) => paraId),
      ).toEqual([finalId]);
    }),
    { numRuns: 2_000 },
  );
});
