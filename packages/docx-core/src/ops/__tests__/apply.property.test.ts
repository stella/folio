/**
 * The laws every direct operation keeps, over synthetic documents built from
 * the model (tracked operations: `review.property.test.ts`).
 *
 * 1. **Inverse.** Applying an operation and then its recorded inverse gives
 *    back a document structurally equal to the input, every unmodelled and
 *    captured field included, and so does the inverse after a JSON
 *    round-trip: it holds data, not references. Applying the inverse's own
 *    inverse redoes the operation. Structural inverses restore authored run
 *    boundaries when the parser's seam merge would lose them.
 * 2. **Sequences.** For a random sequence of operations, the inverses applied
 *    in reverse order restore the input exactly; a batch of the same
 *    operations gives the same document, and its inverse restores the input.
 * 3. **Determinism.** The same operation on equal documents gives equal
 *    results, including after the operation has been through JSON.
 * 4. **Locality.** A paragraph the operation does not report as touched is
 *    the same object afterwards, and so is every top-level block holding
 *    none, every package part besides the story, and the parser's section
 *    view of untouched blocks. The touched set names only ids the operation
 *    names.
 * 5. **Offsets.** Each operation changes the paragraph's logical text and run
 *    properties exactly as its definition says.
 */

import { assertExactModel } from "../../../../../test/exactModel";
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { panic } from "better-result";

import {
  assertProperty,
  propertyConfig,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import { assertFreshIdentityEquivalent } from "./freshIdentityOracle";
import { projectReview } from "../../../../../test/reviewProjection";
import type {
  BlockContent,
  Document,
  Paragraph,
  Run,
  SectionProperties,
  TextFormatting,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps, type AppliedDocumentOp } from "../apply";
import { sameBlockList, storyBody, storyParagraphs } from "../blocks";
import { contractViolation, normalizeForOps } from "../contract";
import { paragraphIdsIn, identityKeysIn, packageIdentityKeys, IDENTITY_SPACES } from "../ids";
import { sameRunFormatting } from "../inline";
import { compareGaps, isCommentAnchor, leafSpans, zeroWidthLeavesAt } from "../leaves";
import { packageResourcesOf, packageResourcesOpOf } from "../packageResources";
import {
  allocateEditorIntentIds,
  compileEditorIntent,
  editorParagraphGroups,
  physicalPositionAtEditorOffset,
  type EditorIntent,
} from "../editorIntent";
import {
  childrenOf,
  isInlineContainer,
  isRemovedRevision,
  paragraphLogicalText,
  paragraphLength,
  runContentWidth,
} from "../offsets";
import { applyFormattingPatch } from "../patch";
import { DOCUMENT_OP_REFUSAL_REASONS, type DocumentOpRefusalReason } from "../refusal";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type DocumentOpType,
  INHERIT_RUN_PROPS,
  OP_STORIES,
  REVISION_DECISIONS,
  SPLIT_HALVES,
} from "../types";
import {
  documentArbitrary,
  GENERATED_OP_KINDS,
  independentCopy,
  type OpSeed,
  opFor,
  opSeedArbitrary,
} from "./documentArbitraries";

setDefaultTimeout(propertyTestTimeout(240_000));

const expectCompactIntentIds = (document: Document, ops: readonly DocumentOp[]): void => {
  let current = document;
  for (const op of ops) {
    const applied = applyDocumentOp(current, op).unwrap();
    if ("newIds" in op && op.newIds !== undefined) {
      const changed = new Set([...applied.touched.modified, ...applied.touched.inserted]);
      const records = storyParagraphs(storyBody(applied.document, OP_STORIES.MAIN))
        .filter(({ paragraph }) => changed.has(paragraph.paraId ?? ""))
        .map(({ paragraph }) => paragraph);
      const identities = new Set(identityKeysIn(records));
      for (const id of op.newIds.revision ?? [])
        expect(
          identities.has(`${IDENTITY_SPACES.REVISION}:${id}`),
          JSON.stringify({ id, op, identities: [...identities] }),
        ).toBe(true);
      for (const id of op.newIds.control ?? [])
        expect(identities.has(`${IDENTITY_SPACES.CONTROL}:${id}`)).toBe(true);
    }
    current = applied.document;
  }
};

const NUM_RUNS = 10_000;

type Tally = Map<DocumentOpType | "refused", number>;

const count = (tally: Tally, key: DocumentOpType | "refused"): void => {
  tally.set(key, (tally.get(key) ?? 0) + 1);
};

/**
 * A property that passes because nothing applied proves nothing: every
 * generated kind must have applied in a real share of the runs.
 */
const expectEveryKindApplied = (tally: Tally, runs: number): void => {
  for (const kind of GENERATED_OP_KINDS) {
    expect({ kind, applied: tally.get(kind) ?? 0 }).toEqual({
      kind,
      applied: expect.any(Number),
    });
    expect(tally.get(kind) ?? 0).toBeGreaterThan(runs / 1000);
  }
};

/**
 * What each operation's inverse is made of: the table in `apply.ts`. Every
 * inverse is one operation, except a run patch's, which restores one stretch
 * of prior values per operation, and a split's or join's, which may give the
 * paragraph keeping its id its own review fields back and restore authored
 * content that the plain-run seam merge would otherwise coalesce.
 */
const INVERSE_KINDS = {
  setPackageResources: ["setPackageResources"],
  createComment: ["restoreCommentState"],
  updateCommentContent: ["restoreCommentState"],
  setCommentResolution: ["restoreCommentState"],
  deleteComment: ["restoreCommentState"],
  restoreCommentState: ["restoreCommentState"],
  createHeaderFooter: ["restoreStoryParts"],
  removeHeaderFooter: ["restoreStoryParts"],
  addNote: ["restoreStoryParts"],
  removeNote: ["restoreStoryParts"],
  setSectionProps: ["restoreStoryParts"],
  restoreStoryParts: ["restoreStoryParts"],
  deleteBlocks: ["insertBlocks", "replaceBlocks", "replaceInline", "setParagraphReview"],
  insertBlocks: ["replaceBlocks"],
  insertTable: ["setContainerBlocks"],
  deleteTable: ["setContainerBlocks", "replaceInline", "setParagraphReview"],
  setContainerBlocks: ["setContainerBlocks"],
  insertRow: ["setTableRows"],
  deleteRow: ["setTableRows", "replaceInline", "setParagraphReview"],
  setTableRows: ["setTableRows"],
  insertText: ["deleteRange"],
  insertContent: ["deleteRange"],
  deleteRange: ["insertContent"],
  splitInline: ["joinInline"],
  joinInline: ["splitInline"],
  setRunProps: ["setRunProps"],
  setParagraphProps: ["setParagraphProps"],
  splitBlock: ["joinBlocks", "replaceInline", "setParagraphReview"],
  joinBlocks: ["splitBlock", "replaceInline", "setParagraphReview"],
  replaceBlocks: ["replaceBlocks"],
  setParagraphReview: ["setParagraphReview"],
  replaceInline: ["replaceInline"],
  resolveRevision: ["replaceInline", "setParagraphReview", "joinBlocks", "replaceBlocks"],
  createNumberingInstance: ["deleteNumberingInstance"],
  deleteNumberingInstance: ["createNumberingInstance"],
  setSectionEndpoint: ["setSectionEndpoint"],
} as const satisfies Record<DocumentOpType, readonly DocumentOpType[]>;

const paragraphsById = (document: Document): Map<string, Paragraph> =>
  new Map(
    storyParagraphs(document.package.document).map(({ paragraph }) => [
      paragraph.paraId ?? "",
      paragraph,
    ]),
  );

const orderedIds = (document: Document): string[] =>
  storyParagraphs(document.package.document).map(({ paragraph }) => paragraph.paraId ?? "");

const blockHolds = (block: BlockContent, ids: ReadonlySet<string>): boolean => {
  switch (block.type) {
    case "paragraph":
      return block.paraId !== undefined && ids.has(block.paraId);
    case "table":
      return block.rows.some((row) =>
        row.cells.some((cell) => cell.content.some((child) => blockHolds(child, ids))),
      );
    case "blockSdt":
      return block.content.some((child) => blockHolds(child, ids));
    default:
      return false;
  }
};

const namedIds = (op: DocumentOp): Set<string> => {
  switch (op.type) {
    case DOCUMENT_OP_TYPES.DELETE_BLOCKS:
      return new Set(op.blockIds);
    case DOCUMENT_OP_TYPES.INSERT_BLOCKS:
      return new Set([
        op.at.blockId,
        ...op.blocks.flatMap(({ paraId }) => (paraId === undefined ? [] : [paraId])),
      ]);
    case DOCUMENT_OP_TYPES.INSERT_TEXT:
    case DOCUMENT_OP_TYPES.INSERT_CONTENT:
    case DOCUMENT_OP_TYPES.SPLIT_INLINE:
    case DOCUMENT_OP_TYPES.JOIN_INLINE:
      return new Set([op.at.blockId]);
    case DOCUMENT_OP_TYPES.DELETE_RANGE:
    case DOCUMENT_OP_TYPES.SET_RUN_PROPS:
      return new Set([op.from.blockId, op.to.blockId]);
    case DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS:
    case DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW:
    case DOCUMENT_OP_TYPES.REPLACE_INLINE:
      return new Set([op.blockId]);
    case DOCUMENT_OP_TYPES.RESOLVE_REVISION:
      return new Set();
    case DOCUMENT_OP_TYPES.SPLIT_BLOCK:
      return new Set([op.at.blockId, op.newBlockId]);
    case DOCUMENT_OP_TYPES.JOIN_BLOCKS:
      return new Set([op.blockId, op.nextBlockId]);
    case DOCUMENT_OP_TYPES.REPLACE_BLOCKS:
      return new Set(
        [...op.expected, ...op.blocks].flatMap(({ paraId }) =>
          paraId === undefined ? [] : [paraId],
        ),
      );
    case DOCUMENT_OP_TYPES.INSERT_TABLE:
      return new Set([op.at.blockId, ...paragraphIdsIn(op.table)]);
    case DOCUMENT_OP_TYPES.DELETE_TABLE:
      return new Set([op.blockId, ...paragraphIdsIn(op.expected ?? [])]);
    case DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS:
      return new Set([op.blockId, ...paragraphIdsIn([op.expected, op.blocks])]);
    case DOCUMENT_OP_TYPES.INSERT_ROW:
      return new Set([op.blockId, ...paragraphIdsIn(op.row)]);
    case DOCUMENT_OP_TYPES.DELETE_ROW:
      return new Set([op.blockId, ...paragraphIdsIn(op.expected ?? [])]);
    case DOCUMENT_OP_TYPES.SET_TABLE_ROWS:
      return new Set([op.blockId, ...paragraphIdsIn([op.expected, op.rows])]);
    case DOCUMENT_OP_TYPES.CREATE_NUMBERING_INSTANCE:
    case DOCUMENT_OP_TYPES.DELETE_NUMBERING_INSTANCE:
      return new Set();
    case DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT:
      return new Set(op.endpoint.type === "paragraph" ? [op.endpoint.blockId] : []);
    case DOCUMENT_OP_TYPES.SET_PACKAGE_RESOURCES:
    case DOCUMENT_OP_TYPES.CREATE_COMMENT:
    case DOCUMENT_OP_TYPES.UPDATE_COMMENT_CONTENT:
    case DOCUMENT_OP_TYPES.SET_COMMENT_RESOLUTION:
    case DOCUMENT_OP_TYPES.DELETE_COMMENT:
    case DOCUMENT_OP_TYPES.RESTORE_COMMENT_STATE:
    case DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER:
    case DOCUMENT_OP_TYPES.REMOVE_HEADER_FOOTER:
    case DOCUMENT_OP_TYPES.ADD_NOTE:
    case DOCUMENT_OP_TYPES.REMOVE_NOTE:
    case DOCUMENT_OP_TYPES.SET_SECTION_PROPS:
    case DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS:
      return new Set();
    default: {
      const unreachable: never = op;
      return unreachable;
    }
  }
};

const touchedIds = ({ touched }: AppliedDocumentOp): Set<string> =>
  new Set([...touched.modified, ...touched.inserted, ...touched.removed]);

type Unit = { formatting: TextFormatting | undefined; inRun: boolean; removed: boolean };

/** The run properties and revision state of every unit, in offset order. */
const unitsOf = (paragraph: Paragraph): Unit[] => {
  const out: Unit[] = [];
  const walk = (items: Paragraph["content"], removed: boolean): void => {
    for (const item of items) {
      if (item.type === "run") {
        for (const content of item.content) {
          const width = runContentWidth(content);
          for (let unit = 0; unit < width; unit += 1) {
            out.push({ formatting: item.formatting, inRun: true, removed });
          }
        }
        continue;
      }
      if (isInlineContainer(item)) {
        walk(childrenOf(item), removed || isRemovedRevision(item));
        continue;
      }
      if (paragraphLogicalText({ type: "paragraph", content: [item] }).length === 1) {
        out.push({ formatting: undefined, inRun: false, removed });
      }
    }
  };
  walk(paragraph.content, false);
  return out;
};

const applyAll = (document: Document, ops: readonly DocumentOp[]): Document => {
  const applied = applyDocumentOps(document, ops);
  if (applied.isErr()) {
    throw applied.error;
  }
  return applied.value.document;
};

const expectRestores = (applied: AppliedDocumentOp, original: Document): void => {
  // An operation leaves the seed contract holding: checked afresh, not from memory.
  expect(contractViolation(applied.document)).toBeUndefined();
  const restored = applyDocumentOps(applied.document, applied.inverse);
  if (restored.isErr()) {
    throw restored.error;
  }
  expect(restored.value.document).toStrictEqual(original);
  expect(contractViolation(restored.value.document)).toBeUndefined();
  // SAFETY: operations are plain data; this is the journal's round-trip.
  const replayed = JSON.parse(JSON.stringify(applied.inverse)) as DocumentOp[];
  expect(applyAll(applied.document, replayed)).toStrictEqual(original);
  expect(applyAll(restored.value.document, restored.value.inverse)).toStrictEqual(applied.document);
};

const plainInputArbitrary = fc
  .array(fc.constantFrom("a", " ", "é", "ß", "😀", "𐐀", "e\u0301"), {
    minLength: 1,
    maxLength: 8,
  })
  .map((parts) => parts.join(""));

const PLAIN_INPUT_KINDS = ["insert", "delete", "replace", "reject"] as const;
const replaceInputArbitrary = fc.record({
  kind: fc.constantFrom(...PLAIN_INPUT_KINDS),
  anchor: fc.nat(),
  head: fc.nat(),
  text: plainInputArbitrary,
});

/** Canonical text input addresses UTF-16 gaps between complete code points. */
const codePointGaps = (text: string): number[] => {
  const gaps = [0];
  let offset = 0;
  for (const point of text) {
    offset += point.length;
    gaps.push(offset);
  }
  return gaps;
};

describe("document operations", () => {
  test("rejecting generated run formatting restores authored boundaries inside wrappers and pending review", () => {
    assertProperty(
      fc.property(
        fc.array(fc.string({ unit: fc.constantFrom("a", "b", "é"), minLength: 2, maxLength: 5 }), {
          minLength: 2,
          maxLength: 4,
        }),
        fc.boolean(),
        fc.boolean(),
        fc.boolean(),
        (texts, wrapped, pending, partial) => {
          const runs = texts.map(
            (text, index) =>
              ({
                type: "run",
                formatting: { italic: true },
                content: [{ type: "text", text }],
                ...(pending
                  ? {
                      propertyChanges: [
                        {
                          type: "runPropertyChange",
                          info: { id: index + 1, author: "Earlier" },
                          previousFormatting: { bold: true },
                        },
                      ],
                    }
                  : {}),
              }) satisfies Run,
          );
          const original = normalizeForOps({
            package: {
              document: {
                content: [
                  {
                    type: "paragraph",
                    paraId: "00000001",
                    content: wrapped
                      ? [{ type: "hyperlink", anchor: "target", children: runs }]
                      : runs,
                  },
                ],
              },
            },
          } satisfies Document);
          const at = { story: OP_STORIES.MAIN, blockId: "00000001", offset: partial ? 1 : 0 };
          const intent = {
            type: "formatRun",
            from: at,
            to: { ...at, offset: texts.join("").length - (partial ? 1 : 0) },
            patch: { bold: true },
          } as const satisfies EditorIntent;
          const allocation = allocateEditorIntentIds(original, intent);
          const plan = compileEditorIntent(original, {
            intent,
            mode: {
              type: "suggesting",
              revision: { id: allocation.revisionId, author: "New", date: "2026-10-02T00:00:00Z" },
              newIds: allocation.newIds,
            },
          }).unwrap();
          const edit = applyDocumentOps(original, plan.ops).unwrap();
          expect(edit.revisions.length).toBeGreaterThan(0);
          const rejected = applyDocumentOp(edit.document, {
            type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
            story: OP_STORIES.MAIN,
            revisionIds: edit.revisions,
            decision: REVISION_DECISIONS.REJECT,
          }).unwrap();
          assertExactModel(rejected.document, original);
          const undone = applyDocumentOps(edit.document, edit.inverse).unwrap();
          assertExactModel(undone.document, original);
          const redone = applyDocumentOps(undone.document, undone.inverse).unwrap();
          assertExactModel(redone.document, edit.document);
          const firstText = texts.at(0);
          const secondText = texts.at(1);
          if (firstText === undefined || secondText === undefined)
            panic("Boundary sequence lost its middle run.");
          const laterIntent = {
            type: "formatRun",
            from: { ...at, offset: firstText.length },
            to: {
              ...at,
              offset:
                firstText.length + secondText.length - (texts.length === 2 && partial ? 1 : 0),
            },
            patch: { underline: { style: "single" } },
          } as const satisfies EditorIntent;
          const laterIds = allocateEditorIntentIds(edit.document, laterIntent);
          const laterPlan = compileEditorIntent(edit.document, {
            intent: laterIntent,
            mode: {
              type: "suggesting",
              revision: { id: laterIds.revisionId, author: "Later", date: "2026-10-02T00:00:00Z" },
              newIds: laterIds.newIds,
            },
          }).unwrap();
          const later = applyDocumentOps(edit.document, laterPlan.ops).unwrap();
          const rejectEarlier = applyDocumentOp(later.document, {
            type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
            story: OP_STORIES.MAIN,
            revisionIds: edit.revisions,
            decision: REVISION_DECISIONS.REJECT,
          }).unwrap();
          const rejectLater = applyDocumentOp(rejectEarlier.document, {
            type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
            story: OP_STORIES.MAIN,
            revisionIds: later.revisions,
            decision: REVISION_DECISIONS.REJECT,
          }).unwrap();
          assertExactModel(rejectLater.document, original);
        },
      ),
      { numRuns: 100 },
    );
  });

  test("each formatting author gets a separately rejectable record over pending run review", () => {
    assertProperty(
      fc.property(
        fc.array(fc.constantFrom("Author A", "Author B", "Author C"), {
          minLength: 2,
          maxLength: 5,
        }),
        (authors) => {
          const original = normalizeForOps({
            package: {
              document: {
                content: [
                  {
                    type: "paragraph",
                    paraId: "00000001",
                    content: [
                      {
                        type: "run",
                        formatting: { bold: true },
                        content: [{ type: "text", text: "a" }],
                        propertyChanges: [
                          {
                            type: "runPropertyChange",
                            info: { id: 1, author: "Earlier A" },
                            previousFormatting: { italic: true },
                          },
                        ],
                      },
                      {
                        type: "run",
                        formatting: { italic: true },
                        content: [{ type: "text", text: "b" }],
                        propertyChanges: [
                          {
                            type: "runPropertyChange",
                            info: { id: 2, author: "Earlier B" },
                            previousFormatting: { bold: true },
                          },
                        ],
                      },
                    ],
                  },
                ],
              },
            },
          } satisfies Document);
          const at = { story: OP_STORIES.MAIN, blockId: "00000001", offset: 0 };
          let document = original;
          const journal: AppliedDocumentOp[] = [];
          for (const [index, author] of authors.entries()) {
            const before = document;
            const intent = {
              type: "formatRun",
              from: at,
              to: { ...at, offset: 2 },
              patch: { underline: { style: index % 2 === 0 ? "single" : "double" } },
            } as const satisfies EditorIntent;
            const ids = allocateEditorIntentIds(document, intent);
            const compiled = compileEditorIntent(document, {
              intent,
              mode: {
                type: "suggesting",
                revision: { id: ids.revisionId, author, date: "2026-10-02T00:00:00Z" },
                newIds: ids.newIds,
              },
            }).unwrap();
            const applied = applyDocumentOps(document, compiled.ops).unwrap();
            expect(applied.revisions).toHaveLength(2);
            for (const paragraph of storyParagraphs(applied.document.package.document)) {
              for (const run of paragraph.paragraph.content) {
                if (run.type !== "run") panic("Formatting changed the authored run structure.");
                expect(run.propertyChanges?.at(-1)?.info.author).toBe(author);
                expect(run.propertyChanges).toHaveLength(index + 2);
              }
            }
            const rejected = applyDocumentOp(applied.document, {
              type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
              story: OP_STORIES.MAIN,
              revisionIds: applied.revisions,
              decision: REVISION_DECISIONS.REJECT,
            }).unwrap();
            expect(rejected.document).toStrictEqual(before);
            expectRestores(applied, before);
            document = applied.document;
            journal.push(applied);
          }
          const final = document;
          const redo: (readonly DocumentOp[])[] = [];
          for (const entry of journal.toReversed()) {
            const undone = applyDocumentOps(document, entry.inverse).unwrap();
            redo.push(undone.inverse);
            document = undone.document;
          }
          expect(document).toStrictEqual(original);
          for (const ops of redo.toReversed())
            document = applyDocumentOps(document, ops).unwrap().document;
          expect(document).toStrictEqual(final);
        },
      ),
      { numRuns: 30 },
    );
  });
  test("run formatting intent patches preserve authored fields in direct and tracked insertion", () => {
    assertProperty(
      fc.property(fc.boolean(), fc.boolean(), (atom, italic) => {
        const formatting = {
          bold: true,
          boldCs: false,
          noProof: true,
          fontFamily: { ascii: "Folio Custom", cs: "Folio Script" },
          language: { val: "fr-FR", bidi: "ar-SA" },
        };
        const document = normalizeForOps({
          package: {
            document: {
              content: [
                {
                  type: "paragraph",
                  paraId: "00000001",
                  content: [
                    { type: "run", formatting, content: [{ type: "text", text: "authored" }] },
                  ],
                },
              ],
            },
          },
        } satisfies Document);
        const at = { story: OP_STORIES.MAIN, blockId: "00000001", offset: 0 };
        const runPropsPatch = { bold: null, italic } as const;
        const intent = atom
          ? ({
              type: "insertAtom",
              from: at,
              to: at,
              atom: { type: "tab" },
              runPropsPatch,
            } as const satisfies EditorIntent)
          : ({
              type: "replaceText",
              from: at,
              to: at,
              text: "X",
              runPropsPatch,
            } as const satisfies EditorIntent);
        const allocation = allocateEditorIntentIds(document, intent);
        const outcomes = [
          { type: "editing", newIds: allocation.newIds } as const,
          {
            type: "suggesting",
            revision: {
              id: allocation.revisionId,
              author: "Property",
              date: "2026-10-02T00:00:00Z",
            },
            newIds: allocation.newIds,
          } as const,
        ].map((mode) => {
          const compiled = compileEditorIntent(document, { intent, mode }).unwrap();
          const insertion = compiled.ops.findLast(
            (op) => op.type === DOCUMENT_OP_TYPES.INSERT_CONTENT,
          );
          if (insertion?.type !== DOCUMENT_OP_TYPES.INSERT_CONTENT)
            panic("An insertion intent must author content.");
          const authored = insertion.slice.content.at(0);
          if (authored?.type !== "run") panic("An insertion intent must author a run.");
          expect(authored.formatting).toStrictEqual({
            boldCs: false,
            noProof: true,
            fontFamily: formatting.fontFamily,
            language: formatting.language,
            italic,
          });
          const applied = applyDocumentOps(document, compiled.ops).unwrap();
          expect(
            applyDocumentOps(applied.document, applied.inverse).unwrap().document,
          ).toStrictEqual(document);
          if (mode.type === "editing") return applied.document;
          expect(
            applyDocumentOp(applied.document, {
              type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
              story: OP_STORIES.MAIN,
              revisionIds: applied.revisions,
              decision: REVISION_DECISIONS.REJECT,
            }).unwrap().document,
          ).toStrictEqual(document);
          return applyDocumentOp(applied.document, {
            type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
            story: OP_STORIES.MAIN,
            revisionIds: applied.revisions,
            decision: REVISION_DECISIONS.ACCEPT,
          }).unwrap().document;
        });
        expect(outcomes.at(1)).toStrictEqual(outcomes.at(0));
      }),
      { numRuns: 20 },
    );
  });
  test("editor intent sequences preserve accepted editing, rejected baseline and exact journal undo", () => {
    const kinds = [
      "insert",
      "delete",
      "replace",
      "split",
      "join",
      "formatRun",
      "formatParagraph",
      "setList",
    ] as const;
    const tally = new Set<string>();
    assertProperty(
      fc.property(
        documentArbitrary,
        plainInputArbitrary,
        fc.array(
          fc.record({
            kind: fc.constantFrom(...kinds),
            anchor: fc.nat(),
            head: fc.nat(),
            paragraph: fc.nat(),
            text: plainInputArbitrary,
          }),
          { minLength: 8, maxLength: 24 },
        ),
        (generated, text, inputs) => {
          const sectionProperties = generated.package.document.content.find(
            (item) => item.type === "paragraph",
          )?.sectionProperties;
          const original = normalizeForOps({
            ...generated,
            package: {
              ...generated.package,
              numbering: {
                nums: [{ numId: 1, abstractNumId: 1 }],
                abstractNums: [
                  { abstractNumId: 1, levels: [{ ilvl: 0, numFmt: "decimal", lvlText: "%1." }] },
                ],
              },
              document: {
                ...generated.package.document,
                sections: undefined,
                content: ["00000001", "00000002"].map((paraId, index) => {
                  const paragraph = {
                    type: "paragraph",
                    paraId,
                    formatting: {
                      alignment: index === 0 ? "start" : "end",
                      runProperties: { italic: index === 0 },
                      runInWithNext: index === 0,
                    },
                    content: [
                      {
                        type: "run",
                        formatting: index === 0 ? { bold: true } : { italic: true },
                        content: [{ type: "text", text }],
                      },
                    ],
                  } satisfies Paragraph;
                  // Model an absent XML property by omission, as the parser does.
                  return index === 0 && sectionProperties !== undefined
                    ? Object.assign(paragraph, { sectionProperties })
                    : paragraph;
                }),
              },
            },
          });
          let direct = original;
          let tracked = original;
          const journal: AppliedDocumentOp[] = [];
          for (const input of inputs) {
            const paragraphs = storyParagraphs(direct.package.document).map(
              ({ paragraph }) => paragraph,
            );
            const source = paragraphs.at(input.paragraph % paragraphs.length);
            if (source === undefined) panic("Generated editor paragraph disappeared.");
            const group = editorParagraphGroups(tracked, OP_STORIES.MAIN).find(
              ({ blockId }) => blockId === source.paraId,
            );
            if (group === undefined) panic("Generated editor tracked paragraph disappeared.");
            const visible = paragraphLogicalText(source);
            expect(group.text).toBe(visible);
            const gaps = codePointGaps(visible);
            const anchor = gaps.at(input.anchor % gaps.length) ?? 0;
            const head = gaps.at(input.head % gaps.length) ?? 0;
            const blockId = source.paraId ?? "";
            const position = (offset: number) => ({ story: OP_STORIES.MAIN, blockId, offset });
            let intent: EditorIntent;
            let suggested: EditorIntent;
            if (input.kind === "split") {
              intent = {
                type: "splitParagraph",
                at: position(Math.min(anchor, head)),
                to: position(Math.max(anchor, head)),
                newBlockId: "00000003",
              };
              suggested = {
                ...intent,
                at: physicalPositionAtEditorOffset(tracked, intent.at),
                to: physicalPositionAtEditorOffset(tracked, intent.to ?? intent.at),
              };
            } else if (input.kind === "formatRun") {
              intent = {
                type: "formatRun",
                from: position(Math.min(anchor, head)),
                to: position(Math.max(anchor, head)),
                patch: { bold: input.anchor % 2 === 0, italic: input.head % 2 === 0 },
              };
              suggested = {
                ...intent,
                from: physicalPositionAtEditorOffset(tracked, intent.from),
                to: physicalPositionAtEditorOffset(tracked, intent.to),
              };
            } else if (input.kind === "formatParagraph") {
              intent = {
                type: "formatParagraph",
                at: position(anchor),
                patch: {
                  alignment: input.anchor % 2 === 0 ? "center" : "right",
                  keepNext: input.head % 2 === 0,
                },
              };
              suggested = intent;
            } else if (input.kind === "setList") {
              intent = {
                type: "setList",
                items: [{ at: position(anchor), ilvl: 0 }],
                target: { type: "existing", numId: 1 },
              };
              suggested = intent;
            } else if (input.kind === "join" && paragraphs.length > 1) {
              // Join accepted-view neighbors even when prior deletions leave
              // several physical paragraphs in either group.
              const joinIndex = input.paragraph % (paragraphs.length - 1);
              const first = paragraphs.at(joinIndex);
              const second = paragraphs.at(joinIndex + 1);
              if (first === undefined || second === undefined)
                panic("Generated join has no endpoints.");
              intent = {
                type: "joinParagraphs",
                story: OP_STORIES.MAIN,
                blockId: first.paraId ?? "",
                nextBlockId: second.paraId ?? "",
              };
              const firstGroup = editorParagraphGroups(tracked, OP_STORIES.MAIN).find(
                ({ blockId: groupBlockId }) => groupBlockId === first.paraId,
              );
              const secondGroup = editorParagraphGroups(tracked, OP_STORIES.MAIN).find(
                ({ blockId: groupBlockId }) => groupBlockId === second.paraId,
              );
              suggested = {
                ...intent,
                blockId: firstGroup?.paragraphs.at(-1)?.paraId ?? "",
                nextBlockId: secondGroup?.paragraphs.at(0)?.paraId ?? "",
              };
            } else {
              const from = input.kind === "insert" ? anchor : Math.min(anchor, head);
              const to = input.kind === "insert" ? anchor : Math.max(anchor, head);
              intent = {
                type: "replaceText",
                from: position(from),
                to: position(to),
                text: input.kind === "delete" ? "" : input.text,
              };
              suggested = {
                ...intent,
                from: physicalPositionAtEditorOffset(tracked, position(from)),
                to: physicalPositionAtEditorOffset(tracked, position(to)),
              };
            }
            const ids = allocateEditorIntentIds(tracked, suggested);
            if (intent.type === "splitParagraph" && suggested.type === "splitParagraph") {
              intent = { ...intent, newBlockId: ids.newBlockId };
              suggested = { ...suggested, newBlockId: ids.newBlockId };
            }
            const editPlan = compileEditorIntent(direct, { intent, mode: { type: "editing" } });
            const trackedPlan = compileEditorIntent(tracked, {
              intent: suggested,
              mode: {
                type: "suggesting",
                revision: { id: ids.revisionId, author: "Editor", date: "2026-10-01T12:00:00Z" },
                newIds: ids.newIds,
              },
            });
            if (editPlan.isErr()) throw editPlan.error;
            if (trackedPlan.isErr()) throw trackedPlan.error;
            expectCompactIntentIds(direct, editPlan.value.ops);
            expectCompactIntentIds(tracked, trackedPlan.value.ops);
            const edited = applyDocumentOps(direct, editPlan.value.ops);
            const suggestedEdit = applyDocumentOps(tracked, trackedPlan.value.ops);
            if (edited.isErr()) throw edited.error;
            if (suggestedEdit.isErr()) throw suggestedEdit.error;
            expectRestores(suggestedEdit.value, tracked);
            journal.push(suggestedEdit.value);
            direct = edited.value.document;
            tracked = suggestedEdit.value.document;
            tally.add(
              intent.type === "replaceText" && input.kind === "join" ? "replace" : input.kind,
            );
            expect(contractViolation(tracked)).toBeUndefined();
            const keys = packageIdentityKeys(tracked.package);
            expect(new Set(keys).size).toBe(keys.length);
            const paragraphIds = paragraphIdsIn({
              ...tracked.package,
              document: { ...tracked.package.document, sections: undefined },
            });
            expect(new Set(paragraphIds).size).toBe(paragraphIds.length);
            const invalid = applyDocumentOps(tracked, [
              {
                type: DOCUMENT_OP_TYPES.INSERT_TEXT,
                at: { story: OP_STORIES.MAIN, blockId: "7FFFFFFE", offset: 0 },
                text: "invalid",
              },
            ]);
            if (invalid.isOk()) panic("An absent-block insertion unexpectedly applied.");
            expect(invalid.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND);
          }
          const prefix = `${IDENTITY_SPACES.REVISION}:`;
          const revisionIds = identityKeysIn({
            ...tracked.package,
            document: { ...tracked.package.document, sections: undefined },
          })
            .filter((key) => key.startsWith(prefix))
            .map((key) => Number(key.slice(prefix.length)));
          const resolve = (
            decision: (typeof REVISION_DECISIONS)[keyof typeof REVISION_DECISIONS],
          ) => {
            const result = applyDocumentOps(tracked, [
              {
                type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
                story: OP_STORIES.MAIN,
                revisionIds,
                decision,
              },
            ]);
            if (result.isErr()) throw result.error;
            expectRestores(result.value, tracked);
            return result.value.document;
          };
          expect(
            projectReview({ document: resolve(REVISION_DECISIONS.ACCEPT), projection: "π" }),
          ).toStrictEqual(projectReview({ document: direct, projection: "π" }));
          expect(
            projectReview({ document: resolve(REVISION_DECISIONS.REJECT), projection: "π" }),
          ).toStrictEqual(projectReview({ document: original, projection: "π" }));
          for (const entry of journal.toReversed()) tracked = applyAll(tracked, entry.inverse);
          expect(tracked).toStrictEqual(original);
        },
      ),
      { numRuns: 300 },
    );
    expect([...tally].sort()).toEqual([...kinds].sort());
  });

  test("pasted bookmark identities and copied link targets stay together", () => {
    const content = [
      { type: "bookmarkStart", id: 4, name: "target" },
      { type: "run", content: [{ type: "text", text: "source" }] },
      { type: "bookmarkEnd", id: 4 },
    ] satisfies Paragraph["content"];
    const document: Document = {
      package: { document: { content: [{ type: "paragraph", paraId: "00000001", content }] } },
      warnings: [],
    };
    const at = { story: OP_STORIES.MAIN, blockId: "00000001", offset: 6 };
    const compiled = compileEditorIntent(document, {
      intent: {
        type: "replaceFragment",
        from: at,
        to: at,
        openStart: 1,
        openEnd: 1,
        paragraphs: [
          {
            type: "paragraph",
            paraId: "00000001",
            content: [
              ...content,
              {
                type: "hyperlink",
                anchor: "target",
                children: [{ type: "run", content: [{ type: "text", text: "link" }] }],
              },
            ],
          },
        ],
      },
      mode: { type: "editing" },
    });
    if (compiled.isErr()) throw compiled.error;
    const insertion = compiled.value.ops.find((op) => op.type === DOCUMENT_OP_TYPES.INSERT_CONTENT);
    if (insertion?.type !== DOCUMENT_OP_TYPES.INSERT_CONTENT)
      panic("A clipboard fragment must insert its content.");
    expect(insertion.slice.content).toContainEqual({
      type: "bookmarkStart",
      id: 1,
      name: "target_paste1",
    });
    expect(insertion.slice.content).toContainEqual({ type: "bookmarkEnd", id: 1 });
    expect(insertion.slice.content).toContainEqual({
      type: "hyperlink",
      anchor: "target_paste1",
      children: [{ type: "run", content: [{ type: "text", text: "link" }] }],
    });
  });
  test("marker-only cut and move distinguish explicit gaps at equal text offsets", () => {
    const original = normalizeForOps({
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              paraId: "00000001",
              content: [
                { type: "run", content: [{ type: "text", text: "a" }] },
                { type: "bookmarkStart", id: 1, name: "A" },
                { type: "bookmarkEnd", id: 1 },
                { type: "bookmarkStart", id: 2, name: "B" },
                { type: "bookmarkEnd", id: 2 },
                { type: "run", content: [{ type: "text", text: "b" }] },
              ],
            },
          ],
        },
      },
    } satisfies Document);
    const from = { story: OP_STORIES.MAIN, blockId: "00000001", offset: 1, zeroWidthBefore: 0 };
    const to = { ...from, zeroWidthBefore: 2 };
    const copied = [
      {
        type: "paragraph",
        content: [
          { type: "bookmarkStart", id: 1, name: "A" },
          { type: "bookmarkEnd", id: 1 },
        ],
      },
    ] satisfies Paragraph[];
    for (const intent of [
      { type: "replaceText", from, to, text: "" },
      {
        type: "moveFragment",
        from,
        to,
        target: { ...from, zeroWidthBefore: 4 },
        paragraphs: copied,
        openStart: 1,
        openEnd: 1,
      },
    ] as const satisfies readonly EditorIntent[]) {
      const allocation = allocateEditorIntentIds(original, intent);
      const plan = compileEditorIntent(original, {
        intent,
        mode: { type: "editing", newIds: allocation.newIds },
      }).unwrap();
      expect(plan.ops.length).toBeGreaterThan(0);
      const edit = applyDocumentOps(original, plan.ops).unwrap();
      const paragraph = storyParagraphs(edit.document.package.document).at(0)?.paragraph;
      if (!paragraph) panic("Marker operation lost its paragraph.");
      expect(paragraphLogicalText(paragraph)).toBe("ab");
      const names = paragraph.content.flatMap((item) =>
        item.type === "bookmarkStart" ? [item.name] : [],
      );
      expect(names).toEqual(intent.type === "replaceText" ? ["B"] : ["B", "A"]);
      const undo = applyDocumentOps(edit.document, edit.inverse).unwrap();
      assertExactModel(undo.document, original);
      const redo = applyDocumentOps(undo.document, undo.inverse).unwrap();
      assertExactModel(redo.document, edit.document);
    }
  });
  test("clipboard paragraph allocation handles the final available id and exhausted inline paste atomically", () => {
    for (const scenario of [
      { existing: "7FFFFFFF", firstBlockId: 0x80000000, count: 1, allowed: true },
      { existing: "7FFFFFFE", firstBlockId: 0x7fffffff, count: 2, allowed: true },
      { existing: "7FFFFFFE", firstBlockId: 0x7fffffff, count: 3, allowed: false },
      { existing: "7FFFFFFF", firstBlockId: 0x80000000, count: 2, allowed: false },
    ]) {
      const original = normalizeForOps({
        package: {
          document: {
            content: [
              {
                type: "paragraph",
                paraId: scenario.existing,
                content: [{ type: "run", content: [{ type: "text", text: "ab" }] }],
              },
            ],
          },
        },
      } satisfies Document);
      const at = { story: OP_STORIES.MAIN, blockId: scenario.existing, offset: 1 };
      const intent = {
        type: "replaceFragment",
        from: at,
        to: at,
        openStart: 1,
        openEnd: 1,
        paragraphs: Array.from(
          { length: scenario.count },
          (_, index) =>
            ({
              type: "paragraph",
              paraId: "00000001",
              content: [{ type: "run", content: [{ type: "text", text: `paste${index}` }] }],
            }) satisfies Paragraph,
        ),
      } as const satisfies EditorIntent;
      const snapshot = structuredClone(original);
      const plan = compileEditorIntent(original, {
        intent,
        mode: { type: "editing" },
        firstBlockId: scenario.firstBlockId,
      });
      assertExactModel(original, snapshot);
      expect(plan.isOk()).toBe(scenario.allowed);
      if (plan.isErr()) {
        expect(plan.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.INVALID_BLOCK_ID);
      } else {
        const edit = applyDocumentOps(original, plan.value.ops).unwrap();
        const ids = paragraphIdsIn(edit.document.package.document);
        expect(ids).toHaveLength(scenario.count);
        expect(ids).toContain(scenario.existing);
        expect(ids).not.toContain("00000001");
        if (scenario.count === 2) expect(ids).toContain("7FFFFFFF");
        const undo = applyDocumentOps(edit.document, edit.inverse).unwrap();
        assertExactModel(undo.document, original);
        const redo = applyDocumentOps(undo.document, undo.inverse).unwrap();
        assertExactModel(redo.document, edit.document);
      }
    }
  });
  test("clipboard paragraph review and section geometry survive closed insertion while open edges retain destination facts", () => {
    const original = normalizeForOps({
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              paraId: "00000001",
              formatting: { alignment: "right" },
              content: [{ type: "run", content: [{ type: "text", text: "ab" }] }],
            },
          ],
        },
      },
    } satisfies Document);
    const sectionProperties = { pageWidth: 12240, pageHeight: 15840 } satisfies SectionProperties;
    const paragraphs = [
      {
        type: "paragraph",
        paraId: "00000001",
        formatting: { alignment: "center" },
        propertyChanges: [
          {
            type: "paragraphPropertyChange",
            info: { id: 12, author: "Source" },
            previousFormatting: { alignment: "end" },
          },
        ],
        sectionProperties,
        content: [{ type: "run", content: [{ type: "text", text: "clip" }] }],
      },
    ] satisfies Paragraph[];
    const at = { story: OP_STORIES.MAIN, blockId: "00000001", offset: 1 };
    for (const openStart of [0, 1] as const)
      for (const openEnd of [0, 1] as const) {
        const intent = {
          type: "replaceFragment",
          from: at,
          to: at,
          paragraphs,
          openStart,
          openEnd,
        } as const satisfies EditorIntent;
        const allocation = allocateEditorIntentIds(original, intent);
        const plan = compileEditorIntent(original, {
          intent,
          mode: { type: "editing", newIds: allocation.newIds },
        }).unwrap();
        const edit = applyDocumentOps(original, plan.ops).unwrap();
        const pasted = storyParagraphs(edit.document.package.document)
          .map(({ paragraph }) => paragraph)
          .find((paragraph) => paragraphLogicalText(paragraph).includes("clip"));
        if (!pasted) panic("Clipboard metadata fixture disappeared.");
        if (openStart === 0 && openEnd === 0) {
          expect(pasted.formatting?.alignment).toBe("center");
          expect(pasted.sectionProperties).toEqual(sectionProperties);
          expect(pasted.propertyChanges?.at(0)?.info.id).not.toBe(12);
          expect(pasted.propertyChanges?.at(0)?.info.author).toBe("Source");
          expect(pasted.propertyChanges?.at(0)?.previousFormatting.alignment).toBe("end");
        } else {
          expect(pasted.propertyChanges).toBeUndefined();
          expect(pasted.sectionProperties).toBeUndefined();
          expect(pasted.formatting?.alignment).toBe("right");
        }
        const undo = applyDocumentOps(edit.document, edit.inverse).unwrap();
        assertExactModel(undo.document, original);
        const redo = applyDocumentOps(undo.document, undo.inverse).unwrap();
        assertExactModel(redo.document, edit.document);
      }
  });
  test("tracked clipboard preserves leading reviews and geometry and refuses combined paragraph marks atomically", () => {
    const original = normalizeForOps({
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              paraId: "00000001",
              formatting: { alignment: "right" },
              content: [{ type: "run", content: [{ type: "text", text: "a" }] }],
            },
          ],
        },
      },
    } satisfies Document);
    const at = { story: OP_STORIES.MAIN, blockId: "00000001", offset: 1 };
    for (const metadata of ["review", "section", "mark"] as const) {
      const copied = {
        type: "paragraph",
        content: [{ type: "run", content: [{ type: "text", text: "metadata" }] }],
        ...(metadata === "review"
          ? {
              propertyChanges: [
                {
                  type: "paragraphPropertyChange",
                  info: { id: 12, author: "Source" },
                  previousFormatting: { alignment: "center" },
                },
              ],
            }
          : {}),
        ...(metadata === "section"
          ? { sectionProperties: { pageWidth: 12240, pageHeight: 15840 } }
          : {}),
        ...(metadata === "mark"
          ? { pPrMark: { kind: "ins", info: { id: 12, author: "Source" } } }
          : {}),
      } satisfies Paragraph;
      for (const leading of [false, true]) {
        const intent = {
          type: "replaceFragment",
          from: at,
          to: at,
          openStart: 0,
          openEnd: 0,
          paragraphs: leading
            ? [
                copied,
                {
                  type: "paragraph",
                  content: [{ type: "run", content: [{ type: "text", text: "tail" }] }],
                },
              ]
            : [copied],
        } as const satisfies EditorIntent;
        const allocation = allocateEditorIntentIds(original, intent);
        const snapshot = structuredClone(original);
        const plan = compileEditorIntent(original, {
          intent,
          mode: {
            type: "suggesting",
            revision: {
              id: allocation.revisionId,
              author: "Clipboard",
              date: "2026-10-02T00:00:00Z",
            },
            newIds: allocation.newIds,
          },
        });
        assertExactModel(original, snapshot);
        expect(plan.isOk()).toBe(leading && metadata !== "mark");
        if (plan.isErr()) {
          expect(plan.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE);
        } else {
          const edit = applyDocumentOps(original, plan.value.ops).unwrap();
          const undo = applyDocumentOps(edit.document, edit.inverse).unwrap();
          assertExactModel(undo.document, original);
          const redo = applyDocumentOps(undo.document, undo.inverse).unwrap();
          assertExactModel(redo.document, edit.document);
        }
      }
    }
  });

  test("generated clipboard fragments preserve history and accepted tracked content", () => {
    assertProperty(
      fc.property(
        fc.array(fc.string({ minLength: 1, maxLength: 8 }), { minLength: 1, maxLength: 4 }),
        fc.constantFrom(0 as const, 1 as const),
        fc.constantFrom(0 as const, 1 as const),
        fc.boolean(),
        (texts, openStart, openEnd, replace) => {
          const document: Document = {
            package: {
              document: {
                content: [
                  {
                    type: "paragraph",
                    paraId: "00000001",
                    formatting: { alignment: "right" },
                    content: [
                      {
                        type: "run",
                        formatting: { italic: true },
                        content: [{ type: "text", text: "prefixsuffix" }],
                      },
                    ],
                  },
                ],
              },
            },
            warnings: [],
          };
          const paragraphs = texts.map(
            (text) =>
              ({
                type: "paragraph",
                paraId: "00000001",
                formatting: { alignment: "center" },
                content: [
                  { type: "run", formatting: { bold: true }, content: [{ type: "text", text }] },
                ],
              }) satisfies Paragraph,
          );
          const from = { story: OP_STORIES.MAIN, blockId: "00000001", offset: 6 };
          const intent = {
            type: "replaceFragment",
            from,
            to: { ...from, offset: replace ? 9 : 6 },
            paragraphs,
            openStart,
            openEnd,
          } as const satisfies EditorIntent;
          const allocation = allocateEditorIntentIds(document, intent);
          const outcomes = [
            { type: "editing", newIds: allocation.newIds } as const,
            {
              type: "suggesting",
              revision: {
                id: allocation.revisionId,
                author: "Property",
                date: "2026-10-02T00:00:00Z",
              },
              newIds: allocation.newIds,
            } as const,
          ].map((mode) => {
            const compiled = compileEditorIntent(document, { intent, mode });
            if (compiled.isErr()) throw compiled.error;
            const applied = applyDocumentOps(document, compiled.value.ops);
            if (applied.isErr()) throw applied.error;
            const undo = applyDocumentOps(applied.value.document, applied.value.inverse);
            if (undo.isErr()) throw undo.error;
            assertExactModel(undo.value.document, document);
            const redo = applyDocumentOps(undo.value.document, undo.value.inverse);
            if (redo.isErr()) throw redo.error;
            assertExactModel(redo.value.document, applied.value.document);
            const expectedCount =
              texts.length + (openStart === 0 ? 1 : 0) + (openEnd === 0 ? 1 : 0);
            expect(paragraphIdsIn(applied.value.document.package.document)).toHaveLength(
              expectedCount,
            );
            expect(new Set(paragraphIdsIn(applied.value.document.package.document)).size).toBe(
              expectedCount,
            );
            const failed = applyDocumentOps(document, [
              ...compiled.value.ops,
              {
                type: DOCUMENT_OP_TYPES.INSERT_TEXT,
                at: { ...from, blockId: "FFFFFFFF" },
                text: "invalid",
              },
            ]);
            expect(failed.isErr()).toBe(true);
            expect(
              storyParagraphs(document.package.document).map(({ paragraph }) =>
                paragraphLogicalText(paragraph),
              ),
            ).toEqual(["prefixsuffix"]);
            if (mode.type === "editing") return applied.value.document;
            const reject = applyDocumentOp(applied.value.document, {
              type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
              story: OP_STORIES.MAIN,
              revisionIds: applied.value.revisions,
              decision: REVISION_DECISIONS.REJECT,
            });
            if (reject.isErr()) throw reject.error;
            expect(reject.value.document).toStrictEqual(document);
            const accept = applyDocumentOp(applied.value.document, {
              type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
              story: OP_STORIES.MAIN,
              revisionIds: applied.value.revisions,
              decision: REVISION_DECISIONS.ACCEPT,
            });
            if (accept.isErr()) throw accept.error;
            return accept.value.document;
          });
          expect(outcomes[1]).toStrictEqual(outcomes[0]);
          expect(
            storyParagraphs(outcomes[0]!.package.document)
              .map(({ paragraph }) => paragraphLogicalText(paragraph))
              .join("\n"),
          ).toBe(
            "prefix" +
              (openStart === 0 ? "\n" : "") +
              texts.join("\n") +
              (openEnd === 0 ? "\n" : "") +
              (replace ? "fix" : "suffix"),
          );
        },
      ),
      { numRuns: 100 },
    );
  });
  test("id-only copied controls and retained identity provenance remint one consistent imported graph", () => {
    const destination = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              paraId: "00000001",
              content: [
                {
                  type: "inlineSdt",
                  properties: { id: 7 },
                  content: [{ type: "run", content: [{ type: "text", text: "owned" }] }],
                },
                {
                  type: "insertion",
                  info: { id: 10, author: "Owned" },
                  content: [{ type: "run", content: [{ type: "text", text: "review" }] }],
                },
              ],
            },
          ],
        },
      },
    } satisfies Document;
    const paragraphs = [
      {
        type: "paragraph",
        paraId: "00000001",
        content: [
          {
            type: "inlineSdt",
            properties: { id: 7 },
            content: [
              {
                type: "insertion",
                info: { id: 10, author: "Foreign" },
                resolutionJoins: {
                  before: 0,
                  after: 0,
                  remove: 0,
                  retainedAfter: [
                    {
                      depth: 1,
                      source: [
                        { space: "control", id: 7 },
                        { space: "revision", id: 10 },
                      ],
                      target: [
                        { space: "control", id: 8 },
                        { space: "revision", id: 11 },
                      ],
                    },
                  ],
                },
                content: [{ type: "run", content: [{ type: "text", text: "first" }] }],
              },
            ],
          },
          {
            type: "inlineSdt",
            properties: { id: 8 },
            content: [
              {
                type: "insertion",
                info: { id: 11, author: "Foreign" },
                content: [{ type: "run", content: [{ type: "text", text: "second" }] }],
              },
            ],
          },
        ],
      },
    ] satisfies Paragraph[];
    const intent = {
      type: "replaceFragment",
      from: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 0 },
      to: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 0 },
      paragraphs,
      openStart: 0,
      openEnd: 0,
    } as const satisfies EditorIntent;
    const original = structuredClone(destination);
    const incoming = structuredClone(paragraphs);
    const plan = compileEditorIntent(destination, { intent, mode: { type: "editing" } }).unwrap();
    const edit = applyDocumentOps(destination, plan.ops).unwrap();
    const copied = storyParagraphs(edit.document.package.document).find(
      ({ paragraph }) => paragraphLogicalText(paragraph) === "firstsecond",
    )?.paragraph;
    const first = copied?.content.at(0);
    const second = copied?.content.at(1);
    if (first?.type !== "inlineSdt" || second?.type !== "inlineSdt")
      panic("Imported id-only controls disappeared.");
    const firstReview = first.content.at(0);
    const secondReview = second.content.at(0);
    if (firstReview?.type !== "insertion" || secondReview?.type !== "insertion")
      panic("Imported control reviews disappeared.");
    expect(first.properties.id).not.toBe(7);
    expect(second.properties.id).not.toBe(8);
    expect(first.properties.id).not.toBe(second.properties.id);
    expect(firstReview.info.id).not.toBe(10);
    expect(secondReview.info.id).not.toBe(11);
    assertExactModel(firstReview.resolutionJoins?.retainedAfter, [
      {
        depth: 1,
        source: [
          { space: "control", id: first.properties.id },
          { space: "revision", id: firstReview.info.id },
        ],
        target: [
          { space: "control", id: second.properties.id },
          { space: "revision", id: secondReview.info.id },
        ],
      },
    ]);
    const undo = applyDocumentOps(edit.document, edit.inverse).unwrap();
    assertExactModel(undo.document, original);
    const redo = applyDocumentOps(undo.document, undo.inverse).unwrap();
    assertExactModel(redo.document, edit.document);
    assertExactModel(paragraphs, incoming);
  });

  test("generated clipboard sequences preserve arbitrary ranges, marked content and every inverse", () => {
    const applied = new Map<string, number>();
    const refusals = new Map<DocumentOpRefusalReason, number>();
    const payloadText = fc.string({
      unit: fc.constantFrom("x", "é", "😀"),
      minLength: 1,
      maxLength: 5,
    });
    assertProperty(
      fc.property(
        fc.array(
          fc.record({
            source: fc.nat(),
            end: fc.nat(),
            anchor: fc.nat(),
            head: fc.nat(),
            openStart: fc.constantFrom(0 as const, 1 as const),
            openEnd: fc.constantFrom(0 as const, 1 as const),
            texts: fc.array(payloadText, { minLength: 1, maxLength: 3 }),
            bold: fc.boolean(),
            italic: fc.boolean(),
            resourceByte: fc.integer({ min: 0, max: 255 }),
          }),
          { minLength: 8, maxLength: 16 },
        ),
        fc.boolean(),
        fc.boolean(),
        (steps, reverseDestinationFormatting, priorReview) => {
          const original = normalizeForOps({
            package: {
              document: {
                content: [
                  {
                    type: "paragraph",
                    paraId: "00000001",
                    formatting: { alignment: reverseDestinationFormatting ? "right" : "left" },
                    content: [
                      {
                        type: "run",
                        formatting: { italic: true },
                        content: [{ type: "text", text: "first😀" }],
                      },
                    ],
                  },
                  {
                    type: "paragraph",
                    paraId: "00000002",
                    formatting: { alignment: reverseDestinationFormatting ? "left" : "right" },
                    content: [{ type: "run", content: [{ type: "text", text: "secondé" }] }],
                  },
                ],
              },
            },
          } satisfies Document);
          if (priorReview) {
            const paragraph = original.package.document.content.at(0);
            const run = paragraph?.type === "paragraph" ? paragraph.content.at(0) : undefined;
            if (run?.type !== "run") panic("Prior clipboard review fixture disappeared.");
            run.propertyChanges = [
              {
                type: "runPropertyChange",
                info: { id: 13, author: "Prior" },
                previousFormatting: { italic: false },
                currentFormatting: { italic: true },
              },
            ];
          }
          let document = original;
          const journal: {
            edit: AppliedDocumentOp;
            before: Document;
            ops: readonly DocumentOp[];
          }[] = [];
          const optionalFixture = storyParagraphs(original.package.document).at(1)?.paragraph;
          if (optionalFixture?.formatting === undefined)
            panic("Optional-field fixture disappeared.");
          Object.defineProperty(optionalFixture.formatting, "keepNext", {
            value: undefined,
            enumerable: true,
            configurable: true,
            writable: true,
          });
          for (const [stepIndex, step] of steps.entries()) {
            const paragraphs = storyParagraphs(document.package.document).map(
              ({ paragraph }) => paragraph,
            );
            const sourceIndex = Math.min(
              step.source % paragraphs.length,
              step.end % paragraphs.length,
            );
            const endIndex = Math.max(
              step.source % paragraphs.length,
              step.end % paragraphs.length,
            );
            const source = paragraphs.at(sourceIndex);
            const end = paragraphs.at(endIndex);
            if (!source?.paraId || !end?.paraId) panic("Generated clipboard range disappeared.");
            const sourceText = paragraphLogicalText(source);
            const endText = paragraphLogicalText(end);
            const sourceGaps = codePointGaps(sourceText);
            const endGaps = codePointGaps(endText);
            const anchor = sourceGaps.at(step.anchor % sourceGaps.length) ?? 0;
            const head = endGaps.at(step.head % endGaps.length) ?? 0;
            const fromOffset = sourceIndex === endIndex ? Math.min(anchor, head) : anchor;
            const toOffset = sourceIndex === endIndex ? Math.max(anchor, head) : head;
            const intent = {
              type: "replaceFragment",
              from: { story: OP_STORIES.MAIN, blockId: source.paraId, offset: fromOffset },
              to: { story: OP_STORIES.MAIN, blockId: end.paraId, offset: toOffset },
              paragraphs: step.texts.map(
                (text) =>
                  ({
                    type: "paragraph",
                    paraId: source.paraId,
                    content: [
                      {
                        type: "run",
                        formatting: { bold: step.bold, italic: step.italic },
                        content: [{ type: "text", text: `paste${text}` }],
                      },
                    ],
                  }) satisfies Paragraph,
              ),
              openStart: step.openStart,
              openEnd: step.openEnd,
            } as const satisfies EditorIntent;
            const allocation = allocateEditorIntentIds(document, intent);
            const before = document;
            const snapshot = structuredClone(before);
            const path = `word/media/clipboard-step${stepIndex}.png`;
            const resourceOp = packageResourcesOpOf({
              before,
              after: {
                ...before,
                package: {
                  ...before.package,
                  media: new Map([
                    ...(before.package.media ?? []),
                    [
                      path,
                      {
                        path,
                        mimeType: "image/png",
                        data: Uint8Array.from([137, 80, 78, 71, step.resourceByte]).buffer,
                      },
                    ],
                  ]),
                },
              },
            });
            const resourceBaseline = applyDocumentOp(before, resourceOp).unwrap().document;
            const outcomes = [
              { type: "editing", newIds: allocation.newIds } as const,
              {
                type: "suggesting",
                newIds: allocation.newIds,
                revision: {
                  id: allocation.revisionId,
                  author: "Clipboard sequence",
                  date: "2026-10-02T00:00:00Z",
                },
              } as const,
            ].map((mode) => {
              const plan = compileEditorIntent(resourceBaseline, { intent, mode });
              if (plan.isErr()) {
                refusals.set(plan.error.reason, (refusals.get(plan.error.reason) ?? 0) + 1);
                throw plan.error;
              }
              const ops = [resourceOp, ...plan.value.ops];
              const edit = applyDocumentOps(before, ops);
              if (edit.isErr()) {
                refusals.set(edit.error.reason, (refusals.get(edit.error.reason) ?? 0) + 1);
                throw edit.error;
              }
              assertExactModel(before, snapshot);
              const undo = applyDocumentOps(edit.value.document, edit.value.inverse);
              if (undo.isErr()) throw undo.error;
              assertExactModel(undo.value.document, before);
              const redo = applyDocumentOps(undo.value.document, undo.value.inverse);
              if (redo.isErr()) throw redo.error;
              assertExactModel(redo.value.document, edit.value.document);
              if (mode.type === "editing") return Object.assign({}, edit.value, { ops });
              const rejected = applyDocumentOp(edit.value.document, {
                type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
                story: OP_STORIES.MAIN,
                revisionIds: edit.value.revisions,
                decision: REVISION_DECISIONS.REJECT,
              });
              if (rejected.isErr()) throw rejected.error;
              assertExactModel(rejected.value.document, resourceBaseline);
              const accepted = applyDocumentOp(edit.value.document, {
                type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
                story: OP_STORIES.MAIN,
                revisionIds: edit.value.revisions,
                decision: REVISION_DECISIONS.ACCEPT,
              });
              if (accepted.isErr()) throw accepted.error;
              return Object.assign({}, edit.value, { ops, document: accepted.value.document });
            });
            const direct = outcomes.at(0);
            const accepted = outcomes.at(1);
            if (!direct || !accepted) panic("Clipboard sequence lost a compilation mode.");
            assertExactModel(accepted.document, direct.document);
            assertExactModel(packageResourcesOf(direct.document), resourceOp.resources);
            assertExactModel(direct.document.package.media, resourceBaseline.package.media);
            const expected = step.texts.map((text) => `paste${text}`);
            const expectedAlignments: NonNullable<Paragraph["formatting"]>["alignment"][] =
              step.texts.map(() => undefined);
            const prefix = sourceText.slice(0, fromOffset);
            const suffix = endText.slice(toOffset);
            if (step.openStart === 1) {
              expected[0] = prefix + expected[0];
              expectedAlignments[0] = source.formatting?.alignment;
            } else if (prefix !== "") {
              expected.unshift(prefix);
              expectedAlignments.unshift(source.formatting?.alignment);
            }
            if (step.openEnd === 1) {
              expected[expected.length - 1] += suffix;
              const leadingWins = step.texts.length === 1 && step.openStart === 1;
              expectedAlignments[expectedAlignments.length - 1] = leadingWins
                ? source.formatting?.alignment
                : end.formatting?.alignment;
            } else if (suffix !== "") {
              expected.push(suffix);
              expectedAlignments.push(end.formatting?.alignment);
            }
            const oracle = [
              ...paragraphs.slice(0, sourceIndex).map(paragraphLogicalText),
              ...expected,
              ...paragraphs.slice(endIndex + 1).map(paragraphLogicalText),
            ];
            expect(
              storyParagraphs(direct.document.package.document).map(({ paragraph }) =>
                paragraphLogicalText(paragraph),
              ),
            ).toEqual(oracle);
            expect(
              storyParagraphs(direct.document.package.document).map(
                ({ paragraph }) => paragraph.formatting?.alignment,
              ),
            ).toEqual([
              ...paragraphs
                .slice(0, sourceIndex)
                .map((paragraph) => paragraph.formatting?.alignment),
              ...expectedAlignments,
              ...paragraphs.slice(endIndex + 1).map((paragraph) => paragraph.formatting?.alignment),
            ]);
            const stale = applyDocumentOps(before, [
              ...direct.ops,
              {
                type: DOCUMENT_OP_TYPES.INSERT_TEXT,
                at: { story: OP_STORIES.MAIN, blockId: "7FFFFFFF", offset: 0 },
                text: "invalid",
              },
            ]);
            if (stale.isOk()) panic("Clipboard sequence accepted a missing target.");
            refusals.set(stale.error.reason, (refusals.get(stale.error.reason) ?? 0) + 1);
            expect(stale.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND);
            assertExactModel(before, snapshot);
            journal.push({ edit: direct, before, ops: direct.ops });
            document = direct.document;
            applied.set(intent.type, (applied.get(intent.type) ?? 0) + 1);
          }
          for (const entry of journal.toReversed()) {
            const undo = applyDocumentOps(document, entry.edit.inverse);
            if (undo.isErr()) throw undo.error;
            assertExactModel(undo.value.document, entry.before);
            document = undo.value.document;
          }
          assertExactModel(document, original);
          for (const entry of journal) {
            const redo = applyDocumentOps(document, entry.ops);
            if (redo.isErr()) throw redo.error;
            assertExactModel(redo.value.document, entry.edit.document);
            document = redo.value.document;
          }
        },
      ),
      { numRuns: 50 },
    );
    expect(applied.get("replaceFragment") ?? 0).toBeGreaterThan(0);
    expect([...refusals.keys()]).toEqual([DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND]);
    expect(refusals.get(DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND) ?? 0).toBeGreaterThan(0);
  });

  test("generated move intents rebase the original target and refuse atomically", () => {
    assertProperty(
      fc.property(fc.integer({ min: 0, max: 12 }), fc.boolean(), (targetOffset, tracked) => {
        const document: Document = {
          package: {
            document: {
              content: [
                {
                  type: "paragraph",
                  paraId: "00000001",
                  content: [{ type: "run", content: [{ type: "text", text: "prefixsuffix" }] }],
                },
              ],
            },
          },
          warnings: [],
        };
        const from = { story: OP_STORIES.MAIN, blockId: "00000001", offset: 3 };
        const target = { ...from, offset: targetOffset };
        const intent = {
          type: "moveFragment",
          from,
          to: { ...from, offset: 6 },
          target,
          openStart: 1,
          openEnd: 1,
          paragraphs: [
            {
              type: "paragraph",
              content: [{ type: "run", content: [{ type: "text", text: "fix" }] }],
            },
          ],
        } as const satisfies EditorIntent;
        const allocation = allocateEditorIntentIds(document, intent);
        const mode = tracked
          ? ({
              type: "suggesting",
              revision: {
                id: allocation.revisionId,
                author: "Property",
                date: "2026-10-02T00:00:00Z",
              },
              newIds: allocation.newIds,
            } as const)
          : ({ type: "editing", newIds: allocation.newIds } as const);
        const compiled = compileEditorIntent(document, { intent, mode });
        if (compiled.isErr()) throw compiled.error;
        const edit = applyDocumentOps(document, compiled.value.ops);
        if (edit.isErr()) throw edit.error;
        const undo = applyDocumentOps(edit.value.document, edit.value.inverse);
        if (undo.isErr()) throw undo.error;
        assertExactModel(undo.value.document, document);
        const redo = applyDocumentOps(undo.value.document, undo.value.inverse);
        if (redo.isErr()) throw redo.error;
        expect(redo.value.document).toStrictEqual(edit.value.document);
        if (targetOffset >= 3 && targetOffset <= 6) {
          expect(compiled.value.ops).toEqual([]);
          expect(edit.value.document).toStrictEqual(document);
          return;
        }
        let accepted = edit.value.document;
        if (tracked) {
          const accept = applyDocumentOp(accepted, {
            type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
            story: OP_STORIES.MAIN,
            revisionIds: edit.value.revisions,
            decision: REVISION_DECISIONS.ACCEPT,
          });
          if (accept.isErr()) throw accept.error;
          accepted = accept.value.document;
          const reject = applyDocumentOp(edit.value.document, {
            type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
            story: OP_STORIES.MAIN,
            revisionIds: edit.value.revisions,
            decision: REVISION_DECISIONS.REJECT,
          });
          if (reject.isErr()) throw reject.error;
          expect(reject.value.document).toStrictEqual(document);
        }
        const rebased = targetOffset > 6 ? targetOffset - 3 : targetOffset;
        const remaining = "presuffix";
        expect(
          storyParagraphs(accepted.package.document)
            .map(({ paragraph }) => paragraphLogicalText(paragraph))
            .join("\n"),
        ).toBe(remaining.slice(0, rebased) + "fix" + remaining.slice(rebased));
        const failed = compileEditorIntent(document, {
          intent: { ...intent, target: { ...target, blockId: "FFFFFFFF" } },
          mode,
        });
        expect(failed.isErr()).toBe(true);
      }),
      { numRuns: 100 },
    );
  });
  test("generated comment-anchor move sequences preserve marker gaps and refuse tracked transfer atomically", () => {
    assertProperty(
      fc.property(
        fc.array(fc.record({ target: fc.nat(), gap: fc.nat(), tracked: fc.boolean() }), {
          minLength: 8,
          maxLength: 16,
        }),
        (steps) => {
          const original = {
            package: {
              document: {
                content: [
                  {
                    type: "paragraph",
                    paraId: "00000001",
                    content: [
                      { type: "run", content: [{ type: "text", text: "a" }] },
                      { type: "commentRangeStart", id: 7 },
                      { type: "commentRangeEnd", id: 7 },
                      { type: "commentReference", id: 7 },
                      { type: "run", content: [{ type: "text", text: "bcd" }] },
                    ],
                  },
                ],
                comments: [{ id: 7, author: "Owner", content: [] }],
              },
            },
          } satisfies Document;
          let document: Document = original;
          const journal: {
            before: Document;
            edit: AppliedDocumentOp;
            ops: readonly DocumentOp[];
          }[] = [];
          const refusals = new Map<DocumentOpRefusalReason, number>();
          let expectedRefusals = 0;
          for (const [index, step] of steps.entries()) {
            const paragraph = storyParagraphs(document.package.document).at(0)?.paragraph;
            if (!paragraph) panic("Comment move sequence lost its paragraph.");
            const anchors = leafSpans(paragraph.content).filter(({ node }) =>
              isCommentAnchor(node),
            );
            const start = anchors.at(0);
            const end = anchors.at(-1);
            if (!start || !end) panic("Comment move sequence lost its anchors.");
            let targetOffset = step.target % (paragraphLength(paragraph) + 1);
            let tracked = step.tracked;
            if (index === 0) {
              targetOffset = 0;
              tracked = false;
            }
            if (index === 1) {
              targetOffset = paragraphLength(paragraph);
              tracked = true;
            }
            const targetGap =
              step.gap % (zeroWidthLeavesAt(paragraph.content, targetOffset).length + 1);
            const at = { story: OP_STORIES.MAIN, blockId: "00000001" };
            const intent = {
              type: "moveFragment",
              from: { ...at, ...start.before },
              to: { ...at, ...end.after },
              target: { ...at, offset: targetOffset, zeroWidthBefore: targetGap },
              openStart: 1,
              openEnd: 1,
              paragraphs: [
                {
                  type: "paragraph",
                  content: [
                    { type: "commentRangeStart", id: 7 },
                    { type: "commentRangeEnd", id: 7 },
                    { type: "commentReference", id: 7 },
                  ],
                },
              ],
            } as const satisfies EditorIntent;
            const allocation = allocateEditorIntentIds(document, intent);
            const mode = tracked
              ? ({
                  type: "suggesting",
                  revision: {
                    id: allocation.revisionId,
                    author: "Property",
                    date: "2026-10-02T00:00:00Z",
                  },
                  newIds: allocation.newIds,
                } as const)
              : ({ type: "editing", newIds: allocation.newIds } as const);
            const before = document;
            const snapshot = independentCopy(before);
            const inside =
              compareGaps(intent.target, start.before) >= 0 &&
              compareGaps(intent.target, end.after) <= 0;
            const compiled = compileEditorIntent(document, { intent, mode });
            if (tracked && !inside) {
              if (compiled.isOk()) panic("Tracked comment transfer unexpectedly compiled.");
              refusals.set(compiled.error.reason, (refusals.get(compiled.error.reason) ?? 0) + 1);
              expectedRefusals += 1;
              expect(compiled.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE);
              assertExactModel(document, snapshot);
            } else {
              if (compiled.isErr()) throw compiled.error;
              const edit = applyDocumentOps(document, compiled.value.ops).unwrap();
              if (inside) expect(compiled.value.ops).toEqual([]);
              const moved = storyParagraphs(edit.document.package.document).at(0)?.paragraph;
              if (!moved) panic("Comment move lost its result paragraph.");
              const text = paragraphLogicalText(paragraph);
              const remaining = text.slice(0, start.before.offset) + text.slice(end.after.offset);
              const rebased = targetOffset > end.after.offset ? targetOffset - 1 : targetOffset;
              expect(paragraphLogicalText(moved)).toBe(
                inside
                  ? text
                  : remaining.slice(0, rebased) +
                      text.slice(start.before.offset, end.after.offset) +
                      remaining.slice(rebased),
              );
              assertExactModel(
                leafSpans(moved.content)
                  .filter(({ node }) => isCommentAnchor(node))
                  .map(({ node }) => node),
                [
                  { type: "commentRangeStart", id: 7 },
                  { type: "commentRangeEnd", id: 7 },
                  { type: "commentReference", id: 7 },
                ],
              );
              const undo = applyDocumentOps(
                edit.document,
                JSON.parse(JSON.stringify(edit.inverse)),
              ).unwrap();
              assertExactModel(undo.document, before);
              const redo = applyDocumentOps(undo.document, undo.inverse).unwrap();
              assertExactModel(redo.document, edit.document);
              journal.push({ before, edit, ops: compiled.value.ops });
              document = redo.document;
            }
          }
          expect([...refusals.keys()]).toEqual([DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE]);
          expect(refusals.get(DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE)).toBe(expectedRefusals);
          for (const entry of journal.toReversed()) {
            document = applyDocumentOps(document, entry.edit.inverse).unwrap().document;
            assertExactModel(document, entry.before);
          }
          assertExactModel(document, original);
          for (const entry of journal) {
            document = applyDocumentOps(document, entry.ops).unwrap().document;
            assertExactModel(document, entry.edit.document);
          }
        },
      ),
      { numRuns: 30 },
    );
  });
  test("plain-text input sequences preserve exact inverse, redo and rejection atomicity", () => {
    const tally = new Map<string, number>();
    fc.assert(
      fc.property(
        documentArbitrary,
        fc.tuple(plainInputArbitrary, plainInputArbitrary).map((parts) => parts.join("")),
        fc.array(replaceInputArbitrary, { minLength: 8, maxLength: 24 }),
        (generated, initialText, inputs) => {
          const first = storyParagraphs(generated.package.document).at(0);
          if (!first?.paragraph.paraId) panic("Generated paragraph id unavailable");
          const blockId = first.paragraph.paraId;
          // Retain the existing generator's package metadata and secondary-story IDs.
          const document = normalizeForOps({
            ...generated,
            package: {
              ...generated.package,
              document: {
                ...generated.package.document,
                sections: undefined,
                content: [
                  {
                    type: "paragraph",
                    paraId: blockId,
                    content: [...initialText].map(
                      (point, index) =>
                        ({
                          type: "run",
                          formatting: index % 2 === 0 ? { bold: true } : { italic: true },
                          content: [{ type: "text", text: point }],
                        }) satisfies Run,
                    ),
                  },
                ],
              },
            },
          });
          const original = structuredClone(document);
          let current = document;
          let text = initialText;
          const journal: AppliedDocumentOp[] = [];
          for (const input of inputs) {
            const gaps = codePointGaps(text);
            const anchor = gaps.at(input.anchor % gaps.length) ?? 0;
            const head = gaps.at(input.head % gaps.length) ?? 0;
            const from = input.kind === "insert" ? anchor : Math.min(anchor, head);
            const to = input.kind === "insert" ? anchor : Math.max(anchor, head);
            const replacement = input.kind === "delete" ? "" : input.text;
            const position = (offset: number) => ({ story: OP_STORIES.MAIN, blockId, offset });
            const source = storyParagraphs(current.package.document).at(0)?.paragraph;
            if (!source) panic("Generated authored input paragraph disappeared");
            const authoredUnit = from === to && from > 0 ? from - 1 : from;
            let runProps: TextFormatting = {};
            let end = 0;
            for (const run of source.content) {
              if (run.type !== "run") panic("Generated authored input encountered a wrapper");
              end += run.content.reduce((width, child) => width + runContentWidth(child), 0);
              if (authoredUnit < end) {
                runProps = run.formatting ?? {};
                break;
              }
            }
            const ops: DocumentOp[] = [];
            if (to > from) {
              ops.push({
                type: DOCUMENT_OP_TYPES.DELETE_RANGE,
                from: position(from),
                to: position(to),
              });
            }
            if (replacement !== "") {
              ops.push({
                type: DOCUMENT_OP_TYPES.INSERT_TEXT,
                at: position(from),
                text: replacement,
                runProps,
              });
            }
            if (input.kind === "reject") {
              // Refuse after a valid insertion in the same batch: partial writes are forbidden.
              ops.push({
                type: DOCUMENT_OP_TYPES.DELETE_RANGE,
                from: position(text.length + replacement.length + 1),
                to: position(text.length + replacement.length + 2),
              });
            }
            const before = structuredClone(current);
            const applied = applyDocumentOps(current, ops);
            expect(current).toStrictEqual(before);
            if (input.kind === "reject") {
              expect(applied.isErr()).toBe(true);
              tally.set(input.kind, (tally.get(input.kind) ?? 0) + 1);
              continue;
            }
            if (applied.isErr()) throw applied.error;
            tally.set(input.kind, (tally.get(input.kind) ?? 0) + 1);
            text = text.slice(0, from) + replacement + text.slice(to);
            current = applied.value.document;
            expect(orderedIds(current)).toEqual([blockId]);
            const paragraph = storyParagraphs(current.package.document).at(0)?.paragraph;
            if (!paragraph) panic("Generated paragraph disappeared");
            expect(paragraphLogicalText(paragraph)).toBe(text);
            // Every inserted UTF-16 unit carries the pre-deletion authored formatting.
            let runOffset = 0;
            for (const run of paragraph.content) {
              if (run.type !== "run") panic("Generated authored input produced a wrapper");
              const runEnd =
                runOffset + run.content.reduce((width, child) => width + runContentWidth(child), 0);
              if (runEnd > from && runOffset < from + replacement.length)
                expect(run.formatting ?? {}).toStrictEqual(runProps);
              runOffset = runEnd;
            }
            expectRestores(applied.value, before);
            journal.push(applied.value);
          }
          const final = structuredClone(current);
          const redo: (readonly DocumentOp[])[] = [];
          for (const entry of journal.toReversed()) {
            const undone = applyDocumentOps(current, entry.inverse);
            if (undone.isErr()) throw undone.error;
            current = undone.value.document;
            redo.push(undone.value.inverse);
          }
          expect(current).toStrictEqual(original);
          for (const ops of redo.toReversed()) current = applyAll(current, ops);
          expect(current).toStrictEqual(final);
        },
      ),
      propertyConfig({ numRuns: NUM_RUNS }),
    );
    expect([...tally.keys()].sort()).toEqual([...PLAIN_INPUT_KINDS].sort());
    for (const applied of tally.values()) expect(applied).toBeGreaterThan(NUM_RUNS / 1000);
  });

  test("an operation's inverse restores the document exactly", () => {
    const tally: Tally = new Map();
    assertProperty(
      fc.property(documentArbitrary, opSeedArbitrary, (document, seed) => {
        const op = opFor(document, seed);
        const original = structuredClone(document);
        const applied = applyDocumentOp(document, op);
        // Applying, refused or not, never writes into its input.
        expect(document).toStrictEqual(original);
        if (applied.isErr()) {
          count(tally, "refused");
          return;
        }
        count(tally, op.type);
        const allowed: readonly DocumentOpType[] = INVERSE_KINDS[op.type];
        for (const inverse of applied.value.inverse) {
          expect(allowed).toContain(inverse.type);
        }
        if (
          op.type === DOCUMENT_OP_TYPES.SPLIT_BLOCK ||
          op.type === DOCUMENT_OP_TYPES.JOIN_BLOCKS
        ) {
          // One structural inverse, at most one review restoration, and at most
          // one content restoration per original paragraph; no duplicate targets.
          const restoring = applied.value.inverse.filter(
            (inverse) => inverse.type === DOCUMENT_OP_TYPES.REPLACE_INLINE,
          );
          const review = applied.value.inverse.filter(
            (inverse) => inverse.type === DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW,
          );
          expect(new Set(restoring.map(({ blockId }) => blockId)).size).toBe(restoring.length);
          expect(restoring.length).toBeLessThanOrEqual(
            op.type === DOCUMENT_OP_TYPES.SPLIT_BLOCK ? 1 : 2,
          );
          expect(review.length).toBeLessThanOrEqual(1);
          expect(applied.value.inverse.length - restoring.length - review.length).toBe(1);
        } else if (op.type !== DOCUMENT_OP_TYPES.SET_RUN_PROPS) {
          expect(applied.value.inverse.length).toBeLessThanOrEqual(1);
        }
        expectRestores(applied.value, original);
      }),
      { numRuns: NUM_RUNS },
    );
    expectEveryKindApplied(tally, NUM_RUNS);
  });

  test("a sequence's inverses in reverse, and a batch's, restore the document exactly", () => {
    const tally: Tally = new Map();
    assertProperty(
      fc.property(
        documentArbitrary,
        fc.array(opSeedArbitrary, { minLength: 2, maxLength: 8 }),
        (document, seeds: OpSeed[]) => {
          const original = structuredClone(document);
          let current = document;
          const ops: DocumentOp[] = [];
          const inverses: (readonly DocumentOp[])[] = [];
          for (const seed of seeds) {
            const op = opFor(current, seed);
            const applied = applyDocumentOp(current, op);
            if (applied.isErr()) {
              count(tally, "refused");
              continue;
            }
            count(tally, op.type);
            ops.push(op);
            inverses.push(applied.value.inverse);
            current = applied.value.document;
          }
          expect(applyAll(current, inverses.toReversed().flat())).toStrictEqual(original);

          const batch = applyDocumentOps(document, ops);
          if (batch.isErr()) {
            throw batch.error;
          }
          expect(batch.value.document).toStrictEqual(current);
          expectRestores(batch.value, original);

          // JSON is the journal wire format; IDs and section/list model state survive both ways.
          const wireOps = JSON.parse(JSON.stringify(ops)) as DocumentOp[];
          const wireBatch = applyDocumentOps(independentCopy(document), wireOps);
          if (wireBatch.isErr()) throw wireBatch.error;
          expect(wireBatch.value.document).toStrictEqual(current);
          const wireInverse = JSON.parse(JSON.stringify(wireBatch.value.inverse)) as DocumentOp[];
          const wireUndo = applyDocumentOps(independentCopy(wireBatch.value.document), wireInverse);
          if (wireUndo.isErr()) throw wireUndo.error;
          expect(wireUndo.value.document).toStrictEqual(original);
        },
      ),
      { numRuns: NUM_RUNS },
    );
    expectEveryKindApplied(tally, NUM_RUNS);
  });

  test("editor joins select properties from visible content and preserve trailing mark formatting", () => {
    assertProperty(
      fc.property(
        fc.constantFrom("visible", "deleted", "empty"),
        fc.constantFrom(undefined, {}, { alignment: "center" as const }),
        fc.constantFrom(undefined, {}, { keepNext: true }),
        fc.constantFrom(undefined, {}, { bold: false }),
        (contentKind, leadingProperties, trailingProperties, markProperties) => {
          const run = { type: "run", content: [{ type: "text", text: "a" }] } satisfies Run;
          const content = (() => {
            if (contentKind === "empty") return [];
            if (contentKind === "deleted")
              return [
                { type: "deletion", info: { id: 1, author: "Source" }, content: [run] },
              ] satisfies Paragraph["content"];
            return [run];
          })();
          const leading = {
            type: "paragraph",
            paraId: "00000001",
            content,
            ...(leadingProperties === undefined ? {} : { formatting: leadingProperties }),
          } satisfies Paragraph;
          const trailingFormatting =
            markProperties === undefined
              ? trailingProperties
              : { ...trailingProperties, runProperties: markProperties };
          const trailing = {
            type: "paragraph",
            paraId: "00000002",
            content: [{ type: "run", content: [{ type: "text", text: "b" }] }],
            ...(trailingFormatting === undefined ? {} : { formatting: trailingFormatting }),
          } satisfies Paragraph;
          const document = {
            package: { document: { content: [leading, trailing] } },
          } satisfies Document;
          const intent = {
            type: "joinParagraphs",
            story: OP_STORIES.MAIN,
            blockId: leading.paraId,
            nextBlockId: trailing.paraId,
          } as const satisfies EditorIntent;
          const sourceProperties =
            contentKind === "visible" ? leadingProperties : trailingProperties;
          // Mark-only formatting still owns an empty paragraph-property set.
          const expectedFormatting = (() => {
            if (contentKind !== "visible") return trailingFormatting;
            if (sourceProperties === undefined && markProperties === undefined) return undefined;
            return {
              ...sourceProperties,
              ...(markProperties === undefined ? {} : { runProperties: markProperties }),
            };
          })();
          const allocation = allocateEditorIntentIds(document, intent);
          for (const mode of [
            { type: "editing", newIds: allocation.newIds },
            {
              type: "suggesting",
              revision: { id: allocation.revisionId, author: "Property" },
              newIds: allocation.newIds,
            },
          ] as const) {
            const plan = compileEditorIntent(document, { intent, mode }).unwrap();
            const edit = applyDocumentOps(document, plan.ops).unwrap();
            const accepted =
              mode.type === "editing"
                ? edit.document
                : applyDocumentOp(edit.document, {
                    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
                    story: OP_STORIES.MAIN,
                    revisionIds: edit.revisions,
                    decision: REVISION_DECISIONS.ACCEPT,
                  }).unwrap().document;
            const survivor = storyParagraphs(accepted.package.document).at(0)?.paragraph;
            expect(survivor?.paraId).toBe(trailing.paraId);
            assertExactModel(survivor?.formatting, expectedFormatting);
            expect(Object.hasOwn(survivor ?? {}, "formatting")).toBe(
              expectedFormatting !== undefined,
            );
            const inverse = applyDocumentOps(edit.document, edit.inverse).unwrap();
            assertExactModel(inverse.document, document);
          }
        },
      ),
      { numRuns: 100 },
    );
  });

  test("generated editor intents agree in direct and accepted tracked mode across structural shapes", () => {
    const tallies = new Map<EditorIntent["type"], number>();
    const refusals = new Map<DocumentOpRefusalReason, number>();
    let attemptedPlans = 0;
    let appliedMultiParagraphFormatting = 0;
    assertProperty(
      fc.property(documentArbitrary, (document) => {
        const paragraphs = storyParagraphs(document.package.document);
        const first = paragraphs.at(0);
        const paragraph = first?.paragraph;
        if (paragraph?.paraId === undefined) return;
        const at = { story: OP_STORIES.MAIN, blockId: paragraph.paraId, offset: 0 };
        const initialAllocation = allocateEditorIntentIds(document, { type: "splitParagraph", at });
        const length = paragraphLogicalText(paragraph).length;
        const intents: EditorIntent[] = [
          { type: "formatParagraph", at, patch: { alignment: "center", keepNext: true } },
          {
            type: "formatRun",
            from: at,
            to: { ...at, offset: length },
            patch: { bold: true, italic: false },
          },
          { type: "insertAtom", from: at, to: at, atom: { type: "tab" } },
          { type: "replaceText", from: at, to: { ...at, offset: length }, text: "edited" },
          { type: "splitParagraph", at, newBlockId: initialAllocation.newBlockId },
          {
            type: "setList",
            items: [{ at, ilvl: 0 }],
            target: {
              type: "new",
              num: {
                numId: 1000 + initialAllocation.revisionId,
                abstractNumId: 2000 + initialAllocation.revisionId,
              },
              abstractNum: {
                abstractNumId: 2000 + initialAllocation.revisionId,
                levels: [{ ilvl: 0, numFmt: "decimal", lvlText: "%1." }],
              },
            },
          },
        ];
        if (document.package.numbering?.nums.at(0) !== undefined) {
          intents.push({
            type: "setList",
            items: [{ at, ilvl: 0 }],
            target: { type: "existing", numId: document.package.numbering.nums[0]!.numId },
          });
        }
        const nextInList = paragraphs.find(
          ({ list, index }) =>
            first !== undefined && sameBlockList(first.list, list) && index > first.index,
        );
        if (nextInList?.paragraph.paraId !== undefined) {
          const next = nextInList.paragraph;
          intents.push({
            type: "formatRun",
            from: at,
            to: {
              story: OP_STORIES.MAIN,
              blockId: next.paraId,
              offset: paragraphLogicalText(next).length,
            },
            patch: { underline: { style: "single" } },
          });
        }
        const adjacent = paragraphs.find(
          ({ list, index }) =>
            first !== undefined && sameBlockList(first.list, list) && index === first.index + 1,
        );
        if (adjacent?.paragraph.paraId !== undefined) {
          intents.push({
            type: "joinParagraphs",
            story: OP_STORIES.MAIN,
            blockId: paragraph.paraId,
            nextBlockId: adjacent.paragraph.paraId,
          });
        }
        for (const intent of intents) {
          const allocation = allocateEditorIntentIds(document, intent);
          const direct = compileEditorIntent(document, {
            intent,
            mode: { type: "editing", newIds: allocation.newIds },
          });
          const tracked = compileEditorIntent(document, {
            intent,
            mode: {
              type: "suggesting",
              revision: {
                id: allocation.revisionId,
                author: "Property",
                date: "2026-10-02T00:00:00Z",
              },
              newIds: allocation.newIds,
            },
          });
          attemptedPlans += 2;
          for (const result of [direct, tracked]) {
            if (result.isErr())
              refusals.set(result.error.reason, (refusals.get(result.error.reason) ?? 0) + 1);
          }
          if (direct.isErr() || tracked.isErr()) continue;
          const directEdit = applyDocumentOps(document, direct.value.ops);
          const trackedEdit = applyDocumentOps(document, tracked.value.ops);
          if (directEdit.isErr()) throw directEdit.error;
          if (trackedEdit.isErr()) throw trackedEdit.error;
          const originalChanges = storyParagraphs(document.package.document).flatMap(
            ({ paragraph: sourceParagraph }) => sourceParagraph.propertyChanges ?? [],
          );
          const foldedIds = storyParagraphs(trackedEdit.value.document.package.document).flatMap(
            ({ paragraph: editedParagraph }) =>
              (editedParagraph.propertyChanges ?? []).flatMap((change) => {
                const previous = originalChanges.find(({ info }) => info.id === change.info.id);
                if (previous === undefined || change.info.author !== "Property") return [];
                expect(["joinParagraphs", "formatParagraph", "setList"]).toContain(intent.type);
                expect(change.previousFormatting).toStrictEqual(previous.previousFormatting);
                expect(change.info).toStrictEqual({
                  ...previous.info,
                  author: "Property",
                  date: "2026-10-02T00:00:00Z",
                });
                expect(trackedEdit.value.revisions).not.toContain(change.info.id);
                return [change.info.id];
              }),
          );
          const resolveFolded = (
            source: Document,
            decision: (typeof REVISION_DECISIONS)[keyof typeof REVISION_DECISIONS],
          ) =>
            foldedIds.length === 0
              ? source
              : applyDocumentOp(source, {
                  type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
                  story: OP_STORIES.MAIN,
                  revisionIds: foldedIds,
                  decision,
                }).unwrap().document;
          let acceptedDocument = trackedEdit.value.document;
          if (trackedEdit.value.revisions.length > 0 || foldedIds.length > 0) {
            const accepted = applyDocumentOp(acceptedDocument, {
              type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
              story: OP_STORIES.MAIN,
              revisionIds: [...trackedEdit.value.revisions, ...foldedIds],
              decision: REVISION_DECISIONS.ACCEPT,
            });
            if (accepted.isErr()) throw accepted.error;
            acceptedDocument = accepted.value.document;
            const rejected = applyDocumentOp(trackedEdit.value.document, {
              type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
              story: OP_STORIES.MAIN,
              revisionIds: [...trackedEdit.value.revisions, ...foldedIds],
              decision: REVISION_DECISIONS.REJECT,
            });
            if (rejected.isErr()) throw rejected.error;
            if (intent.type === "setList" && intent.target.type === "new") {
              const rejectedPackage = Object.fromEntries(
                Object.entries(rejected.value.document.package).filter(
                  ([key]) => key !== "numbering",
                ),
              );
              const baselinePackage = Object.fromEntries(
                Object.entries(resolveFolded(document, REVISION_DECISIONS.REJECT).package).filter(
                  ([key]) => key !== "numbering",
                ),
              );
              assertExactModel(rejectedPackage, baselinePackage);
              expect(rejected.value.document.package.numbering?.nums).toContainEqual(
                intent.target.num,
              );
              if (intent.target.abstractNum !== undefined) {
                expect(rejected.value.document.package.numbering?.abstractNums).toContainEqual(
                  intent.target.abstractNum,
                );
              }
            } else {
              assertExactModel(
                rejected.value.document,
                resolveFolded(document, REVISION_DECISIONS.REJECT),
              );
            }
          }
          assertFreshIdentityEquivalent({
            actual: acceptedDocument,
            expected: resolveFolded(directEdit.value.document, REVISION_DECISIONS.ACCEPT),
            original: document,
            allocated: allocation.newIds,
          });
          const directUndo = applyDocumentOps(directEdit.value.document, directEdit.value.inverse);
          const trackedUndo = applyDocumentOps(
            trackedEdit.value.document,
            trackedEdit.value.inverse,
          );
          if (directUndo.isErr()) throw directUndo.error;
          if (trackedUndo.isErr()) throw trackedUndo.error;
          assertExactModel(directUndo.value.document, document);
          assertExactModel(trackedUndo.value.document, document);
          const directRedo = applyDocumentOps(directUndo.value.document, directUndo.value.inverse);
          const trackedRedo = applyDocumentOps(
            trackedUndo.value.document,
            trackedUndo.value.inverse,
          );
          if (directRedo.isErr()) throw directRedo.error;
          if (trackedRedo.isErr()) throw trackedRedo.error;
          assertExactModel(directRedo.value.document, directEdit.value.document);
          assertExactModel(trackedRedo.value.document, trackedEdit.value.document);
          tallies.set(intent.type, (tallies.get(intent.type) ?? 0) + 1);
          if (intent.type === "formatRun" && intent.from.blockId !== intent.to.blockId) {
            appliedMultiParagraphFormatting += 1;
          }
        }
      }),
      { numRuns: 100 },
    );
    for (const kind of [
      "formatParagraph",
      "formatRun",
      "insertAtom",
      "replaceText",
      "splitParagraph",
      "joinParagraphs",
      "setList",
    ] as const) {
      expect(tallies.get(kind) ?? 0).toBeGreaterThan(0);
    }
    expect(appliedMultiParagraphFormatting).toBeGreaterThan(0);
    const expectedRefusals: readonly DocumentOpRefusalReason[] = [
      DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
      DOCUMENT_OP_REFUSAL_REASONS.INSIDE_TRACKED_DELETION,
      DOCUMENT_OP_REFUSAL_REASONS.SECTION_BOUNDARY,
      DOCUMENT_OP_REFUSAL_REASONS.CONTAINER_FINAL_MARK,
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
    ];
    for (const reason of refusals.keys()) expect(expectedRefusals).toContain(reason);
    expect(
      [...refusals.values()].reduce((total, refusalCount) => total + refusalCount, 0),
    ).toBeLessThan(attemptedPlans / 2);
  });

  test("the same operation on equal documents gives equal results", () => {
    const outcome = (result: ReturnType<typeof applyDocumentOp>) =>
      result.isOk() ? result.value : { refused: result.error.reason };
    assertProperty(
      fc.property(documentArbitrary, opSeedArbitrary, (document, seed) => {
        const op = opFor(document, seed);
        // SAFETY: the operation is plain data; this is the journal's round-trip.
        const replayed = JSON.parse(JSON.stringify(op)) as DocumentOp;
        const first = outcome(applyDocumentOp(structuredClone(document), op));
        expect(outcome(applyDocumentOp(structuredClone(document), replayed))).toStrictEqual(first);
        expect(outcome(applyDocumentOp(document, op))).toStrictEqual(first);
        // Rebuilt record by record, sharing nothing: the result must not depend on sharing.
        expect(outcome(applyDocumentOp(independentCopy(document), replayed))).toStrictEqual(first);
      }),
      { numRuns: NUM_RUNS },
    );
  });

  test("every generated document meets the seed contract", () => {
    fc.assert(
      fc.property(documentArbitrary, (document) => {
        const copy = independentCopy(document);
        expect(copy).toStrictEqual(document);
        expect(copy).not.toBe(document);
        expect(copy.package.document.content).not.toBe(document.package.document.content);
        expect(document.package.document.sections?.flatMap(({ content }) => content)).toStrictEqual(
          document.package.document.content,
        );
        expect(contractViolation(document)).toBeUndefined();
        expect(contractViolation(copy)).toBeUndefined();
      }),
      propertyConfig({ numRuns: NUM_RUNS }),
    );
  });

  test("blocks an operation does not touch are the same objects", () => {
    const tally: Tally = new Map();
    assertProperty(
      fc.property(documentArbitrary, opSeedArbitrary, (document, seed) => {
        const op = opFor(document, seed);
        const applied = applyDocumentOp(document, op);
        if (applied.isErr()) {
          count(tally, "refused");
          return;
        }
        count(tally, op.type);
        const next = applied.value.document;
        const touched = touchedIds(applied.value);
        const named = namedIds(op);
        for (const id of touched) {
          expect(named.has(id)).toBe(true);
        }

        const before = paragraphsById(document);
        const after = paragraphsById(next);
        for (const [id, paragraph] of before) {
          if (!touched.has(id)) {
            expect(after.get(id)).toBe(paragraph);
          }
        }
        for (const [id, paragraph] of after) {
          if (!touched.has(id)) {
            expect(before.get(id)).toBe(paragraph);
          }
        }
        for (const id of applied.value.touched.modified) {
          expect(after.get(id)).not.toBe(before.get(id));
          expect(after.has(id)).toBe(true);
        }
        for (const id of applied.value.touched.inserted) {
          expect(before.has(id)).toBe(false);
          expect(after.has(id)).toBe(true);
        }
        for (const id of applied.value.touched.removed) {
          expect(after.has(id)).toBe(false);
        }
        const untouchedOrder = (ids: string[]) => ids.filter((id) => !touched.has(id));
        expect(untouchedOrder(orderedIds(next))).toEqual(untouchedOrder(orderedIds(document)));

        const untouchedBlocks = (target: Document) =>
          target.package.document.content.filter((block) => !blockHolds(block, touched));
        const beforeBlocks = untouchedBlocks(document);
        const afterBlocks = untouchedBlocks(next);
        expect(afterBlocks.length).toBe(beforeBlocks.length);
        for (const [index, block] of afterBlocks.entries()) {
          expect(block).toBe(beforeBlocks[index]!);
        }

        // The section view is derived from the body's blocks; where the input's
        // view shared the body's records, an untouched section stays put.
        const body = next.package.document;
        const beforeBody = document.package.document;
        const beforeSections = beforeBody.sections ?? [];
        const beforeContent = new Set(beforeBody.content);
        const shared = beforeSections.every((section) =>
          section.content.every((block) => beforeContent.has(block)),
        );
        const derived = body.sections?.flatMap(({ content }) => content) ?? [];
        expect(derived).toStrictEqual(body.content);
        // An edit derives the view from the body, so it then holds the body's records.
        if (next !== document) {
          for (const [index, block] of derived.entries()) {
            expect(block).toBe(body.content[index]!);
          }
        }
        for (const [index, section] of (body.sections ?? []).entries()) {
          if (shared && !section.content.some((block) => blockHolds(block, touched))) {
            expect(section).toBe(beforeSections[index]!);
          }
        }

        for (const key of Object.keys(document.package)) {
          if (key !== "document") {
            expect(Reflect.get(next.package, key)).toBe(Reflect.get(document.package, key));
          }
        }
        expect(body.finalSectionProperties).toBe(document.package.document.finalSectionProperties!);
        expect(body.comments).toBe(document.package.document.comments!);
        expect(next.warnings).toBe(document.warnings!);
      }),
      { numRuns: NUM_RUNS },
    );
    expectEveryKindApplied(tally, NUM_RUNS);
  });

  test("operations change logical text and run properties as defined", () => {
    fc.assert(
      fc.property(documentArbitrary, opSeedArbitrary, (document, seed) => {
        const op = opFor(document, seed);
        const applied = applyDocumentOp(document, op);
        if (applied.isErr()) {
          return;
        }
        const before = paragraphsById(document);
        const after = paragraphsById(applied.value.document);
        const paragraphOf = (paragraphs: Map<string, Paragraph>, id: string): Paragraph => {
          const paragraph = paragraphs.get(id);
          if (paragraph === undefined) throw new Error(`${id} is missing`);
          return paragraph;
        };
        const textOf = (paragraphs: Map<string, Paragraph>, id: string): string =>
          paragraphLogicalText(paragraphOf(paragraphs, id));
        switch (op.type) {
          case DOCUMENT_OP_TYPES.INSERT_TEXT: {
            const { blockId, offset } = op.at;
            const old = textOf(before, blockId);
            expect(textOf(after, blockId)).toBe(old.slice(0, offset) + op.text + old.slice(offset));
            const inserted = unitsOf(paragraphOf(after, blockId)).slice(
              offset,
              offset + op.text.length,
            );
            for (const unit of inserted) {
              expect(unit.inRun).toBe(true);
              expect(unit.removed).toBe(false);
              if (op.runProps !== INHERIT_RUN_PROPS) {
                expect(sameRunFormatting(unit.formatting, op.runProps)).toBe(true);
              }
            }
            break;
          }
          case DOCUMENT_OP_TYPES.INSERT_CONTENT: {
            const { blockId, offset } = op.at;
            const old = textOf(before, blockId);
            const added = paragraphLogicalText({
              type: "paragraph",
              content: [...op.slice.content],
            });
            expect(textOf(after, blockId)).toBe(old.slice(0, offset) + added + old.slice(offset));
            break;
          }
          case DOCUMENT_OP_TYPES.DELETE_RANGE: {
            const { blockId } = op.from;
            const old = textOf(before, blockId);
            expect(textOf(after, blockId)).toBe(
              old.slice(0, op.from.offset) + old.slice(op.to.offset),
            );
            break;
          }
          case DOCUMENT_OP_TYPES.SPLIT_INLINE:
          case DOCUMENT_OP_TYPES.JOIN_INLINE: {
            const { blockId } = op.at;
            expect(textOf(after, blockId)).toBe(textOf(before, blockId));
            expect(unitsOf(paragraphOf(after, blockId))).toEqual(
              unitsOf(paragraphOf(before, blockId)),
            );
            break;
          }
          case DOCUMENT_OP_TYPES.SET_RUN_PROPS: {
            const { blockId } = op.from;
            expect(textOf(after, blockId)).toBe(textOf(before, blockId));
            const oldUnits = unitsOf(paragraphOf(before, blockId));
            const newUnits = unitsOf(paragraphOf(after, blockId));
            for (const [index, unit] of newUnits.entries()) {
              const old = oldUnits[index]!;
              const inRange = index >= op.from.offset && index < op.to.offset;
              const expected =
                inRange && old.inRun
                  ? applyFormattingPatch(old.formatting, op.patch)
                  : old.formatting;
              expect(sameRunFormatting(unit.formatting, expected)).toBe(true);
            }
            break;
          }
          case DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS: {
            expect(after.get(op.blockId)?.content).toBe(paragraphOf(before, op.blockId).content);
            break;
          }
          case DOCUMENT_OP_TYPES.SPLIT_BLOCK: {
            const { blockId, offset } = op.at;
            const old = textOf(before, blockId);
            // By default the new paragraph follows at the end, and takes the text before otherwise.
            const newHalf =
              op.newHalf ?? (offset === old.length ? SPLIT_HALVES.SECOND : SPLIT_HALVES.FIRST);
            const [firstId, secondId] =
              newHalf === SPLIT_HALVES.FIRST ? [op.newBlockId, blockId] : [blockId, op.newBlockId];
            expect(textOf(after, firstId)).toBe(old.slice(0, offset));
            expect(textOf(after, secondId)).toBe(old.slice(offset));
            break;
          }
          case DOCUMENT_OP_TYPES.JOIN_BLOCKS: {
            const survivor = op.survivor === SPLIT_HALVES.FIRST ? op.blockId : op.nextBlockId;
            expect(textOf(after, survivor)).toBe(
              textOf(before, op.blockId) + textOf(before, op.nextBlockId),
            );
            break;
          }
          case DOCUMENT_OP_TYPES.SET_PACKAGE_RESOURCES:
          case DOCUMENT_OP_TYPES.CREATE_COMMENT:
          case DOCUMENT_OP_TYPES.UPDATE_COMMENT_CONTENT:
          case DOCUMENT_OP_TYPES.SET_COMMENT_RESOLUTION:
          case DOCUMENT_OP_TYPES.DELETE_COMMENT:
          case DOCUMENT_OP_TYPES.RESTORE_COMMENT_STATE:
          case DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER:
          case DOCUMENT_OP_TYPES.REMOVE_HEADER_FOOTER:
          case DOCUMENT_OP_TYPES.ADD_NOTE:
          case DOCUMENT_OP_TYPES.REMOVE_NOTE:
          case DOCUMENT_OP_TYPES.SET_SECTION_PROPS:
          case DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS:
          case DOCUMENT_OP_TYPES.DELETE_BLOCKS:
          case DOCUMENT_OP_TYPES.INSERT_BLOCKS:
          case DOCUMENT_OP_TYPES.REPLACE_BLOCKS:
          case DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW:
          case DOCUMENT_OP_TYPES.REPLACE_INLINE:
          case DOCUMENT_OP_TYPES.RESOLVE_REVISION:
          case DOCUMENT_OP_TYPES.INSERT_TABLE:
          case DOCUMENT_OP_TYPES.DELETE_TABLE:
          case DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS:
          case DOCUMENT_OP_TYPES.INSERT_ROW:
          case DOCUMENT_OP_TYPES.DELETE_ROW:
          case DOCUMENT_OP_TYPES.SET_TABLE_ROWS:
            break;
          default: {
            const unreachable: never = op;
            return unreachable;
          }
        }
      }),
      propertyConfig({ numRuns: NUM_RUNS }),
    );
  });
});
