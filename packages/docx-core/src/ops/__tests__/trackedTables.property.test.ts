/** Tracked whole-table and terminal-carrier laws across container, cell and text shapes. */
import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { assertProperty, propertyTestTimeout } from "../../../../../test/property-testing";
import type { BlockContent, Document, Paragraph, Table } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { storyParagraphs } from "../blocks";
import { contractViolation } from "../contract";
import { identityKeysIn, paragraphIdsIn } from "../ids";
import { asParagraphContent, childNodes, rebuildNode, type InlineNode } from "../leaves";
import { paragraphLength } from "../offsets";
import { revisionIdDemand } from "../plan";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import { mergeAtSeam } from "../resolve";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
  type DocumentOp,
  type RevisionDecision,
} from "../types";
import { independentCopy } from "./documentArbitraries";

const NUM_RUNS = 2_000;
const shapes = fc.record({
  container: fc.constantFrom("body", "cell", "customXml", "sdt", "wrapperCell"),
  rows: fc.integer({ min: 1, max: 4 }),
  cells: fc.integer({ min: 1, max: 3 }),
  paragraphs: fc.integer({ min: 1, max: 3 }),
  seed: fc.integer({ min: 0, max: 0xffff_ffff }),
  text: fc
    .array(fc.constantFrom("a", "é", "ž", "漢", "😀", " "), { minLength: 0, maxLength: 16 })
    .map((tokens) => tokens.join("")),
  placement: fc.constantFrom("before", "after", "terminal"),
});
type Shape = fc.ArbitraryValue<typeof shapes>;
const id = (value: number) => value.toString(16).toUpperCase().padStart(8, "0");
const makeParagraph = (value: number, shape: Shape): Paragraph => ({
  type: "paragraph",
  paraId: id(value),
  textId: id(value + 0x10000),
  preservedAttributes: [
    {
      namespace: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
      name: "rsidR",
      value: id(shape.seed % 0x10000),
    },
  ],
  formatting: {
    alignment: shape.seed % 2 === 0 ? "start" : "end",
    runProperties: { bold: shape.seed % 3 === 0 },
  },
  content:
    shape.text === ""
      ? []
      : [
          {
            type: "run",
            formatting: { italic: shape.seed % 2 === 0 },
            content: [{ type: "text", text: shape.text }],
          },
        ],
});
const makeTable = (shape: Shape, base = 0x100): Table => ({
  type: "table",
  formatting: { justification: shape.seed % 2 === 0 ? "center" : "end" },
  rows: Array.from({ length: shape.rows }, (_rowValue, row) => ({
    type: "tableRow",
    formatting: { cantSplit: row % 2 === 0 },
    cells: Array.from({ length: 1 + ((shape.cells + row) % 3) }, (_cellValue, cell) => ({
      type: "tableCell",
      content: Array.from(
        { length: 1 + ((shape.paragraphs + cell) % 3) },
        (_paragraphValue, index) =>
          makeParagraph(base + row * 27 + cell * 9 + index, {
            ...shape,
            text: (shape.seed + row + cell + index) % 3 === 0 ? "" : shape.text,
          }),
      ),
    })),
  })),
});
const wrap = (shape: Shape, content: BlockContent[]): BlockContent[] => {
  const custom: BlockContent = {
    type: "blockCustomXml",
    openingXml: '<w:customXml w:element="clause">',
    closingXml: "</w:customXml>",
    content,
  };
  switch (shape.container) {
    case "body":
      return content;
    case "customXml":
      return shape.placement === "terminal"
        ? [makeParagraph(0x90, shape), custom]
        : [custom, makeParagraph(0x90, shape)];
    case "sdt": {
      const control: BlockContent = {
        type: "blockSdt",
        properties: { sdtType: "group", id: 9 },
        content,
      };
      return shape.placement === "terminal"
        ? [makeParagraph(0x90, shape), control]
        : [control, makeParagraph(0x90, shape)];
    }
    case "cell":
    case "wrapperCell": {
      let cellContent = content;
      if (shape.container === "wrapperCell") {
        cellContent =
          shape.placement === "terminal"
            ? [makeParagraph(0x91, shape), custom]
            : [custom, makeParagraph(0x91, shape)];
      }
      return [
        {
          type: "table",
          rows: [
            {
              type: "tableRow",
              cells: [
                {
                  type: "tableCell",
                  content: cellContent,
                },
              ],
            },
          ],
        },
        makeParagraph(0x90, shape),
      ];
    }
    default: {
      const unreachable: never = shape.container;
      return unreachable;
    }
  }
};
const documentOf = (shape: Shape, content: BlockContent[]): Document => ({
  package: { document: { content: wrap(shape, content) } },
});
const stamp = (value = 1000) => ({ id: value, author: "Reviewer", date: "2026-05-06T07:08:09Z" });
const freshIds = (value = 1000) => ({
  revision: Array.from({ length: 128 }, (_revisionValue, index) => value + index + 1),
});
const applied = (document: Document, op: DocumentOp) => {
  const result = applyDocumentOp(document, op);
  if (result.isErr()) throw result.error;
  expect(contractViolation(result.value.document)).toBeUndefined();
  return result.value;
};
const resolved = (document: Document, revisionIds: readonly number[], decision: RevisionDecision) =>
  applied(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds,
    decision,
  });
const exactInverse = (document: Document, result: ReturnType<typeof applied>) => {
  for (const inverse of [result.inverse, independentCopy(result.inverse)]) {
    const undone = applyDocumentOps(result.document, inverse);
    if (undone.isErr()) throw undone.error;
    expect(undone.value.document).toStrictEqual(document);
    expect(contractViolation(undone.value.document)).toBeUndefined();
  }
};
const canonicalList = (nodes: readonly InlineNode[]): InlineNode[] => {
  const out: InlineNode[] = [];
  for (const node of nodes) {
    const children = childNodes(node);
    const current = children === undefined ? node : rebuildNode(node, canonicalList(children));
    const last = out.at(-1);
    if (last === undefined) out.push(current);
    else out.splice(-1, 1, ...mergeAtSeam(last, current));
  }
  return out;
};
const canonicalBlocks = (blocks: readonly BlockContent[]): BlockContent[] =>
  blocks.map((block): BlockContent => {
    switch (block.type) {
      case "paragraph":
        return { ...block, content: asParagraphContent(canonicalList(block.content)) };
      case "table":
        return {
          ...block,
          rows: block.rows.map((row) => ({
            ...row,
            cells: row.cells.map((cell) => ({ ...cell, content: canonicalBlocks(cell.content) })),
          })),
        };
      case "blockSdt":
      case "blockCustomXml":
        return { ...block, content: canonicalBlocks(block.content) };
      default:
        return block;
    }
  });
// π merges only inline seams; paragraph ids, formatting and carrier counts stay observable.
const canonical = (document: Document): Document => ({
  ...document,
  package: {
    ...document.package,
    document: {
      ...document.package.document,
      content: canonicalBlocks(document.package.document.content),
    },
  },
});
const insertionFor = (shape: Shape) =>
  ({
    type: DOCUMENT_OP_TYPES.INSERT_TABLE,
    story: OP_STORIES.MAIN,
    at: {
      type: shape.placement === "before" ? "before" : "after",
      blockId: shape.placement === "terminal" ? id(3) : id(1),
    },
    table: makeTable(shape),
    ...(shape.placement === "terminal" ? { terminal: { beforeBlockId: id(4) } } : {}),
  }) as const satisfies DocumentOp;
const targetTable = (document: Document, blockId = id(0x100)): Table => {
  const visit = (blocks: readonly BlockContent[]): Table | undefined => {
    for (const block of blocks) {
      switch (block.type) {
        case "table": {
          if (
            block.rows.some((row) =>
              row.cells.some((cell) =>
                cell.content.some(
                  (child) => child.type === "paragraph" && child.paraId === blockId,
                ),
              ),
            )
          )
            return block;
          for (const row of block.rows)
            for (const cell of row.cells) {
              const found = visit(cell.content);
              if (found !== undefined) return found;
            }
          break;
        }
        case "blockSdt":
        case "blockCustomXml": {
          const found = visit(block.content);
          if (found !== undefined) return found;
          break;
        }
        default:
          break;
      }
    }
    return undefined;
  };
  const table = visit(document.package.document.content);
  if (table === undefined) throw new Error("Generated target table remains present.");
  return table;
};
const demandFor = (value: Table) => {
  let records = value.rows.length;
  for (const row of value.rows)
    for (const cell of row.cells) {
      records += 1;
      for (const block of cell.content)
        if (block.type === "paragraph" && block.content.length !== 0) records += 1;
    }
  return records - 1;
};
const assertLocality = (before: Document, result: ReturnType<typeof applied>) => {
  const touched = new Set([
    ...result.touched.modified,
    ...result.touched.inserted,
    ...result.touched.removed,
  ]);
  const after = new Map(
    storyParagraphs(result.document.package.document).map(({ paragraph }) => [
      paragraph.paraId,
      paragraph,
    ]),
  );
  for (const { paragraph } of storyParagraphs(before.package.document))
    if (!touched.has(paragraph.paraId ?? "")) expect(after.get(paragraph.paraId)).toBe(paragraph);
  const beforeIds = paragraphIdsIn(before.package.document);
  const afterIds = paragraphIdsIn(result.document.package.document);
  expect(new Set(result.touched.inserted)).toEqual(
    new Set(afterIds.filter((value) => !beforeIds.includes(value))),
  );
  expect(new Set(result.touched.removed)).toEqual(
    new Set(beforeIds.filter((value) => !afterIds.includes(value))),
  );
};
const assertReviewShape = (value: Table, kind: "ins" | "del") => {
  for (const row of value.rows) {
    expect(row.structuralChange?.type).toBe(
      kind === "ins" ? "tableRowInsertion" : "tableRowDeletion",
    );
    for (const cell of row.cells) {
      const final = cell.content.at(-1);
      if (final?.type !== "paragraph") throw new Error("Generated cells end in paragraphs.");
      expect(final.pPrMark?.kind).toBe(kind);
      for (const block of cell.content) {
        if (block.type !== "paragraph") throw new Error("Generated cells contain paragraphs.");
        if (block !== final) expect(block.pPrMark).toBeUndefined();
        for (const wrapper of block.content)
          expect(wrapper.type).toBe(kind === "ins" ? "insertion" : "deletion");
      }
    }
  }
};

const families = ["insertion", "deletion"] as const;
describe("tracked whole table properties", () => {
  for (const family of families) {
    test(
      `${family}: L1 accept, L2 reject, L4 inverse, L5 determinism, L6 locality and S8 review shape`,
      () => {
        assertProperty(
          fc.property(shapes, (shape) => {
            const anchors = [
              makeParagraph(1, shape),
              makeParagraph(2, shape),
              makeParagraph(3, shape),
            ] as const;
            const directOp =
              family === "insertion"
                ? insertionFor(shape)
                : ({
                    type: DOCUMENT_OP_TYPES.DELETE_TABLE,
                    story: OP_STORIES.MAIN,
                    blockId: id(0x100),
                  } as const satisfies DocumentOp);
            const original = documentOf(
              shape,
              family === "insertion"
                ? anchors
                : [anchors[0], makeTable(shape), anchors[1], anchors[2]],
            );
            const snapshot = structuredClone(original);
            const op = { ...directOp, revision: stamp(), newIds: freshIds() };
            const tracked = applied(original, op);
            expect(original).toStrictEqual(snapshot);
            exactInverse(original, tracked);
            assertLocality(original, tracked);
            const repeat = applied(independentCopy(original), independentCopy(op));
            expect(repeat).toStrictEqual(tracked);
            assertReviewShape(
              targetTable(tracked.document),
              family === "insertion" ? "ins" : "del",
            );
            const final = storyParagraphs(tracked.document.package.document).find(
              ({ paragraph }) => paragraph.paraId === id(3),
            );
            expect(final?.paragraph.pPrMark).toBeUndefined();
            const demand = revisionIdDemand(original, op);
            if (demand.isErr()) throw demand.error;
            expect(demand.value).toBe(
              demandFor(makeTable(shape)) +
                (family === "insertion" && shape.placement === "terminal" ? 1 : 0),
            );
            const keys = identityKeysIn(targetTable(tracked.document));
            expect(new Set(keys).size).toBe(keys.length);
            expect(tracked.revisions.length).toBe(demand.value + 1);
            expect(new Set(tracked.revisions).size).toBe(tracked.revisions.length);
            const exact = applied(original, {
              ...op,
              newIds: { revision: freshIds().revision.slice(0, demand.value) },
            });
            expect(exact.document).toStrictEqual(tracked.document);
            if (demand.value > 0) {
              const short = applyDocumentOp(original, {
                ...op,
                newIds: { revision: freshIds().revision.slice(0, demand.value - 1) },
              });
              if (short.isOk())
                throw new Error(
                  "One missing physical revision id must refuse the whole operation.",
                );
              expect(short.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS);
              expect(original).toStrictEqual(snapshot);
            }
            const accepted = resolved(
              tracked.document,
              tracked.revisions,
              REVISION_DECISIONS.ACCEPT,
            );
            const rejected = resolved(
              tracked.document,
              tracked.revisions,
              REVISION_DECISIONS.REJECT,
            );
            exactInverse(tracked.document, accepted);
            exactInverse(tracked.document, rejected);
            expect(canonical(accepted.document)).toStrictEqual(
              canonical(applied(original, directOp).document),
            );
            expect(canonical(rejected.document)).toStrictEqual(canonical(original));
            expect(identityKeysIn(accepted.document)).toEqual(identityKeysIn(original));
            expect(identityKeysIn(rejected.document)).toEqual(identityKeysIn(original));
            const again = resolved(accepted.document, tracked.revisions, REVISION_DECISIONS.ACCEPT);
            expect(again.document).toStrictEqual(accepted.document);
            expect(again.inverse).toEqual([]);
            if (family === "insertion" && shape.placement === "terminal") {
              const source = anchors.at(-1);
              const restored = storyParagraphs(rejected.document.package.document).find(
                ({ paragraph }) => paragraph.paraId === id(3),
              );
              expect(restored?.paragraph).toStrictEqual(source);
              expect(paragraphIdsIn(rejected.document)).not.toContain(id(4));
              expect(paragraphIdsIn(accepted.document)).toContain(id(4));
            }
          }),
          { numRuns: NUM_RUNS },
        );
      },
      propertyTestTimeout(240_000),
    );
  }

  test(
    "L3 batches: table insertion, original-span deletion and other-table deletion resolve together or separately",
    () => {
      assertProperty(
        fc.property(shapes, (shape) => {
          const textShape = { ...shape, text: shape.text || "source" };
          const original = documentOf(shape, [
            makeParagraph(1, textShape),
            makeTable(shape, 0x500),
            makeParagraph(2, textShape),
            makeParagraph(3, textShape),
          ]);
          const insert = insertionFor(shape);
          const paragraphToDelete = makeParagraph(2, textShape);
          const deleteSpan = {
            type: DOCUMENT_OP_TYPES.DELETE_RANGE,
            from: { story: OP_STORIES.MAIN, blockId: id(2), offset: 0 },
            to: {
              story: OP_STORIES.MAIN,
              blockId: id(2),
              offset: paragraphLength(paragraphToDelete),
            },
          } as const satisfies DocumentOp;
          const deleteTable = {
            type: DOCUMENT_OP_TYPES.DELETE_TABLE,
            story: OP_STORIES.MAIN,
            blockId: id(0x500),
          } as const satisfies DocumentOp;
          const directOps = [insert, deleteSpan, deleteTable] as const;
          const batches: number[][] = [];
          let tracked = original;
          let direct = original;
          const inverses: DocumentOp[][] = [];
          for (const [index, op] of directOps.entries()) {
            const result = applied(tracked, {
              ...op,
              revision: stamp(1000 + index * 1000),
              newIds: freshIds(1000 + index * 1000),
            });
            exactInverse(tracked, result);
            inverses.unshift([...result.inverse]);
            batches.push([...result.revisions]);
            tracked = result.document;
            direct = applied(direct, op).document;
          }
          const all = batches.flat();
          expect(new Set(all).size).toBe(all.length);
          const accepted = resolved(tracked, all, REVISION_DECISIONS.ACCEPT);
          const rejected = resolved(tracked, all, REVISION_DECISIONS.REJECT);
          exactInverse(tracked, accepted);
          exactInverse(tracked, rejected);
          expect(canonical(accepted.document)).toStrictEqual(canonical(direct));
          expect(canonical(rejected.document)).toStrictEqual(canonical(original));
          let separatelyAccepted = tracked;
          let separatelyRejected = tracked;
          for (const ids of batches)
            separatelyAccepted = resolved(
              separatelyAccepted,
              ids,
              REVISION_DECISIONS.ACCEPT,
            ).document;
          for (const ids of batches.toReversed())
            separatelyRejected = resolved(
              separatelyRejected,
              ids,
              REVISION_DECISIONS.REJECT,
            ).document;
          expect(canonical(separatelyAccepted)).toStrictEqual(canonical(accepted.document));
          expect(canonical(separatelyRejected)).toStrictEqual(canonical(original));
          const undone = applyDocumentOps(tracked, inverses.flat());
          if (undone.isErr()) throw undone.error;
          expect(undone.value.document).toStrictEqual(original);
        }),
        { numRuns: NUM_RUNS },
      );
    },
    propertyTestTimeout(240_000),
  );

  test(
    "unsupported nested tables, cell records, moves and conflicting revisions refuse atomically",
    () => {
      assertProperty(
        fc.property(
          shapes,
          fc.constantFrom(
            "nested",
            "cellRevision",
            "rowRevision",
            "moveFrom",
            "moveTo",
            "finalMark",
          ),
          (shape, unsupported) => {
            const source = makeTable(shape);
            const row = source.rows.at(0);
            const cell = row?.cells.at(0);
            if (row === undefined || cell === undefined)
              throw new Error("Generated table has a row and cell.");
            const foreign = { id: 50, author: "Other" };
            let value: Table;
            let reason: string;
            switch (unsupported) {
              case "nested":
                value = {
                  ...source,
                  rows: [
                    {
                      ...row,
                      cells: [
                        {
                          ...cell,
                          content: [
                            makeTable({ ...shape, rows: 1 }, 0x900),
                            makeParagraph(0x100, shape),
                          ],
                        },
                      ],
                    },
                  ],
                };
                reason = DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE;
                break;
              case "cellRevision":
                value = {
                  ...source,
                  rows: [
                    {
                      ...row,
                      cells: [
                        {
                          ...cell,
                          structuralChange: { type: "tableCellInsertion", info: foreign },
                        },
                      ],
                    },
                  ],
                };
                reason = DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE;
                break;
              case "rowRevision":
                value = {
                  ...source,
                  rows: [{ ...row, structuralChange: { type: "tableRowDeletion", info: foreign } }],
                };
                reason = DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT;
                break;
              case "moveFrom":
              case "moveTo":
                value = {
                  ...source,
                  rows: [
                    {
                      ...row,
                      cells: [
                        {
                          ...cell,
                          content: [
                            {
                              ...makeParagraph(0x100, shape),
                              content: [
                                {
                                  type: unsupported,
                                  info: foreign,
                                  content: [
                                    {
                                      type: "run",
                                      content: [{ type: "text", text: shape.text || "moved" }],
                                    },
                                  ],
                                },
                              ],
                            },
                          ],
                        },
                      ],
                    },
                  ],
                };
                reason = DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE;
                break;
              case "finalMark":
                value = {
                  ...source,
                  rows: [
                    {
                      ...row,
                      cells: [
                        {
                          ...cell,
                          content: [
                            {
                              ...makeParagraph(0x100, shape),
                              pPrMark: {
                                kind: shape.seed % 2 === 0 ? "ins" : "del",
                                info: foreign,
                              },
                            },
                          ],
                        },
                      ],
                    },
                  ],
                };
                reason = DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT;
                break;
              default: {
                const unreachable: never = unsupported;
                return unreachable;
              }
            }
            for (const family of families) {
              const original = documentOf(
                shape,
                family === "insertion"
                  ? [makeParagraph(3, shape)]
                  : [value, makeParagraph(3, shape)],
              );
              const snapshot = structuredClone(original);
              const op =
                family === "insertion"
                  ? ({
                      type: DOCUMENT_OP_TYPES.INSERT_TABLE,
                      story: OP_STORIES.MAIN,
                      at: { type: "before", blockId: id(3) },
                      table: value,
                      revision: stamp(),
                      newIds: freshIds(),
                    } as const satisfies DocumentOp)
                  : ({
                      type: DOCUMENT_OP_TYPES.DELETE_TABLE,
                      story: OP_STORIES.MAIN,
                      blockId: id(0x100),
                      revision: stamp(),
                      newIds: freshIds(),
                    } as const satisfies DocumentOp);
              const result = applyDocumentOp(original, op);
              if (result.isOk())
                throw new Error(`Unsupported ${unsupported} table must be refused.`);
              expect(result.error.reason).toBe(reason);
              expect(original).toStrictEqual(snapshot);
            }
          },
        ),
        { numRuns: NUM_RUNS },
      );
    },
    propertyTestTimeout(240_000),
  );

  test(
    "partial resolution retains cell final marks after keeping rows and restores them exactly through inverses",
    () => {
      assertProperty(
        fc.property(shapes, fc.constantFrom("insertion", "deletion"), (shape, family) => {
          const value = makeTable(shape);
          const original = documentOf(
            shape,
            family === "insertion"
              ? [makeParagraph(1, shape), makeParagraph(3, shape)]
              : [value, makeParagraph(3, shape)],
          );
          const op =
            family === "insertion"
              ? ({
                  type: DOCUMENT_OP_TYPES.INSERT_TABLE,
                  story: OP_STORIES.MAIN,
                  at: { type: "before", blockId: id(3) },
                  table: value,
                  revision: stamp(),
                  newIds: freshIds(),
                } as const satisfies DocumentOp)
              : ({
                  type: DOCUMENT_OP_TYPES.DELETE_TABLE,
                  story: OP_STORIES.MAIN,
                  blockId: id(0x100),
                  revision: stamp(),
                  newIds: freshIds(),
                } as const satisfies DocumentOp);
          const tracked = applied(original, op);
          const pending = targetTable(tracked.document);
          const rowIds = pending.rows.map((row) => {
            if (row.structuralChange === undefined)
              throw new Error("Generated rows carry revisions.");
            return row.structuralChange.info.id;
          });
          const decision =
            family === "insertion" ? REVISION_DECISIONS.ACCEPT : REVISION_DECISIONS.REJECT;
          const rowsKept = resolved(tracked.document, rowIds, decision);
          exactInverse(tracked.document, rowsKept);
          const kept = targetTable(rowsKept.document);
          expect(kept.rows.every((row) => row.structuralChange === undefined)).toBe(true);
          const markIds = kept.rows.flatMap((row) =>
            row.cells.map((cell) => {
              const final = cell.content.at(-1);
              if (final?.type !== "paragraph" || final.pPrMark === undefined)
                throw new Error("Cell final marks survive independent row resolution.");
              return final.pPrMark.info.id;
            }),
          );
          const marksKept = resolved(rowsKept.document, markIds, decision);
          exactInverse(rowsKept.document, marksKept);
          const contentIds = tracked.revisions.filter(
            (revisionId) => !rowIds.includes(revisionId) && !markIds.includes(revisionId),
          );
          const completed = resolved(marksKept.document, contentIds, decision);
          exactInverse(marksKept.document, completed);
          expect(identityKeysIn(completed.document)).toEqual(identityKeysIn(original));
        }),
        { numRuns: NUM_RUNS },
      );
    },
    propertyTestTimeout(240_000),
  );
});
