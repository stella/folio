/**
 * The snapshot's cost, and the table coordinates that pay for it.
 *
 * `createFolioAIEditSnapshot` walks the document once. Every fact it needs
 * about a block's surroundings — the row that hides it, the cell it sits in —
 * is on the path the walk already took, so it never resolves a position.
 * `doc.resolve` re-descends from the root and finds each level's child by
 * scanning that level's fragment from index 0, so one resolve per block is
 * O(blocks^2) on a flat document. The guard below is the invariant rather than
 * a stopwatch: the snapshot may not resolve at all.
 */

import { describe, expect, test } from "bun:test";
import { type Node as PMNode, Schema } from "prosemirror-model";

import type { RunStyleResolver } from "../prosemirror/runStyleFormatting";
import { schema as folioSchema } from "../prosemirror/schema";
import { resolveSequentialBlockAnchor } from "./blockRange";
import type { CleanTextStructuralBoundary } from "./clean-text";
import {
  createFolioAIEditSnapshot,
  createFolioAIEditSnapshotWithStyleResolver,
  folioStoryTables,
  hashFolioAIBlockStructuralBoundaries,
  hashFolioAIBlockText,
  isFolioAIContentBlock,
  projectFolioAIBlockStructuralBoundaries,
  storyTablesOf,
} from "./snapshot";

const schema = new Schema({
  nodes: {
    doc: { content: "block+" },
    paragraph: {
      group: "block",
      content: "inline*",
      attrs: { paraId: { default: null } },
    },
    text: { group: "inline" },
    table: { content: "tableRow+", group: "block", tableRole: "table" },
    tableRow: {
      content: "tableCell+",
      tableRole: "row",
      attrs: { hidden: { default: false } },
    },
    tableCell: {
      content: "block+",
      tableRole: "cell",
      attrs: {
        colspan: { default: 1 },
        rowspan: { default: 1 },
        colwidth: { default: null },
      },
    },
  },
});

const paragraph = (text: string): PMNode =>
  schema.node("paragraph", { paraId: null }, text.length === 0 ? [] : [schema.text(text)]);

const cell = (content: PMNode[]): PMNode => schema.node("tableCell", null, content);

const row = (cells: PMNode[], hidden = false): PMNode => schema.node("tableRow", { hidden }, cells);

const table = (rows: PMNode[]): PMNode => schema.node("table", null, rows);

/**
 * Make the document refuse to resolve a position, so a snapshot that reaches
 * for one fails here instead of quietly reintroducing the quadratic term.
 */
const refusingToResolve = (doc: PMNode): PMNode => {
  Object.defineProperty(doc, "resolve", {
    configurable: true,
    value: () => {
      throw new Error("The snapshot resolved a position; walk the document instead.");
    },
  });
  return doc;
};

describe("createFolioAIEditSnapshot", () => {
  test("keeps structural projection and hashing at the same fixed point", () => {
    type BoundaryOptions = Pick<
      CleanTextStructuralBoundary,
      "clear" | "offset" | "presentInCleanView"
    >;
    const boundary = ({
      offset,
      presentInCleanView,
      clear,
    }: BoundaryOptions): CleanTextStructuralBoundary => ({
      type: "pageBreakRun",
      offset,
      from: offset + 1,
      to: offset + 2,
      ...(clear !== undefined ? { clear } : {}),
      presentInCleanView,
    });
    const emptyCleanProjection = { structuralBoundaries: [] };
    const deletionOnlyCleanProjection = {
      structuralBoundaries: [boundary({ offset: 1, presentInCleanView: false, clear: "all" })],
    };
    const cleanProjections = [
      emptyCleanProjection,
      deletionOnlyCleanProjection,
      { structuralBoundaries: [boundary({ offset: 1, presentInCleanView: true })] },
      {
        structuralBoundaries: [
          boundary({ offset: 1, presentInCleanView: true, clear: "none" }),
          boundary({ offset: 1, presentInCleanView: false, clear: "left" }),
          boundary({ offset: 4, presentInCleanView: true, clear: "right" }),
        ],
      },
    ];

    for (const cleanProjection of cleanProjections) {
      const projected = projectFolioAIBlockStructuralBoundaries(cleanProjection);
      expect(projectFolioAIBlockStructuralBoundaries(cleanProjection)).toEqual(projected);
      expect(hashFolioAIBlockStructuralBoundaries(cleanProjection)).toBe(
        hashFolioAIBlockText(JSON.stringify(projected)),
      );
    }

    const empty = projectFolioAIBlockStructuralBoundaries(emptyCleanProjection);
    const deletionOnly = projectFolioAIBlockStructuralBoundaries(deletionOnlyCleanProjection);
    expect(empty).toBe(deletionOnly);
    expect(Object.isFrozen(empty)).toBe(true);
  });

  test("locates a block in a table inside a table", () => {
    const nested = table([row([cell([paragraph("nested first"), paragraph("nested second")])])]);
    const doc = schema.node("doc", null, [
      paragraph("opening"),
      table([
        row([cell([paragraph("r0c0")]), cell([paragraph("r0c1")])]),
        row([cell([paragraph("r1c0")]), cell([paragraph("before nested"), nested])]),
      ]),
      paragraph("closing"),
    ]);

    const { blocks } = createFolioAIEditSnapshot(doc);

    expect(blocks.map(({ text }) => text)).toEqual([
      "opening",
      "r0c0",
      "r0c1",
      "r1c0",
      "before nested",
      "nested first",
      "nested second",
      "closing",
    ]);
    expect(blocks.at(0)?.table).toBeUndefined();
    expect(blocks.at(-1)?.table).toBeUndefined();
    expect(blocks.at(2)?.table).toEqual({
      outerTableIndex: 0,
      tableIndex: 0,
      rowIndex: 0,
      cellIndex: 1,
      gridColumnIndex: 1,
      columnSpan: 1,
      rowSpan: 1,
      paragraphIndex: 0,
    });
    // The nested table is the second table in document order, and its own
    // coordinates are relative to itself, not to the table containing it.
    // The nested table's own index, with the outer table it belongs to.
    expect(blocks.at(5)?.table).toEqual({
      outerTableIndex: 0,
      tableIndex: 1,
      rowIndex: 0,
      cellIndex: 0,
      gridColumnIndex: 0,
      columnSpan: 1,
      rowSpan: 1,
      paragraphIndex: 0,
    });
    expect(blocks.at(6)?.table).toEqual({
      outerTableIndex: 0,
      tableIndex: 1,
      rowIndex: 0,
      cellIndex: 0,
      gridColumnIndex: 0,
      columnSpan: 1,
      rowSpan: 1,
      paragraphIndex: 1,
    });
  });

  test("derives grid coordinates and spans independently of physical cell indexes", () => {
    const doc = schema.node("doc", null, [
      table([
        row([
          schema.node("tableCell", { colspan: 2, rowspan: 1, colwidth: null }, [
            paragraph("merged"),
          ]),
          cell([paragraph("right")]),
        ]),
        row([
          cell([paragraph("left")]),
          cell([paragraph("middle")]),
          cell([paragraph("lower right")]),
        ]),
      ]),
    ]);

    const { blocks } = createFolioAIEditSnapshot(doc);

    expect(blocks.find(({ text }) => text === "merged")?.table).toMatchObject({
      cellIndex: 0,
      gridColumnIndex: 0,
      columnSpan: 2,
    });
    expect(blocks.find(({ text }) => text === "right")?.table).toMatchObject({
      cellIndex: 1,
      gridColumnIndex: 2,
      columnSpan: 1,
    });
  });

  test("omits a hidden row's text, including a nested table inside it", () => {
    const doc = schema.node("doc", null, [
      table([
        row([cell([paragraph("visible")])]),
        row(
          [cell([paragraph("hidden"), table([row([cell([paragraph("hidden nested")])])])])],
          true,
        ),
      ]),
    ]);

    expect(createFolioAIEditSnapshot(doc).blocks.map(({ text }) => text)).toEqual(["visible"]);
  });

  test("resolves no positions, so its cost stays linear in block count", () => {
    const rows = Array.from({ length: 40 }, (_unused, index) =>
      row([
        cell([paragraph(`row ${String(index)} left`)]),
        cell([paragraph(`row ${String(index)} right`)]),
      ]),
    );
    const paragraphs = Array.from({ length: 200 }, (_unused, index) =>
      paragraph(`body paragraph ${String(index)}`),
    );
    const doc = refusingToResolve(
      schema.node("doc", null, [...paragraphs, table(rows), ...paragraphs]),
    );

    expect(createFolioAIEditSnapshot(doc).blocks).toHaveLength(480);
  });

  test("collects blocks and tables in one whole-document walk", () => {
    const doc = schema.node("doc", null, [
      paragraph("opening"),
      table([row([cell([paragraph("table cell")])])]),
      paragraph("closing"),
    ]);
    const walk = doc.descendants.bind(doc);
    let wholeDocumentWalks = 0;
    Object.defineProperty(doc, "descendants", {
      configurable: true,
      value: (callback: Parameters<PMNode["descendants"]>[0]) => {
        wholeDocumentWalks += 1;
        return walk(callback);
      },
    });

    const snapshot = createFolioAIEditSnapshot(doc);

    expect(wholeDocumentWalks).toBe(1);
    expect(
      storyTablesOf(snapshot).map(({ index, node }) => ({ index, text: node.textContent })),
    ).toEqual([{ index: 0, text: "table cell" }]);
  });

  test("keeps the direct and snapshot table censuses identical across nested and hidden rows", () => {
    const visibleNested = table([row([cell([paragraph("visible nested")])])]);
    const hiddenNested = table([row([cell([paragraph("hidden nested")])])]);
    const doc = schema.node("doc", null, [
      paragraph("opening"),
      table([
        row([cell([paragraph("outer visible"), visibleNested])]),
        row([cell([paragraph("hidden"), hiddenNested])], true),
        row([cell([paragraph("outer trailing")])]),
      ]),
      table([row([cell([paragraph("body trailing")])])]),
    ]);

    const direct = folioStoryTables(doc);
    const projected = storyTablesOf(createFolioAIEditSnapshot(doc));

    expect(projected.map(({ index, start }) => ({ index, start }))).toEqual(
      direct.map(({ index, start }) => ({ index, start })),
    );
    expect(projected).toHaveLength(3);
    for (const [index, tableProjection] of projected.entries()) {
      expect(tableProjection.node).toBe(direct.at(index)?.node);
    }
  });

  test("binds each table census to the immutable document that produced its snapshot", () => {
    const before = schema.node("doc", null, [table([row([cell([paragraph("before mutation")])])])]);
    const after = schema.node("doc", null, [
      table([row([cell([paragraph("after mutation")])])]),
      table([row([cell([paragraph("new table")])])]),
    ]);

    const beforeSnapshot = createFolioAIEditSnapshot(before);
    const afterSnapshot = createFolioAIEditSnapshot(after);

    expect(storyTablesOf(beforeSnapshot).map(({ node }) => node.textContent)).toEqual([
      "before mutation",
    ]);
    expect(storyTablesOf(afterSnapshot).map(({ node }) => node.textContent)).toEqual([
      "after mutation",
      "new table",
    ]);
  });

  test("projects carrierless run marks without consulting the style package", () => {
    const refuseStyleResolution = (): never => {
      throw new Error("Carrierless run marks already own their effective formatting");
    };
    const styleResolver = {
      getDefaultCharacterStyle: refuseStyleResolution,
      getDocDefaults: refuseStyleResolution,
      getRunStyleOwnProperties: refuseStyleResolution,
      resolveParagraphStyle: refuseStyleResolution,
    } satisfies RunStyleResolver;
    const inheritedMarks = [folioSchema.mark("bold"), folioSchema.mark("fontSize", { size: 22 })];
    const doc = folioSchema.node("doc", null, [
      folioSchema.node(
        "paragraph",
        {
          paraId: "A1000001",
          defaultTextFormatting: { bold: true, fontSize: 22 },
        },
        [
          folioSchema.text("Inherited", inheritedMarks),
          folioSchema.text(" direct", [...inheritedMarks, folioSchema.mark("italic")]),
        ],
      ),
    ]);

    expect(
      createFolioAIEditSnapshotWithStyleResolver(doc, styleResolver).blocks.at(0)?.previewRuns,
    ).toEqual([
      { text: "Inherited", effectiveFormatting: { bold: true, fontSize: 22 } },
      {
        text: " direct",
        effectiveFormatting: { bold: true, italic: true, fontSize: 22 },
        authoredFormatting: { italic: true },
      },
    ]);
  });

  test("projects styled hidden text out of both block text and preview runs", () => {
    const bold = folioSchema.mark("bold");
    const hidden = folioSchema.mark("hidden");
    const doc = folioSchema.node("doc", null, [
      folioSchema.node("paragraph", null, [
        folioSchema.text("Shown ", [bold]),
        folioSchema.text("secret", [bold, hidden]),
      ]),
    ]);

    const snapshot = createFolioAIEditSnapshot(doc);
    const block = snapshot.blocks.at(0);

    expect(block?.text).toBe("Shown ");
    expect(block?.previewRuns).toEqual([
      {
        text: "Shown ",
        effectiveFormatting: { bold: true },
        authoredFormatting: { bold: true },
      },
    ]);
    expect(block?.previewRuns?.map(({ text }) => text).join("")).toBe(block?.text);
  });

  test("the seq- ids are the same whether or not blank paragraphs are there", () => {
    // The published contract: `seq-NNNN` counts the paragraphs that carry
    // text. A host extractor derives the same numbers, and a stored citation
    // has to keep naming its paragraph, so blank paragraphs may never take a
    // position in that sequence however many of them a document grows.
    const contentTexts = ["first", "second", "third", "fourth"];
    const withoutBlanks = schema.node(
      "doc",
      null,
      contentTexts.map((text) => paragraph(text)),
    );
    const withBlanks = schema.node("doc", null, [
      paragraph(""),
      paragraph("first"),
      paragraph(""),
      paragraph(""),
      paragraph("second"),
      paragraph("third"),
      paragraph(""),
      paragraph("fourth"),
      paragraph(""),
    ]);

    const seqIdsOf = (doc: PMNode): string[] =>
      createFolioAIEditSnapshot(doc)
        .blocks.filter(isFolioAIContentBlock)
        .map(({ id }) => id);

    expect(seqIdsOf(withoutBlanks)).toEqual(["seq-0001", "seq-0002", "seq-0003", "seq-0004"]);
    expect(seqIdsOf(withBlanks)).toEqual(seqIdsOf(withoutBlanks));

    // The blanks are addressable, and in their own sequence.
    expect(
      createFolioAIEditSnapshot(withBlanks)
        .blocks.filter((block) => !isFolioAIContentBlock(block))
        .map(({ id }) => id),
    ).toEqual(["blank-0001", "blank-0002", "blank-0003", "blank-0004", "blank-0005"]);
  });

  test("a seq- id resolves to the paragraph it numbered, past any blanks", () => {
    const doc = schema.node("doc", null, [
      paragraph(""),
      paragraph("first"),
      paragraph(""),
      paragraph("second"),
    ]);
    const snapshot = createFolioAIEditSnapshot(doc);

    expect(resolveSequentialBlockAnchor("seq-0002", snapshot)?.text).toBe("second");
    expect(resolveSequentialBlockAnchor("seq-0003", snapshot)).toBeUndefined();
  });
});
