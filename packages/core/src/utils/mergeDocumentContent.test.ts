// mergeDocumentContent is the general replacement for hand-merging a parsed
// document's content blocks onto another document (e.g. `fromMarkdown`'s
// output onto a styled preset): appending `source.package.document.content`
// directly is unsafe once `source` carries its own numbering, because
// `fromMarkdown` mints small numIds (1, 2, …) that collide with a preset's
// reserved numbering (see stellaStyle.ts's `createLegalNumbering`, which
// reserves numId 1-5 for clause/definitions/recitals/parties/bullet). A
// colliding markdown list silently renders with the preset's clause/
// definition marker instead of a plain bullet/number.

import { describe, expect, test } from "bun:test";

import { docxToMarkdown } from "../docx/server/docxToMarkdown";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { createStellaStyleDocumentPreset } from "../style-sets/stellaStyle";
import type { BlockContent, Document, Paragraph, Table, TableCell } from "../types/document";
import { createEmptyDocument } from "./createDocument";
import { mergeDocumentContent } from "./mergeDocumentContent";

const CLEAN = {
  annotations: "strip",
  trackedChanges: "clean",
  comments: "strip",
  hyperlinks: "inline",
  footnotes: "strip",
} as const;

const clauseParagraph = (text: string): Paragraph => ({
  type: "paragraph",
  formatting: { numPr: { numId: 1, ilvl: 0 }, styleId: "ClauseParagraph1" },
  content: [{ type: "run", formatting: {}, content: [{ type: "text", text }] }],
});

const stellaTargetWithClause = (): Document => {
  const target = createEmptyDocument({ preset: createStellaStyleDocumentPreset() });
  target.package.document.content = [clauseParagraph("Clause text.")];
  return target;
};

describe("mergeDocumentContent", () => {
  test("appends source content and merges numbering definitions", () => {
    const target = stellaTargetWithClause();
    const merged = mergeDocumentContent(target, fromMarkdown("1. first\n2. second"));

    expect(merged.package.document.content).toHaveLength(3);
    // Target's own reserved numbering (numId 1-5) is untouched.
    expect(merged.package.numbering?.nums.filter((num) => num.numId <= 5)).toEqual(
      target.package.numbering?.nums,
    );
  });

  test("does not mutate either input document", () => {
    const target = stellaTargetWithClause();
    const source = fromMarkdown("1. first\n2. second");
    const targetContentLength = target.package.document.content.length;
    const sourceNumIds = source.package.numbering?.nums.map((num) => num.numId);

    mergeDocumentContent(target, source);

    expect(target.package.document.content).toHaveLength(targetContentLength);
    expect(source.package.numbering?.nums.map((num) => num.numId)).toEqual(sourceNumIds);
  });

  test("remaps a colliding source numId above the target's existing range", () => {
    const target = stellaTargetWithClause();
    const source = fromMarkdown("1. first\n2. second");
    // fromMarkdown mints numId 1 for its first list — the same id Stella's
    // preset reserves for clause numbering.
    expect(source.package.numbering?.nums.map((num) => num.numId)).toEqual([1]);

    const merged = mergeDocumentContent(target, source);
    const mergedListParagraph = merged.package.document.content.find(
      (block) => block.type === "paragraph" && block.listRendering,
    );
    expect(mergedListParagraph?.type).toBe("paragraph");
    const remappedNumId =
      mergedListParagraph?.type === "paragraph"
        ? mergedListParagraph.formatting?.numPr?.numId
        : undefined;
    expect(remappedNumId).toBeGreaterThan(5);
  });

  test("createDocx(merged) succeeds and round-trips both lists correctly", async () => {
    const target = stellaTargetWithClause();
    const merged = mergeDocumentContent(target, fromMarkdown("1. first\n2. second\n\n- bullet"));

    const bytes = await createDocx(merged);
    const markdown = await docxToMarkdown(bytes, CLEAN);

    // The markdown list keeps its own identity (plain "1."/"- "), not
    // Stella's clause "(a)"/definitions marker.
    expect(markdown).toContain("1. first\n2. second");
    expect(markdown).toContain("- bullet");

    // The preset's clause paragraph still renders with clause numbering
    // (decimal at level 0), not a plain unnumbered paragraph.
    const clauseLine = markdown.split("\n").find((line) => line.includes("Clause text."));
    expect(clauseLine).toMatch(/^1[.)]?\s+Clause text\.$/);
  });

  test("remaps a colliding numId inside a nested table (table cell containing a table)", () => {
    const target = stellaTargetWithClause();
    const listItem = fromMarkdown("1. first").package.document.content[0];
    if (!listItem || listItem.type !== "paragraph") {
      throw new Error("expected fromMarkdown to produce a list paragraph");
    }

    const nestedTable: Table = {
      type: "table",
      rows: [{ type: "tableRow", cells: [{ type: "tableCell", content: [listItem] }] }],
    };
    const outerTable: Table = {
      type: "table",
      rows: [{ type: "tableRow", cells: [{ type: "tableCell", content: [nestedTable] }] }],
    };
    const source: Document = {
      ...fromMarkdown("1. first"),
      package: {
        ...fromMarkdown("1. first").package,
        document: { ...fromMarkdown("1. first").package.document, content: [outerTable] },
      },
    };

    const merged = mergeDocumentContent(target, source);
    const mergedOuterTable = merged.package.document.content.at(-1);
    if (mergedOuterTable?.type !== "table") {
      throw new Error("expected the merged content to end with the outer table");
    }
    const mergedNestedTable = mergedOuterTable.rows[0]?.cells[0]?.content[0];
    if (mergedNestedTable?.type !== "table") {
      throw new Error("expected the outer cell to contain the nested table");
    }
    const mergedParagraph = mergedNestedTable.rows[0]?.cells[0]?.content[0];

    expect(mergedParagraph?.type).toBe("paragraph");
    const remappedNumId =
      mergedParagraph?.type === "paragraph" ? mergedParagraph.formatting?.numPr?.numId : undefined;
    // Remapped above the target's reserved numId range (1-5), not left at
    // the colliding numId 1 fromMarkdown minted.
    expect(remappedNumId).toBeGreaterThan(5);
  });

  test("a table cell holding an unexpected (non paragraph/table) block does not crash", () => {
    // `TableCell.content` is statically typed as `(Paragraph | Table)[]` —
    // folio's own DOCX parser flattens a `w:sdt` inside a cell into its
    // children rather than keeping a `blockSdt` wrapper (see
    // `tableParser.ts`) — but `mergeDocumentContent` is a general helper
    // that also has to tolerate hand-built `Document` values that don't
    // honor that invariant. Build one here to lock in the crash fix: passing
    // a non-table block straight into `remapTable` used to throw on
    // `table.rows`.
    const foreignBlock = {
      type: "blockSdt",
      properties: { sdtType: "richText" },
      content: [] as BlockContent[],
    } satisfies Extract<BlockContent, { type: "blockSdt" }>;
    // SAFETY: simulating an untrusted/hand-built Document whose cell content
    // does not conform to `TableCell.content`'s static type — see the test
    // description above.
    const cell = { type: "tableCell", content: [foreignBlock] } as unknown as TableCell;
    const table: Table = { type: "table", rows: [{ type: "tableRow", cells: [cell] }] };

    const target = stellaTargetWithClause();
    const source: Document = {
      ...fromMarkdown("1. first"),
      package: {
        ...fromMarkdown("1. first").package,
        document: { ...fromMarkdown("1. first").package.document, content: [table] },
      },
    };

    expect(() => mergeDocumentContent(target, source)).not.toThrow();

    const merged = mergeDocumentContent(target, source);
    const mergedTable = merged.package.document.content.at(-1);
    if (mergedTable?.type !== "table") {
      throw new Error("expected the merged content to end with the table");
    }
    // The foreign block is preserved unchanged, not dropped or crashed on.
    expect(mergedTable.rows[0]?.cells[0]?.content[0]).toEqual(foreignBlock);
  });

  test("recurses into a top-level blockSdt content control", () => {
    const target = stellaTargetWithClause();
    const listItem = fromMarkdown("1. first").package.document.content[0];
    if (!listItem || listItem.type !== "paragraph") {
      throw new Error("expected fromMarkdown to produce a list paragraph");
    }

    const sdt: Extract<BlockContent, { type: "blockSdt" }> = {
      type: "blockSdt",
      properties: { sdtType: "richText" },
      content: [listItem],
    };
    const source: Document = {
      ...fromMarkdown("1. first"),
      package: {
        ...fromMarkdown("1. first").package,
        document: { ...fromMarkdown("1. first").package.document, content: [sdt] },
      },
    };

    const merged = mergeDocumentContent(target, source);
    const mergedSdt = merged.package.document.content.at(-1);
    if (mergedSdt?.type !== "blockSdt") {
      throw new Error("expected the merged content to end with the blockSdt");
    }
    const mergedParagraph = mergedSdt.content[0];
    expect(mergedParagraph?.type).toBe("paragraph");
    const remappedNumId =
      mergedParagraph?.type === "paragraph" ? mergedParagraph.formatting?.numPr?.numId : undefined;
    expect(remappedNumId).toBeGreaterThan(5);
  });
});
