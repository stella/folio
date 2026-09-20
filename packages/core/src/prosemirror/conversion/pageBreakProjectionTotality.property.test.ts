/**
 * Opening a package Word opens is not optional.
 *
 * `w:br w:type="page"` is an ordinary run child, so Word writes one anywhere a
 * run may go: beside a text box, inside a bordered or outlined paragraph, in a
 * table cell, inside a tracked wrapper or a content control. Folio used to
 * refuse several of those shapes at `toProseDoc`, and a refusal there is total:
 * the document cannot be opened, laid out or exported at all.
 *
 * The law this pins is the one `dispositions.ts` states — a refusal is a worse
 * loss than a drop — as three properties over the placements: the projection
 * completes, the break survives a save, and the editor round trip keeps it.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig } from "../../../../../test/property-testing";
import { parseDocx } from "../../docx/parser";
import { createDocx, repackDocx } from "../../docx/rezip";
import type {
  BlockContent,
  Paragraph,
  ParagraphContent,
  ParagraphFormatting,
  Run,
  RunContent,
} from "../../types/document";
import { createEmptyDocument } from "../../utils/createDocument";
import { fromProseDoc } from "./fromProseDoc";
import { toProseDoc } from "./toProseDoc";

const PAGE_BREAK: RunContent = { type: "break", breakType: "page" };

/** Inline kinds Word writes beside a break; each one used to be a refusal. */
const SIBLINGS = {
  none: undefined,
  text: { type: "text", text: "beside" },
  tab: { type: "tab" },
  softHyphen: { type: "softHyphen" },
  noBreakHyphen: { type: "noBreakHyphen" },
  symbol: { type: "symbol", font: "Wingdings", char: "F0E0" },
  fieldChar: { type: "fieldChar", charType: "begin" },
  instrText: { type: "instrText", text: " PAGE " },
  textBoxShape: {
    type: "shape",
    shape: {
      type: "shape",
      shapeType: "rect",
      size: { width: 914_400, height: 457_200 },
      textBody: { content: [{ type: "paragraph", content: [] }] },
    },
  },
} as const satisfies Record<string, RunContent | undefined>;

/** Paragraph properties whose page-break projection layout approximates. */
const FORMATTINGS = {
  plain: undefined,
  borders: { borders: { bottom: { style: "single", size: 8 } } },
  outline: { outlineLevel: { kind: "heading", level: 0 } },
  frame: { frame: { width: 720 } },
} as const satisfies Record<string, ParagraphFormatting | undefined>;

type Placement = "only" | "first" | "last" | "middle";

const runFor = (placement: Placement, sibling: RunContent | undefined): Run => {
  if (sibling === undefined) {
    return { type: "run", content: [PAGE_BREAK] };
  }
  switch (placement) {
    case "only":
      return { type: "run", content: [PAGE_BREAK] };
    case "first":
      return { type: "run", content: [PAGE_BREAK, sibling] };
    case "last":
      return { type: "run", content: [sibling, PAGE_BREAK] };
    case "middle":
      return { type: "run", content: [sibling, PAGE_BREAK, sibling] };
  }
};

const WRAPPERS = ["direct", "insertion", "deletion", "inlineSdt"] as const;
type Wrapper = (typeof WRAPPERS)[number];

const REVISION = { id: 91, author: "Reviewer", date: "2026-09-09T00:00:00.000Z" } as const;

const wrapRun = (wrapper: Wrapper, run: Run): ParagraphContent => {
  switch (wrapper) {
    case "direct":
      return run;
    case "insertion":
      return { type: "insertion", info: REVISION, content: [run] };
    case "deletion":
      return { type: "deletion", info: REVISION, content: [run] };
    case "inlineSdt":
      return { type: "inlineSdt", properties: { sdtType: "richText" }, content: [run] };
  }
};

const CONTAINERS = ["body", "tableCell"] as const;
type Container = (typeof CONTAINERS)[number];

const containerise = (container: Container, paragraph: Paragraph): BlockContent => {
  if (container === "body") {
    return paragraph;
  }
  return {
    type: "table",
    rows: [
      {
        type: "tableRow",
        cells: [{ type: "tableCell", content: [{ type: "paragraph", content: [] }, paragraph] }],
      },
    ],
  };
};

type Shape = {
  placement: Placement;
  sibling: keyof typeof SIBLINGS;
  formatting: keyof typeof FORMATTINGS;
  wrapper: Wrapper;
  container: Container;
};

const shapeArbitrary = (siblings: readonly (keyof typeof SIBLINGS)[]): fc.Arbitrary<Shape> =>
  fc.record<Shape>({
    placement: fc.constantFrom<Placement>("only", "first", "last", "middle"),
    sibling: fc.constantFrom(...siblings),
    formatting: fc.constantFrom(...(Object.keys(FORMATTINGS) as (keyof typeof FORMATTINGS)[])),
    wrapper: fc.constantFrom(...WRAPPERS),
    container: fc.constantFrom(...CONTAINERS),
  });

const ALL_SIBLINGS = Object.keys(SIBLINGS) as (keyof typeof SIBLINGS)[];

/**
 * A lone `w:fldChar` opens a complex field the generator never closes, and the
 * parser re-groups the whole run into that field. What a run holding an
 * unterminated field keeps is a separate question from where a page break may
 * sit, so the survival properties leave that to the parser's own tests.
 */
const SAVEABLE_SIBLINGS = ALL_SIBLINGS.filter((name) => name !== "fieldChar");

const documentFor = ({ placement, sibling, formatting, wrapper, container }: Shape) => {
  const document = createEmptyDocument();
  const paragraph: Paragraph = {
    type: "paragraph",
    ...(FORMATTINGS[formatting] === undefined ? {} : { formatting: FORMATTINGS[formatting] }),
    content: [
      { type: "run", content: [{ type: "text", text: "before" }] },
      wrapRun(wrapper, runFor(placement, SIBLINGS[sibling])),
      { type: "run", content: [{ type: "text", text: "after" }] },
    ],
  };
  document.package.document.content = [containerise(container, paragraph)];
  return document;
};

const countProjectedPageBreaks = (doc: ReturnType<typeof toProseDoc>): number => {
  let count = 0;
  doc.descendants((node) => {
    if (node.type.name === "pageBreakRun" || node.type.name === "pageBreak") {
      count += 1;
    }
    return true;
  });
  return count;
};

const countModelPageBreaks = (blocks: readonly BlockContent[]): number => {
  let count = 0;
  const visitRun = (run: Run): void => {
    for (const content of run.content) {
      if (content.type === "break" && content.breakType === "page") count += 1;
    }
  };
  const visitParagraphContent = (content: ParagraphContent): void => {
    switch (content.type) {
      case "run":
        visitRun(content);
        return;
      case "hyperlink":
        for (const child of content.children) if (child.type === "run") visitRun(child);
        return;
      case "simpleField":
      case "inlineSdt":
      case "insertion":
      case "deletion":
      case "moveFrom":
      case "moveTo":
      case "bidiWrapper":
        for (const child of content.content) visitParagraphContent(child);
        return;
      case "complexField":
        for (const run of content.fieldResult) visitRun(run);
        return;
      default:
        return;
    }
  };
  const visitBlock = (block: BlockContent): void => {
    if (block.type === "paragraph") {
      for (const content of block.content) visitParagraphContent(content);
      return;
    }
    if (block.type === "table") {
      for (const row of block.rows) {
        for (const cell of row.cells) for (const child of cell.content) visitBlock(child);
      }
      return;
    }
    for (const child of block.content) visitBlock(child);
  };
  for (const block of blocks) visitBlock(block);
  return count;
};

describe("page-break projection is total", () => {
  test("every placement projects without throwing", () => {
    fc.assert(
      fc.property(shapeArbitrary(ALL_SIBLINGS), (shape) => {
        const document = documentFor(shape);
        const projected = toProseDoc(document);
        expect(countProjectedPageBreaks(projected)).toBeGreaterThan(0);
      }),
      propertyConfig({ numRuns: 200 }),
    );
  });

  test("every placement survives a save", async () => {
    const shapes = fc.sample(shapeArbitrary(SAVEABLE_SIBLINGS), { numRuns: 60, seed: 0x9e3779b9 });
    for (const shape of shapes) {
      const document = documentFor(shape);
      const saved = await createDocx(document);
      const reparsed = await parseDocx(saved, { preloadFonts: false });
      expect({
        shape,
        breaks: countModelPageBreaks(reparsed.package.document.content) > 0,
      }).toEqual({ shape, breaks: true });
    }
  });

  test("every placement survives the editor round trip", async () => {
    const shapes = fc.sample(shapeArbitrary(SAVEABLE_SIBLINGS), { numRuns: 60, seed: 0x9e3779b9 });
    for (const shape of shapes) {
      const opened = await parseDocx(await createDocx(documentFor(shape)), {
        preloadFonts: false,
      });
      const rebuilt = fromProseDoc(toProseDoc(opened), opened);
      const reparsed = await parseDocx(await repackDocx(rebuilt, { updateModifiedDate: false }), {
        preloadFonts: false,
      });
      expect({
        shape,
        breaks: countModelPageBreaks(reparsed.package.document.content) > 0,
      }).toEqual({ shape, breaks: true });
    }
  });
});
