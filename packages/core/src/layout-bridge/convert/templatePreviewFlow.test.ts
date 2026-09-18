import { describe, expect, test } from "bun:test";

import { measureParagraph } from "../../layout-engine/measure";
import {
  fixedCharWidth,
  withFakeTextMeasure,
} from "../../layout-engine/measure/__tests__/fakeTextMeasure";
import type {
  FlowBlock,
  ParagraphBlock,
  TableBlock,
  TextBoxBlock,
  TextRun,
} from "../../layout-engine/types";
import { applyTemplatePreviewToBlocks, templatePreviewDirtyRange } from "./templatePreviewFlow";
import type { TemplatePreviewFlowEntry, TemplatePreviewFlowState } from "./templatePreviewFlow";
import type { TemplatePreviewHiddenRange } from "../../prosemirror/plugins/templatePreviewValues";

const textRun = (text: string, pmStart: number, extra: Partial<TextRun> = {}): TextRun => ({
  kind: "text",
  text,
  pmStart,
  pmEnd: pmStart + text.length,
  ...extra,
});

const paragraph = (id: string, pmStart: number, runs: ParagraphBlock["runs"]): ParagraphBlock => {
  const last = runs.at(-1);
  return {
    kind: "paragraph",
    id,
    runs,
    pmStart,
    pmEnd: (last?.pmEnd ?? pmStart) + 1,
  };
};

const runTexts = (block: FlowBlock): string[] => {
  if (block.kind !== "paragraph") {
    throw new Error("expected a paragraph block");
  }
  return block.runs.map((run) => (run.kind === "text" ? run.text : run.kind));
};

describe("applyTemplatePreviewToBlocks", () => {
  test("replaces the marker range inside a single run with the value", () => {
    // "before {{x}} after" — content starts at PM position 1.
    const source = paragraph("p1", 0, [textRun("before {{x}} after", 1)]);
    const [block] = applyTemplatePreviewToBlocks([source], {
      entries: [{ from: 8, to: 13, value: "1234" }],
      hidden: [],
      mode: "plain",
    });

    expect(runTexts(block!)).toEqual(["before ", "1234", " after"]);
    if (block!.kind !== "paragraph") {
      throw new Error("expected paragraph");
    }
    const [prefix, value, suffix] = block!.runs as [TextRun, TextRun, TextRun];
    expect([prefix.pmStart, prefix.pmEnd]).toEqual([1, 8]);
    // The value run keeps the marker's PM range, not the value's length.
    expect([value.pmStart, value.pmEnd]).toEqual([8, 13]);
    expect(value.templatePreview).toBe("plain");
    expect([suffix.pmStart, suffix.pmEnd]).toEqual([13, 19]);
  });

  test("carries the hosting run's formatting onto the value run", () => {
    const source = paragraph("p1", 0, [
      textRun("{{client.name}}", 1, { bold: true, fontSize: 14 }),
    ]);
    const [block] = applyTemplatePreviewToBlocks([source], {
      entries: [{ from: 1, to: 16, value: "Maciej Kur" }],
      hidden: [],
      mode: "highlighted",
    });

    if (block!.kind !== "paragraph") {
      throw new Error("expected paragraph");
    }
    const value = block!.runs[0] as TextRun;
    expect(value.text).toBe("Maciej Kur");
    expect(value.bold).toBe(true);
    expect(value.fontSize).toBe(14);
    expect(value.templatePreview).toBe("highlighted");
  });

  test("a rich value emits one run per span, layering bold/italic over the host formatting", () => {
    const source = paragraph("p1", 0, [textRun("{{company}}", 1, { bold: true, fontSize: 14 })]);
    const [block] = applyTemplatePreviewToBlocks([source], {
      entries: [
        {
          from: 1,
          to: 12,
          value: {
            runs: [
              { text: "Acme", bold: true },
              { text: ", seat in " },
              { text: "Poznań", italic: true },
            ],
          },
        },
      ],
      hidden: [],
      mode: "highlighted",
    });

    expect(runTexts(block!)).toEqual(["Acme", ", seat in ", "Poznań"]);
    if (block!.kind !== "paragraph") {
      throw new Error("expected paragraph");
    }
    const [first, middle, last] = block!.runs as [TextRun, TextRun, TextRun];
    // Host bold stays on every span; the span's own flags OR in.
    expect(first.bold).toBe(true);
    expect(middle.bold).toBe(true);
    expect(middle.italic).toBeUndefined();
    expect(last.bold).toBe(true);
    expect(last.italic).toBe(true);
    for (const run of [first, middle, last]) {
      // Host formatting and the marker's PM range carry onto each span run.
      expect(run.fontSize).toBe(14);
      expect([run.pmStart, run.pmEnd]).toEqual([1, 12]);
      expect(run.templatePreview).toBe("highlighted");
    }
  });

  test("collapses a marker split across formatting boundaries into one value run", () => {
    const source = paragraph("p1", 0, [
      textRun("{{cli", 1, { italic: true }),
      textRun("ent.name}}", 6),
      textRun(" tail", 16),
    ]);
    const [block] = applyTemplatePreviewToBlocks([source], {
      entries: [{ from: 1, to: 16, value: "Acme" }],
      hidden: [],
      mode: "highlighted",
    });

    expect(runTexts(block!)).toEqual(["Acme", " tail"]);
    if (block!.kind !== "paragraph") {
      throw new Error("expected paragraph");
    }
    const value = block!.runs[0] as TextRun;
    // Formatting comes from the run hosting the marker start.
    expect(value.italic).toBe(true);
    expect([value.pmStart, value.pmEnd]).toEqual([1, 16]);
  });

  test("drops non-text inline nodes swallowed by the marker range", () => {
    const source = paragraph("p1", 0, [
      textRun("{{a", 1),
      { kind: "lineBreak", pmStart: 4, pmEnd: 5 },
      textRun("b}}", 5),
    ]);
    const [block] = applyTemplatePreviewToBlocks([source], {
      entries: [{ from: 1, to: 8, value: "v" }],
      hidden: [],
      mode: "plain",
    });

    expect(runTexts(block!)).toEqual(["v"]);
  });

  test("substitutes markers inside a text box table cell", () => {
    const table: TableBlock = {
      kind: "table",
      id: "t1",
      pmStart: 0,
      pmEnd: 30,
      rows: [
        {
          id: "r1",
          cells: [
            {
              id: "c1",
              blocks: [paragraph("p1", 2, [textRun("{{x}}", 3)])],
            },
          ],
        },
      ],
    };
    const textBox: TextBoxBlock = {
      kind: "textBox",
      id: "tb1",
      width: 200,
      content: [table],
    };
    const [block] = applyTemplatePreviewToBlocks([textBox], {
      entries: [{ from: 3, to: 8, value: "cell value" }],
      hidden: [],
      mode: "plain",
    });

    if (block?.kind !== "textBox") {
      throw new Error("expected text box");
    }
    const contentBlock = block.content.at(0);
    if (contentBlock?.kind !== "table") {
      throw new Error("expected table in text box");
    }
    const cellParagraph = contentBlock.rows[0]!.cells[0]!.blocks[0]!;
    expect(runTexts(cellParagraph)).toEqual(["cell value"]);
  });

  test("returns untouched blocks by reference and the same array when nothing matches", () => {
    const touched = paragraph("p1", 0, [textRun("{{x}} text", 1)]);
    const untouched = paragraph("p2", 20, [textRun("plain text", 21)]);

    const unchanged = applyTemplatePreviewToBlocks([touched, untouched], {
      entries: [],
      hidden: [],
      mode: "plain",
    });
    expect(unchanged[0]).toBe(touched);
    expect(unchanged[1]).toBe(untouched);

    const transformed = applyTemplatePreviewToBlocks([touched, untouched], {
      entries: [{ from: 1, to: 6, value: "v" }],
      hidden: [],
      mode: "plain",
    });
    expect(transformed[0]).not.toBe(touched);
    expect(transformed[1]).toBe(untouched);
    // The source paragraph is never mutated — clearing the preview is just
    // laying out the original blocks again.
    expect(runTexts(touched)).toEqual(["{{x}} text"]);
  });

  test("drops the blocks a hidden span swallows whole", () => {
    // `{% if premium %}` / body / `{% endif %}` as three paragraphs, with a
    // tail paragraph outside the span. PM positions: a paragraph at `pmStart`
    // holds its text from `pmStart + 1`.
    const opener = paragraph("p1", 0, [textRun("{% if premium %}", 1)]);
    const body = paragraph("p2", 18, [textRun("Premium terms.", 19)]);
    const closer = paragraph("p3", 34, [textRun("{% endif %}", 35)]);
    const tail = paragraph("p4", 47, [textRun("Tail.", 48)]);

    const blocks = applyTemplatePreviewToBlocks([opener, body, closer, tail], {
      entries: [],
      hidden: [{ from: 1, to: 46, expr: "premium" }],
      mode: "plain",
    });

    expect(blocks.map((block) => block.id)).toEqual(["p4"]);
    // The surviving block is returned by reference, unmeasured and unchanged.
    expect(blocks[0]).toBe(tail);
  });

  test("keeps a block the hidden span only partly covers", () => {
    // An inline `{% if %}` inside running text: dropping the paragraph would
    // take the authored text around it, so the flow leaves it as written.
    const inline = paragraph("p1", 0, [textRun("Fee {% if waived %}waived{% endif %} due.", 1)]);
    const blocks = applyTemplatePreviewToBlocks([inline], {
      entries: [],
      hidden: [{ from: 5, to: 37, expr: "waived" }],
      mode: "plain",
    });

    expect(blocks).toHaveLength(1);
    expect(blocks[0]).toBe(inline);
  });

  test("keeps a block that carries no PM positions", () => {
    const orphan: ParagraphBlock = { kind: "paragraph", id: "p1", runs: [textRun("Text", 1)] };
    const blocks = applyTemplatePreviewToBlocks([orphan], {
      entries: [],
      hidden: [{ from: 0, to: 999, expr: "premium" }],
      mode: "plain",
    });

    expect(blocks).toHaveLength(1);
    expect(blocks[0]).toBe(orphan);
  });

  test("substituted values reflow the paragraph instead of keeping the marker's width", () => {
    withFakeTextMeasure(
      () => {
        // 5px per char, 150px wide: the 25-char marker line fits in one
        // 125px line.
        const source = paragraph("p1", 0, [
          textRun("Name: ", 1),
          textRun("{{client.name}}", 7),
          textRun(" end", 22),
        ]);
        const sourceMeasure = measureParagraph(source, 150);
        expect(sourceMeasure.lines).toHaveLength(1);
        expect(sourceMeasure.lines[0]!.width).toBe(125);

        // Short value: the line shrinks to the value's width — no dead
        // space where the marker used to be.
        const [short] = applyTemplatePreviewToBlocks([source], {
          entries: [{ from: 7, to: 22, value: "1234" }],
          hidden: [],
          mode: "plain",
        });
        const shortMeasure = measureParagraph(short as ParagraphBlock, 150);
        expect(shortMeasure.lines).toHaveLength(1);
        // "Name: " (6) + "1234" (4) + " end" (4) = 14 chars * 5px.
        expect(shortMeasure.lines[0]!.width).toBe(70);

        // Long value: the paragraph wraps onto a second line instead of
        // overlapping the following text.
        const [long] = applyTemplatePreviewToBlocks([source], {
          entries: [{ from: 7, to: 22, value: "An Unusually Long Company Name Ltd." }],
          hidden: [],
          mode: "plain",
        });
        const longMeasure = measureParagraph(long as ParagraphBlock, 150);
        expect(longMeasure.lines.length).toBeGreaterThan(1);
      },
      { charWidth: fixedCharWidth(5) },
    );
  });

  test("breaks a value's lines around line break runs", () => {
    const source = paragraph("p1", 0, [textRun("Intro {{terms}} end.", 1)]);
    const [block] = applyTemplatePreviewToBlocks([source], {
      entries: [{ from: 7, to: 16, value: "First line.\nSecond line." }],
      hidden: [],
      mode: "plain",
    });

    expect(runTexts(block!)).toEqual([
      "Intro ",
      "First line.",
      "lineBreak",
      "Second line.",
      " end.",
    ]);
    if (block!.kind !== "paragraph") {
      throw new Error("expected paragraph");
    }
    // Every run of the value, the break included, keeps the marker's PM range,
    // so click-to-position still resolves into the marker.
    for (const run of block!.runs.slice(1, 4)) {
      expect([run.pmStart, run.pmEnd]).toEqual([7, 16]);
    }
  });

  test("counts each newline form once and keeps a blank line blank", () => {
    const source = paragraph("p1", 0, [textRun("{{terms}}", 1)]);
    const [block] = applyTemplatePreviewToBlocks([source], {
      entries: [{ from: 1, to: 10, value: "a\r\nb\rc\n\nd" }],
      hidden: [],
      mode: "plain",
    });

    // `\r\n` is one break, a bare `\r` is one, and the empty segment between
    // two newlines contributes a second break instead of an empty run.
    expect(runTexts(block!)).toEqual([
      "a",
      "lineBreak",
      "b",
      "lineBreak",
      "c",
      "lineBreak",
      "lineBreak",
      "d",
    ]);
  });

  test("breaks the lines of a rich value's span", () => {
    const source = paragraph("p1", 0, [textRun("{{terms}}", 1)]);
    const [block] = applyTemplatePreviewToBlocks([source], {
      entries: [
        {
          from: 1,
          to: 10,
          value: { runs: [{ text: "Bold first.\nBold second.", bold: true }, { text: " tail" }] },
        },
      ],
      hidden: [],
      mode: "plain",
    });

    expect(runTexts(block!)).toEqual(["Bold first.", "lineBreak", "Bold second.", " tail"]);
    if (block!.kind !== "paragraph") {
      throw new Error("expected paragraph");
    }
    // The span's formatting survives the split.
    const [first, , second] = block!.runs as [TextRun, unknown, TextRun];
    expect(first.bold).toBe(true);
    expect(second.bold).toBe(true);
  });

  test("a multi-line value measures one line per value line", () => {
    withFakeTextMeasure(
      () => {
        const source = paragraph("p1", 0, [textRun("{{terms}}", 1)]);
        const [block] = applyTemplatePreviewToBlocks([source], {
          entries: [{ from: 1, to: 10, value: "First line.\nSecond line.\nThird line." }],
          hidden: [],
          mode: "plain",
        });

        // Wide enough that nothing wraps: every line comes from a newline. The
        // marker used to measure as one line and paint as three, so whatever
        // the layout placed below it was painted over.
        const measure = measureParagraph(block as ParagraphBlock, 6000);
        expect(measure.lines).toHaveLength(3);
      },
      { charWidth: fixedCharWidth(5) },
    );
  });
});

describe("templatePreviewDirtyRange", () => {
  const entryA = { from: 5, to: 12, value: "a" };
  const entryB = { from: 40, to: 55, value: "b" };
  const hiddenA: TemplatePreviewHiddenRange = { from: 60, to: 90, expr: "premium" };

  const previewOf = (
    entries: readonly TemplatePreviewFlowEntry[],
    hidden: readonly TemplatePreviewHiddenRange[] = [],
  ): TemplatePreviewFlowState => ({ entries, hidden });

  test("returns null when the substituted content is identical", () => {
    expect(
      templatePreviewDirtyRange(previewOf([entryA, entryB]), previewOf([entryA, entryB])),
    ).toBe(null);
  });

  test("covers only the changed entry", () => {
    expect(
      templatePreviewDirtyRange(
        previewOf([entryA, entryB]),
        previewOf([entryA, { ...entryB, value: "b2" }]),
      ),
    ).toEqual({ from: 40, to: 55 });
  });

  test("distinguishes rich values by formatting, not just text", () => {
    const richBold = {
      from: 5,
      to: 12,
      value: { runs: [{ text: "Acme", bold: true }] },
    };
    expect(templatePreviewDirtyRange(previewOf([richBold]), previewOf([richBold]))).toBe(null);
    // Same text, different formatting → the marker's blocks must re-lay out.
    expect(
      templatePreviewDirtyRange(
        previewOf([richBold]),
        previewOf([{ ...richBold, value: { runs: [{ text: "Acme", italic: true }] } }]),
      ),
    ).toEqual({ from: 5, to: 12 });
    // A plain string and a rich value with the same text are not identical.
    expect(
      templatePreviewDirtyRange(
        previewOf([{ from: 5, to: 12, value: "Acme" }]),
        previewOf([{ from: 5, to: 12, value: { runs: [{ text: "Acme" }] } }]),
      ),
    ).toEqual({ from: 5, to: 12 });
  });

  test("covers added and removed entries", () => {
    expect(templatePreviewDirtyRange(previewOf([]), previewOf([entryA]))).toEqual({
      from: 5,
      to: 12,
    });
    expect(templatePreviewDirtyRange(previewOf([entryA, entryB]), previewOf([entryB]))).toEqual({
      from: 5,
      to: 12,
    });
    expect(templatePreviewDirtyRange(previewOf([entryA]), previewOf([entryB]))).toEqual({
      from: 5,
      to: 55,
    });
  });

  test("covers a hidden span that appears, disappears, or moves", () => {
    expect(templatePreviewDirtyRange(previewOf([], [hiddenA]), previewOf([], [hiddenA]))).toBe(
      null,
    );
    expect(templatePreviewDirtyRange(previewOf([]), previewOf([], [hiddenA]))).toEqual({
      from: 60,
      to: 90,
    });
    expect(templatePreviewDirtyRange(previewOf([], [hiddenA]), previewOf([]))).toEqual({
      from: 60,
      to: 90,
    });
    // Same span, different condition: a second `{% if %}` now owns it.
    expect(
      templatePreviewDirtyRange(
        previewOf([], [hiddenA]),
        previewOf([], [{ ...hiddenA, expr: "vip" }]),
      ),
    ).toEqual({ from: 60, to: 90 });
    // Entry and hidden changes widen one range together.
    expect(templatePreviewDirtyRange(previewOf([entryA]), previewOf([], [hiddenA]))).toEqual({
      from: 5,
      to: 90,
    });
  });
});
